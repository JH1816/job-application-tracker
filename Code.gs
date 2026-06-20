// ============================================================
// JOB TRACKER — Google Apps Script
// ============================================================
//
// SETUP: Replace YOUR_SPREADSHEET_ID below with your own Google
// Sheets ID. You can find it in your spreadsheet's URL:
// https://docs.google.com/spreadsheets/d/YOUR_SPREADSHEET_ID/edit
//
// ============================================================

var SHEET_NAME = 'Job Tracker';
var SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID'; // ← Replace this

var NON_JOB_SUBJECTS = [
  'arrival card', 'flash deal', '<adv>', '[adv]', 'transaction alert',
  'order #', 'order id', 'have been cancelled', 'passcode',
  'password reset', 'password updated', 'password has been',
  'nus l3ap', 'data visualization continues', 'power bi', 'placement offer',
  'visa application', 'credit card', 'uob application sg-',
  'priority pass', 'booking id', 'check in', 'lumosity', 'ticketmaster'
];

var NON_JOB_SENDERS = [
  'shopee', 'lazada', 'grab', 'youtrip', 'you.co',
  'uobgroup.com', 'hsbc', 'dbs.com', 'ocbc', 'maybank', 'citibank',
  'ica.gov', 'immigration.go.th', 'tdacservices',
  'healthxchange', 'maribank', 'travelsure', 'singhealth.com.sg',
  'stratascratch', 'noreply@email.jobstreet', 'noreply@email.seek',
  'mycareersfuture', 'se.nus.edu.sg', 'nus.edu.sg',
  'l3ap_noreply', 'fos.shortcourses', 'lumosity', 'ticketmaster',
  'agoda', 'notifications.'
];

var JOB_REQUIRED_PHRASES = [
  'your application for the role',
  'your application for the position',
  'your application for',
  'we have received your application',
  'application has been received',
  'application received',
  'thank you for applying',
  'thank you for your application',
  'application was sent to',
  'application was viewed by',
  'application confirmation',
  'we regret to inform you',
  'not been selected',
  'not been shortlisted',
  'decided not to take your candidacy',
  'not to take your candidacy further',
  'invite you to interview',
  'schedule.*interview',
  'pleased to invite you',
  'offer of employment',
  'job offer'
];

var STATUS_RULES = [
  { status: 'Offer',     phrases: ['offer of employment', 'job offer', 'pleased to offer', 'formal offer', 'offer letter'] },
  { status: 'Interview', phrases: ['invite you to interview', 'schedule an interview', 'pleased to invite you', 'would like to invite you', 'interview with us on', 'interview scheduled'] },
  { status: 'Rejected',  phrases: ['regret to inform', 'not been selected', 'not been shortlisted', 'not moving forward', 'will not be progressing', 'position has been filled', 'decided not to take your candidacy', 'not to take your candidacy further', 'unsuccessful in', 'unable to move forward', 'not progress your application'] },
  { status: 'Reviewing', phrases: ['application was viewed by', 'carefully reviewed', 'under review', 'shortlisted candidates', 'we will review', 'reviewing your application', 'review your application'] },
  { status: 'Applied',   phrases: ['application received', 'we have received your application', 'application has been received', 'thank you for applying', 'thank you for your application', 'application was sent to', 'application confirmation', 'your application for the role', 'your application for the position', 'received your application for'] },
];

var STATUS_ORDER = ['Applied', 'Reviewing', 'Interview', 'Offer', 'Rejected'];

// ── Sheet ─────────────────────────────────────────────────────

function getSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    var sheets = ss.getSheets();
    sheet = sheets.length > 0 ? sheets[0] : ss.insertSheet(SHEET_NAME);
    sheet.setName(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    var h = ['Company', 'Role', 'Status', 'Applied Date', 'Last Updated', 'Notes', 'Source'];
    sheet.getRange(1, 1, 1, h.length).setValues([h]);
    sheet.getRange(1, 1, 1, h.length).setFontWeight('bold').setBackground('#f3f4f6');
    sheet.setFrozenRows(1);
    [160,200,110,110,110,280,80].forEach(function(w,i){ sheet.setColumnWidth(i+1,w); });
  }
  return sheet;
}

// ── Entry points ──────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Job Tracker')
    .addItem('Sync Gmail now', 'syncGmail')
    .addItem('Set up auto-sync (daily)', 'setupTrigger')
    .addItem('Remove auto-sync', 'removeTrigger')
    .addToUi();
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Job Tracker')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function syncGmail() {
  SpreadsheetApp.getUi().alert(syncGmailAndReturn());
}

// ── Core sync ─────────────────────────────────────────────────

function syncGmailAndReturn() {
  var sheet = getSheet();
  // Build TWO lookup maps: exact (company+role) and fuzzy (company only)
  var existing = loadExisting(sheet);
  var startDate = new Date('2026-01-01');
  var threads = GmailApp.search(buildQuery(startDate), 0, 200);
  var added = 0, updated = 0;

  threads.forEach(function(thread) {
    try {
      thread.getMessages().forEach(function(msg) {
        try {
          if (msg.getDate() < startDate) return;
          var subject = msg.getSubject() || '';
          var sender = msg.getFrom() || '';
          var body = '';
          try { body = msg.getPlainBody() || ''; } catch(e) {}

          if (!isJobEmail(sender, subject, body)) return;

          var parsed = parseJobEmail(msg, subject, sender, body);
          if (!parsed) return;

          var companyKey = parsed.company.toLowerCase();
          var exactKey = companyKey + '||' + parsed.role.toLowerCase();

          // Try exact match first, then company-only match
          var match = existing.exact[exactKey] || existing.byCompany[companyKey];

          if (match) {
            var curr = STATUS_ORDER.indexOf(match.status);
            var next = STATUS_ORDER.indexOf(parsed.status);
            // Update if status progresses OR if it's a rejection (always final)
            if (next > curr || (parsed.status === 'Rejected' && match.status !== 'Rejected')) {
              updateRow(sheet, match.row, parsed.status, parsed.notes);
              // Update both maps
              match.status = parsed.status;
              updated++;
            }
          } else {
            appendRow(sheet, parsed, msg.getDate());
            var newRow = sheet.getLastRow();
            existing.exact[exactKey] = { status: parsed.status, row: newRow };
            existing.byCompany[companyKey] = { status: parsed.status, row: newRow };
            added++;
          }
        } catch(e) {}
      });
    } catch(e) {}
  });

  return 'Sync complete: ' + added + ' added, ' + updated + ' updated.';
}

// ── Filters ───────────────────────────────────────────────────

function isJobEmail(sender, subject, body) {
  var senderL = sender.toLowerCase();
  var subjectL = subject.toLowerCase();
  var combined = (body + ' ' + subject).toLowerCase();

  for (var i = 0; i < NON_JOB_SENDERS.length; i++) {
    if (senderL.indexOf(NON_JOB_SENDERS[i]) !== -1) return false;
  }
  for (var j = 0; j < NON_JOB_SUBJECTS.length; j++) {
    if (subjectL.indexOf(NON_JOB_SUBJECTS[j]) !== -1) return false;
  }
  for (var k = 0; k < JOB_REQUIRED_PHRASES.length; k++) {
    if (combined.indexOf(JOB_REQUIRED_PHRASES[k]) !== -1) return true;
  }
  return false;
}

// ── Parser ────────────────────────────────────────────────────

function parseJobEmail(msg, subject, sender, body) {
  var combined = (body + ' ' + subject).toLowerCase();

  var status = null;
  for (var i = 0; i < STATUS_RULES.length; i++) {
    for (var j = 0; j < STATUS_RULES[i].phrases.length; j++) {
      if (combined.indexOf(STATUS_RULES[i].phrases[j]) !== -1) {
        status = STATUS_RULES[i].status;
        break;
      }
    }
    if (status) break;
  }
  if (!status) status = 'Applied';

  var company = extractCompany(subject, sender, body);
  if (!company) return null;

  var role = extractRole(subject, body);
  return { company: company, role: role || 'Not specified', status: status, notes: subject };
}

// ── Company extraction ────────────────────────────────────────

function extractCompany(subject, sender, body) {
  var m;

  // LinkedIn: "sent to [Company]" or "viewed by [Company]"
  m = subject.match(/(?:sent to|viewed by)\s+(.+?)$/i);
  if (m) return m[1].trim();

  // "at [Company]" in body
  m = body.match(/\bat\s+([A-Z][A-Za-z0-9\s&.'()-]{2,50}?)(?=\s*[!.,\n]|\s+has been|\s+is\b)/);
  if (m) return m[1].trim();

  // "joining [Company]"
  m = body.match(/joining\s+([A-Z][A-Za-z0-9\s&.'()-]{2,50}?)(?=[!.,\n])/);
  if (m) return m[1].trim();

  // "interest in [Company]"
  m = body.match(/interest in\s+([A-Z][A-Za-z0-9\s&.'()-]{2,50}?)(?=[!.,\n])/);
  if (m) return m[1].trim();

  // "from [Company]!"
  m = body.match(/from\s+([A-Z][A-Za-z0-9\s&.'()-]{2,50}?)(?=!)/);
  if (m) return m[1].trim();

  // Subject: "position at [Company]"
  m = subject.match(/\bat\s+([A-Z][A-Za-z0-9\s&.'()-]{2,40}?)(?:\s*$|[!.,])/);
  if (m) return m[1].trim();

  // Subject: "with [Company]: Role"
  m = subject.match(/with\s+([A-Z][A-Za-z0-9\s&.'()-]{2,50}?)(?=:\s|\s*$)/);
  if (m) return m[1].trim();

  // Fall back to sender domain (skip ATS platforms)
  var ATS = ['myworkday','successfactors','teamtailor','greenhouse','lever','workday',
    'linkedin','psd.gov','hrp_admin','sap.hr','hr.ext','jobs-noreply','noreply','no-reply','donotreply'];
  m = sender.match(/@([a-zA-Z0-9.-]+)\b/);
  if (m) {
    var domain = m[1].toLowerCase();
    var isAts = ATS.some(function(d){ return domain.indexOf(d) !== -1; });
    if (!isAts) {
      var part = domain.split('.')[0];
      if (part.length > 2) return part.charAt(0).toUpperCase() + part.slice(1);
    }
  }

  return null;
}

// ── Role extraction ───────────────────────────────────────────

function extractRole(subject, body) {
  var m;

  // Subject: "- Role Title (ID)" e.g. "Seagate ... - Data Scientist I (14328)"
  m = subject.match(/[-–]\s*([A-Za-z][A-Za-z0-9\s\/,&()\-]{3,60}?)\s*(?:\(\d+\))?\s*$/);
  if (m) {
    var r = m[1].trim();
    if (!/^(singapore|technology|confirm|receiv|applic|update|career|interest)/i.test(r)) return r;
  }

  // Subject: ": Role" e.g. "Application Status with Coca-Cola: Data Analyst"
  m = subject.match(/:\s*([A-Za-z][A-Za-z0-9\s\/,&()\-]{3,60}?)\s*$/);
  if (m) return m[1].trim();

  // Subject: "our [Role] position at" e.g. GIC "our Associate/AVP, Data Engineer, AI Alpha Group position at GIC"
  m = subject.match(/our\s+([A-Za-z][A-Za-z0-9\s\/,&()\-]{3,80}?)\s+position\s+at\s+/i);
  if (m) return m[1].trim();

  // Subject: "received for [Role]" — strip trailing job IDs e.g. "Senior / Statistician 17647017"
  m = subject.match(/(?:received for|application for)\s+([A-Za-z\/][A-Za-z0-9\s\/,&\-]{3,60}?)(?:\s*$)/i);
  if (m) return m[1].trim().replace(/\s+\d{5,}\s*$/, '').trim();

  // Subject: "Job Title: [Role]" — handles zero-width spaces (SIA)
  m = subject.match(/Job Title\s*:\s*[\u200b\u200c]*([\w][A-Za-z0-9\s\/,&()\-]{3,60}?)[\u200b\u200c]*(?:,\s*Req|$)/i);
  if (m) return m[1].trim();

  // Body: "application for the role of [Role]" / "for the role: [Role]"
  m = body.match(/(?:application for the role(?:\s+of)?|for the role)\s*:?\s*([A-Za-z][A-Za-z0-9\s\/,&()\-]{3,60}?)(?:\s*[.,\n]|$)/i);
  if (m) return m[1].trim();

  // Body: "application for the position of [Role]"
  m = body.match(/application for the position(?:\s+of)?\s*:?\s*([A-Za-z][A-Za-z0-9\s\/,&()\-]{3,60}?)(?:\s*[.,\n]|$)/i);
  if (m) return m[1].trim();

  // Body: "applying to the [Role] position" e.g. GIC rejection "applying to the Associate, Data Analyst (Contract) position"
  m = body.match(/applying to the\s+([A-Za-z][A-Za-z0-9\s\/,&()\-]{3,60}?)\s+position/i);
  if (m) return m[1].trim();

  // Body: "job application for [Role]" e.g. NUHS
  m = body.match(/job application for\s+([A-Za-z(][A-Za-z0-9\s\/,&()\-.]{3,80}?)(?:,\s*Dept|\s*-\s*[A-Z]{2,}|\s*\.|$)/i);
  if (m) return m[1].trim();

  // Body: "application for the [Role] role" — GIC pattern
  m = body.match(/application for the\s+([A-Za-z][A-Za-z0-9\s\/,&()\-]{3,60}?)\s+(?:role|position)\b/i);
  if (m) return m[1].trim();

  return null;
}

// ── Gmail query ───────────────────────────────────────────────

function buildQuery(startDate) {
  var d = Utilities.formatDate(startDate, 'UTC', 'yyyy/MM/dd');
  return 'after:' + d + ' ' +
    '("your application for" OR "application received" OR "application confirmation" OR ' +
    '"thank you for applying" OR "thank you for your application" OR ' +
    '"we have received your application" OR "application was sent to" OR ' +
    '"application was viewed by" OR "we regret to inform" OR "not been selected" OR ' +
    '"not been shortlisted" OR "decided not to take your candidacy" OR ' +
    '"invite you to interview" OR "offer of employment" OR "job offer") ' +
    '-subject:"arrival card" -subject:"<ADV>" -subject:"transaction alert" ' +
    '-subject:"order" -subject:"passcode" -subject:"password" -subject:"nus l3ap" ' +
    '-subject:"power bi" -subject:"data visualization" -subject:"placement offer" ' +
    '-subject:"booking id" -subject:"check in" -subject:"priority pass"';
}

// ── Sheet helpers ─────────────────────────────────────────────

function loadExisting(sheet) {
  var data = sheet.getDataRange().getValues();
  var exact = {}, byCompany = {};
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var company = String(data[i][0]).toLowerCase();
    var role = String(data[i][1] || '').toLowerCase();
    var status = String(data[i][2]);
    var row = i + 1;
    var entry = { status: status, row: row };
    exact[company + '||' + role] = entry;
    // byCompany keeps the LATEST row for that company
    byCompany[company] = entry;
  }
  return { exact: exact, byCompany: byCompany };
}

function appendRow(sheet, parsed, date) {
  var tz = Session.getScriptTimeZone();
  sheet.appendRow([parsed.company, parsed.role, parsed.status,
    Utilities.formatDate(date, tz, 'dd MMM yyyy'),
    Utilities.formatDate(new Date(), tz, 'dd MMM yyyy'),
    parsed.notes, 'Gmail']);
  colorStatusCell(sheet, sheet.getLastRow(), parsed.status);
}

function updateRow(sheet, row, status, notes) {
  sheet.getRange(row, 3).setValue(status);
  sheet.getRange(row, 5).setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy'));
  var existing = String(sheet.getRange(row, 6).getValue() || '');
  if (notes && existing.indexOf(notes) === -1) {
    sheet.getRange(row, 6).setValue(existing ? existing + ' | ' + notes : notes);
  }
  colorStatusCell(sheet, row, status);
}

function colorStatusCell(sheet, row, status) {
  var colors = {
    'Applied':   {bg:'#dbeafe', fg:'#1e3a8a'},
    'Reviewing': {bg:'#ede9fe', fg:'#4c1d95'},
    'Interview': {bg:'#fef3c7', fg:'#78350f'},
    'Offer':     {bg:'#dcfce7', fg:'#14532d'},
    'Rejected':  {bg:'#fee2e2', fg:'#7f1d1d'}
  };
  var c = colors[status] || {bg:'#f3f4f6', fg:'#111827'};
  sheet.getRange(row, 3).setBackground(c.bg).setFontColor(c.fg).setFontWeight('bold');
}

// ── Web app functions ─────────────────────────────────────────

function getSheetData() {
  var sheet = getSheet();
  var tz = Session.getScriptTimeZone();
  var data = sheet.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var row = [];
    for (var j = 0; j < data[i].length; j++) {
      var val = data[i][j];
      row.push(val instanceof Date ? Utilities.formatDate(val, tz, 'dd MMM yyyy') : (val ? String(val) : ''));
    }
    row.push(i + 1);
    rows.push(row);
  }
  return rows;
}

function saveEntry(data) {
  var sheet = getSheet();
  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy');
  if (data.row) {
    sheet.getRange(data.row, 1).setValue(data.company);
    sheet.getRange(data.row, 2).setValue(data.role);
    sheet.getRange(data.row, 3).setValue(data.status);
    sheet.getRange(data.row, 5).setValue(now);
    sheet.getRange(data.row, 6).setValue(data.notes);
    colorStatusCell(sheet, data.row, data.status);
  } else {
    appendRow(sheet, {company:data.company, role:data.role, status:data.status, notes:data.notes}, new Date());
  }
}

function deleteEntry(row) {
  var sheet = getSheet();
  sheet.deleteRow(row);
}

// ── Triggers ──────────────────────────────────────────────────

function setupTrigger() {
  removeTrigger();
  ScriptApp.newTrigger('syncGmailAndReturn').timeBased().everyDays(1).atHour(8).create();
  SpreadsheetApp.getUi().alert('Auto-sync set up! Gmail will sync every day at 8am.');
}

function removeTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'syncGmailAndReturn' || t.getHandlerFunction() === 'syncGmail') ScriptApp.deleteTrigger(t);
  });
}
