# 💼 Job Tracker

> A vibe-coded Google Apps Script that automatically syncs your Gmail inbox to track job applications — no more manual spreadsheet updates.

## What it does

Job Tracker scans your Gmail for job-related emails and populates a Google Sheet with the company, role, and current application status. It also ships a clean web UI (served as a Google Apps Script Web App) so you can view, add, edit, and delete applications without ever touching the spreadsheet directly.

**Status tracking is automatic.** Every time you sync, the script reads your emails and advances each application through the pipeline:

`Applied → Reviewing → Interview → Offer → Rejected`

---

## Features

- **Gmail sync** — parses application confirmation, status update, interview invite, and rejection emails
- **Smart deduplication** — matches by company + role (exact) or company alone (fuzzy), and only advances status forward
- **Web UI** — filterable table with stats dashboard, add/edit/delete modals, and a Sync button
- **Auto-sync** — optional daily trigger that runs at 8 AM in your timezone
- **Color-coded statuses** — both in the sheet and the web UI

---

## Setup

### 1. Create a Google Sheet

Open [Google Sheets](https://sheets.google.com) and create a new spreadsheet. Copy the spreadsheet ID from the URL:

```
https://docs.google.com/spreadsheets/d/YOUR_SPREADSHEET_ID/edit
```

### 2. Open Apps Script

In your spreadsheet, go to **Extensions → Apps Script**.

### 3. Add the files

Create two files in the Apps Script editor:

| File | Content |
|------|---------|
| `Code.gs` | Paste the contents of `Code.gs` from this repo |
| `Index.html` | Paste the contents of `Index.html` from this repo |

### 4. Set your Spreadsheet ID

In `Code.gs`, replace the placeholder at the top:

```javascript
var SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID'; // ← Replace this
```

### 5. Authorize permissions

Run any function (e.g. `getSheet`) from the Apps Script editor. Google will prompt you to grant permissions for Gmail and Sheets access.

### 6. Deploy as a Web App (optional)

To use the web UI:

1. Click **Deploy → New deployment**
2. Choose **Web app**
3. Set "Execute as" to **Me** and "Who has access" to **Only myself**
4. Click **Deploy** and copy the URL

### 7. Enable auto-sync (optional)

From the spreadsheet menu, go to **Job Tracker → Set up auto-sync (daily)**. This creates a trigger that syncs your Gmail every day at 8 AM.

---

## Customization

### Change the sync date range

In `Code.gs`, update the `startDate` in `syncGmailAndReturn()`:

```javascript
var startDate = new Date('2026-01-01'); // ← Change to your job search start date
```

### Add non-job senders/subjects to filter out

Extend the `NON_JOB_SENDERS` or `NON_JOB_SUBJECTS` arrays at the top of `Code.gs` with any domains or keywords you want ignored.

### Add status-detection phrases

Extend the `STATUS_RULES` array to catch phrasing from specific companies or job boards that aren't already covered.

---

## How it works

```
Gmail Search
    │
    ▼
isJobEmail()        ← filters out non-job senders/subjects, requires at least one job phrase
    │
    ▼
parseJobEmail()     ← determines status from email body keywords
    │
    ├── extractCompany()   ← regex patterns across subject, body, sender domain
    └── extractRole()      ← regex patterns across subject and body
    │
    ▼
Sheet update        ← appends new row or advances status of existing entry
```

---

## File structure

```
├── Code.gs       # Backend: Gmail sync, sheet logic, web app functions
├── Index.html    # Frontend: Web UI served via Google Apps Script HtmlService
└── README.md
```

---

## Limitations

- **Gmail only** — does not sync applications submitted through LinkedIn Easy Apply or other portals unless you receive a confirmation email
- **English emails only** — status detection phrases are English-language
- **~200 threads per sync** — Apps Script's `GmailApp.search()` is capped per call; adjust the limit in `syncGmailAndReturn()` if needed
- **ATS platforms** — emails from Workday, Greenhouse, Lever, etc. fall back to sender domain for company name, which may be less accurate

---

## Tech stack

- **Google Apps Script** (backend + hosting)
- **Google Sheets** (data store)
- **Gmail API** (via Apps Script's `GmailApp`)
- **Vanilla HTML/CSS/JS** (web UI, no frameworks)
