# Workforce Attendance & Reporting Automation

Two Google Apps Script projects that turn raw attendance/manpower spreadsheets into
polished, auto-generated HTML email reports for field workforce operations —
no manual report-building, no copy-pasting numbers into slides or emails.

| Project | What it does | Sends to |
|---|---|---|
| [`gmr-monthly-comparison/`](./gmr-monthly-comparison) | Month-over-month attendance comparison dashboard for GMR field projects (Kashi, Triveni), with fair like-for-like day-range comparison | HR Operations / Project Heads |
| [`apraava-workforce-report/`](./apraava-workforce-report) | State-wise workforce mandate vs. achieved headcount, deployment pipeline tracking, and daily attendance trend chart for Apraava Energy field technicians (Gujarat, Rajasthan, MP) | Workforce Ops / Management |

---

## Why this exists (business context)

Field operations teams manage hundreds of technicians spread across states/circles,
tracked in Google Sheets updated daily by ground staff. Historically, producing a
management report meant someone manually:

- Pulling numbers out of raw attendance sheets
- Calculating attendance %, deficits, and month-over-month deltas by hand
- Building tables/charts in Excel or PowerPoint
- Emailing them out, often late and inconsistently formatted

Both scripts remove that manual step entirely. They run **inside the spreadsheet
itself** (Google Apps Script, bound to the sheet), read the raw data on demand,
compute all KPIs, and generate a client-ready HTML report — either on a custom
menu click or automatically on a schedule (e.g. every Monday 8 AM).

**Impact:**
- Eliminates hours of manual report preparation per reporting cycle
- Removes human error in KPI calculation (attendance %, deltas, deficits)
- Guarantees consistent, professional formatting every time
- Gives leadership same-day visibility instead of waiting on manual compilation
- Fully auditable — GMR version logs every sent email to a "Comparison Log" sheet

## How it works (technical overview)

Both projects follow the same pattern, native to Google Apps Script:

```
Google Sheet (raw data tabs)
        │
        ▼
 Data layer  — reads sheet ranges, aggregates KPIs (attendance %, deltas, deficits)
        │
        ▼
 Render layer — builds a self-contained HTML email (inline CSS, no external assets)
        │
        ▼
 Delivery layer — MailApp / GmailApp sendEmail(), optionally via a custom
                  dialog (recipient/CC picker) or a time-based trigger
```

Key engineering details worth noting:

- **No external dependencies.** Runs entirely on Apps Script's built-in
  `SpreadsheetApp`, `MailApp`/`GmailApp`, and `HtmlService` — no npm packages,
  no servers, no API keys.
- **Email-client-safe HTML.** All charts are built as HTML/CSS (table-based bars)
  rather than SVG, since Gmail and Outlook strip `<svg>` from emails. The Apraava
  script also demonstrates the alternative approach — rendering a real chart image
  via `Charts` service and embedding it as a `cid:` inline image.
- **Fair period comparison** (GMR): when comparing "this month" to "last month,"
  the script automatically trims the previous month's data to the same number of
  days as the current (in-progress) month, so an 8-day-old month is never compared
  against a full 30-day one.
- **Config-driven** (Apraava): sheet names, recipients, and status tokens live in
  a single `CONFIG` object at the top of the file, so adapting the script to a new
  spreadsheet/state doesn't require touching the logic.

## Setup (both projects)

1. Open the target Google Sheet.
2. `Extensions → Apps Script`.
3. Paste the relevant `.gs` file's contents into the script editor.
4. Update the config block (sheet names, recipient emails, theme) to match your sheet.
5. Run the send/preview function once from the editor to authorize Gmail/Sheets access.
6. Optionally set up a time-based trigger for automatic sending (see each project's README).

See each project folder for full setup instructions and configuration details.

## Repo structure

```
.
├── README.md                          ← you are here
├── gmr-monthly-comparison/
│   ├── GMR_Monthly_Comparison.gs
│   └── README.md
├── apraava-workforce-report/
│   ├── Apraava_Report_AppsScript.gs
│   └── README.md
└── docs/
    └── presentation.md                ← slide-deck source
```

## License

Internal tooling — adapt freely for your own workforce/attendance reporting needs.
