---
name: timesheet-workflow-automation
description: "Automate timesheet CSV to weekly sheet, report, drafts."
category: productivity
version: 1.0.0
author: Hermes Agent
license: MIT
metadata:
  hermes:
    tags: [timesheet, attendance, google-workspace, drive, sheets, word, email]
    related_skills: [google-workspace]
---

# Timesheet workflow automation

Automate a daily/weekly attendance review loop:

## When to Use

Use this skill when the user runs a recurring attendance/time-tracking review loop: a daily CSV export (e.g. FlowAce "Workforce Efficiency Daily Breakdown") must be merged into a weekly tracking workbook, idle/undertime violations analyzed against policy thresholds, a Word report produced, and email notification drafts prepared for pre-send review.

### How this workflow is started

The user initiates it by saying "**Generate me the TimeSheet Report for Today**" (or words to that effect). When they do, **always ask them to provide/share the FlowAce daily CSV report first** before running anything. Do not run the automation without the CSV in hand. Once the CSV is provided, run the full procedure below (Steps 1–8).

The skill owns the end-to-end business workflow: Drive backup → weekly tab fill → violation analysis → Word report → `.eml` drafts. It delegates the underlying OAuth and Drive/Gmail/Sheets API mechanics to the `google-workspace` skill; load that skill for credential setup and API command reference.

Do not use this skill for one-off CSV conversions, single-email composition, or ad-hoc spreadsheet edits that are not part of the recurring review loop.

1. Parse the daily timesheet export (FlowAce-style CSV).
2. Copy the master workbook on Drive **before** updating (user requirement).
3. Fill the current week's Mon–Fri tab with the day's records.
4. Upload the updated workbook back to Drive, preserving the same file ID and share link.
5. Analyze idle (>15 min) and undertime (>5 min) violations against scheduled hours.
6. Generate a Word report (exec summary, findings table, attendance overview, exception details, management notes).
7. Fill email templates from `.eml` sources as `X-Unsent: 1` drafts for **the user to review before sending** — never send automatically.

## Prerequisites

- Google Workspace OAuth authenticated for the active Hermes profile: `python <HERMES_HOME>/profiles/<profile>/skills/productivity/google-workspace/scripts/setup.py --check` must print `AUTHENTICATED`.
- Python packages `openpyxl` and `python-docx` installed (`pip install openpyxl python-docx`).
- The daily export CSV and the email templates (`template_idle.eml`, `template_undertime.eml`) available on disk.

## Procedure

### Step 1 — Auth check

Run once per session:

```bash
python "$HERMES/profiles/<profile>/skills/productivity/google-workspace/scripts/setup.py" --check
```

Expect `AUTHENTICATED` or `AUTHENTICATED (partial)`. Partial is acceptable if the missing scopes are not needed for this task.

### Step 2 — Parse the daily CSV

Locate the most recent `Workforce_Efficiency_Daily_Breakdown*.csv` in `~/Downloads`, or accept a path argument. Parse rows from the header row where column A is `Date` and column D is `User Name`. Skip rows where User Name is `-`, `n/a`, or empty. Extract per employee: name, email, team, start/end datetimes (columns Work Start Time / Work End Time), and logged/active/idle/matter/non-matter durations (columns Logged Hours / Idle Hours, format `HH:MM`).

The report date comes from the `Start Date` metadata row at the top of the CSV (format `="YYYY-MM-DD"`).

### Step 3 — Determine the weekly tab

Weekly tabs are named `MM-DD to MM-DD` (Monday to Friday). Compute Monday = report_date − weekday(), Friday = Monday + 4 days. If the tab does not exist in the workbook, stop and tell the user — do not guess a tab name or create one silently.

### Step 4 — Back up before updating (hard gate)

**Always** create a Drive copy of the master workbook **before** modifying it. Name it `Weekly Time Tracking Employee Report (backup YYYY-MM-DD).xlsx`. This is a user requirement — never skip it, even on re-runs.

```python
drive.files().copy(fileId=MASTER_FILE_ID, body={"name": f"... (backup {date}).xlsx"}, fields="id,name,webViewLink").execute()
```

### Step 5 — Download, fill, upload the workbook

Download the master xlsx via `drive.files().get_media(fileId=MASTER_FILE_ID)` into a scratch path. Load with `openpyxl.load_workbook`.

**Workbook layout (per employee block):**
- Name row: column B = employee name, column D = `(Scheduled hours - Xh)`, column F = `Shift (...)`
- Header row: name_row + 2, columns C..P = Scheduled Hours, Login Time, Logout Time, Actual Time, Matter Time, Non Mater Time, Manual Entries, Idle Time, Total hrs rendered, Undertime, OverTime Work, ECC Report Submitted, ...
- Day rows: name_row + 3 .. + 7 (Mon..Fri). Column C = scheduled (datetime.time), D = login, E = logout, F = actual, H = non-matter, J = idle, N = ECC Report Submitted.

For each CSV employee, find their block by name (case-insensitive contains match as fallback). Within the block, find the open day row whose Login (D) and Logout (E) are empty but Scheduled (C) is set. Write:
- D = start.time(), E = end.time()
- F = logged as a `datetime.time`, H = same
- J = idle as a `datetime.time`
- N = `"No"` if empty

If the employee has zero submitted hours, write `"No Submitted Hours"` in column C and leave N empty instead.

Save locally, then upload in place preserving the file ID:

```python
media = MediaFileUpload(str(local_xlsx), mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", resumable=True)
drive.files().update(fileId=MASTER_FILE_ID, media_body=media, fields="id,name,webViewLink,modifiedTime").execute()
```

`files().update` preserves the file ID and share link — do not use `files().create`, which would break the sheet URL.

### Step 6 — Analyze violations

Per employee present in the CSV:
- **Idle violation:** idle > 15 minutes → flagged.
- **Undertime:** scheduled − logged > 5 minutes → flagged. Only evaluate when logged > 0.
- **Missing hours:** logged is None or 0 → missing, verify before action.

Employees in the CSV but not on the weekly sheet are noted separately (add them to the sheet before next review).

### Step 7 — Generate the Word report

Use python-docx. Sections:
- Title: `Work & Attendance Report` with subtitle `<Weekday>, <Month D, YYYY> | Food For Good | Employee Time Tracking Review`
- Executive Summary: counts reviewed / idle violations / undertime violations / missing records.
- Rule / category findings table: Idle time / Undertime / Missing hours, each with violation count and review standard.
- Employee Attendance Overview table: Employee, Scheduled, Actual, Idle, Undertime, Finding. Missing-time rows show Undertime as `–`, not a computed value. Zero-hours rows show Finding = `Missing hours`.
- Exception Details table: Employee, Exception, Recorded evidence, Required management action, Status.
- Management Review Notes: caution that missing/discrepancy is not automatically unapproved absence; list anyone exceeding schedule by >1 hour; source line naming the workbook and tab.
- Appendix note: employees on the CSV but not on the weekly sheet.

Font size 9 for footnote-style paragraphs.

### Step 8 — Fill email drafts (review before send)

Load `template_idle.eml` and `template_undertime.eml`. Replace placeholders:
- `[Employee Name]` → employee name
- `[Date]` → report date as `November 13, 2026`
- Idle template: `[XX minutes]` → idle minutes (integer)
- Undertime template: two `[X hours and X minutes]` occurrences — replace 1st with recorded working hours, 2nd with undertime duration
- Refresh the `Date:` header to now (UTC)
- Keep `X-Unsent: 1`, From `eerecords@ffgp.ca`, Mike BCC (idle) / CC (undertime), and signature block intact.

Save as `draft_<idle|undertime>_<Name>_<YYYY-MM-DD>.eml`. **Do not send** — the user reviews drafts first.

## Pitfalls

- **Time cells are `datetime.time`, not `datetime`, `timedelta`, or `openpyxl.utils.datetime.to_excel()`.** openpyxl reads/writes sheet time-formatted columns (C scheduled, D login, E logout, F actual, H non-matter, J idle) as `datetime.time` objects. `to_excel()` writes an Excel serial date number, which breaks time-formatted cells and the sheet's duration formulas. Writing a `timedelta` or `datetime` also corrupts the cell. Convert durations with `time(hour, minute, second)` derived from `total_seconds() % 86400`.
- **Always back up before updating.** The user requires a Drive copy before any modification. Skipping it on re-runs or "small" edits violates a standing preference.
- **Preserve the Drive file ID.** Use `files().update` with `media_body`, not `files().create`, so the sheet URL and sharing stay stable.
- **Windows native paths.** The `google_api.py` Python client and native tools (git, node, python) need `C:/Users/...`-style forward-slash native paths, not MSYS `/c/...` paths. Bash builtins (`cd`) accept MSYS paths; native executables do not.
- **Undertime template placeholder order.** The undertime `.eml` contains `[X hours and X minutes]` twice: first is recorded working hours, second is undertime. Replace them positionally (1st → recorded, 2nd → undertime), not with a blanket replace.
- **Zero-hours rows.** An employee with no submitted FlowAce hours is `Missing hours`, not undertime. Do not compute undertime from a zero baseline.
- **Employees in CSV but not on sheet.** New hires or transfers appear in the daily export but lack a block on the weekly tab. Flag them explicitly rather than silently skipping.

## User preferences (standing)

- Backup the workbook on Drive before every update.
- Email drafts are `X-Unsent: 1` — review before sending; never auto-send.
- One sheet per week, named `MM-DD to MM-DD`.
- Do not delete or overwrite existing files without asking.

## Related

- `google-workspace` — underlying Drive/Gmail/Sheets API mechanics and OAuth setup. Load that skill for credential setup and API command reference; this skill defines the business workflow that calls it.

## Example invocation

```bash
python workflow_automation/flowace_weekly_automation.py "Downloads/Workforce_Efficiency_Daily_Breakdown_*.csv"
```

Outputs to `workflow_automation/output/`: `Work_Attendance_Report_YYYY-MM-DD.docx` and `draft_*_YYYY-MM-DD.eml` files.
