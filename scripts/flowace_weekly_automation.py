#!/usr/bin/env python3
"""
FlowAce -> Weekly Time Tracking Report automation.

Daily flow (run once per day after downloading the FlowAce CSV):
  1. Parse the FlowAce "Workforce Efficiency Daily Breakdown" CSV.
  2. Copy the master workbook on Google Drive (backup BEFORE updating).
  3. Fill the current week's tab (Mon-Fri) with the day's records.
  4. Upload the updated workbook back to Drive (same file ID/link).
  5. Analyse idle (>15 min) and undertime (>5 min) violations.
  6. Generate a Word report (same format as your Thursday report).
  7. Fill email templates (idle / undertime) as .eml drafts for review.

Usage:
    python flowace_weekly_automation.py [path_to_flowace_csv.csv]

If no CSV path is given, the script looks for the most recent
Workforce_Efficiency_Daily_Breakdown*.csv in ~/Downloads.
"""

from __future__ import annotations

import csv
import io
import os
import re
import sys
from datetime import date, datetime, time, timedelta
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils.datetime import to_excel
from docx import Document
from docx.shared import Pt

# --- Google Workspace (Hermes-managed OAuth) -------------------------------
SCRIPTS_DIR = r"C:\Users\Administrator\AppData\Local\hermes\profiles\coding-sensei\skills\productivity\google-workspace\scripts"
sys.path.insert(0, SCRIPTS_DIR)
from google_api import build_service  # noqa: E402
from googleapiclient.http import MediaFileUpload  # noqa: E402

# --- Constants -------------------------------------------------------------
HOME = Path.home()  # C:\Users\Administrator
DOWNLOADS = HOME / "Downloads"
PROJECT_DIR = HOME / "workflow_automation"
OUTPUT_DIR = PROJECT_DIR / "output"
TEMPLATES = {
    "idle": DOWNLOADS / "template_idle.eml",
    "undertime": DOWNLOADS / "template_undertime.eml",
}

MASTER_FILE_ID = "1pTciNn-haOW6UcwtqVMCAE5ZTploaFc5"  # Weekly Time Tracking Employee Report.xlsx
MASTER_NAME = "Weekly Time Tracking Employee Report.xlsx"

IDLE_LIMIT = timedelta(minutes=15)
UNDERTIME_LIMIT = timedelta(minutes=5)

CSV_DATE_RE = re.compile(r"(\d{2}-\d{2}-\d{4})")


# --- Time helpers ----------------------------------------------------------
def parse_hhmm(text: str) -> timedelta | None:
    """Parse 'HH:MM' (or 'HH:MM:SS') into a timedelta. Empty/invalid -> None."""
    if text is None:
        return None
    text = str(text).strip()
    if not text or text in {"00:00", "0", "0:00"}:
        return timedelta(0)
    parts = text.split(":")
    try:
        if len(parts) == 2:
            h, m = int(parts[0]), int(parts[1])
            return timedelta(hours=h, minutes=m)
        if len(parts) == 3:
            h, m, s = int(parts[0]), int(parts[1]), int(parts[2])
            return timedelta(hours=h, minutes=m, seconds=s)
    except ValueError:
        return None
    return None


def parse_csv_time(text: str) -> timedelta | None:
    """FlowAce 'Logged Hours'/'Idle Hours' are 'HH:MM'. Full datetimes fall back to None here."""
    return parse_hhmm(text)


def td_str(td: timedelta | None) -> str:
    if td is None:
        return "-"
    total = int(td.total_seconds())
    h, rem = divmod(abs(total), 3600)
    m, s = divmod(rem, 60)
    sign = "-" if total < 0 else ""
    return f"{sign}{h:02d}:{m:02d}:{s:02d}"


def td_hm(td: timedelta | None) -> str:
    """'X hours and Y minutes' style for email templates."""
    if td is None:
        return "0 hours and 0 minutes"
    total = int(abs(td.total_seconds()))
    h, m = divmod(total, 3600)
    m //= 60
    return f"{h} hours and {m} minutes"


def td_minutes(td: timedelta | None) -> float:
    if td is None:
        return 0.0
    return td.total_seconds() / 60.0


def td_to_time(td: timedelta | None) -> time | None:
    """Convert a timedelta to a time-of-day object (for sheet time-formatted cells)."""
    if td is None:
        return None
    total = int(td.total_seconds()) % 86400
    h, rem = divmod(total, 3600)
    m, s = divmod(rem, 60)
    return time(h, m, s)


# --- Step 1: parse CSV -----------------------------------------------------
def find_latest_csv() -> Path:
    candidates = sorted(DOWNLOADS.glob("Workforce_Efficiency_Daily_Breakdown*.csv"))
    if not candidates:
        raise FileNotFoundError(
            f"No FlowAce CSV found in {DOWNLOADS}. "
            "Pass the CSV path as an argument."
        )
    return candidates[-1]


def parse_flowace_csv(path: Path) -> dict:
    """Return {report_date, start, end, employees: [dict]}"""
    raw = path.read_text(encoding="utf-8-sig", errors="replace")
    reader = csv.reader(io.StringIO(raw))
    rows = [r for r in reader if any(c.strip() for c in r)]

    header_idx = None
    for i, r in enumerate(rows):
        if r and r[0].strip() == "Date" and "User Name" in r:
            header_idx = i
            break
    if header_idx is None:
        raise ValueError("Could not find the CSV header row (Date, User Name, ...).")

    header = [c.strip() for c in rows[header_idx]]
    def col(name):
        return header.index(name)

    report_date = None
    for r in rows[:header_idx]:
        if len(r) >= 2 and r[0].strip() == "Start Date":
            m = re.search(r"(\d{4}-\d{2}-\d{2})", r[1])
            if m:
                report_date = datetime.strptime(m.group(1), "%Y-%m-%d").date()
            break
    if report_date is None:
        # fall back to first data row date
        m = re.search(r"(\d{4}-\d{2}-\d{2})", rows[header_idx + 1][col("Date")])
        report_date = datetime.strptime(m.group(1), "%Y-%m-%d").date()

    employees = []
    for r in rows[header_idx + 1:]:
        if len(r) <= col("User Name"):
            continue
        name = r[col("User Name")].strip()
        if not name or name.lower() in {"-", "n/a", "none"}:
            continue
        start_raw = r[col("Work Start Time")].strip() if "Work Start Time" in header else ""
        end_raw = r[col("Work End Time")].strip() if "Work End Time" in header else ""
        employees.append({
            "name": name,
            "email": r[col("User Email")].strip() if "User Email" in header else "",
            "team": r[col("Teams")].strip() if "Teams" in header else "",
            "logged_days": r[col("Timesheet Logged Days")].strip() if "Timesheet Logged Days" in header else "",
            "start": _parse_datetime(start_raw),
            "end": _parse_datetime(end_raw),
            "logged": parse_csv_time(r[col("Logged Hours")]) if "Logged Hours" in header else None,
            "active": parse_csv_time(r[col("Active Hours")]) if "Active Hours" in header else None,
            "idle": parse_csv_time(r[col("Idle Hours")]) if "Idle Hours" in header else None,
            "matter": parse_csv_time(r[col("Matter Hours")]) if "Matter Hours" in header else None,
            "non_matter": parse_csv_time(r[col("Non Matter Hours")]) if "Non Matter Hours" in header else None,
        })
    return {"report_date": report_date, "employees": employees}


def _parse_datetime(raw: str):
    raw = raw.strip().strip('="')
    if not raw or raw == "0000-00-00 00:00:00":
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M"):
        try:
            return datetime.strptime(raw, fmt)
        except ValueError:
            pass
    return None


# --- Step 2/3: workbook helpers --------------------------------------------
def week_tab_name(report_date: date) -> str:
    """'MM-DD to MM-DD' Monday->Friday of the week containing report_date."""
    monday = report_date - timedelta(days=report_date.weekday())
    friday = monday + timedelta(days=4)
    return f"{monday:%m-%d} to {friday:%m-%d}"


def find_employee_blocks(ws):
    """Map employee name -> block info in a weekly sheet tab.

    Block layout:
      name row  : B=name, D='(Scheduled hours - Xh)', F='Shift (...)'
      header row: name_row + 2  (C..P headers)
      day rows  : name_row + 3 .. +7  (Mon..Fri); C=scheduled, D=login, E=logout,
                  F=actual, H=non-matter, J=idle, N=ECC
    Returns list of dicts: name, name_row, header_row, day_rows[5], scheduled.
    """
    blocks = []
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=6):
        name_cell = row[1].value  # column B
        if not isinstance(name_cell, str):
            continue
        name = name_cell.strip()
        if not name or name in {"Summary", "NOTES:"}:
            continue
        if name.startswith(("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday")):
            continue
        name_row = row[0].row
        sched_cell = ws.cell(row=name_row, column=4).value  # D: '(Scheduled hours - 8h)'
        shift_cell = ws.cell(row=name_row, column=6).value  # F
        header_row = name_row + 2
        day_rows = [name_row + 3 + i for i in range(5)]
        # scheduled time from first day row that has column C filled
        scheduled = None
        for dr in day_rows:
            v = ws.cell(row=dr, column=3).value
            if isinstance(v, (datetime, time)):
                scheduled = timedelta(hours=v.hour, minutes=v.minute, seconds=v.second)
                break
        blocks.append({
            "name": name,
            "name_row": name_row,
            "header_row": header_row,
            "day_rows": day_rows,
            "scheduled": scheduled,
            "sched_cell": sched_cell,
            "shift_cell": shift_cell,
        })
    return blocks


def find_open_day_row(ws, block) -> int | None:
    """Among the 5 day rows, return the one whose Login (D) is empty but Scheduled (C) set."""
    for dr in block["day_rows"]:
        login = ws.cell(row=dr, column=4).value
        logout = ws.cell(row=dr, column=5).value
        sched = ws.cell(row=dr, column=3).value
        if login is None and logout is None and isinstance(sched, (datetime, time)):
            return dr
    return None


def fill_day_row(ws, dr, rec: dict, block: dict):
    """Write one FlowAce day record into a sheet day row."""
    if rec["start"]:
        ws.cell(row=dr, column=4, value=rec["start"].time())  # D Login Time
    if rec["end"]:
        ws.cell(row=dr, column=5, value=rec["end"].time())  # E Logout Time
    if rec["logged"] is not None:
        ws.cell(row=dr, column=6, value=td_to_time(rec["logged"]))  # F Actual Time
        ws.cell(row=dr, column=8, value=td_to_time(rec["logged"]))  # H Non Mater Time
    if rec["idle"] is not None:
        ws.cell(row=dr, column=10, value=td_to_time(rec["idle"]))  # J Idle Time
    # ECC Report Submitted default 'No' when not already set
    ecc = ws.cell(row=dr, column=14).value
    if ecc in (None, ""):
        ws.cell(row=dr, column=14, value="No")


def mark_missing(ws, dr):
    ws.cell(row=dr, column=3, value="No Submitted Hours")
    ws.cell(row=dr, column=14, value=None)


# --- Google Drive ----------------------------------------------------------
def drive_backup_and_upload(local_xlsx: Path):
    drive = build_service("drive", "v3")
    timestamp = date.today().strftime("%Y-%m-%d")

    # 1. Backup copy BEFORE updating (user requirement)
    backup = drive.files().copy(
        fileId=MASTER_FILE_ID,
        body={"name": f"{MASTER_NAME} (backup {timestamp}).xlsx"},
        fields="id, name, webViewLink",
    ).execute()
    print(f"[Drive] Backup copy: {backup['name']} ({backup['id']})")

    # 2. Upload updated workbook in place (preserves file ID + share link)
    media = MediaFileUpload(
        str(local_xlsx),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        resumable=True,
    )
    updated = drive.files().update(
        fileId=MASTER_FILE_ID,
        media_body=media,
        fields="id, name, webViewLink, modifiedTime",
    ).execute()
    print(f"[Drive] Updated master: {updated['name']} ({updated['id']})")
    return backup, updated


# --- Step 5: analysis ------------------------------------------------------
def analyse(records: list[dict], blocks: list[dict], report_date: date):
    """Return violations + summary stats using CSV values vs sheet scheduled."""
    sched_by_name = {b["name"]: b["scheduled"] for b in blocks if b["scheduled"]}
    present = {r["name"]: r for r in records}

    idle_violations = []
    undertime_violations = []
    missing = []

    for r in records:
        name = r["name"]
        sched = sched_by_name.get(name)
        logged = r["logged"]
        idle = r["idle"]

        if logged is None or logged == timedelta(0):
            missing.append(name)
            continue

        if idle is not None and idle > IDLE_LIMIT:
            idle_violations.append({
                "employee": name,
                "email": r["email"],
                "idle": idle,
                "date": report_date,
            })

        if sched is not None and logged < sched:
            under = sched - logged
            if under > UNDERTIME_LIMIT:
                undertime_violations.append({
                    "employee": name,
                    "email": r["email"],
                    "scheduled": sched,
                    "actual": logged,
                    "undertime": under,
                    "date": report_date,
                })

    return {
        "report_date": report_date,
        "total_reviewed": len(records),
        "idle_violations": idle_violations,
        "undertime_violations": undertime_violations,
        "missing": missing,
    }


# --- Step 6: Word report ---------------------------------------------------
def write_word_report(analysis: dict, blocks: list[dict], records: list[dict], out_path: Path):
    sched_by_name = {b["name"]: b["scheduled"] for b in blocks if b["scheduled"]}
    present = {r["name"]: r for r in records}

    doc = Document()

    rd = analysis["report_date"]
    weekday = rd.strftime("%A")
    doc.add_heading("Work & Attendance Report", level=0)
    p = doc.add_paragraph()
    r = p.add_run(f"{weekday}, {rd:%B %d, %Y}  |  Food For Good  |  Employee Time Tracking Review")
    r.italic = True

    doc.add_heading("Executive Summary", level=1)
    n_idle = len(analysis["idle_violations"])
    n_under = len(analysis["undertime_violations"])
    n_missing = len(analysis["missing"])
    p = doc.add_paragraph(
        f"{analysis['total_reviewed']} employees reviewed. {n_idle} employees exceeded the "
        f"15-minute daily idle-time allowance, and {n_under} exceeded the five-minute "
        f"undertime tolerance. {n_missing} records show no submitted FlowAce hours "
        f"(missing hours / possible FlowAce or fieldwork issue - verify before action)."
    )

    doc.add_heading("Rule / category findings", level=1)
    table = doc.add_table(rows=1, cols=3)
    table.style = "Light Grid Accent 1"
    hdr = table.rows[0].cells
    for i, t in enumerate(["Rule / category", "Finding", "Review standard"]):
        hdr[i].text = t

    def add_find(row_cat, finding, std):
        cells = table.add_row().cells
        cells[0].text = row_cat
        cells[1].text = finding
        cells[2].text = std

    add_find("Idle time", f"{n_idle} violations", "Flag only if greater than 15m/day")
    add_find("Undertime", f"{n_under} violations", "Flag only if greater than 5m/day")
    add_find("Missing hours", f"{n_missing} no submitted hours",
             "Collect approved timesheet and verify FlowAce issue / fieldwork before assessment")

    doc.add_heading("Employee Attendance Overview", level=1)
    t2 = doc.add_table(rows=1, cols=6)
    t2.style = "Light Grid Accent 1"
    for i, t in enumerate(["Employee", "Scheduled", "Actual", "Idle", "Undertime", "Finding"]):
        t2.rows[0].cells[i].text = t

    def fmt_td(td):
        return td_str(td) if td is not None else "-"

    for b in blocks:
        name = b["name"]
        r = present.get(name)
        sched = sched_by_name.get(name)
        row = t2.add_row().cells
        row[0].text = name
        row[1].text = fmt_td(sched)
        if r is None:
            row[2].text = "-"
            row[3].text = "-"
            row[4].text = "-"
            row[5].text = "Not in daily report"
            continue
        logged = r["logged"] or timedelta(0)
        idle = r["idle"] or timedelta(0)
        row[2].text = fmt_td(logged)
        row[3].text = fmt_td(idle)
        if logged == timedelta(0):
            row[4].text = "-"
        else:
            under = sched - logged if (sched is not None and logged < sched) else timedelta(0)
            row[4].text = fmt_td(under) if under > timedelta(0) else "00:00"
        flags = []
        if idle > IDLE_LIMIT:
            flags.append("Idle")
        if logged > timedelta(0):
            under = sched - logged if (sched is not None and logged < sched) else timedelta(0)
            if under > UNDERTIME_LIMIT:
                flags.append("Undertime")
        if logged == timedelta(0):
            flags.append("Missing hours")
        row[5].text = "/".join(flags) if flags else "Within thresholds*"

    doc.add_paragraph(
        "* Within thresholds refers to recorded idle and undertime values only; it does not confirm "
        "ECC submission, schedule approval, or accuracy of time records."
    ).runs[0].font.size = Pt(9)

    sheet_names = {b["name"] for b in blocks}
    not_on_sheet = [r["name"] for r in records if r["name"] not in sheet_names]
    if not_on_sheet:
        p = doc.add_paragraph(
            "On the daily FlowAce report but not on the weekly tracking sheet: "
            + ", ".join(not_on_sheet)
            + ". Add these employees to the weekly sheet before next review."
        )
        p.runs[0].font.size = Pt(9)

    doc.add_heading("Exception Details and Required Actions", level=1)
    t3 = doc.add_table(rows=1, cols=5)
    t3.style = "Light Grid Accent 1"
    for i, t in enumerate(["Employee / scope", "Exception", "Recorded evidence", "Required management action", "Status"]):
        t3.rows[0].cells[i].text = t

    def add_exception(employee, exception, evidence, action, status):
        cells = t3.add_row().cells
        cells[0].text = employee
        cells[1].text = exception
        cells[2].text = evidence
        cells[3].text = action
        cells[4].text = status

    for v in analysis["idle_violations"]:
        add_exception(
            v["employee"], "Idle time",
            f"{td_str(v['idle'])} (limit 15m) on {v['date']:%m/%d/%Y}",
            "Review activity logs and seek explanation; document outcome.",
            "Flagged",
        )
    for v in analysis["undertime_violations"]:
        add_exception(
            v["employee"], "Undertime",
            f"{td_str(v['undertime'])} (limit 5m) on {v['date']:%m/%d} "
            f"(scheduled {td_str(v['scheduled'])}, recorded {td_str(v['actual'])})",
            "Confirm approved schedule, any fieldwork/manual entries and reason for shortfall.",
            "Flagged",
        )
    for name in analysis["missing"]:
        add_exception(
            name, "Missing hours",
            f"No FlowAce hours or manual entries recorded on {rd:%m/%d/%Y}",
            "Collect approved timesheet and verify FlowAce issue/fieldwork before assessment.",
            "Verify",
        )

    doc.add_heading("Management Review Notes", level=1)
    doc.add_paragraph(
        "A missing entry or discrepancy should not automatically be treated as an unapproved "
        "absence or unauthorized shift change. Review supporting records first. Idle/undertime "
        "flags apply only when the recorded value exceeds the daily tolerance (15m idle / 5m "
        "undertime)."
    )
    over = [r for r in records if r["logged"] and r["start"] and sched_by_name.get(r["name"])
            and r["logged"] > sched_by_name[r["name"]] + timedelta(hours=1)]
    if over:
        over_txt = "; ".join(f"{r['name']} (+{td_str(r['logged'] - sched_by_name[r['name']])})" for r in over)
        doc.add_paragraph(f"For information, recorded time exceeded the schedule by more than one hour for: {over_txt}.")

    doc.add_paragraph(
        f'Source: "{MASTER_NAME}", sheet "{week_tab_name(rd)}", {weekday} {rd:%m/%d} records. '
        "Values reflect the spreadsheet as provided; no attendance or approval data was altered."
    )

    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(out_path)
    print(f"[Report] Saved: {out_path}")

# --- Step 7: email drafts --------------------------------------------------
def load_template(kind: str) -> str:
    return TEMPLATES[kind].read_text(encoding="utf-8", errors="replace")


def fill_eml(template_text: str, employee: str, date_str: str) -> str:
    out = template_text
    out = out.replace("[Employee Name]", employee)
    out = out.replace("[Date]", date_str)
    # freshen the Date header to today
    out = re.sub(r"^Date: .*$", f"Date: {datetime.utcnow():%a, %d %b %Y %H:%M:%S +0000}", out, flags=re.M)
    return out


def replace_alternate(text: str, placeholder: str, odd_val: str, even_val: str) -> str:
    """Replace 1st, 3rd, ... occurrences with odd_val; 2nd, 4th, ... with even_val."""
    parts = text.split(placeholder)
    if len(parts) < 2:
        return text
    out = [parts[0]]
    for i in range(1, len(parts)):
        out.append(odd_val if i % 2 == 1 else even_val)
        out.append(parts[i])
    return "".join(out)


def write_email_drafts(analysis: dict, records: list[dict], out_dir: Path):
    out_dir.mkdir(parents=True, exist_ok=True)
    email_by_name = {r["name"]: r["email"] for r in records}
    written = []

    for v in analysis["idle_violations"]:
        tmpl = load_template("idle")
        mins = int(round(td_minutes(v["idle"])))
        filled = fill_eml(tmpl, v["employee"], v["date"].strftime("%B %d, %Y"))
        filled = filled.replace("[XX minutes]", f"{mins} minutes")
        path = out_dir / f"draft_idle_{v['employee'].replace(' ', '_')}_{v['date']:%Y-%m-%d}.eml"
        path.write_text(filled, encoding="utf-8")
        written.append(path)

    for v in analysis["undertime_violations"]:
        tmpl = load_template("undertime")
        recorded = td_hm(v["actual"])
        under = td_hm(v["undertime"])
        filled = replace_alternate(tmpl, "[X hours and X minutes]", recorded, under)
        filled = fill_eml(filled, v["employee"], v["date"].strftime("%B %d, %Y"))
        path = out_dir / f"draft_undertime_{v['employee'].replace(' ', '_')}_{v['date']:%Y-%m-%d}.eml"
        path.write_text(filled, encoding="utf-8")
        written.append(path)

    return written


# --- Main ------------------------------------------------------------------
def main(argv: list[str]) -> int:
    csv_path = Path(argv[1]) if len(argv) > 1 else find_latest_csv()
    print(f"[1/7] Parsing FlowAce CSV: {csv_path}")
    data = parse_flowace_csv(csv_path)
    report_date = data["report_date"]
    records = data["employees"]
    print(f"    Report date: {report_date}  |  {len(records)} employees in CSV")

    tab = week_tab_name(report_date)
    print(f"[2/7] Weekly tab: {tab}")

    scratch = Path(r"C:\Users\Administrator\AppData\Local\hermes\profiles\coding-sensei\cache\scratch")
    src_xlsx = scratch / f"Weekly_Time_Tracking_Employee_Report_{report_date:%Y-%m-%d}.xlsx"

    print(f"[3/7] Downloading master workbook from Drive ({MASTER_FILE_ID})")
    drive = build_service("drive", "v3")
    request = drive.files().get_media(fileId=MASTER_FILE_ID)
    with io.FileIO(src_xlsx, "wb") as fh:
        downloader = __import__("googleapiclient.http", fromlist=["MediaIoBaseDownload"]).MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
    print(f"    Downloaded {src_xlsx.stat().st_size} bytes")

    print(f"[4/7] Filling {tab} tab with {report_date:%A} records")
    wb = load_workbook(src_xlsx)
    if tab not in wb.sheetnames:
        print(f"    ERROR: week tab '{tab}' not in workbook. Tabs: {wb.sheetnames[:10]} ...")
        return 1
    ws = wb[tab]
    blocks = find_employee_blocks(ws)
    print(f"    Found {len(blocks)} employee blocks")

    matched = 0
    missing_in_sheet = []
    for rec in records:
        name = rec["name"]
        block = next((b for b in blocks if b["name"] == name), None)
        if block is None:
            # fuzzy: case-insensitive contains
            block = next((b for b in blocks if name.lower() in b["name"].lower()
                          or b["name"].lower() in name.lower()), None)
        if block is None:
            missing_in_sheet.append(name)
            continue
        dr = find_open_day_row(ws, block)
        if dr is None:
            print(f"    WARNING: no open day row for {name} ({block['name_row']})")
            continue
        if rec["logged"] is None or rec["logged"] == timedelta(0):
            mark_missing(ws, dr)
            print(f"    {name}: no submitted hours -> marked missing")
        else:
            fill_day_row(ws, dr, rec, block)
            print(f"    {name}: login={rec['start']:%H:%M} logout={rec['end']:%H:%M} "
                  f"logged={td_str(rec['logged'])} idle={td_str(rec['idle'])}")
        matched += 1

    if missing_in_sheet:
        print(f"    CSV employees not in sheet: {missing_in_sheet}")

    wb.save(src_xlsx)

    print(f"[5/7] Uploading updated workbook to Drive (backup first)")
    drive_backup_and_upload(src_xlsx)

    print(f"[6/7] Analysing violations (idle > 15m, undertime > 5m vs scheduled)")
    analysis = analyse(records, blocks, report_date)
    print(f"    Idle violations: {len(analysis['idle_violations'])}")
    print(f"    Undertime violations: {len(analysis['undertime_violations'])}")
    print(f"    Missing hours: {len(analysis['missing'])}")

    print(f"[7/7] Generating Word report and email drafts")
    report_path = OUTPUT_DIR / f"Work_Attendance_Report_{report_date:%Y-%m-%d}.docx"
    write_word_report(analysis, blocks, records, report_path)

    drafts = write_email_drafts(analysis, records, OUTPUT_DIR)
    print(f"    Email drafts: {len(drafts)}")

    print("\nDONE.")
    print(f"  Report : {report_path}")
    for d in drafts:
        print(f"  Draft  : {d}")
    print("\nNext: open the draft .eml files, review, then send from your mail client.")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
