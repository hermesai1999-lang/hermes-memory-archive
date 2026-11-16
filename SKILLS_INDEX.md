# Hermes Skill & Memory Index

This index tracks every skill, script, memory archive entry, and workflow that Coding Sensei has created or offloaded. It lives in the GitHub repo so it survives the ~2,200 character persistent-memory cap.

Use this index before starting any task: if a skill or workflow is already listed here, reuse it rather than rebuilding from scratch.

## Skills (in this repo)

| Skill | Path in repo | Description | Status | Date |
|-------|-------------|-------------|--------|------|
| `memory-archive` | `skills/memory-archive/SKILL.md` | Repository for facts evicted from persistent memory | Active | 2026-11-16 |
| `timesheet-workflow-automation` | `skills/timesheet-workflow-automation/SKILL.md` | FlowAce CSV → weekly sheet → Word report → email drafts → zip drafts → send via Gmail | Active | 2026-11-16 |
| `drive-upload-document` | `skills/drive-upload-document/SKILL.md` | Upload a Word document to Google Drive via CLI scripts | Active | 2026-11-16 |

## Scripts (in this repo)

| Script | Path in repo | Description | Status | Date |
|--------|-------------|-------------|--------|------|
| `flowace_weekly_automation.py` | `scripts/flowace_weekly_automation.py` | End-to-end timesheet automation: CSV parse → Drive backup → weekly tab fill → violation analysis → Word report → email drafts → zip drafts → send report + zip via Gmail API (verified end-to-end with 2026-11-13 CSV) | Active | 2026-11-16 |
| `Timesheet_Automation_User_Guide.docx` | `scripts/Timesheet_Automation_User_Guide.docx` | Layman-friendly user guide with step-by-step instructions and screenshots | Active | 2026-11-16 |

## Persistent memory entries (active, in Hermes memory store)

| Entry | Summary |
|-------|---------|
| Timesheet report trigger | "Generate me the TimeSheet Report for Today" → ask for FlowAce CSV first |
| Timesheet automation script path | `C:/Users/Administrator/workflow_automation/flowace_weekly_automation.py` |
| Master workbook Drive file ID | `1pTciNn-haOW6UcwtqVMCAE5ZTploaFc5` |
| Email templates | `Downloads/template_idle.eml`, `Downloads/template_undertime.eml` |
| GitHub PAT location | Local git credential helper (`~/.git-credentials`, repo+user scope) |
| GitHub repo | `hermesai1999-lang/hermes-memory-archive` |

## Archived memory entries (evicted, in `skills/memory-archive/SKILL.md`)

| Entry | Summary |
|-------|---------|
| Discord connector slug | `'discordbot'`, CREATE_MESSAGE send-only only |
| Docker | Not installed on Windows host; use bare Python |
| computer_use capture | Fails with 'Connection closed'; PIL ImageGrab works as fallback |
| MEDIA:/ delivery | No delivery confirmation on Telegram |
| Desktop screenshot path | `C:/Users/Administrator/Desktop/` more visible than temp scratch |

## Key external references

| Item | Value |
|------|-------|
| Master workbook (Drive) | `Weekly Time Tracking Employee Report.xlsx`, ID `1pTciNn-haOW6UcwtqVMCAE5ZTploaFc5` |
| Weekly tab naming | `MM-DD to MM-DD` (Mon–Fri) |
| Output folder | `C:/Users/Administrator/workflow_automation/output/` |
| Local clone (for repo updates) | `C:/Users/Administrator/hermes-repo-push/` |

## How to use this index

When you ask Coding Sensei to do something, it should:
1. Check this index first — does a skill/script already exist for this?
2. If yes, reuse it (load with `skill_view(name=...)` or run the script).
3. If no, build it and add a new row to this index.

Repo: https://github.com/hermesai1999-lang/hermes-memory-archive
