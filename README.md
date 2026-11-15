# Hermes Memory Archive

This repository is a long-term archive for facts, lessons, and context that were
evicted from the Hermes **Coding Sensei** persistent memory store because of its
character limit (~2,200 characters). Nothing here is deleted — it is simply
moved out of the per-turn memory budget so it can be reloaded when relevant.

## Contents

| Path | Description |
|------|-------------|
| `skills/memory-archive/SKILL.md` | The memory archive skill itself — instructions for loading and using archived facts |
| `skills/timesheet-workflow-automation/SKILL.md` | The timesheet workflow automation skill (FlowAce CSV → weekly sheet → Word report → email drafts) |
| `scripts/flowace_weekly_automation.py` | The working automation script (verified end-to-end) |

## How to use the memory archive

Load it in Hermes with:

```
skill_view(name='memory-archive')
```

If you need a specific archived fact back in active memory, tell Coding Sensei
and it can be swapped in (evicting something else if needed).

## Timesheet workflow automation

The `timesheet-workflow-automation` skill automates your daily attendance review
loop:

1. Ask for the FlowAce daily CSV (trigger: "Generate me the TimeSheet Report for Today")
2. Back up the master workbook on Google Drive before updating
3. Fill the current week's Mon–Fri tab with the day's records
4. Upload back to Drive preserving the same file ID/share link
5. Analyze idle (>15 min) and undertime (>5 min) violations
6. Generate a Word report
7. Fill `.eml` email drafts (review before send — never auto-send)

## How this repo was created

Pushed from the Coding Sensei profile on Windows via `git` + GitHub PAT
(`gh` CLI not required). Repo: `hermesai1999-lang/hermes-memory-archive`
