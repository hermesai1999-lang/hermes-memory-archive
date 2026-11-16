---
name: drive-upload-document
description: "Upload a Word document to Google Drive via CLI scripts."
category: productivity
version: 1.0.0
author: Hermes Agent
license: MIT
metadata:
  hermes:
    tags: [google-drive, upload, word-document]
    related_skills: [google-workspace]
---

# Drive Document Upload

## When to Use

Use when you need to upload a locally-created .docx document to Google Drive using the Hermes google-workspace skill's Python scripts.

Reusable process for uploading a Word document to Google Drive using the Hermes google-workspace skill's Python scripts.

## Prerequisites

- OAuth credentials at `C:/Users/Administrator/AppData/Local/hermes/google_token.json` and `google_client_secret.json`
- Auth: `setup.py --check` returns AUTHENTICATED

## Script paths

- Setup: `C:/Users/Administrator/AppData/Local/hermes/profiles/coding-sensei/skills/productivity/google-workspace/scripts/setup.py`
- API: `C:/Users/Administrator/AppData/Local/hermes/profiles/coding-sensei/skills/productivity/google-workspace/scripts/google_api.py`

## Steps

1. **Verify auth:**

   ```bash
   python setup.py --check
   ```

   Should return `AUTHENTICATED`.

2. **Upload the file:**

   ```bash
   python google_api.py drive upload <absolute_path_to_docx>
   ```

   Returns JSON with `id`, `name`, `webViewLink`.

3. **Result fields:** `id` = Drive file ID, `webViewLink` = shareable link

## Pitfalls

- Use absolute Windows paths (`C:/Users/...`), never MSYS `/c/...` paths.
- Upload creates a new file each time; does not update existing by default.
- If auth shows partial (missing Drive scope), run `$GSETUP --revoke` then redo OAuth.
- Upload is fast (~1 second for a small .docx). If it seems hung, the process may be waiting on token refresh.

## Last known good

- File: `C:/Users/Administrator/AppData/Local/hermes/profiles/coding-sensei/cache/scratch/daily-routes-process-flow.docx`
- Drive ID: `1QN9gPOujxY_eGwPu3jPHKKee9hQomaTH`
- Link: `https://docs.google.com/document/d/1QN9gPOujxY_eGwPu3jPHKKee9hQomaTH/edit`

---

*Maintained by Coding Sensei (profile `coding-sensei`).*
