# Runbook: Initial Sheets Bootstrap

**Audience:** Developer (you). Run this once, before handing the workbook to the
data collector.

**When to run:** When going from "no Sheets workbook" to "live Publish flow."
If recovering from a botched bootstrap, follow the same steps and skip the ones
already done (the bootstrap script is idempotent).

**End state:** A Google Sheets workbook fully structured with data, conditional
formatting, and protected ranges; a GitHub Actions pipeline that syncs changes
and redeploys the dashboard; and an Apps Script "Publish" menu the Thai-speaking
data collector can operate without any dev involvement.

**Related docs:**
- `docs/architecture.md` — end-to-end data flow and security design
- `docs/data-collector-guide-th.md` — Thai guide to hand to the data collector

---

## Checklist

- [ ] 0. Schema-cleanup commit on main (already done in Phase 0.5)
- [ ] 1. Create empty Google Sheets workbook
- [ ] 2. GCP service account + share Sheet (Editor)
- [ ] 3. Generate GitHub PAT (Contents: write, Actions: read)
- [ ] 4. Add GitHub Secret + Variable
- [ ] 5. Run bootstrap script
- [ ] 6. (Optional) Downgrade SA to Viewer
- [ ] 7. Paste Apps Script + PublishModal
- [ ] 8. Set Script Properties
- [ ] 9. Verify Publish menu appears
- [ ] 10. Sanity-check dry-run (no diff)
- [ ] 11. Edit cell + dry-run (1 diff)
- [ ] 12. Full Publish → verify deploy completes
- [ ] 13. Test recovery (no-op Publish)
- [ ] 14. Test hard protection
- [ ] 15. Confirm credential-trust-boundary acceptance
- [ ] 16. Hand off URLs to data collector

---

### Step 0 — Schema-cleanup commit on main (Phase 0.5)

**Goal:** Confirm the `slide` field has been stripped from all JSON data files
before bootstrap runs. Bootstrap reads `web/src/data/*.json` as the source of
truth; stale fields cause spurious diffs on the first sync.

**Verify:**

```bash
grep '"slide"' web/src/data/*.json
# Expected: no output.
# Exit code 1 from grep means "no matches found" — that is the success case.
```

If any matches are found, the Phase 0.5 strip commit has not landed on `main`.
Cherry-pick or merge it before proceeding.

---

### Step 1 — Create empty Google Sheets workbook

**Goal:** Obtain a Sheet ID to pass to subsequent steps.

1. Go to [sheets.google.com](https://sheets.google.com) and sign in with your
   Google account.
2. Click **Blank** to create a new workbook.
3. Rename it to something recognisable, e.g. `KMUTT Trends Data 2568`.
4. Copy the Sheet ID from the URL bar. The URL looks like:

   ```
   https://docs.google.com/spreadsheets/d/1ABCdefGHIjklMNOpqrSTUvwxYZ/edit
                                          ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
                                          this is the Sheet ID
   ```

5. Save the Sheet ID — you will use it in Steps 4, 5, and 8.

**Verify:** The workbook opens and shows a single blank tab (`Sheet1`).

---

### Step 2 — GCP service account + share Sheet (Editor)

**Goal:** Create a non-human identity that GitHub Actions uses to read the Sheet,
and share the workbook with it.

### 2a. Create a GCP project

1. Go to [console.cloud.google.com](https://console.cloud.google.com).
2. Create a new project (or reuse an existing one). Note the project ID.
3. Enable the **Google Sheets API** and **Google Drive API**:
   - Navigation menu → **APIs & Services** → **Library** → search for each API
     → click **Enable**.

### 2b. Create a service account

1. Navigation menu → **IAM & Admin** → **Service Accounts**.
2. Click **Create Service Account**.
3. Name: `kmutt-trends-sync` (or similar). Click **Create and Continue**.
4. Skip the optional role assignment on the GCP project level (the SA needs
   Sheets access, not broad GCP access). Click **Done**.
5. Click the service account you just created → **Keys** tab → **Add Key** →
   **Create new key** → **JSON**.
6. Save the downloaded JSON file securely, e.g. `~/.config/kmutt-sa.json`.

   > **Windows path:** `%USERPROFILE%\.config\kmutt-sa.json`

7. Note the service account's email address from the service account list
   (format: `kmutt-trends-sync@<project-id>.iam.gserviceaccount.com`).

### 2c. Share the Sheet with the service account

1. Open the Sheets workbook created in Step 1.
2. Click **Share** (top-right).
3. Paste the service account email. Role: **Editor**. Uncheck "Notify people".
4. Also add the data collector's Google account as **Editor**.
5. Also add any other dev maintainers as **Editor**.

> **Why Editor for SA now?** The bootstrap script creates and deletes tabs and
> runs `batchUpdate` — those operations require Editor. After bootstrap you can
> downgrade to Viewer (Step 6).

**Verify:** The service account email appears in the Share dialog with Editor
role.

---

### Step 3 — Generate GitHub PAT (Contents: write, Actions: read)

**Goal:** Create a fine-grained Personal Access Token (PAT) scoped to this
repository that Apps Script uses to trigger `repository_dispatch`.

1. Go to **github.com** → your profile → **Settings** → **Developer settings**
   → **Personal access tokens** → **Fine-grained tokens** → **Generate new
   token**.
2. Token name: `kmutt-trends-appsscript` (or similar).
3. Expiration: choose a date appropriate for your maintenance cycle (e.g. 1 year).
4. Resource owner: the organisation or user account that owns the repo.
5. Repository access: **Only select repositories** → choose this repo.
6. Permissions — Repository permissions:
   - **Contents: Read and write** (`repository_dispatch` POST requires write;
     read-only is NOT sufficient).
   - **Actions: Read-only** (Apps Script polls workflow run status via the
     Actions API).
7. Click **Generate token** and copy it immediately — GitHub shows it only once.

**Verify:** Token starts with `github_pat_`. Store it securely; it goes into
Script Properties in Step 8, not into any file in this repo.

---

### Step 4 — Add GitHub Secret + Variable

**Goal:** Give GitHub Actions the service account credentials and the Sheet ID.

> The Sheet ID goes into a **Variable** (not a secret) because it is not
> sensitive. The workflow reads `KMUTT_TRENDS_SHEET_ID` from Variables instead
> of the `repository_dispatch` payload — this prevents a PAT holder from
> pointing the workflow at a sheet they control.

1. Go to your repository on GitHub → **Settings** → **Secrets and variables**
   → **Actions**.

### 4a. Add Secret — `GOOGLE_SERVICE_ACCOUNT_JSON`

1. Under **Secrets** tab → **New repository secret**.
2. Name: `GOOGLE_SERVICE_ACCOUNT_JSON`
3. Value: paste the full contents of the service account JSON file (the entire
   JSON object, including curly braces).
4. Click **Add secret**.

### 4b. Add Variable — `KMUTT_TRENDS_SHEET_ID`

1. Under **Variables** tab → **New repository variable**.
2. Name: `KMUTT_TRENDS_SHEET_ID`
3. Value: the Sheet ID from Step 1 (just the ID string, no URL).
4. Click **Add variable**.

**Verify:** Both appear in their respective lists. If the sync workflow runs
without `KMUTT_TRENDS_SHEET_ID`, it exits early with the error message
`"GitHub Variable KMUTT_TRENDS_SHEET_ID is not set — contact the developer"`.

---

### Step 5 — Run bootstrap script

**Goal:** Populate the workbook from `web/src/data/*.json` — creates all chart
tabs, the INDEX and STYLE tabs, conditional formatting, protected ranges, and
data-validation rules.

**Pre-conditions:**
- Service account has Editor access to the Sheet (Step 2c).
- Python dependencies are installed (`pip install -r requirements-dev.txt`).

**Command (macOS / Linux):**

```bash
python scripts/bootstrap_sheets.py \
  --sheet-id 1ABCdefGHIjklMNOpqrSTUvwxYZ \
  --credentials ~/.config/kmutt-sa.json \
  --dev-email yourname@example.com
```

**Command (Windows PowerShell):**

```powershell
python scripts\bootstrap_sheets.py `
  --sheet-id 1ABCdefGHIjklMNOpqrSTUvwxYZ `
  --credentials "$env:USERPROFILE\.config\kmutt-sa.json" `
  --dev-email yourname@example.com
```

Replace:
- `1ABCdefGHIjklMNOpqrSTUvwxYZ` → your actual Sheet ID.
- `~/.config/kmutt-sa.json` → path to your downloaded SA key file.
- `yourname@example.com` → the Google account you use to open the Sheet.
  **Without `--dev-email` you will be locked out of protected ranges.**

**What it does:**
- Drops all existing protected ranges and conditional formats (idempotent).
- Deletes all existing tabs and rebuilds them from JSON.
- Writes INDEX, STYLE_CHARTS, STYLE_SERIES, and one tab per chart.
- Applies hard protection on rows 1, 13, 14 and the STYLE tabs (editors: dev
  email + service account only).
- Applies conditional formatting for data validation feedback.

**Typical runtime:** 2–4 minutes for ~20 chart tabs.

**Verify:**

```bash
# Should print no errors and exit 0
python scripts/bootstrap_sheets.py \
  --sheet-id 1ABCdefGHIjklMNOpqrSTUvwxYZ \
  --credentials ~/.config/kmutt-sa.json \
  --dev-email yourname@example.com
```

Then open the Sheet: you should see an INDEX tab followed by tabs named like
`EDU-01`, `PER-01`, `RES-01`, etc., each populated with chart metadata and a
data table.

**If it fails:**
- `PERMISSION_DENIED` — service account does not have Editor access. Return to
  Step 2c.
- `File not found` for credentials — check the path to the JSON key file.
- `ModuleNotFoundError` — run `pip install -r requirements-dev.txt` first.
- Network errors — check your internet connection; the script uses the Sheets
  API directly.

---

### Step 6 — (Optional) Downgrade SA to Viewer

**Goal:** Apply least-privilege to the service account. The sync script only
reads the Sheet; Editor access is no longer needed post-bootstrap.

1. Open the Sheet → **Share** dialog.
2. Find the service account email → change role from **Editor** to **Viewer**.
3. Click **Save**.

> **Future re-bootstraps:** restore SA to Editor → re-run the bootstrap script →
> downgrade back to Viewer. The script is fully idempotent (deletes and rebuilds
> everything).

**Verify:** SA email shows "Viewer" in the Share dialog. The sync workflow still
succeeds (it only reads).

---

### Step 7 — Paste Apps Script + PublishModal

**Goal:** Install the "Publish to Dashboard" menu in the Sheet.

1. Open the Sheet → **Extensions** → **Apps Script**.
2. In the editor you will see a default `Code.gs` file with `function myFunction()`.
3. Select all content in `Code.gs` and delete it.
4. Open `apps_script/Code.gs` from this repo and paste its entire contents.
5. Click **+** (Add a file) → **HTML** → name it `PublishModal` (without `.html`
   extension — Apps Script appends it automatically).
6. Delete the default HTML content in the new file.
7. Open `apps_script/PublishModal.html` from this repo and paste its entire
   contents.
8. Click the floppy-disk icon (Save project) or press `Ctrl+S` / `Cmd+S`.

**Verify:** Both `Code.gs` and `PublishModal.html` appear in the left sidebar of
the Apps Script editor with content.

---

### Step 8 — Set Script Properties

**Goal:** Provide the four runtime secrets/config values that `Code.gs` reads via
`PropertiesService`.

1. In the Apps Script editor → **Project Settings** (gear icon in the left
   sidebar) → scroll to **Script Properties** → **Add script property**.
2. Add all four properties:

   | Property   | Value                                                       |
   |------------|-------------------------------------------------------------|
   | `GITHUB_PAT` | The fine-grained PAT from Step 3                          |
   | `SHEET_ID`   | The Sheet ID from Step 1                                  |
   | `REPO`       | `org/repo` — e.g. `tayakorn221/kmutt-trends`             |
   | `HELP_URL`   | Full URL to `docs/data-collector-guide-th.md` on GitHub   |

   For `HELP_URL`, navigate to the file in GitHub and copy the URL from the
   browser, e.g.:
   `https://github.com/org/repo/blob/main/docs/data-collector-guide-th.md`

3. Click **Save script properties**.

**Verify:** All four properties appear in the Script Properties list (values are
masked after saving).

---

### Step 9 — Verify Publish menu appears

**Goal:** Confirm Apps Script is wired up correctly.

1. Close the Apps Script editor tab and return to the Sheet.
2. Reload the Sheet (F5 / Cmd+R).
3. After the sheet loads, a menu labelled **📤 Publish to Dashboard** should
   appear in the menu bar (between `Help` and any other add-on menus).
4. Click it — you should see three items:
   - Check what will change (dry-run)
   - Publish all changes
   - 📖 Help (Thai guide)

**If the menu does not appear:**
- Check that you saved the Apps Script project (Step 7 step 8).
- Look for errors in the Apps Script editor: **Executions** tab shows run history.
- Ensure `onOpen` is the trigger: **Triggers** (clock icon) → should have a
  `onOpen` entry bound to "From spreadsheet → On open".
  If not, click **Add Trigger** → function: `onOpen`, event: "On open" → Save.
  You may be prompted to authorise the script.

---

### Step 10 — Sanity-check dry-run (no diff)

**Goal:** Confirm that immediately after bootstrap the Sheet content matches the
JSON on `main` — i.e., zero changes to publish.

1. In the Sheet menu → **📤 Publish to Dashboard** → **Check what will change
   (dry-run)**.
2. The modal opens, shows a spinner, then reports the result.

**Expected result:** "No changes detected" (or equivalent — the Thai text reads
"ไม่พบการเปลี่ยนแปลง"). Zero files changed.

**If diffs are reported:** The Sheet content diverges from the JSON. Most likely
cause: bootstrap read from a different branch than `main`. Ensure you ran
bootstrap from the repo root with `main` checked out and with the Phase 0.5
commit present (Step 0).

---

### Step 11 — Edit a cell + dry-run (expect 1 diff)

**Goal:** Confirm the diff-detection pipeline is actually working.

1. Open any chart tab, e.g. `EDU-01`.
2. Find the `2568` row in the data table (row 15 or thereabouts).
3. Change the value in the bachelor column by ±1 (e.g. `12345` → `12346`).
4. Menu → **📤 Publish to Dashboard** → **Check what will change (dry-run)**.

**Expected result:** Exactly 1 file shown as changed (the JSON for that chart).
The diff summary should name the modified chart ID.

5. **Undo the edit** (Ctrl+Z / Cmd+Z) after confirming — do not leave test data
   in the Sheet.

**If zero diffs are reported:** The sync script is not reading the cell you
edited. Check the tab name matches a chart in `web/src/data/` and that the row
you edited falls inside the data table (row 15 onward, not the metadata rows).

---

### Step 12 — Full Publish → verify deploy completes

**Goal:** Confirm the entire pipeline — dispatch → sync → commit → deploy →
Pages — works end to end.

1. Make the same small edit as Step 11 (change a value by ±1).
2. Menu → **📤 Publish to Dashboard** → **Publish all changes**.
3. The modal opens and polls GitHub Actions every 5 seconds.

**Expected sequence inside the modal:**
1. "Dispatching…" → dispatched (workflow triggered).
2. "Syncing…" → sync job running.
3. "Deploying…" → deploy job running.
4. "Success" message — **only after the deploy job finishes**.

**Verify on GitHub:**
- Go to the repository → **Actions** tab.
- You should see a run named `Sync from Sheets [<correlation_id>]`.
- Both the `sync` and `deploy` jobs should be green.

**Verify on the dashboard:**
- Open the GitHub Pages URL (shown in the Actions run or in repo Settings →
  Pages).
- The chart you edited should reflect the new value.

**If the modal reports an error:** Open the failing Actions run and read the
step logs. Common causes:
- `GOOGLE_SERVICE_ACCOUNT_JSON` not set → Step 4a.
- `KMUTT_TRENDS_SHEET_ID` not set → Step 4b.
- Service account lacks read access to the Sheet → Step 2c or Step 6.

---

### Step 13 — Test recovery (no-op Publish)

**Goal:** Confirm that clicking Publish when there are no changes is safe and
recoverable (proves "repo updated but Pages stuck" scenario is fixable by
clicking Publish again).

1. Without making any edits to the Sheet, open the Publish menu → **Publish all
   changes**.
2. Watch the modal through completion.

**Expected behaviour:**
- Sync job runs and finds no JSON changes → skips the git commit.
- Deploy job still runs (this is intentional — it re-publishes the current
  `main` to Pages).
- Modal reports **success**.

**Why this matters:** If a previous deploy job failed after a successful sync
commit, clicking Publish again with no edits will re-trigger the deploy without
creating a duplicate commit. The dashboard catches up without manual
intervention.

---

### Step 14 — Test hard protection

**Goal:** Confirm that the protected ranges are enforced as hard blocks (not
just warnings), preventing data-collector edits to metadata rows.

1. Open the Sheet in an incognito window (or use a different browser signed into
   the **data collector's Google account**, not your dev account).
2. Navigate to any chart tab, e.g. `EDU-01`.
3. Try to edit cell `B14` — this is the `series_key` cell for column B, which
   sits in the protected row 14.

**Expected response from Sheets:**

> "You are trying to edit a protected cell or object. Please contact the
> spreadsheet owner to remove protection if you need to edit."

The popup has only a **Close** button — **no "OK to edit anyway" button**.

**If the popup shows a yellow warning with an "OK" button:** The protection was
set to `warningOnly` instead of hard. Re-run the bootstrap script (Step 5) —
this resets all protections. Check that `--dev-email` was passed correctly so
the bootstrap script can set you (not the SA alone) as the authorised editor.

**Also test row 1 and row 13** (chart metadata) in the same incognito session.
All three rows must show the hard-block popup.

---

### Step 15 — Confirm credential-trust-boundary acceptance

**Goal:** Ensure you have read and accept the credential security model before
handing the Sheet to the data collector.

1. Open `docs/architecture.md` → section **Credential trust boundary**.
2. Read it carefully. Key points:
   - The GitHub PAT stored in Script Properties has `Contents: write` access.
     If it leaks, an attacker can commit to this repo (but not run arbitrary
     code — the workflow only runs the pinned sync script from `main`).
   - The Sheet ID in Script Properties is not secret (same value as the GitHub
     Variable). If leaked, an attacker learns which Sheet triggers the workflow —
     but they cannot redirect the workflow to a different Sheet because
     `KMUTT_TRENDS_SHEET_ID` is read from the GitHub Variable, not the dispatch
     payload.
   - The service account JSON (GitHub Secret) is never exposed to Apps Script.
     Only GitHub Actions reads it, and only to authenticate against Sheets.
3. If you accept this model, proceed. If you have concerns, discuss with the
   team before handing off.

---

### Step 16 — Hand off URLs to data collector

**Goal:** Give the data collector everything they need to operate independently.

Provide the following to the data collector (in whatever channel you use):

1. **Sheet URL** — direct link to the workbook, e.g.:
   `https://docs.google.com/spreadsheets/d/1ABCdefGHIjklMNOpqrSTUvwxYZ/edit`
2. **Thai guide URL** — link to `docs/data-collector-guide-th.md` on GitHub, e.g.:
   `https://github.com/org/repo/blob/main/docs/data-collector-guide-th.md`
3. Confirm the data collector's Google account has **Editor** access to the Sheet
   (added in Step 2c). If not, add it now via the Share dialog.
4. Walk them through the Publish menu at least once (or share a screen recording).

**Verify (final check):** Ask the data collector to open the Sheet on their own
device, make a small edit, run a dry-run, undo the edit, and confirm they can
see the diff in the modal. This confirms their account, network, and the Apps
Script are all working before you step away.

---

_Bootstrap complete. The data collector can now edit the Sheet and click
"Publish all changes" to push updates to the live dashboard._
