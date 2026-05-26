// apps_script/Code.gs
//
// KMUTT Trends Dashboard — publish button (server side).
// Setup: set Script Properties GITHUB_PAT, SHEET_ID, REPO, HELP_URL.
//
// All long-running work lives client-side in PublishModal.html, which calls
// these server functions via google.script.run. Each call returns quickly.
//

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📤 Publish to Dashboard')
    .addItem('Check what will change (dry-run)', 'runDryRun')
    .addItem('Publish all changes', 'runPublish')
    .addSeparator()
    .addItem('📖 Help (Thai guide)', 'openHelp')
    .addToUi();
}

function openHelp() {
  const url = PropertiesService.getScriptProperties().getProperty('HELP_URL')
    || 'https://github.com/<org>/<repo>/blob/main/docs/data-collector-guide-th.md';
  const html = HtmlService.createHtmlOutput(
    `<p>เปิดคู่มือผู้รวบรวมข้อมูล:</p>
     <p><a href="${url}" target="_blank">${url}</a></p>`
  ).setWidth(420).setHeight(120);
  SpreadsheetApp.getUi().showModalDialog(html, '📖 คู่มือ');
}

function runDryRun() { _openPublishModal(true); }
function runPublish() { _openPublishModal(false); }

function _openPublishModal(dryRun) {
  // Dispatch first (fast). Then open the modal that polls.
  const dispatch = _dispatchSync(dryRun);
  if (!dispatch.ok) {
    SpreadsheetApp.getUi().alert(`❌ ${dispatch.error}`);
    return;
  }
  const template = HtmlService.createTemplateFromFile('PublishModal');
  template.correlationId = dispatch.correlationId;
  template.dryRun = dryRun;
  template.repo = _props().getProperty('REPO');
  const html = template.evaluate().setWidth(560).setHeight(440);
  SpreadsheetApp.getUi().showModalDialog(
    html, dryRun ? '🔍 ตรวจสอบความเปลี่ยนแปลง' : '📤 Publish to Dashboard'
  );
}

// ---------- functions invoked from the modal via google.script.run ----------

function dispatchSync(dryRun) { return _dispatchSync(dryRun); }

function pollStatus(correlationId) {
  const props = _props();
  const repo = props.getProperty('REPO');
  const pat = props.getProperty('GITHUB_PAT');
  // DEBUG: surface key values to Cloud logs so the modal-poll-finds-nothing
  // mystery can be diagnosed without round-tripping screenshots.
  Logger.log('[poll] correlationId=' + JSON.stringify(correlationId) + ' repo=' + JSON.stringify(repo));
  const url = `https://api.github.com/repos/${repo}/actions/runs?event=repository_dispatch&per_page=100`;
  Logger.log(`[poll] url=${url}`);
  // per_page=100 is GitHub's hard cap. We need this many because
  // `event=repository_dispatch` filter cannot also filter by event_type;
  // other dispatch types or bursty publishes could push our run off the
  // first page with the default per_page=30.
  const r = UrlFetchApp.fetch(
    url,
    { headers: _ghHeaders(pat), muteHttpExceptions: true }
  );
  const code = r.getResponseCode();
  Logger.log(`[poll] response code=${code}`);
  if (code !== 200) {
    const body = r.getContentText();
    Logger.log(`[poll] non-200 body (first 300 chars): ${body.slice(0, 300)}`);
    return { found: false, error: `GitHub API ตอบ HTTP ${code}` };
  }
  const runs = JSON.parse(r.getContentText()).workflow_runs || [];
  Logger.log(`[poll] got ${runs.length} runs; titles (first 5): ${runs.slice(0, 5).map(r => r.display_title).join(' | ')}`);
  // IMPORTANT: filter on display_title (which reflects `run-name:`), NOT name
  // (which is the workflow's static `name:` field).
  const match = runs.find(run =>
    run.display_title && run.display_title.indexOf(correlationId) >= 0
  );
  Logger.log(`[poll] match=${match ? match.id : 'none'}`);
  if (!match) return { found: false };
  return {
    found: true,
    runId: match.id,
    status: match.status,            // queued | in_progress | completed
    conclusion: match.conclusion,    // null until status==completed
    htmlUrl: match.html_url,
  };
}

function fetchResult(runId, kind) {
  // kind is 'sync-result' or 'sync-errors'
  const fileName = kind === 'sync-errors' ? 'errors.json' : 'result.json';
  return _downloadArtifactJson(runId, kind, fileName);
}

// ---------- internals ----------

function _props() { return PropertiesService.getScriptProperties(); }

function _ghHeaders(pat) {
  return { Authorization: `Bearer ${pat}`, Accept: 'application/vnd.github+json' };
}

function _dispatchSync(dryRun) {
  const props = _props();
  const pat = props.getProperty('GITHUB_PAT');
  const repo = props.getProperty('REPO');
  if (!pat || !repo) {
    return { ok: false, error: 'Script Properties ไม่ครบ (GITHUB_PAT, REPO)' };
  }
  // NOTE: sheet_id is NOT sent in client_payload. The workflow reads the
  // sheet ID from a GitHub repository Variable (KMUTT_TRENDS_SHEET_ID)
  // instead. This is defense-in-depth: even if the PAT leaks, an
  // attacker cannot redirect the workflow to a sheet they control.
  // SHEET_ID is still set as a Script Property for the Help menu / future
  // use, but the dispatch path does not depend on it.
  const correlationId = Utilities.getUuid().replace(/-/g, '').slice(0, 16);
  const userEmail = Session.getActiveUser().getEmail() || 'unknown';
  const resp = UrlFetchApp.fetch(
    `https://api.github.com/repos/${repo}/dispatches`,
    {
      method: 'post',
      headers: _ghHeaders(pat),
      contentType: 'application/json',
      payload: JSON.stringify({
        event_type: 'sync-sheets',
        client_payload: {
          user_email: userEmail,
          dry_run: dryRun,
          correlation_id: correlationId,
        },
      }),
      muteHttpExceptions: true,
    }
  );
  if (resp.getResponseCode() !== 204) {
    return { ok: false, error: `dispatch failed (HTTP ${resp.getResponseCode()}): ${resp.getContentText()}` };
  }
  return { ok: true, correlationId };
}

function _downloadArtifactJson(runId, artifactName, fileInsideZip) {
  const props = _props();
  const repo = props.getProperty('REPO');
  const pat = props.getProperty('GITHUB_PAT');
  const listResp = UrlFetchApp.fetch(
    `https://api.github.com/repos/${repo}/actions/runs/${runId}/artifacts`,
    { headers: _ghHeaders(pat), muteHttpExceptions: true }
  );
  if (listResp.getResponseCode() !== 200) return null;
  const artifact = (JSON.parse(listResp.getContentText()).artifacts || [])
    .find(a => a.name === artifactName);
  if (!artifact) return null;
  const zipResp = UrlFetchApp.fetch(artifact.archive_download_url, {
    headers: _ghHeaders(pat), followRedirects: true,
  });
  const blobs = Utilities.unzip(zipResp.getBlob().setContentType('application/zip'));
  const file = blobs.find(b => b.getName() === fileInsideZip);
  return file ? JSON.parse(file.getDataAsString('UTF-8')) : null;
}
