// Google Sheet sync.
export function installSync(app) {
  app.updateSyncStatusLabel = function updateSyncStatusLabel() {
    if (!app.lastSyncedDocId || !app.lastSyncedTime) {
      document.querySelectorAll('.sync-btn-text').forEach(el => el.textContent = 'Sync Spreadsheet');
      document.querySelectorAll('.sync-action-btn').forEach(btn => btn.title = 'Sync spreadsheet');
      return;
    }
    const mins = Math.floor((Date.now() - app.lastSyncedTime) / 60000);
    const statusText = mins < 1 ? 'Synced just now' : `Synced ${mins}m ago`;
    document.querySelectorAll('.sync-btn-text').forEach(el => el.textContent = `Sync Spreadsheet (${statusText})`);
    document.querySelectorAll('.sync-action-btn').forEach(btn => btn.title = `Sync spreadsheet (${statusText})`);
  };

  app.detachSheetSync = function detachSheetSync() {
    if (app.autoSyncInterval) {
      clearInterval(app.autoSyncInterval);
      app.autoSyncInterval = null;
    }
    app.lastSyncedDocId = null;
    app.lastSyncedGid = '0';
    app.lastSyncedTime = null;
    app.linkedSheet = null;
    if (app.syncBtn) app.syncBtn.style.display = 'none';
    if (app.btnClearDoc) app.btnClearDoc.style.display = 'none';
  };

  app.enableSheetSync = function enableSheetSync(docId, gid = '0') {
    app.lastSyncedDocId = docId;
    app.lastSyncedGid = gid;
    app.lastSyncedTime = Date.now();
    document.querySelectorAll('.sync-action-btn').forEach(btn => btn.style.display = 'flex');
    app.updateSyncStatusLabel();
    if (app.autoSyncInterval) clearInterval(app.autoSyncInterval);
    app.autoSyncInterval = setInterval(() => {
      app.updateSyncStatusLabel();
      if (app.lastSyncedDocId && (app.googleTokenFresh() || app.db.get(app.GOOGLE_SIGNED_IN_KEY) === 'true')) {
        app.fetchSheetData(app.lastSyncedDocId, app.h1.textContent, app.lastSyncedGid, true, true);
      }
    }, 5 * 60 * 1000);
  };
}
