// Google Drive import.
export function installDrive(app) {
  app.updateGoogleStatus = function updateGoogleStatus(signedIn, loading = false) {
    if (loading) {
      app.googleStatus.textContent = 'Connecting to Google...';
      app.googleStatus.style.color = 'var(--text-muted)';
      app.btnGoogleSignIn.disabled = true;
      return;
    }
    app.btnGoogleSignIn.disabled = false;
    if (signedIn) {
      app.googleStatus.textContent = 'Signed in';
      app.googleStatus.style.color = 'var(--success)';
      app.btnGoogleSignIn.innerHTML = `<img src="https://fonts.gstatic.com/s/i/productlogos/googleg/v6/24px.svg" alt="" width="18"> Sign out of Google`;
      if (app.driveManualInput) app.driveManualInput.style.display = 'flex';
      if (app.btnDrivePicker) app.btnDrivePicker.style.display = 'flex';
      const warning = document.getElementById('drive-auth-warning');
      if (warning) warning.style.display = 'none';
    } else {
      app.googleStatus.textContent = 'Not signed in';
      app.googleStatus.style.color = 'var(--text-muted)';
      app.btnGoogleSignIn.innerHTML = `<img src="https://fonts.gstatic.com/s/i/productlogos/googleg/v6/24px.svg" alt="" width="18"> Sign in with Google`;
      if (app.driveManualInput) app.driveManualInput.style.display = 'none';
      if (app.btnDrivePicker) app.btnDrivePicker.style.display = 'none';
      const warning = document.getElementById('drive-auth-warning');
      if (warning) warning.style.display = 'block';
      app.accessToken = null;
    }
  };

  app.initTokenClient = function initTokenClient() {
    if (!app.googleClientId || app.tokenClient) return;

    // Check if the Google Identity Services script has finished loading
    if (typeof google === 'undefined' || !google.accounts) {
      return;
    }

    // 1. Initialize Identity (Sign In With Google / One Tap)
    // This handles the "Automatic Login" feature
    google.accounts.id.initialize({
      client_id: app.googleClientId,
      auto_select: true,
      context: 'signin',
      callback: (response) => {
        // Successful automatic or manual identity verification
        if (app.tokenClient && !app.accessToken) {
          app.tokenClient.requestAccessToken({ prompt: '' });
        }
      }
    });

    // 2. Initialize Authorization (Google Drive access)
    app.tokenClient = google.accounts.oauth2.initTokenClient({
      client_id: app.googleClientId,
      scope: 'https://www.googleapis.com/auth/drive.readonly https://www.googleapis.com/auth/spreadsheets',
      callback: (response) => {
        const pending = app.pendingAuth;
        app.pendingAuth = null;
        if (response.error) {
          if (response.error === 'immediate_failed' || response.error === 'interaction_required') {
            app.db.remove('google_signed_in');
            app.updateGoogleStatus(false);
          } else {
            console.warn('Google Auth silent token warning:', response.error);
            app.updateGoogleStatus(false);
          }
          if (pending) {
            const cancelled = response.error === 'access_denied' || response.error === 'popup_closed';
            pending.reject(new Error(cancelled ? 'Google sign-in was cancelled.' : 'Google sign-in failed. Try again.'));
          }
          return;
        }
        app.db.set('google_signed_in', 'true');
        app.accessToken = response.access_token;
        app.updateGoogleStatus(true);
        if (pending) pending.resolve(app.accessToken);
        else app.loadFromUrl();
      },
    });
  };

  app.trackRecentDoc = function trackRecentDoc(fileId, name, gid = '0') {
    const key = `${fileId}_${gid}`;
    const metadata = JSON.parse(app.db.get(app.DOC_METADATA_KEY, '{}'));

    let startDate = metadata[key]?.startDate || null;
    let endDate = metadata[key]?.endDate || null;

    if (app.parsed && app.parsed.dates && app.parsed.dates.length > 1) {
      const validDates = app.parsed.dates.slice(1).filter(d => d && d.trim());
      if (validDates.length > 0) {
        startDate = validDates[0];
        endDate = validDates[validDates.length - 1];
      }
    }

    metadata[key] = {
      fileId,
      name: name || metadata[key]?.name || 'Untitled Itinerary',
      lastAccessed: Date.now(),
      startDate,
      endDate,
      gid
    };
    app.db.set(app.DOC_METADATA_KEY, JSON.stringify(metadata));
    app.renderRecentDocs();
  };

  app.renderRecentDocs = function renderRecentDocs() {
    const container = document.getElementById('recent-docs-list');
    const wrapper = document.getElementById('recent-docs-container');
    const empty = document.getElementById('recent-docs-empty');
    const head = wrapper?.querySelector('.recent-docs-head');
    const sortBy = app.recentDocsSort?.value || 'recent';
    if (!container || !wrapper) return;

    const metadata = JSON.parse(app.db.get(app.DOC_METADATA_KEY, '{}'));
    const entries = Object.entries(metadata)
      .filter(([id, info]) => app.db.get(app.DOC_CACHE_PREFIX + id))
      .sort((a, b) => {
        if (sortBy === 'name') {
          return a[1].name.localeCompare(b[1].name);
        } else if (sortBy === 'date') {
          const tsA = app.parseDateForSort(a[1].startDate);
          const tsB = app.parseDateForSort(b[1].startDate);
          if (tsA === Infinity && tsB === Infinity) return 0;
          if (tsA === Infinity) return 1;
          if (tsB === Infinity) return -1;
          return tsB - tsA; // Newest trip first
        }
        return b[1].lastAccessed - a[1].lastAccessed;
      });

    const tripOpen = app.itinerarySection?.classList.contains('visible');
    if (!tripOpen && app.emptyState) {
      app.emptyState.style.display = entries.length ? 'none' : 'block';
    }

    if (entries.length === 0) {
      if (head) head.hidden = true;
      container.hidden = true;
      container.innerHTML = '';
      if (empty) empty.hidden = false;
      return;
    }

    if (head) head.hidden = false;
    container.hidden = false;
    if (empty) empty.hidden = true;
    container.innerHTML = '';
    entries.slice(0, 5).forEach(([id, info]) => {
      const item = document.createElement('div');
      item.className = 'recent-doc-item';

      const dateRange = info.startDate ? `${info.startDate}${info.endDate && info.endDate !== info.startDate ? ' – ' + info.endDate : ''}` : '';

      item.innerHTML = `
        <div class="recent-doc-info">
          <div class="recent-doc-title">
            ${app.icon('file', 16)}
            <span>${app.escapeHtml(info.name)}</span>
          </div>
          ${dateRange ? `<span class="recent-doc-date">${app.escapeHtml(dateRange)}</span>` : ''}
        </div>
        <button class="delete-recent-btn" type="button" title="Remove from recent list" aria-label="Remove from recent list">
          ${app.icon('trash', 14)}
        </button>
      `;

      item.querySelector('.recent-doc-info').onclick = () => {
        app.updateUrlWithDocId(info.fileId || id.split('_')[0], info.gid || '0');
        app.loadFromUrl();
      };

      const delBtn = item.querySelector('.delete-recent-btn');
      delBtn.onclick = (e) => {
        e.stopPropagation();
        if (confirm(`Remove "${info.name}" from your recent list?`)) {
          const metadata = JSON.parse(app.db.get(app.DOC_METADATA_KEY, '{}'));
          delete metadata[id];
          app.db.set(app.DOC_METADATA_KEY, JSON.stringify(metadata));
          app.db.remove(app.DOC_CACHE_PREFIX + id);
          renderRecentDocs();
        }
      };

      container.appendChild(item);
    });
  };

  app.createPicker = function createPicker() {
    const view = new google.picker.View(google.picker.ViewId.SPREADSHEETS);
    const picker = new google.picker.PickerBuilder()
      .addView(view)
      .setOAuthToken(app.accessToken)
      .setCallback(app.pickerCallback)
      .build();
    picker.setVisible(true);
  };

  app.fetchSheetData = async function fetchSheetData(fileId, fileName, gid = '0', skipHistory = false, silent = false) {
    if (!silent) app.showError(`Fetching "${fileName}"...`);
    try {
      // Use Sheets API v4 to get formatted values and merge metadata
      const url = `https://sheets.googleapis.com/v4/spreadsheets/${fileId}?includeGridData=true&fields=properties.title,sheets(properties(sheetId,title),merges,data(rowData(values(formattedValue))))`;
      const response = await fetch(url, {
        headers: { 'Authorization': `Bearer ${app.accessToken}` }
      });

      if (!response.ok) {
        const errorBody = await response.json().catch(() => ({}));
        const apiMessage = errorBody.error?.message || 'Ensure it is a Google Sheet and you have permission to view it.';
        throw new Error(`API Status ${response.status}: ${apiMessage}`);
      }

      const data = await response.json();
      const spreadsheetTitle = data.properties?.title || fileName;

      const sheets = data.sheets || [];
      const requestedGid = String(gid || '0');
      let sheet = requestedGid !== '0'
        ? sheets.find(item => String(item.properties?.sheetId) === requestedGid)
        : null;
      if (!sheet) sheet = sheets[0];
      if (!sheet) throw new Error('Spreadsheet contains no sheets.');

      const rows = [];
      const rowData = sheet.data[0].rowData || [];

      // 1. Build initial grid from formatted values
      rowData.forEach((row, r) => {
        rows[r] = [];
        (row.values || []).forEach((cell, c) => {
          rows[r][c] = cell.formattedValue || "";
        });
      });

      // 2. Handle Merged Cells: Fill merged ranges with the top-left value
      (sheet.merges || []).forEach(m => {
        const startR = m.startRowIndex || 0;
        const startC = m.startColumnIndex || 0;
        const endR = m.endRowIndex || rows.length;
        // Determine column end based on the first row of the merge or total grid width
        const endC = m.endColumnIndex || (rows[startR] ? rows[startR].length : 0);

        const val = rows[startR] ? rows[startR][startC] : "";
        if (val) {
          for (let r = startR; r < endR; r++) {
            if (!rows[r]) rows[r] = [];
            for (let c = startC; c < endC; c++) {
              rows[r][c] = val;
            }
          }
        }
      });

      app.db.set(app.DOC_CACHE_PREFIX + fileId + '_' + gid, JSON.stringify(rows));
      app.rememberLinkedSheet(fileId, gid, rows, {
        sheetId: sheet.properties?.sheetId ?? 0,
        sheetName: sheet.properties?.title || 'Sheet1',
        title: spreadsheetTitle,
        merges: sheet.merges || []
      });
      app.h1.textContent = spreadsheetTitle;
      app.trackRecentDoc(fileId, spreadsheetTitle, gid);
      if (!skipHistory) {
        app.updateUrlWithDocId(fileId, gid);
      }
      if (app.btnClearDoc) app.btnClearDoc.style.display = 'block';
      app.buildItineraryFromRows(rows, true);
      app.enableSheetSync(fileId, gid);
      if (silent) app.showError('');
    } catch (e) {
      console.error(e);
      const cached = app.db.get(app.DOC_CACHE_PREFIX + fileId + '_' + gid);
      if (cached) {
        if (!silent) app.showError('');
        console.warn('Network fetch failed, but cached data is available:', e);
      } else if (!silent) {
        app.showError('Error loading from Sheets: ' + e.message);
      }
    }
  };

  app.pickerCallback = async function pickerCallback(data) {
    if (data.action === google.picker.Action.PICKED) {
      const doc = data.docs[0];
      // Populate the input with the user-friendly Doc ID URL after picking
      app.driveUrlInput.value = `https://docs.google.com/spreadsheets/d/${doc.id}/edit`;
      app.fetchSheetData(doc.id, doc.name);
    }
  };

  app.rememberLinkedSheet = function rememberLinkedSheet(fileId, gid, grid, meta = {}) {
    const previous = app.linkedSheet;
    app.linkedSheet = {
      fileId,
      gid: String(gid == null ? '0' : gid),
      sheetId: meta.sheetId == null ? (previous && previous.fileId === fileId ? previous.sheetId : null) : meta.sheetId,
      sheetName: meta.sheetName || (previous && previous.fileId === fileId ? previous.sheetName : '') || 'Sheet1',
      title: meta.title || (previous && previous.fileId === fileId ? previous.title : '') || '',
      grid: (grid || []).map(row => (row || []).map(cell => (cell == null ? '' : String(cell)))),
      merges: meta.merges
        ? meta.merges.map(merge => ({
          startRowIndex: merge.startRowIndex || 0,
          endRowIndex: merge.endRowIndex || 0,
          startColumnIndex: merge.startColumnIndex || 0,
          endColumnIndex: merge.endColumnIndex || 0
        }))
        : ((previous && previous.fileId === fileId && previous.merges) || [])
    };
  };

  app.stampSheetId = function stampSheetId(request, sheetId) {
    const copy = JSON.parse(JSON.stringify(request));
    const walk = (node) => {
      if (!node || typeof node !== 'object') return;
      if (Object.prototype.hasOwnProperty.call(node, 'sheetId')) node.sheetId = sheetId;
      Object.values(node).forEach(walk);
    };
    walk(copy);
    return copy;
  };

  app.sheetsRequest = async function sheetsRequest(path, options = {}) {
    const response = await fetch(`https://sheets.googleapis.com/v4/${path}`, {
      method: options.method || 'GET',
      headers: {
        Authorization: `Bearer ${app.accessToken}`,
        ...(options.body ? { 'Content-Type': 'application/json' } : {})
      },
      body: options.body ? JSON.stringify(options.body) : undefined
    });
    const data = await response.json().catch(() => ({}));
    if (!response.ok) {
      const details = data.error?.details || [];
      const error = new Error(data.error?.message || `Sheets request failed (${response.status})`);
      error.status = response.status;
      error.reason = details.map(detail => detail.reason).filter(Boolean).join(' ') || data.error?.status || '';
      throw error;
    }
    return data;
  };

  app.isSheetsScopeError = function isSheetsScopeError(error) {
    return error.status === 403 && /ACCESS_TOKEN_SCOPE_INSUFFICIENT|insufficient authentication scopes/i.test(`${error.reason || ''} ${error.message || ''}`);
  };

  app.sheetsPermissionMessage = function sheetsPermissionMessage(error) {
    if (error.status === 403 && !app.isSheetsScopeError(error)) {
      return 'Google did not allow editing this spreadsheet. Fork a copy instead, or ask the owner for edit access.';
    }
    return error.message || 'Could not save to Google Sheets.';
  };

  app.ensureSheetsAccess = function ensureSheetsAccess(options = {}) {
    if (app.accessToken && !options.force) return Promise.resolve(app.accessToken);
    return new Promise((resolve, reject) => {
      app.initTokenClient();
      if (!app.tokenClient) {
        reject(new Error('Google sign-in is still loading. Wait a moment and try again.'));
        return;
      }
      app.pendingAuth = { resolve, reject };
      app.tokenClient.requestAccessToken({ prompt: options.prompt || (options.force ? 'consent' : 'select_account') });
    });
  };

  app.refreshLinkedSheetMeta = async function refreshLinkedSheetMeta() {
    const link = app.linkedSheet;
    const data = await app.sheetsRequest(`spreadsheets/${encodeURIComponent(link.fileId)}?fields=properties.title,sheets(properties(sheetId,title),merges)`);
    const sheets = data.sheets || [];
    let sheet = null;
    if (link.gid && link.gid !== '0') {
      sheet = sheets.find(item => String(item.properties?.sheetId) === String(link.gid));
    }
    if (!sheet && link.sheetId != null) {
      sheet = sheets.find(item => item.properties?.sheetId === link.sheetId);
    }
    if (!sheet) sheet = sheets[0];
    if (!sheet) throw new Error('Spreadsheet contains no sheets.');
    link.sheetId = sheet.properties?.sheetId ?? 0;
    link.sheetName = sheet.properties?.title || 'Sheet1';
    if (!link.title) link.title = data.properties?.title || '';
    link.merges = (sheet.merges || []).map(merge => ({
      startRowIndex: merge.startRowIndex || 0,
      endRowIndex: merge.endRowIndex || 0,
      startColumnIndex: merge.startColumnIndex || 0,
      endColumnIndex: merge.endColumnIndex || 0
    }));
    return link;
  };

  app.writeLinkedSheet = async function writeLinkedSheet(days, title) {
    if (!app.linkedSheet?.fileId) throw new Error('This trip is not linked to a Google Sheet.');
    if (!app.accessToken) await app.ensureSheetsAccess();
    let link;
    try {
      link = await app.refreshLinkedSheetMeta();
    } catch (error) {
      if (app.isSheetsScopeError(error) || error.status === 401) {
        await app.ensureSheetsAccess({ force: true, prompt: 'consent' });
        link = await app.refreshLinkedSheetMeta();
      } else {
        throw new Error(app.sheetsPermissionMessage(error));
      }
    }
    const plan = app.planLinkedSheetEdits(link.grid, days, link.merges || []);
    const requests = plan.requests.map(request => app.stampSheetId(request, link.sheetId));
    plan.cells.forEach(cell => {
      requests.push(app.stampSheetId({
        updateCells: {
          range: {
            sheetId: 0,
            startRowIndex: cell.row,
            endRowIndex: cell.row + 1,
            startColumnIndex: cell.col,
            endColumnIndex: cell.col + 1
          },
          rows: [{ values: [{ userEnteredValue: { stringValue: cell.value } }] }],
          fields: 'userEnteredValue'
        }
      }, link.sheetId));
    });
    if (title && title !== link.title) {
      requests.push({
        updateSpreadsheetProperties: {
          properties: { title },
          fields: 'title'
        }
      });
    }
    if (requests.length) {
      try {
        await app.sheetsRequest(`spreadsheets/${encodeURIComponent(link.fileId)}:batchUpdate`, {
          method: 'POST',
          body: { requests }
        });
      } catch (error) {
        if (app.isSheetsScopeError(error) || error.status === 401) {
          await app.ensureSheetsAccess({ force: true, prompt: 'consent' });
          await app.sheetsRequest(`spreadsheets/${encodeURIComponent(link.fileId)}:batchUpdate`, {
            method: 'POST',
            body: { requests }
          });
        } else {
          throw new Error(app.sheetsPermissionMessage(error));
        }
      }
    }
    link.grid = plan.grid;
    link.merges = plan.merges;
    link.title = title || link.title;
    return plan;
  };
}
