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
      scope: 'https://www.googleapis.com/auth/drive.readonly https://www.googleapis.com/auth/spreadsheets.readonly',
      callback: (response) => {
        if (response.error) {
          if (response.error === 'immediate_failed' || response.error === 'interaction_required') {
            app.db.remove('google_signed_in');
            app.updateGoogleStatus(false);
          } else {
            console.warn('Google Auth silent token warning:', response.error);
            app.updateGoogleStatus(false);
          }
          return;
        }
        app.db.set('google_signed_in', 'true');
        app.accessToken = response.access_token;
        app.updateGoogleStatus(true);
        app.loadFromUrl(); // Re-check if we need to load a docId now that we're authed
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

    if (entries.length === 0) {
      wrapper.style.display = 'none';
      return;
    }

    wrapper.style.display = 'block';
    container.innerHTML = '';
    entries.slice(0, 5).forEach(([id, info]) => {
      const item = document.createElement('div');
      item.className = 'recent-doc-item';

      const dateRange = info.startDate ? `${info.startDate}${info.endDate && info.endDate !== info.startDate ? ' – ' + info.endDate : ''}` : '';

      item.innerHTML = `
        <div class="recent-doc-info">
          <div class="recent-doc-title">
            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
              <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
              <polyline points="14 2 14 8 20 8"></polyline>
              <line x1="16" y1="13" x2="8" y2="13"></line>
              <line x1="16" y1="17" x2="8" y2="17"></line>
              <polyline points="10 9 9 9 8 9"></polyline>
            </svg>
            <span>${app.escapeHtml(info.name)}</span>
          </div>
          ${dateRange ? `<span class="recent-doc-date">${app.escapeHtml(dateRange)}</span>` : ''}
        </div>
        <button class="delete-recent-btn" type="button" title="Remove from recent list" aria-label="Remove from recent list">
          <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
            <polyline points="3 6 5 6 21 6"></polyline>
            <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path>
          </svg>
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
      // Use Sheets API v4 to get formatted values and merge metadata from the first tab
      const url = `https://sheets.googleapis.com/v4/spreadsheets/${fileId}?includeGridData=true&fields=properties.title,sheets(merges,data(rowData(values(formattedValue))))`;
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

      // Requirement: Always use the first tab
      const sheet = data.sheets[0];
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
}
