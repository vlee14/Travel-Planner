(function () {
  const csvInput = document.getElementById('csv-input');
  const btnParse = document.getElementById('btn-parse');
  const errorMsg = document.getElementById('error-msg');
  const itinerarySection = document.getElementById('itinerary-section');
  const inputSection = document.querySelector('.input-section');
  const emptyState = document.getElementById('empty-state');
  const listContent = document.getElementById('list-content');
  const tabs = document.querySelectorAll('.tab');
  const panels = document.querySelectorAll('.input-panel');
  const viewBtns = document.querySelectorAll('.view-btn[data-view]');
  const themeToggle = document.getElementById('theme-toggle');
  const homeBtn = document.getElementById('home-btn');
  const h1 = document.querySelector('.header-row h1');
  const shareBtn = document.getElementById('share-btn');
  const searchInput = document.getElementById('search-input');
  const settingsBtn = document.getElementById('settings-btn');
  const settingsModal = document.getElementById('settings-modal');
  const settingsClose = document.getElementById('settings-close');
  const settingApiKey = document.getElementById('setting-api-key');
  const settingVcKey = document.getElementById('setting-vc-key');
  const settingModel = document.getElementById('setting-model');
  const btnGoogleSignIn = document.getElementById('btn-google-signin');
  const btnDriveAuth = document.getElementById('btn-drive-auth');
  const googleStatus = document.getElementById('google-status');
  const btnDrivePicker = document.getElementById('btn-drive-picker');
  const recentDocsSort = document.getElementById('recent-docs-sort');
  const btnClearDoc = document.getElementById('btn-clear-doc');
  const driveUrlInput = document.getElementById('drive-url-input');
  const btnDriveUrlFetch = document.getElementById('btn-drive-url-fetch');
  const driveManualInput = document.getElementById('drive-manual-input');
  const btnSaveSettings = document.getElementById('btn-save-settings');
  const btnClearCache = document.getElementById('btn-clear-cache');
  const aiModal = document.getElementById('ai-modal');
  const aiModalClose = document.getElementById('ai-modal-close');
  const aiModalTitle = document.getElementById('ai-modal-title');
  const aiModalBody = document.getElementById('ai-modal-body');
  const aiFollowUpInput = document.getElementById('ai-follow-up-input');
  const aiFollowUpBtn = document.getElementById('ai-follow-up-btn');
  const packingCountInput = document.getElementById('packing-count-input');
  const btnGenerateChecklist = document.getElementById('btn-generate-checklist');
  const packingContent = document.getElementById('packing-content');
  const packingTabs = document.querySelectorAll('.packing-tab');

  let parsed = null; // { locations, dates, rows: [ { time, cells[] } ] }
  let map = null;
  let mapLayerGroup = null;
  let mapTileLayer = null;
  let polyline = null;
  const mapPoints = []; // Array of {lat, lon}
  const mapMarkers = {}; // Object to track unique locations
  let hasScrolledToToday = false
  const GEMINI_API_KEY_STORAGE = 'gemini_api_key';
  const VC_API_KEY_STORAGE = 'visual_crossing_key';
  const WEATHER_CACHE_PREFIX = 'weather_cache_v1_';
  const WEATHER_CACHE_EXPIRY = 1000 * 60 * 60 * 6; // 6 hours
  const AI_CACHE_PREFIX = 'ai_chat_cache_v1_';
  const AI_CACHE_EXPIRY = 1000 * 60 * 60 * 3; // 3 hours
  const PACKING_COUNT_STORAGE = 'packing_travelers_count';
  const DOC_CACHE_PREFIX = 'doc_cache_v1_';
  const DOC_METADATA_KEY = 'doc_metadata_v1';

  // Initialize settings immediately from storage to prevent empty fields on quick click
  let geminiApiKey = localStorage.getItem(GEMINI_API_KEY_STORAGE);
  let visualCrossingKey = localStorage.getItem(VC_API_KEY_STORAGE);
  // Hardcoded default Client ID for easier setup
  let googleClientId = '663334830027-1182pamonrcagb0sgvg46jserh7v5jqg.apps.googleusercontent.com';
  let tokenClient = null;
  let accessToken = null;
  let geminiModelName = localStorage.getItem('gemini_model') || 'gemini-3.1-flash-lite-preview';
  let aiConversationHistory = [];
  let currentAiCacheKey = null;
  let currentChecklistType = 'packing'; // 'packing' or 'apps'

  // Pre-populate setting inputs so they are ready before the modal is even opened
  if (settingApiKey) settingApiKey.value = geminiApiKey || '';
  if (settingVcKey) settingVcKey.value = visualCrossingKey || '';
  if (settingModel && geminiModelName) {
    settingModel.innerHTML = `<option value="${geminiModelName}">${geminiModelName}</option>`;
    settingModel.value = geminiModelName;
  }

  // Theme Logic
  let currentTheme = localStorage.getItem('theme') || 'light';
  const htmlEl = document.documentElement;

  function renderMarkdown(text, keyOverride) {
    const keyToUse = keyOverride || currentAiCacheKey;

    // Use marked library
    // breaks: true allows newlines to be rendered as <br>
    const htmlContent = marked.parse(text, { breaks: true });

    const div = document.createElement('div');
    div.innerHTML = htmlContent;

    // Enable checkboxes and restore state
    div.querySelectorAll('input[type="checkbox"]').forEach(checkbox => {
      checkbox.removeAttribute('disabled');
      const li = checkbox.closest('li');
      if (li) {
        li.style.listStyleType = 'none'; // Hide bullet for task items
        const label = li.textContent.trim();
        const storageKey = keyToUse ? (keyToUse + '_chk_' + encodeURIComponent(label)) : null;

        if (storageKey) {
          checkbox.dataset.key = storageKey;
          const stored = localStorage.getItem(storageKey);
          if (stored !== null) {
            checkbox.checked = stored === 'true';
          }
        }
      }
    });

    // Enforce secure links
    div.querySelectorAll('a').forEach(a => {
      a.target = '_blank';
      a.rel = 'noopener noreferrer';
    });

    return div.innerHTML;
  }

  function updateTheme(theme) {
    currentTheme = theme;
    htmlEl.setAttribute('data-theme', theme);
    localStorage.setItem('theme', theme);
    if (mapTileLayer) {
      const isLight = theme === 'light';
      const url = isLight
        ? 'https://{s}.basemaps.cartocdn.com/rastertiles/voyager/{z}/{x}/{y}{r}.png'
        : 'https://{s}.basemaps.cartocdn.com/dark_all/{z}/{x}/{y}{r}.png';
      mapTileLayer.setUrl(url);
    }
  }

  // Initialize theme
  if (currentTheme) updateTheme(currentTheme);

  // Prevent layout shift when switching between short and long pages (scrollbar jump)
  document.documentElement.style.overflowY = 'scroll';

  // Add Home Button to the left of Theme Toggle if not already in HTML
  if (themeToggle && !document.getElementById('home-btn')) {
    // Ensure the button container stays pinned to the right regardless of title length
    if (themeToggle.parentNode) {
      themeToggle.parentNode.style.display = 'flex';
      themeToggle.parentNode.style.alignItems = 'center';
      themeToggle.parentNode.style.marginLeft = 'auto';
      themeToggle.parentNode.style.flexShrink = '0';
    }

    const btn = document.createElement('button');
    btn.id = 'home-btn';
    btn.className = themeToggle.className;
    btn.style.display = 'flex'; // Ensure consistent vertical alignment
    btn.title = 'Home / Start Over';
    btn.innerHTML = `<svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="m3 9 9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"></path><polyline points="9 22 9 12 15 12 15 22"></polyline></svg>`;
    themeToggle.parentNode.insertBefore(btn, themeToggle);
    btn.onclick = resetApp;
  }

  themeToggle?.addEventListener('click', () => {
    const newTheme = currentTheme === 'dark' ? 'light' : 'dark';
    updateTheme(newTheme);
  });

  function showError(msg) {
    errorMsg.textContent = msg || '';
    errorMsg.classList.toggle('visible', !!msg);
  }

  function resetApp(skipHistory = false) {
    const skip = skipHistory === true;
    // Clear data model
    parsed = null;

    // Clear inputs
    csvInput.value = '';
    if (driveUrlInput) driveUrlInput.value = '';
    showError('');

    // Reset UI visibility
    itinerarySection.classList.remove('visible');
    inputSection.style.display = 'block';
    emptyState.style.display = 'block';
    h1.textContent = 'Travel Itinerary';
    shareBtn.style.display = 'none';
    if (btnClearDoc) btnClearDoc.style.display = 'none';

    // Clear rendered content
    listContent.innerHTML = '';
    document.getElementById('itinerary-calendar').innerHTML = '';
    if (map) {
      resetMapData();
    }

    // Refresh the recent docs list on the home page
    renderRecentDocs();

    // Reset input tabs to default
    tabs.forEach(t => t.classList.remove('active'));
    panels.forEach(p => p.classList.remove('active'));
    document.querySelector('.tab[data-tab="drive"]').classList.add('active');
    document.getElementById('panel-drive').classList.add('active');

    // Clear data from URL
    if (!skip) {
      const url = new URL(window.location);
      if (url.searchParams.has('data') || url.searchParams.has('docId')) {
        url.searchParams.delete('data');
        url.searchParams.delete('docId');
        url.searchParams.delete('gid');
        window.history.pushState({}, '', url);
      }
    }
  }

  function parseCSV(text) {
    // Detect delimiter properly by finding the first *unquoted* newline
    let firstRowEnd = text.length;
    let q = false;
    for (let i = 0; i < text.length; i++) {
      if (text[i] === '"') q = !q;
      else if (!q && (text[i] === '\n' || text[i] === '\r')) {
        firstRowEnd = i;
        break;
      }
    }

    const firstLine = text.slice(0, firstRowEnd);
    const tabCount = (firstLine.match(/\t/g) || []).length;
    const commaCount = (firstLine.match(/,/g) || []).length;
    const delimiter = commaCount > tabCount ? ',' : '\t';

    const rows = [];
    let current = '';
    let inQuotes = false;
    const row = [];
    const pushRow = () => {
      rows.push(row.slice());
      row.length = 0;
    };
    const pushCell = () => {
      row.push(current.trim());
      current = '';
    };
    for (let i = 0; i < text.length; i++) {
      const c = text[i];
      const next = text[i + 1];

      if (inQuotes) {
        if (c === '"') {
          if (next === '"') {
            current += '"';
            i++;
          } else {
            inQuotes = false;
          }
        } else {
          current += c;
        }
        continue;
      }

      if (c === '"') {
        inQuotes = true;
        continue;
      }

      if (c === delimiter) {
        pushCell();
        continue;
      }

      if (c === '\n' || c === '\r') {
        pushCell();
        pushRow();
        if (c === '\r' && text[i + 1] === '\n') i++;
        continue;
      }
      current += c;
    }

    if (current || row.length > 0) {
      pushCell();
      pushRow();
    }

    return rows;
  }

  function looksLikeHotelHeader(str) {
    const s = (str || '').trim();
    return /\bhotel\b/i.test(s);
  }

  function parseSheetRows(rows) {
    if (!rows.length) return null;
    const first = rows[0];
    const locations = first.map((c, i) => (i === 0 ? getTextContent(c || 'Time') : getTextContent(c || `Location ${i}`)));
    const dates = rows[1] ? rows[1].map((c, i) => (i === 0 ? '' : getTextContent(c || ''))) : first.map(() => '');
    const dataRows = rows.slice(2);
    let hotelCells = null;
    const data = [];
    dataRows.forEach(row => {
      const firstCell = getTextContent(row[0] || '');
      if (looksLikeHotelHeader(firstCell)) {
        hotelCells = row.slice(0, locations.length);
        while (hotelCells.length < locations.length) hotelCells.push('');
      } else {
        // Keep row if it's a time or potential note (empty time)
        const time = firstCell.trim();
        const cells = row.slice(1);
        data.push({ time, cells });
      }
    });
    return { locations, dates, rows: data, hotelCells };
  }

  function buildItineraryFromRows(rows, fromUrl = false, csvTextForUrl = '') {
    showError('');
    if (!rows.length) {
      showError('No data found. Paste your sheet content (first row = locations, second = dates, first column = times).');
      return;
    }
    parsed = parseSheetRows(rows);
    if (!parsed || !parsed.rows.length) {
      showError('Could not parse structure. Ensure first row = locations, second row = dates, first column = times.');
      return;
    }
    render();
    itinerarySection.classList.add('visible');
    emptyState.style.display = 'none';
    inputSection.style.display = 'none';
    shareBtn.style.display = 'flex';

    // Switch to calendar view by default, or list view on small screens
    if (window.innerWidth < 768) {
      document.querySelector('.view-btn[data-view="list"]').click();
    } else {
      document.querySelector('.view-btn[data-view="calendar"]').click();
    }

    initMap();
    fetchAllWeather();

    if (!fromUrl && csvTextForUrl) {
      updateUrlWithData(csvTextForUrl);
    }

    // Update recent docs metadata with date range if it's a docId-based load
    const urlParams = new URLSearchParams(window.location.search);
    const docId = urlParams.get('docId');
    if (docId) {
      trackRecentDoc(docId);
    }
  }

  function buildItinerary(csvText, fromUrl = false) {
    const rows = parseCSV(csvText);
    buildItineraryFromRows(rows, fromUrl, csvText);
  }

  function render() {
    if (!parsed) return;
    resetMapData();
    hasScrolledToToday = false;
    const { locations, dates, rows, hotelCells } = parsed;

    // Valid day columns: must have both date and location (non-empty after trim).
    const validDayCols = [];
    for (let i = 1; i < locations.length; i++) {
      const d = (dates[i] || '').toString().trim();
      const loc = (locations[i] || '').toString().trim();
      if (d && loc) validDayCols.push(i);
    }
    const lastValidColIndex = validDayCols.length ? validDayCols[validDayCols.length - 1] : -1;

    // Hotel stays: backfill empty cells with previous hotel; "through" = date before next hotel change.
    let effectiveHotel = [];
    let throughDateForCol = [];
    if (hotelCells && locations.length) {
      const pad = Array(locations.length).fill('');
      const cells = hotelCells.length >= locations.length ? hotelCells : [...hotelCells, ...pad.slice(hotelCells.length)];
      effectiveHotel = cells.map((c, i) => (i === 0 ? '' : (c || '').trim()));
      for (let i = 1; i < effectiveHotel.length; i++) {
        if (!effectiveHotel[i]) {
          const prevLoc = (locations[i - 1] || '').toString().trim();
          const currLoc = (locations[i] || '').toString().trim();
          if (currLoc === prevLoc) {
            effectiveHotel[i] = effectiveHotel[i - 1] || '';
          }
        }
      }
      throughDateForCol = effectiveHotel.map((h, i) => {
        if (!h || i === 0) return null;
        const nextDiff = effectiveHotel.findIndex((h2, j) => j > i && h2 !== h);
        if (nextDiff === -1) return null;
        const throughDate = (dates[nextDiff - 1] || '').toString().trim();
        return throughDate || null;
      });
    }

    // List view: only valid day columns. Group by date, then by time. Hotel from effectiveHotel (through next hotel).
    const byDate = {};
    const locationsByDateKey = {};
    validDayCols.forEach(colIndex => {
      const dateKey = (dates[colIndex] || '').toString().trim() || (locations[colIndex] || '').toString().trim();
      const loc = (locations[colIndex] || '').toString().trim();
      if (!byDate[dateKey]) byDate[dateKey] = [];
      if (!locationsByDateKey[dateKey]) locationsByDateKey[dateKey] = new Set();
      if (loc) locationsByDateKey[dateKey].add(loc);
      rows.forEach(({ time, cells }) => {
        const content = (cells[colIndex - 1] || '').trim();
        if (content) byDate[dateKey].push({ time, location: locations[colIndex], content });
      });
    });
    listContent.innerHTML = '';
    // Group valid columns by location, then split each location into continuous date ranges (one card per run).
    const locationToCols = {};
    validDayCols.forEach(colIndex => {
      const loc = (locations[colIndex] || '').toString().trim();
      if (!loc) return;
      if (!locationToCols[loc]) locationToCols[loc] = [];
      locationToCols[loc].push(colIndex);
    });
    const ONE_DAY_MS = 86400000;
    const locationRuns = [];
    Object.keys(locationToCols).forEach(locName => {
      const colIndices = locationToCols[locName];
      const entries = colIndices.map(i => ({
        colIndex: i,
        dateKey: (dates[i] || '').toString().trim(),
        ts: parseDateForSort((dates[i] || '').toString().trim())
      })).filter(e => e.ts !== Infinity);
      if (!entries.length) return;
      entries.sort((a, b) => a.ts - b.ts);
      const runs = [];
      let run = [entries[0]];
      for (let i = 1; i < entries.length; i++) {
        const gap = entries[i].ts - entries[i - 1].ts;
        if (gap > ONE_DAY_MS * 1.5) {
          runs.push(run);
          run = [entries[i]];
        } else {
          run.push(entries[i]);
        }
      }
      if (run.length) runs.push(run);
      runs.forEach(runEntries => {
        const dateKeys = runEntries.map(e => e.dateKey);
        const runColIndices = runEntries.map(e => e.colIndex);
        const hotelsInRun = [];
        runColIndices.forEach(colIndex => {
          const hotel = effectiveHotel[colIndex] || '';
          const through = throughDateForCol[colIndex] || null;
          if (!hotel) return;
          const last = hotelsInRun[hotelsInRun.length - 1];
          if (last && last.hotel === hotel) {
            if (through) last.through = through;
          } else {
            hotelsInRun.push({ hotel, through });
          }
        });
        locationRuns.push({
          locName,
          dateKeys,
          colIndices: runColIndices,
          hotels: hotelsInRun
        });
      });
    });
    locationRuns.sort((a, b) => parseDateForSort(a.dateKeys[0]) - parseDateForSort(b.dateKeys[0]));

    const now = new Date();
    now.setHours(0, 0, 0, 0);
    const todayYmd = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;

    searchInput.value = ''; // Reset search

    locationRuns.forEach(run => {
      // Create a detailed plan for the entire location run to provide context to the AI
      const fullLocationPlanParts = [];
      run.dateKeys.forEach(dKey => {
        const allItems = (byDate[dKey] || []).filter(e => e.location === run.locName);

        let lastTimeIndex = -1;
        for (let i = allItems.length - 1; i >= 0; i--) {
          if (allItems[i].time) { lastTimeIndex = i; break; }
        }

        let eventsList = allItems.slice(0, lastTimeIndex + 1).filter(e => e.time);

        let sortTracker = -1;
        const events = eventsList.map(e => {
          if (e.time) sortTracker = parseTimeForSort(e.time);
          return { ...e, _sort: sortTracker };
        }).sort((a, b) => a._sort - b._sort);

        if (events.length > 0) {
          const dayPlan = events.map(e => `  ${e.time} - ${e.content}`).join('\n');
          fullLocationPlanParts.push(`On ${dKey}:\n${dayPlan}`);
        }
      });
      const fullLocationPlan = fullLocationPlanParts.join('\n\n');
      const hotelInfo = run.hotels.map(h => h.hotel + (h.through ? ` (through ${h.through})` : '')).join(', ');

      const dateRangeStr = run.dateKeys.length === 1 ? run.dateKeys[0] : `${run.dateKeys[0]} – ${run.dateKeys[run.dateKeys.length - 1]}`;
      const cardEl = document.createElement('div');
      cardEl.className = 'location-card';

      let isCurrent = false;
      if (run.dateKeys.some(k => parseSmartDate(k) === todayYmd)) {
        isCurrent = true;
      }

      const startTs = parseDateForSort(run.dateKeys[0]);
      const endTs = parseDateForSort(run.dateKeys[run.dateKeys.length - 1]);
      if (startTs !== Infinity && endTs !== Infinity) {
        const checkDate = new Date(now);
        const startYear = new Date(startTs).getFullYear();
        // If parsed year is old (e.g. 2000 from default), match year to ignore it
        if (startYear < 2020) checkDate.setFullYear(startYear);
        if (checkDate.getTime() >= startTs && checkDate.getTime() <= endTs) isCurrent = true;
      }

      // Store metadata for weather fetching
      cardEl.dataset.location = run.locName;
      cardEl.dataset.start = run.dateKeys[0];
      cardEl.dataset.end = run.dateKeys[run.dateKeys.length - 1];

      cardEl.className = isCurrent ? 'location-card expanded' : 'location-card collapsed';

      const summary = document.createElement('div');
      summary.className = 'location-card-summary';
      summary.innerHTML = `<h3 class="location-card-name">${escapeHtml(run.locName)}</h3>`;

      const meta = document.createElement('div');
      meta.className = 'location-card-meta';

      const dateDiv = document.createElement('div');
      dateDiv.innerHTML = `<span class="date-range">${escapeHtml(dateRangeStr)}</span>`;
      meta.appendChild(dateDiv);

      run.hotels.forEach(({ hotel, through }) => {
        const hotelEl = document.createElement('div');
        hotelEl.className = 'hotel';
        hotelEl.textContent = through ? `🏨 ${hotel} (through ${through})` : `🏨 ${hotel}`;
        hotelEl.title = 'Open in Google Maps';
        hotelEl.onclick = (e) => {
          e.stopPropagation();
          const query = `${hotel} ${run.locName}`;
          window.open(`https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(query)}`, '_blank');
        };
        meta.appendChild(hotelEl);
      });

      summary.appendChild(meta);
      summary.addEventListener('click', () => cardEl.classList.toggle('collapsed'));
      cardEl.appendChild(summary);

      const daysContainer = document.createElement('div');
      daysContainer.className = 'location-card-days';
      const daysInner = document.createElement('div');
      daysInner.className = 'location-card-days-inner';
      run.dateKeys.forEach((dateLabel, i) => {
        const colIndex = run.colIndices[i];
        let hotelForDay = effectiveHotel[colIndex];

        // No hotel on the very last day of the trip unless explicitly specified 
        // (which means you're staying that night)
        const isLastDayOfTrip = colIndex === lastValidColIndex;
        const hasExplicitHotel = hotelCells && (hotelCells[colIndex] || '').trim();
        if (isLastDayOfTrip && !hasExplicitHotel) {
          hotelForDay = null;
        }

        const allItems = (byDate[dateLabel] || []).filter(e => e.location === run.locName);

        // Notes are only rows below the last timestamped row
        let lastTimeIndex = -1;
        for (let i = allItems.length - 1; i >= 0; i--) {
          if (allItems[i].time) {
            lastTimeIndex = i;
            break;
          }
        }

        let eventsList = allItems.slice(0, lastTimeIndex + 1).filter(e => e.time);
        const notes = allItems.slice(lastTimeIndex + 1);

        let sortTracker = -1;
        const events = eventsList.map(e => {
          if (e.time) sortTracker = parseTimeForSort(e.time);
          return { ...e, _sort: sortTracker };
        });
        events.sort((a, b) => a._sort - b._sort);

        const dailyPlan = events.map(e => {
          let planItem = '';
          if (e.time) planItem += `${e.time} - `;
          planItem += e.content;
          return planItem;
        }).join('\n');

        const dayDiv = document.createElement('div');
        dayDiv.className = 'day-group';
        if (parseSmartDate(dateLabel) === todayYmd) {
          dayDiv.classList.add('today-active');
        }

        if (hotelForDay) {
          const hotelDiv = document.createElement('div');
          hotelDiv.className = 'day-group-hotel';
          hotelDiv.textContent = `🏨 ${hotelForDay}`;
          dayDiv.appendChild(hotelDiv);
        }

        const titleDiv = document.createElement('div');
        titleDiv.className = 'day-group-title';
        titleDiv.dataset.date = dateLabel;
        titleDiv.innerHTML = `<span>${escapeHtml(dateLabel)}</span>`;

        if (notes.length > 0) {
          const btn = document.createElement('button');
          btn.className = 'day-notes-btn';
          btn.innerHTML = `<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><line x1="16" y1="13" x2="8" y2="13"></line><line x1="16" y1="17" x2="8" y2="17"></line><polyline points="10 9 9 9 8 9"></polyline></svg>`;
          btn.title = 'Show notes';

          const notesDiv = document.createElement('div');
          notesDiv.className = 'day-notes';
          notes.forEach(n => {
            const p = document.createElement('div');
            p.className = 'day-note-item';
            p.innerHTML = linkifyAndSanitize(n.content);
            notesDiv.appendChild(p);
          });

          btn.onclick = () => {
            notesDiv.classList.toggle('visible');
            btn.classList.toggle('active');
          };

          titleDiv.appendChild(btn);
          dayDiv.appendChild(titleDiv);
          dayDiv.appendChild(notesDiv);
        } else {
          dayDiv.appendChild(titleDiv);
        }

        const suggestBtn = document.createElement('button');
        suggestBtn.className = 'btn-ai-suggest';
        suggestBtn.title = 'Get AI suggestions for this day';
        suggestBtn.innerHTML = '✨';
        suggestBtn.onclick = (e) => {
          e.stopPropagation();
          if (!geminiApiKey) {
            alert('Please set your Google AI API Key in the "AI Settings" tab first.');
            return;
          }
          const weather = titleDiv.dataset.weatherSummary;
          getAiSuggestions(run.locName, dateLabel, dailyPlan, fullLocationPlan, hotelInfo, weather);
        };
        titleDiv.appendChild(suggestBtn);

        events.forEach(({ time, content }) => {
          const eventCard = document.createElement('div');
          eventCard.className = 'event-card';
          eventCard.innerHTML = `
            <span class="event-time">${escapeHtml(time)}</span>
            <span class="event-content">${linkifyAndSanitize(content)}</span>
          `;

          dayDiv.appendChild(eventCard);
        });
        daysContainer.appendChild(dayDiv);
        daysInner.appendChild(dayDiv);
      });
      daysContainer.appendChild(daysInner);
      cardEl.appendChild(daysContainer);
      listContent.appendChild(cardEl);
    });
    renderCalendar();
  }

  function renderCalendar() {
    const cal = document.getElementById('itinerary-calendar');
    cal.innerHTML = '';
    if (!parsed) return;
    const { locations, dates, rows } = parsed;

    // Backfill hotel logic for calendar (similar to render/list view)
    let effectiveHotel = [];
    if (parsed.hotelCells && parsed.locations.length) {
      const pad = Array(parsed.locations.length).fill('');
      const cells = parsed.hotelCells.length >= parsed.locations.length ? parsed.hotelCells : [...parsed.hotelCells, ...pad.slice(parsed.hotelCells.length)];
      effectiveHotel = cells.map((c, i) => (i === 0 ? '' : (c || '').trim()));
      for (let i = 1; i < effectiveHotel.length; i++) {
        if (!effectiveHotel[i]) effectiveHotel[i] = effectiveHotel[i - 1] || '';
      }
      // Clear hotel if last col has no explicit hotel? 
      // (Matching list logic: "no hotel unless the last valid column has an explicit hotel")
      // For calendar, we map daily. If effectiveHotel[i] exists, that day has a hotel.
      // We'll handle the "last day" logic inside the mapping loop if needed, 
      // but generally if effectiveHotel has a value, it applies to that night.
    }

    const today = new Date();
    const todayYmd = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-${String(today.getDate()).padStart(2, '0')}`;

    const eventsByYmd = {};
    let minTs = Infinity;
    let maxTs = -Infinity;

    // 1. Map data from columns to dates
    for (let i = 1; i < locations.length; i++) {
      const rawDate = (dates[i] || '').trim();
      const loc = (locations[i] || '').trim();
      if (!rawDate || !loc) continue;

      const ymd = parseSmartDate(rawDate); // "YYYY-MM-DD"
      if (!ymd) continue;

      if (!eventsByYmd[ymd]) eventsByYmd[ymd] = { locs: new Set(), events: [], hotel: null };
      eventsByYmd[ymd].locs.add(loc);

      if (effectiveHotel[i]) {
        eventsByYmd[ymd].hotel = effectiveHotel[i];
      }

      rows.forEach(r => {
        const content = (r.cells[i - 1] || '').trim();
        if (content && r.time) {
          eventsByYmd[ymd].events.push({ time: r.time, content });
        }
      });

      const parts = ymd.split('-').map(Number);
      const d = new Date(parts[0], parts[1] - 1, parts[2]);
      const ts = d.getTime();
      if (ts < minTs) minTs = ts;
      if (ts > maxTs) maxTs = ts;
    }

    if (minTs === Infinity) return;

    // 2. Generate Continuous Grid
    // Start from Sunday of the first week
    const startDate = new Date(minTs);
    startDate.setDate(startDate.getDate() - startDate.getDay());

    // We'll iterate until we pass maxTs and finish the week
    const wrapper = document.createElement('div');
    wrapper.className = 'calendar-month'; // Use existing class for container styling

    const monthTitle = document.createElement('div');
    monthTitle.className = 'calendar-month-title';
    const startM = new Date(minTs).toLocaleDateString('default', { month: 'long', year: 'numeric' });
    const endM = new Date(maxTs).toLocaleDateString('default', { month: 'long', year: 'numeric' });
    monthTitle.textContent = startM === endM ? startM : `${startM} – ${endM}`;
    wrapper.appendChild(monthTitle);

    const grid = document.createElement('div');
    grid.className = 'calendar-grid';

    // Headers
    ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'].forEach(d => {
      const h = document.createElement('div');
      h.className = 'calendar-header';
      h.textContent = d;
      grid.appendChild(h);
    });

    let curr = new Date(startDate);
    let isFirstVisibleDay = true;
    let loops = 0;

    while (loops < 1000) { // Safety break
      // Build a week
      const week = [];
      for (let i = 0; i < 7; i++) {
        const ymd = `${curr.getFullYear()}-${String(curr.getMonth() + 1).padStart(2, '0')}-${String(curr.getDate()).padStart(2, '0')}`;
        week.push({
          date: new Date(curr),
          ymd: ymd,
          dayNum: curr.getDate(),
          data: eventsByYmd[ymd]
        });
        curr.setDate(curr.getDate() + 1);
      }

      // Check if we are done (past maxTs and empty week)
      const weekStartTs = week[0].date.getTime();
      // data check: locs, events, hotel
      const hasData = week.some(d => d.data && (d.data.locs.size > 0 || d.data.events.length > 0 || d.data.hotel));

      if (weekStartTs > maxTs && !hasData) {
        break;
      }

      if (hasData) {
        week.forEach(d => {
          const cell = document.createElement('div');
          cell.className = 'calendar-day';
          if (d.ymd === todayYmd) cell.classList.add('today');
          cell.dataset.date = d.ymd;

          const num = document.createElement('div');
          num.className = 'calendar-day-number';

          // Show Month Name if 1st of month OR first visible day
          if (d.dayNum === 1 || isFirstVisibleDay) {
            num.textContent = d.date.toLocaleDateString('default', { month: 'short', day: 'numeric' });
            num.style.fontWeight = '700';
            num.style.color = 'var(--accent)';
            isFirstVisibleDay = false;
          } else {
            num.textContent = d.dayNum;
          }
          cell.appendChild(num);

          if (d.data) {
            // Locations
            d.data.locs.forEach(l => {
              const pill = document.createElement('div');
              pill.className = 'calendar-event location-pill';
              pill.textContent = l;
              cell.appendChild(pill);
            });

            // Weather placeholder
            const wDiv = document.createElement('div');
            wDiv.className = 'calendar-weather';
            cell.appendChild(wDiv);

            // Hotel
            if (d.data.hotel) {
              const hotelPill = document.createElement('div');
              hotelPill.className = 'calendar-event hotel-pill';
              hotelPill.textContent = `🏨 ${d.data.hotel}`;
              hotelPill.style.cursor = 'pointer';
              hotelPill.title = 'Open in Google Maps';
              hotelPill.onclick = (e) => {
                e.stopPropagation();
                let query = d.data.hotel;
                if (d.data.locs && d.data.locs.size > 0) {
                  query += ` ${Array.from(d.data.locs)[0]}`;
                }
                window.open(`https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(query)}`, '_blank');
              };
              cell.appendChild(hotelPill);
            }

            // Events
            d.data.events.forEach(ev => {
              const pill = document.createElement('div');
              pill.className = 'calendar-event';
              const timePart = ev.time ? `${ev.time} ` : '';
              pill.innerHTML = linkifyAndSanitize(timePart + ev.content);
              pill.title = getTextContent(timePart + ev.content);
              cell.appendChild(pill);
            });
          }
          grid.appendChild(cell);
        });
      }
      loops++;
    }

    wrapper.appendChild(grid);
    cal.appendChild(wrapper);
  }

  function initMap() {
    if (map) return;
    // Initialize map
    map = L.map('itinerary-map').setView([20, 0], 2);

    const isLight = currentTheme === 'light';
    const url = isLight ? 'https://{s}.basemaps.cartocdn.com/rastertiles/voyager/{z}/{x}/{y}{r}.png' : 'https://{s}.basemaps.cartocdn.com/dark_all/{z}/{x}/{y}{r}.png';

    mapTileLayer = L.tileLayer(url, {
      attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors &copy; <a href="https://carto.com/attributions">CARTO</a>',
      maxZoom: 18
    }).addTo(map);
    mapLayerGroup = L.layerGroup().addTo(map);
  }

  function resetMapData() {
    if (!map) return;
    mapLayerGroup.clearLayers();
    mapPoints.length = 0;
    for (const key in mapMarkers) delete mapMarkers[key];
    polyline = L.polyline([], { color: 'var(--accent)', weight: 3 }).addTo(mapLayerGroup);
  }

  function addMapLocation(name, lat, lon) {
    if (!map) return;

    // Add marker if this specific location hasn't been marked yet
    if (!mapMarkers[name]) {
      const marker = L.marker([lat, lon]).addTo(mapLayerGroup);
      marker.bindPopup(`<b>${escapeHtml(name)}</b>`);
      mapMarkers[name] = marker;
    }

    // Add to polyline path (allows duplicate locations to show route)
    mapPoints.push([lat, lon]);
    if (polyline) {
      polyline.setLatLngs(mapPoints);
    }

    // Fit bounds if we have points
    if (mapPoints.length > 0) {
      const bounds = L.latLngBounds(mapPoints);
      // Add some padding
      map.fitBounds(bounds, { padding: [50, 50], maxZoom: 12 });
    }
  }

  async function updateUrlWithData(csvText) {
    if (!window.CompressionStream) {
      console.warn('CompressionStream API not supported. Cannot update URL with data.');
      return;
    }
    try {
      const stream = new Blob([new TextEncoder().encode(csvText)]).stream().pipeThrough(new CompressionStream('deflate'));
      const compressed = await new Response(stream).arrayBuffer();

      const bytes = new Uint8Array(compressed);
      let binaryString = '';
      for (let i = 0; i < bytes.length; i += 8192) {
        binaryString += String.fromCharCode.apply(null, bytes.subarray(i, i + 8192));
      }
      const base64 = btoa(binaryString);
      const urlSafeBase64 = base64.replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');

      const url = new URL(window.location);
      url.searchParams.delete('docId');
      url.searchParams.delete('gid');
      url.searchParams.set('data', urlSafeBase64);
      window.history.pushState({}, '', url);
    } catch (e) {
      console.error('Failed to compress and update URL:', e);
    }
  }

  function updateUrlWithDocId(fileId, gid = '0') {
    const url = new URL(window.location);
    url.searchParams.delete('data');
    url.searchParams.set('docId', fileId);
    if (gid !== '0') url.searchParams.set('gid', gid);
    else url.searchParams.delete('gid');
    window.history.pushState({}, '', url);
  }

  async function loadFromUrl() {
    const urlParams = new URLSearchParams(window.location.search);
    const data = urlParams.get('data');
    const docId = urlParams.get('docId');
    const gid = urlParams.get('gid') || '0';

    if (data && window.DecompressionStream) {
      try {
        let base64 = data.replace(/-/g, '+').replace(/_/g, '/');
        while (base64.length % 4) {
          base64 += '=';
        }
        const binaryString = atob(base64);
        const len = binaryString.length;
        const bytes = new Uint8Array(len);
        for (let i = 0; i < len; i++) {
          bytes[i] = binaryString.charCodeAt(i);
        }

        let decompressed;
        try {
          const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('deflate'));
          decompressed = await new Response(stream).text();
        } catch (e) {
          // Fallback for legacy gzip-compressed URLs
          const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('gzip'));
          decompressed = await new Response(stream).text();
        }
        buildItinerary(decompressed, true);
      } catch (e) {
        console.error('Failed to decompress data from URL:', e);
        showError('Could not load itinerary from URL. The data might be corrupted.');
      }
    } else if (docId) {
      const metadata = JSON.parse(localStorage.getItem(DOC_METADATA_KEY) || '{}');
      const existingName = metadata[docId]?.name;

      // Attempt to load from cache immediately for better UX
      const cached = localStorage.getItem(DOC_CACHE_PREFIX + docId + '_' + gid);
      if (cached && !parsed) {
        trackRecentDoc(docId, existingName, gid); // Update access timestamp

        if (existingName) h1.textContent = existingName;

        if (cached.startsWith('[[')) {
          // Cached JSON grid from Sheets API
          buildItineraryFromRows(JSON.parse(cached), true);
        } else {
          // Legacy CSV cache
          buildItinerary(cached, true);
        }
        if (btnClearDoc) btnClearDoc.style.display = 'block';
      }

      if (!googleClientId) {
        if (!cached) showError('A Google Sheet was linked, but authentication is required. Please sign in.');
        return;
      }

      // If we don't have an access token yet, we stop here. 
      // The auth callbacks will trigger loadFromUrl() again once a token is acquired.
      if (!accessToken) return;

      fetchSheetData(docId, existingName || 'Linked Spreadsheet', gid, true);
    } else {
      resetApp(true);
    }
  }

  function parseDateForSort(label) {
    const s = (label || '').toString().trim();
    if (!s) return Infinity;
    const d = new Date(s);
    if (!isNaN(d.getTime())) return d.getTime();
    const match = s.match(/^(\w{3,})\s+(\d{1,2})$/i);
    if (match) {
      const d2 = new Date(match[1] + ' ' + match[2] + ', 2000');
      if (!isNaN(d2.getTime())) return d2.getTime();
    }
    return Infinity;
  }

  function parseTimeForSort(timeStr) {
    const s = (timeStr || '').toString().trim();
    const m = s.match(/^(\d{1,2})(?::(\d{2}))?\s*(am|pm)?$/i);
    if (!m) return 0;
    let h = parseInt(m[1], 10);
    const min = m[2] ? parseInt(m[2], 10) : 0;
    const pm = (m[3] || '').toLowerCase() === 'pm';
    const am = (m[3] || '').toLowerCase() === 'am';
    if (pm && h < 12) h += 12;
    if (am && h === 12) h = 0;
    if (!pm && !am && h <= 12) h = h % 12;
    return h * 60 + min;
  }

  function parseSmartDate(str) {
    const s = (str || '').toString().trim();
    if (!s) return null;
    const now = new Date();
    let d = new Date(s);
    // Handle Month Day format without year (defaults to 2001 or current)
    const match = s.match(/^([a-zA-Z]{3,})\s+(\d{1,2})$/i);
    if (match) {
      d = new Date(`${match[1]} ${match[2]}, ${now.getFullYear()}`);
    } else if (!isNaN(d.getTime()) && d.getFullYear() < 2020) {
      d.setFullYear(now.getFullYear());
    }
    if (isNaN(d.getTime())) return null;
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
  }

  async function getCoordinates(location) {
    try {
      // Use Open-Meteo Geocoding API to avoid Nominatim's strict Referer/User-Agent policies
      const url = `https://geocoding-api.open-meteo.com/v1/search?name=${encodeURIComponent(location)}&count=1&language=en&format=json`;
      const res = await fetch(url);
      if (!res.ok) return null;
      const data = await res.json();
      if (data && data.results && data.results.length > 0) {
        return { lat: data.results[0].latitude, lon: data.results[0].longitude };
      }
    } catch (e) {
      console.error('Geocoding error:', e);
    }
    return null;
  }

  function getWmoInfo(code) {
    const map = {
      0: { icon: 'clear-day', label: 'Clear sky' },
      1: { icon: 'partly-cloudy-day', label: 'Mainly clear' },
      2: { icon: 'partly-cloudy-day', label: 'Partly cloudy' },
      3: { icon: 'cloudy', label: 'Overcast' },
      45: { icon: 'fog', label: 'Fog' },
      48: { icon: 'fog', label: 'Depositing rime fog' },
      51: { icon: 'rain', label: 'Light drizzle' },
      53: { icon: 'rain', label: 'Moderate drizzle' },
      55: { icon: 'rain', label: 'Dense drizzle' },
      56: { icon: 'sleet', label: 'Light freezing drizzle' },
      57: { icon: 'sleet', label: 'Dense freezing drizzle' },
      61: { icon: 'rain', label: 'Slight rain' },
      63: { icon: 'rain', label: 'Moderate rain' },
      65: { icon: 'rain', label: 'Heavy rain' },
      66: { icon: 'sleet', label: 'Light freezing rain' },
      67: { icon: 'sleet', label: 'Heavy freezing rain' },
      71: { icon: 'snow', label: 'Slight snow fall' },
      73: { icon: 'snow', label: 'Moderate snow fall' },
      75: { icon: 'snow', label: 'Heavy snow fall' },
      77: { icon: 'snow', label: 'Snow grains' },
      80: { icon: 'showers-day', label: 'Slight rain showers' },
      81: { icon: 'showers-day', label: 'Moderate rain showers' },
      82: { icon: 'showers-day', label: 'Violent rain showers' },
      85: { icon: 'snow-showers-day', label: 'Slight snow showers' },
      86: { icon: 'snow-showers-day', label: 'Heavy snow showers' },
      95: { icon: 'thunder-rain', label: 'Thunderstorm' },
      96: { icon: 'thunder-showers-day', label: 'Thunderstorm with slight hail' },
      99: { icon: 'thunder-showers-day', label: 'Thunderstorm with heavy hail' }
    };
    return map[code] || { icon: 'partly-cloudy-day', label: 'Unknown' };
  }

  async function fetchCardWeatherData(card) {
    const loc = card.dataset.location;
    const start = parseSmartDate(card.dataset.start);
    const end = parseSmartDate(card.dataset.end);
    if (!loc || !start) return;

    const source = visualCrossingKey ? 'vc' : 'om';
    const cacheKey = `${WEATHER_CACHE_PREFIX}${source}_${encodeURIComponent(loc)}_${start}_${end || ''}`;

    try {
      const cachedJson = localStorage.getItem(cacheKey);
      if (cachedJson) {
        const cached = JSON.parse(cachedJson);
        if (Date.now() - cached.timestamp < WEATHER_CACHE_EXPIRY) {
          if (cached.coords) {
            addMapLocation(loc, cached.coords.lat, cached.coords.lon);
          }
          return cached.data;
        }
      }
    } catch (e) {
      console.warn('Cache read error', e);
    }

    // Try Visual Crossing first if key exists
    if (visualCrossingKey) {
      try {
        const endDateStr = end || start;
        const url = `https://weather.visualcrossing.com/VisualCrossingWebServices/rest/services/timeline/${encodeURIComponent(loc)}/${start}/${endDateStr}?unitGroup=us&key=${visualCrossingKey}&include=days&contentType=json`;

        const res = await fetch(url);
        if (!res.ok) throw new Error(`Visual Crossing error: ${res.status}`);

        const data = await res.json();
        if (data) {
          let coords = null;
          if (data.latitude && data.longitude) {
            coords = { lat: data.latitude, lon: data.longitude };
            addMapLocation(loc, data.latitude, data.longitude);
          }
          const timezone = data.timezone;
          const days = (data.days || []).map(d => ({
            datetime: d.datetime,
            tempmax: d.tempmax,
            tempmin: d.tempmin,
            icon: d.icon,
            label: d.conditions,
            isHistorical: false // VC handles mixed forecast/history automatically
          }));
          const result = { days, timezone };
          try { localStorage.setItem(cacheKey, JSON.stringify({ timestamp: Date.now(), data: result, coords })); } catch (e) { }
          return result;
        }
      } catch (e) {
        console.warn('Visual Crossing fetch failed, falling back to Open-Meteo.', e);
      }
    }

    const coords = await getCoordinates(loc);
    if (!coords) return;

    // Update map immediately with coordinates
    addMapLocation(loc, coords.lat, coords.lon);

    // Generate dates to fetch
    const daysToFetch = [];
    const startParts = start.split('-').map(Number);
    const startDate = new Date(startParts[0], startParts[1] - 1, startParts[2]);
    const endDate = end ? new Date(end.split('-')[0], end.split('-')[1] - 1, end.split('-')[2]) : new Date(startDate);

    let curr = new Date(startDate);
    while (curr <= endDate) {
      const y = curr.getFullYear();
      const m = String(curr.getMonth() + 1).padStart(2, '0');
      const d = String(curr.getDate()).padStart(2, '0');
      daysToFetch.push(`${y}-${m}-${d}`);
      curr.setDate(curr.getDate() + 1);
    }

    const now = new Date();
    now.setHours(0, 0, 0, 0);

    const days = [];
    let timezone = null;

    // Fetch per day
    for (const dateStr of daysToFetch) {
      const p = dateStr.split('-').map(Number);
      const dObj = new Date(p[0], p[1] - 1, p[2]);
      const diffDays = (dObj - now) / (1000 * 60 * 60 * 24);
      let isDayHistorical = diffDays > 14;
      let dayData = null;

      // Try Forecast
      if (!isDayHistorical) {
        try {
          const url = `https://api.open-meteo.com/v1/forecast?latitude=${coords.lat}&longitude=${coords.lon}&daily=weathercode,temperature_2m_max,temperature_2m_min&timezone=auto&temperature_unit=fahrenheit&start_date=${dateStr}&end_date=${dateStr}`;
          const res = await fetch(url);
          if (res.ok) {
            const json = await res.json();
            if (json.timezone) timezone = json.timezone; // Capture timezone
            if (json.daily && json.daily.time && json.daily.time.length > 0) {
              dayData = {
                datetime: dateStr,
                tempmax: json.daily.temperature_2m_max[0],
                tempmin: json.daily.temperature_2m_min[0],
                ...getWmoInfo(json.daily.weathercode[0]),
                isHistorical: false
              };
            }
          }
        } catch (e) {
          // Ignore, fallback to historical
        }
      }

      // Fallback to Historical
      if (!dayData) {
        try {
          const prevDate = new Date(p[0] - 1, p[1] - 1, p[2]);
          const py = prevDate.getFullYear();
          const pm = String(prevDate.getMonth() + 1).padStart(2, '0');
          const pd = String(prevDate.getDate()).padStart(2, '0');
          const prevDateStr = `${py}-${pm}-${pd}`;

          const url = `https://archive-api.open-meteo.com/v1/archive?latitude=${coords.lat}&longitude=${coords.lon}&daily=weathercode,temperature_2m_max,temperature_2m_min&timezone=auto&temperature_unit=fahrenheit&start_date=${prevDateStr}&end_date=${prevDateStr}`;
          const res = await fetch(url);
          if (res.ok) {
            const json = await res.json();
            if (json.timezone) timezone = json.timezone;
            if (json.daily && json.daily.time && json.daily.time.length > 0) {
              dayData = {
                datetime: dateStr,
                tempmax: json.daily.temperature_2m_max[0],
                tempmin: json.daily.temperature_2m_min[0],
                ...getWmoInfo(json.daily.weathercode[0]),
                isHistorical: true
              };
            }
          }
        } catch (e) {
          console.error('Historical fetch failed', e);
        }
      }

      if (dayData) {
        days.push(dayData);
      }
    }

    const result = { days, timezone };
    try { localStorage.setItem(cacheKey, JSON.stringify({ timestamp: Date.now(), data: result, coords })); } catch (e) { }
    return result;
  }

  function renderCardWeather(card, { days, timezone }) {
    try {
      if (days && days.length > 0) {
        const maxHigh = Math.round(Math.max(...days.map(d => d.tempmax)));
        const minHigh = Math.round(Math.min(...days.map(d => d.tempmax)));

        const uniqueIcons = [...new Set(days.map(d => d.icon))];
        const iconsHtml = uniqueIcons.map(icon =>
          `<img src="https://raw.githubusercontent.com/visualcrossing/WeatherIcons/refs/heads/main/PNG/2nd%20Set%20-%20Color/${icon}.png" alt="${icon}" title="${icon}" class="weather-icon" style="width: 1.4em; height: 1.4em; vertical-align: middle;">`
        ).join('');

        const dateRangeEl = card.querySelector('.date-range');
        if (dateRangeEl) {
          let sumEl = card.querySelector('.weather-summary-inline');
          if (!sumEl) {
            sumEl = document.createElement('span');
            sumEl.className = 'weather-summary-inline';
            sumEl.style.marginLeft = '0.75rem';
            sumEl.style.display = 'inline-flex';
            sumEl.style.alignItems = 'center';
            sumEl.style.gap = '0.35rem';
            sumEl.style.color = 'var(--text)';
            dateRangeEl.parentNode.insertBefore(sumEl, dateRangeEl.nextSibling);
          }
          const tempStr = minHigh === maxHigh ? `${maxHigh}°F` : `${minHigh}°F–${maxHigh}°F`;
          sumEl.innerHTML = `${iconsHtml} <b>${tempStr}</b>`;

          if (days.some(d => d.isHistorical)) {
            const meta = card.querySelector('.location-card-meta');
            if (meta && !meta.querySelector('.historical-note')) {
              const note = document.createElement('div');
              note.className = 'historical-note';
              note.textContent = 'Includes historical data (last year)';
              note.style.fontSize = '0.85rem';
              note.style.fontStyle = 'italic';
              note.style.marginTop = '0.25rem';
              note.style.opacity = '0.8';
              meta.appendChild(note);
            }
          }
        }
      }

      // Setup Local Time Clock if timezone found
      if (timezone) {
        const meta = card.querySelector('.location-card-meta');
        if (meta && !meta.querySelector('.local-time-display')) {
          const clockEl = document.createElement('div');
          clockEl.className = 'local-time-display';
          clockEl.style.marginTop = '0.5rem';
          clockEl.style.color = 'var(--text-muted)';
          clockEl.style.fontSize = '0.9rem';
          clockEl.innerHTML = `🕒 Local time: <span class="clock-val">--:--</span>`;
          const firstHotel = meta.querySelector('.hotel');
          if (firstHotel) {
            meta.insertBefore(clockEl, firstHotel);
          } else {
            meta.appendChild(clockEl);
          }

          // Store timezone for the interval updater
          clockEl.dataset.tz = timezone;
          updateSingleClock(clockEl);
        }
      }

      const groups = card.querySelectorAll('.day-group-title');
      groups.forEach(title => {
        // data-date is usually "Mon, Jan 1"
        // parseSmartDate converts it to YYYY-MM-DD
        const smartDate = parseSmartDate(title.dataset.date);
        const tDate = smartDate;
        const dayData = days.find(d => d.datetime === tDate);
        if (dayData) {
          title.dataset.weatherSummary = `${Math.round(dayData.tempmax)}°F / ${Math.round(dayData.tempmin)}°F, ${dayData.label}`;
          let badge = title.querySelector('.weather-badge');
          if (!badge) {
            badge = document.createElement('span');
            badge.className = 'weather-badge';
            const firstButton = title.querySelector('.day-notes-btn, .btn-ai-suggest');
            title.insertBefore(badge, firstButton || null);
          }
          const iconUrl = `https://raw.githubusercontent.com/visualcrossing/WeatherIcons/refs/heads/main/PNG/2nd%20Set%20-%20Color/${dayData.icon}.png`;
          badge.innerHTML = `<img src="${iconUrl}" alt="${dayData.icon}" class="weather-icon" style="width: 1.4em; height: 1.4em; margin-right: 0.3em;"> ${Math.round(dayData.tempmax)}°F / ${Math.round(dayData.tempmin)}°F ${dayData.label}`;
        }
      });

      // Update Calendar View cells
      if (days && days.length > 0) {
        days.forEach(dayData => {
          const calCell = document.querySelector(`.calendar-day[data-date="${dayData.datetime}"]`);
          if (calCell) {
            let wDiv = calCell.querySelector('.calendar-weather');
            if (!wDiv) {
              wDiv = document.createElement('div');
              wDiv.className = 'calendar-weather';
              const hotel = calCell.querySelector('.hotel-pill');
              if (hotel) {
                calCell.insertBefore(wDiv, hotel);
              } else {
                const firstEvent = calCell.querySelector('.calendar-event:not(.location-pill)');
                if (firstEvent) calCell.insertBefore(wDiv, firstEvent);
                else calCell.appendChild(wDiv);
              }
            }

            if (!wDiv.innerHTML) {
              const iconUrl = `https://raw.githubusercontent.com/visualcrossing/WeatherIcons/refs/heads/main/PNG/2nd%20Set%20-%20Color/${dayData.icon}.png`;
              const histIcon = dayData.isHistorical ? '<span title="Historical data" style="margin-left:3px; cursor:help;">🕒</span>' : '';
              wDiv.title = dayData.label;
              wDiv.innerHTML = `
              <img src="${iconUrl}" class="weather-icon" style="width:1.2em; height:1.2em;">
              <span>${Math.round(dayData.tempmax)}° / ${Math.round(dayData.tempmin)}°${histIcon}</span>
            `;
              if (dayData.isHistorical) wDiv.classList.add('historical');
            }
          }
        });
      }

    } catch (e) {
      console.error('Error processing weather data:', e);
    }
  }

  function updateSingleClock(el) {
    try {
      const tz = el.dataset.tz;
      if (!tz) return;
      const now = new Date();
      const options = { hour: 'numeric', minute: '2-digit', timeZone: tz, hour12: true };
      // Format: "2:45 PM"
      const timeStr = new Intl.DateTimeFormat('en-US', options).format(now);
      const val = el.querySelector('.clock-val');
      if (val) val.textContent = timeStr;
    } catch (e) { console.error(e); }
  }

  // Update all clocks every minute
  setInterval(() => {
    document.querySelectorAll('.local-time-display').forEach(updateSingleClock);
  }, 60000);

  async function fetchAllWeather() {
    const cards = document.querySelectorAll('.location-card');
    const results = [];
    for (const card of cards) {
      const data = await fetchCardWeatherData(card);
      results.push({ card, data });
    }
    for (const { card, data } of results) {
      if (data) renderCardWeather(card, data);
    }
  }

  async function getAiSuggestions(location, date, dailyPlan, fullLocationPlan, hotelInfo, weather) {
    if (!geminiApiKey) {
      alert('Please set your Google AI API Key in the "AI Settings" tab first.');
      return;
    }

    currentAiCacheKey = `${AI_CACHE_PREFIX}${encodeURIComponent(location)}_${date}`;

    try {
      const cachedJson = localStorage.getItem(currentAiCacheKey);
      if (cachedJson) {
        const cached = JSON.parse(cachedJson);
        if (Date.now() - cached.timestamp < AI_CACHE_EXPIRY) {
          aiConversationHistory = cached.history;
          aiModal.style.display = 'flex';
          aiModalTitle.textContent = `AI Suggestions for ${date}`;
          aiModalBody.innerHTML = '';

          aiConversationHistory.forEach((msg, index) => {
            if (index === 0) return; // Skip initial system prompt

            const div = document.createElement('div');
            div.className = `ai-message ${msg.role}`;

            if (msg.role === 'model') {
              const text = msg.parts[0].text;
              div.innerHTML = renderMarkdown(text);
            } else {
              div.textContent = msg.parts[0].text;
            }
            aiModalBody.appendChild(div);
          });
          aiModalBody.scrollTop = 0;
          return;
        }
      }
    } catch (e) { console.warn('Cache read error', e); }

    aiModal.style.display = 'flex';
    aiModalTitle.textContent = `AI Suggestions for ${date}`;
    aiModalBody.innerHTML = ''; // Clear previous conversation
    aiConversationHistory = []; // Reset history

    const prompt = `You are an expert local travel guide and itinerary planner.
I am planning a trip to **${location}**.
${hotelInfo ? `Accommodation: ${hotelInfo}` : ''}
${weather ? `Expected Weather: ${weather}` : ''}

**Context: My full itinerary for ${location} is:**
${fullLocationPlan}

**Focus Day: ${date}**
My current plan for this specific day is:
${dailyPlan || "No specific activities planned yet."}

**Task:**
Please provide **3 specific, actionable, and distinct suggestions** to enhance or fill my plan for **${date}**.
${weather ? `Take the expected weather (${weather}) into account (e.g. prefer indoor activities if raining).` : ''}
Consider logistical efficiency (places near my current activities), distinct experiences I might be missing, or highly-rated dining/relaxation spots nearby.

**Constraints:**
- Do NOT suggest attractions I am already visiting on this day or other days.
- Ensure suggestions are geographically feasible.
- Keep descriptions concise.

**Format:**
Use Markdown. For each suggestion, use an H3 (###) for the name, followed by a brief description and a specific reason why it fits this itinerary. Include a Markdown link to search for the location on Google Maps (e.g. Open in Maps).`;

    aiConversationHistory.push({
      role: 'user',
      parts: [{ text: prompt }]
    });

    await callGemini();
  }

  async function callGemini() {
    const loadingDiv = document.createElement('div');
    loadingDiv.className = 'ai-message loading';
    loadingDiv.textContent = 'Generating...';
    aiModalBody.appendChild(loadingDiv);
    aiModalBody.scrollTop = aiModalBody.scrollHeight;

    aiFollowUpInput.disabled = true;
    aiFollowUpBtn.disabled = true;

    try {
      const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${geminiModelName}:generateContent?key=${geminiApiKey}`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          contents: aiConversationHistory,
          generationConfig: {
            temperature: 0.7,
            topK: 1,
            topP: 1,
            maxOutputTokens: 800,
          },
        }),
      });

      loadingDiv.remove();

      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.error.message || `API Error: ${response.status}`);
      }

      const data = await response.json();
      let responseText = '';

      if (data.candidates && data.candidates.length > 0 && data.candidates[0].content) {
        responseText = data.candidates[0].content.parts[0].text;
      } else {
        let reason = 'No suggestions returned.';
        if (data.candidates && data.candidates[0].finishReason) {
          reason = `Could not generate suggestions. Reason: ${data.candidates[0].finishReason}`;
          if (data.promptFeedback && data.promptFeedback.blockReason) {
            reason += ` (Block Reason: ${data.promptFeedback.blockReason})`;
          }
        }
        responseText = `<p style="color: #e88;">${reason}</p>`;
      }

      aiConversationHistory.push({ role: 'model', parts: [{ text: responseText }] });

      if (currentAiCacheKey) {
        try {
          localStorage.setItem(currentAiCacheKey, JSON.stringify({
            timestamp: Date.now(),
            history: aiConversationHistory
          }));
        } catch (e) { }
      }

      const messageDiv = document.createElement('div');
      messageDiv.className = 'ai-message model';
      // Basic markdown to HTML conversion
      messageDiv.innerHTML = renderMarkdown(responseText);
      aiModalBody.appendChild(messageDiv);

    } catch (error) {
      loadingDiv.remove();
      console.error('Gemini API call failed:', error);
      const errorDiv = document.createElement('div');
      errorDiv.className = 'ai-message model';
      errorDiv.innerHTML = `<p style="color: #e88;"><strong>Error:</strong> ${error.message}</p><p>Please check your API key and network connection.</p>`;
      aiModalBody.appendChild(errorDiv);
    } finally {
      aiFollowUpInput.disabled = false;
      aiFollowUpBtn.disabled = false;
      aiFollowUpInput.focus();
    }
  }

  // Event delegation for checkbox state storage
  document.addEventListener('change', (e) => {
    if (e.target.matches('input[type="checkbox"][data-key]')) {
      localStorage.setItem(e.target.dataset.key, e.target.checked);
    }
  });

  function getChecklistCacheKey(type, count) {
    if (!parsed) return null;
    const locs = [...new Set(parsed.locations.slice(1).filter(l => l && l.trim()))].join(', ');
    const validDates = parsed.dates.slice(1).filter(d => d && d.trim());
    const start = validDates.length > 0 ? validDates[0] : '';
    return `${AI_CACHE_PREFIX}${type}_${encodeURIComponent(locs)}_${start}_${count}`;
  }

  async function generateChecklist(ignoreCache = false) {
    if (!parsed) {
      packingContent.innerHTML = '<p>Please build an itinerary first.</p>';
      return;
    }

    // Use 1 for apps to normalize cache key, as app suggestions rarely depend on traveler count
    const count = currentChecklistType === 'packing' ? (packingCountInput.value || 1) : 1;
    if (currentChecklistType === 'packing') {
      localStorage.setItem(PACKING_COUNT_STORAGE, count);
    }

    packingContent.innerHTML = '<p>Generating checklist... <span class="ai-message loading">✨</span></p>';

    const locs = parsed ? [...new Set(parsed.locations.slice(1).filter(l => l && l.trim()))].join(', ') : '';
    const validDates = parsed ? parsed.dates.slice(1).filter(d => d && d.trim()) : [];
    const start = validDates.length > 0 ? validDates[0] : '';
    const end = validDates.length > 0 ? validDates[validDates.length - 1] : '';

    let prompt = '';
    if (currentChecklistType === 'packing') {
      const weatherSummaries = [];
      document.querySelectorAll('.location-card').forEach(card => {
        const l = card.dataset.location;
        const wInline = card.querySelector('.weather-summary-inline');
        if (l && wInline) weatherSummaries.push(`${l}: ${wInline.textContent}`);
      });
      const weatherContext = weatherSummaries.join('; ');
      const allContent = parsed ? parsed.rows.map(r => r.cells.join(' ')).join(' ').substring(0, 3000) : '';

      prompt = `Act as a travel expert. Create a comprehensive packing checklist for ${count} person(s) for a trip to ${locs}.\nDates: ${start} to ${end}.\n\nContext:\n- Weather Forecast: ${weatherContext || 'Not available'}\n- Activities overview: ${allContent}\n\nFormat the response as a Markdown checklist categorized by Essentials, Clothing, Toiletries, Electronics, and Misc. Include quantity suggestions based on the trip duration.`;
    } else {
      prompt = `Act as a travel expert. Create a checklist of essential and useful mobile apps to download for a trip to ${locs}.\nInclude categories like Navigation, Transport, Language, Finance/Payment, and Local Guides. Explain briefly why each app is useful for this specific destination.\nFormat as a Markdown checklist.`;
    }

    currentAiCacheKey = getChecklistCacheKey(currentChecklistType, count);

    // Check Cache
    if (!ignoreCache) {
      try {
        const cachedJson = localStorage.getItem(currentAiCacheKey);
        if (cachedJson) {
          const cached = JSON.parse(cachedJson);
          if (Date.now() - cached.timestamp < AI_CACHE_EXPIRY) {
            packingContent.innerHTML = renderMarkdown(cached.text, currentAiCacheKey);
            return;
          }
        }
      } catch (e) { }
    }

    // Fetch from API
    if (!geminiApiKey) {
      packingContent.innerHTML = '<p class="error-msg visible">Please set your Google AI API Key in Settings first.</p>';
      return;
    }

    try {
      const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${geminiModelName}:generateContent?key=${geminiApiKey}`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          contents: [{ role: 'user', parts: [{ text: prompt }] }],
          generationConfig: {
            temperature: 0.7,
            maxOutputTokens: 2000,
          },
        }),
      });

      if (!response.ok) {
        throw new Error(`API Error: ${response.status}`);
      }

      const data = await response.json();
      const text = data.candidates?.[0]?.content?.parts?.[0]?.text || 'No suggestion returned.';

      packingContent.innerHTML = renderMarkdown(text, currentAiCacheKey);

      // Cache result
      localStorage.setItem(currentAiCacheKey, JSON.stringify({
        timestamp: Date.now(),
        text: text,
        history: [] // Not using chat history for packing list
      }));

    } catch (e) {
      console.error(e);
      packingContent.innerHTML = `<p class="error-msg visible">Error generating list: ${e.message}</p>`;
    }
  }

  btnGenerateChecklist?.addEventListener('click', () => generateChecklist(true));

  packingTabs.forEach(tab => {
    tab.addEventListener('click', () => {
      packingTabs.forEach(t => t.classList.remove('active'));
      tab.classList.add('active');
      currentChecklistType = tab.dataset.type;

      // Hide travelers input for apps
      document.getElementById('packing-inputs-wrapper').style.display = currentChecklistType === 'packing' ? 'flex' : 'none';

      // Attempt to load existing content
      if (parsed) {
        const count = currentChecklistType === 'packing' ? (packingCountInput?.value || 1) : 1;
        const key = getChecklistCacheKey(currentChecklistType, count);
        currentAiCacheKey = key;
        let loaded = false;
        try {
          const cachedJson = localStorage.getItem(key);
          if (cachedJson) {
            const cached = JSON.parse(cachedJson);
            if (Date.now() - cached.timestamp < AI_CACHE_EXPIRY) {
              packingContent.innerHTML = renderMarkdown(cached.text, key);
              loaded = true;
            }
          }
        } catch (e) { }

        if (!loaded) {
          packingContent.innerHTML = '<p style="color:var(--text-muted)">Click Regenerate to create this list.</p>';
        }
      }
    });
  });

  // Search Logic
  searchInput?.addEventListener('input', (e) => {
    const term = e.target.value.toLowerCase();
    const cards = document.querySelectorAll('.location-card');

    cards.forEach(card => {
      const days = card.querySelectorAll('.day-group');
      let hasVisibleDay = false;

      days.forEach(day => {
        const text = day.textContent.toLowerCase();
        if (text.includes(term)) {
          day.classList.remove('hidden');
          hasVisibleDay = true;
        } else {
          day.classList.add('hidden');
        }
      });

      // If card matches title, show all days? Or just filter days?
      // Let's hide card if no days match and card title doesn't match
      const titleMatch = card.querySelector('.location-card-name').textContent.toLowerCase().includes(term);

      if (titleMatch) {
        if (term) card.classList.remove('collapsed');
        // If title matches, show matches inside, but ensure card is visible
        // Optional: unhide all days if title matches? Let's keep granular filtering.
        card.classList.remove('hidden');
        if (!hasVisibleDay && term.length > 0) {
          // If title matches but no days, show card but maybe days are hidden. 
          // Let's unhide all days if title matches to avoid empty card
          days.forEach(d => d.classList.remove('hidden'));
        }
      } else {
        if (hasVisibleDay) {
          card.classList.remove('hidden');
          if (term) card.classList.remove('collapsed');
        }
        else card.classList.add('hidden');
      }
    });
  });

  async function handleFollowUp() {
    const prompt = aiFollowUpInput.value.trim();
    if (!prompt) return;

    aiFollowUpInput.value = '';

    // Render user message
    const messageDiv = document.createElement('div');
    messageDiv.className = 'ai-message user';
    messageDiv.textContent = prompt;
    aiModalBody.appendChild(messageDiv);
    aiModalBody.scrollTop = aiModalBody.scrollHeight;

    // Add to history
    aiConversationHistory.push({
      role: 'user',
      parts: [{ text: prompt }]
    });

    await callGemini();
  }

  aiFollowUpBtn?.addEventListener('click', handleFollowUp);
  aiFollowUpInput?.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      handleFollowUp();
    }
  });

  async function populateModels(keyOverride) {
    const key = keyOverride || geminiApiKey;
    if (!key) return;
    try {
      const res = await fetch(`https://generativelanguage.googleapis.com/v1beta/models?key=${key}`);
      if (!res.ok) return;
      const data = await res.json();
      if (data.models) {
        const contentModels = data.models.filter(m =>
          m.supportedGenerationMethods &&
          m.supportedGenerationMethods.includes('generateContent') &&
          !m.name.includes('vision') &&
          !m.name.includes('image') &&
          m.name.includes('gemini')
        );
        contentModels.sort((a, b) => {
          const nA = a.displayName || a.name;
          const nB = b.displayName || b.name;
          return nB.localeCompare(nA);
        });
        settingModel.innerHTML = '';
        contentModels.forEach(m => {
          const opt = document.createElement('option');
          const val = m.name.replace(/^models\//, '');
          opt.value = val;
          opt.textContent = m.displayName ? `${m.displayName} (${val})` : val;
          settingModel.appendChild(opt);
        });

        // Auto-select flash-lite if no user preference
        if (!localStorage.getItem('gemini_model')) {
          const liteModel = contentModels.find(m => m.name.endsWith('flash-lite-preview'));
          if (liteModel) {
            geminiModelName = liteModel.name.replace(/^models\//, '');
          } else {
            geminiModelName = 'gemini-flash-lite';
          }
          if (settingModel.querySelector(`option[value="${geminiModelName}"]`)) {
            settingModel.value = geminiModelName;
          }
        }
      }
    } catch (e) {
      console.error('Failed to list models', e);
    }
  }

  function updateGoogleStatus(signedIn, loading = false) {
    if (loading) {
      googleStatus.textContent = 'Connecting to Google...';
      googleStatus.style.color = 'var(--text-muted)';
      btnGoogleSignIn.disabled = true;
      return;
    }
    btnGoogleSignIn.disabled = false;
    if (signedIn) {
      googleStatus.textContent = 'Signed in successfully';
      googleStatus.style.color = 'var(--success)';
      btnGoogleSignIn.innerHTML = `<img src="https://fonts.gstatic.com/s/i/productlogos/googleg/v6/24px.svg" alt="" width="18"> Sign out of Google`;
      btnDrivePicker.style.display = 'flex';
      driveManualInput.style.display = 'block';
      const warning = document.getElementById('drive-auth-warning');
      if (warning) warning.style.display = 'none';
    } else {
      googleStatus.textContent = 'Not signed in';
      googleStatus.style.color = 'var(--text-muted)';
      btnGoogleSignIn.innerHTML = `<img src="https://fonts.gstatic.com/s/i/productlogos/googleg/v6/24px.svg" alt="" width="18"> Sign in with Google`;
      btnDrivePicker.style.display = 'none';
      driveManualInput.style.display = 'none';
      const warning = document.getElementById('drive-auth-warning');
      if (warning) warning.style.display = 'block';
      accessToken = null;
    }
  }

  function initTokenClient() {
    if (!googleClientId || tokenClient) return;

    // Check if the Google Identity Services script has finished loading
    if (typeof google === 'undefined' || !google.accounts) {
      return;
    }

    // 1. Initialize Identity (Sign In With Google / One Tap)
    // This handles the "Automatic Login" feature
    google.accounts.id.initialize({
      client_id: googleClientId,
      auto_select: true,
      context: 'signin',
      callback: (response) => {
        // Successful automatic or manual identity verification
        if (tokenClient && !accessToken) {
          tokenClient.requestAccessToken({ prompt: '' });
        }
      }
    });

    // 2. Initialize Authorization (Google Drive access)
    tokenClient = google.accounts.oauth2.initTokenClient({
      client_id: googleClientId,
      scope: 'https://www.googleapis.com/auth/drive.readonly https://www.googleapis.com/auth/spreadsheets.readonly',
      callback: (response) => {
        if (response.error) {
          // If silent prompt fails at startup, we just revert to manual sign-in status
          if (response.error === 'immediate_failed') {
            updateGoogleStatus(false);
          } else {
            showError('Google Auth Error: ' + response.error);
          }
          return;
        }
        localStorage.setItem('google_signed_in', 'true');
        accessToken = response.access_token;
        updateGoogleStatus(true);
        loadFromUrl(); // Re-check if we need to load a docId now that we're authed
      },
    });
  }

  function trackRecentDoc(fileId, name, gid = '0') {
    const metadata = JSON.parse(localStorage.getItem(DOC_METADATA_KEY) || '{}');

    let startDate = metadata[fileId]?.startDate || null;
    let endDate = metadata[fileId]?.endDate || null;

    if (parsed && parsed.dates && parsed.dates.length > 1) {
      const validDates = parsed.dates.slice(1).filter(d => d && d.trim());
      if (validDates.length > 0) {
        startDate = validDates[0];
        endDate = validDates[validDates.length - 1];
      }
    }

    metadata[fileId] = {
      name: name || metadata[fileId]?.name || 'Untitled Itinerary',
      lastAccessed: Date.now(),
      startDate,
      endDate,
      gid
    };
    localStorage.setItem(DOC_METADATA_KEY, JSON.stringify(metadata));
    renderRecentDocs();
  }

  function renderRecentDocs() {
    const container = document.getElementById('recent-docs-list');
    const wrapper = document.getElementById('recent-docs-container');
    const sortBy = recentDocsSort?.value || 'recent';
    if (!container || !wrapper) return;

    const metadata = JSON.parse(localStorage.getItem(DOC_METADATA_KEY) || '{}');
    const entries = Object.entries(metadata)
      .filter(([id, info]) => localStorage.getItem(DOC_CACHE_PREFIX + id + '_' + (info.gid || '0')))
      .sort((a, b) => {
        if (sortBy === 'name') {
          return a[1].name.localeCompare(b[1].name);
        } else if (sortBy === 'date') {
          const tsA = parseDateForSort(a[1].startDate);
          const tsB = parseDateForSort(b[1].startDate);
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
      item.style.display = 'flex';
      item.style.alignItems = 'center';
      item.style.justifyContent = 'space-between';
      item.style.paddingRight = '0.5rem';

      const dateRange = info.startDate ? `${info.startDate}${info.endDate && info.endDate !== info.startDate ? ' - ' + info.endDate : ''}` : '';

      item.innerHTML = `
        <div class="recent-doc-info" style="display: flex; flex-direction: column; gap: 0.15rem; flex: 1; overflow: hidden; cursor: pointer;">
          <div style="display: flex; align-items: center; gap: 0.75rem;">
            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
              <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
              <polyline points="14 2 14 8 20 8"></polyline>
              <line x1="16" y1="13" x2="8" y2="13"></line>
              <line x1="16" y1="17" x2="8" y2="17"></line>
              <polyline points="10 9 9 9 8 9"></polyline>
            </svg>
            <span style="white-space: nowrap; overflow: hidden; text-overflow: ellipsis; font-weight: 500;">${escapeHtml(info.name)}</span>
          </div>
          ${dateRange ? `<span style="font-size: 0.75rem; color: var(--text-muted); margin-left: 2.15rem;">${escapeHtml(dateRange)}</span>` : ''}
        </div>
        <button class="delete-recent-btn" title="Remove from recent list" style="background: none; border: none; color: var(--text-muted); cursor: pointer; padding: 6px; display: flex; align-items: center; justify-content: center; opacity: 0.5; transition: all 0.2s;">
          <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <polyline points="3 6 5 6 21 6"></polyline>
            <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path>
          </svg>
        </button>
      `;

      item.querySelector('.recent-doc-info').onclick = () => {
        updateUrlWithDocId(id, info.gid || '0');
        loadFromUrl();
      };

      const delBtn = item.querySelector('.delete-recent-btn');
      delBtn.onmouseover = () => { delBtn.style.opacity = '1'; delBtn.style.color = '#e88'; };
      delBtn.onmouseout = () => { delBtn.style.opacity = '0.5'; delBtn.style.color = ''; };
      delBtn.onclick = (e) => {
        e.stopPropagation();
        if (confirm(`Remove "${info.name}" from your recent list?`)) {
          const metadata = JSON.parse(localStorage.getItem(DOC_METADATA_KEY) || '{}');
          delete metadata[id];
          localStorage.setItem(DOC_METADATA_KEY, JSON.stringify(metadata));
          localStorage.removeItem(DOC_CACHE_PREFIX + id + '_' + (info.gid || '0'));
          renderRecentDocs();
        }
      };

      container.appendChild(item);
    });
  }

  const handleGoogleSignIn = () => {
    if (accessToken) {
      google.accounts.oauth2.revoke(accessToken, () => {
        localStorage.removeItem('google_signed_in');
        updateGoogleStatus(false);
      });
    } else {
      if (!googleClientId) {
        alert('Google Client ID is not configured.');
        return;
      }

      initTokenClient();

      if (!tokenClient) {
        alert('Google Auth library is still loading. Please wait a few seconds and try again.');
        return;
      }
      tokenClient.requestAccessToken({ prompt: 'select_account' });
    }
  };

  btnGoogleSignIn?.addEventListener('click', handleGoogleSignIn);
  btnDriveAuth?.addEventListener('click', handleGoogleSignIn);

  btnDrivePicker?.addEventListener('click', () => {
    if (!accessToken) {
      tokenClient.requestAccessToken({ prompt: '' });
      return;
    }
    gapi.load('picker', { 'callback': createPicker });
  });

  btnDriveUrlFetch?.addEventListener('click', () => {
    const input = driveUrlInput.value.trim();
    if (!input) return;

    // Extract File ID (supports /d/ID or /d/e/ID or just the ID)
    const fileIdMatch = input.match(/\/d\/(?:e\/)?([a-zA-Z0-9-_]+)/);
    const gidMatch = input.match(/[#&]gid=([0-9]+)/);

    const fileId = fileIdMatch ? fileIdMatch[1] : input;
    const gid = gidMatch ? gidMatch[1] : '0';

    // Normalize UI to show the document URL instead of internal export links or just the ID
    driveUrlInput.value = `https://docs.google.com/spreadsheets/d/${fileId}/edit${gid !== '0' ? '#gid=' + gid : ''}`;

    fetchSheetData(fileId, 'Manual Entry', gid);
  });

  function createPicker() {
    const view = new google.picker.View(google.picker.ViewId.SPREADSHEETS);
    const picker = new google.picker.PickerBuilder()
      .addView(view)
      .setOAuthToken(accessToken)
      .setCallback(pickerCallback)
      .build();
    picker.setVisible(true);
  }

  async function fetchSheetData(fileId, fileName, gid = '0', skipHistory = false) {
    showError(`Fetching "${fileName}"...`);
    try {
      // Use Sheets API v4 to get formatted values and merge metadata from the first tab
      const url = `https://sheets.googleapis.com/v4/spreadsheets/${fileId}?includeGridData=true&fields=properties.title,sheets(merges,data(rowData(values(formattedValue))))`;
      const response = await fetch(url, {
        headers: { 'Authorization': `Bearer ${accessToken}` }
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

        const val = rows[startR][startC];
        if (val) {
          for (let r = startR; r < endR; r++) {
            if (!rows[r]) rows[r] = [];
            for (let c = startC; c < endC; c++) {
              rows[r][c] = val;
            }
          }
        }
      });

      localStorage.setItem(DOC_CACHE_PREFIX + fileId + '_' + gid, JSON.stringify(rows));
      h1.textContent = spreadsheetTitle;
      trackRecentDoc(fileId, spreadsheetTitle, gid);
      if (!skipHistory) {
        updateUrlWithDocId(fileId, gid);
      }
      if (btnClearDoc) btnClearDoc.style.display = 'block';
      buildItineraryFromRows(rows, true);
    } catch (e) {
      console.error(e);
      showError('Error loading from Sheets: ' + e.message);
    }
  }

  async function pickerCallback(data) {
    if (data.action === google.picker.Action.PICKED) {
      const doc = data.docs[0];
      // Populate the input with the user-friendly Doc ID URL after picking
      driveUrlInput.value = `https://docs.google.com/spreadsheets/d/${doc.id}/edit`;
      fetchSheetData(doc.id, doc.name);
    }
  }

  settingApiKey?.addEventListener('change', () => {
    populateModels(settingApiKey.value.trim());
  });

  settingsBtn?.addEventListener('click', () => {
    settingsModal.style.display = 'flex';
    settingApiKey.value = geminiApiKey || '';
    settingVcKey.value = visualCrossingKey || '';
    settingModel.value = geminiModelName;
    if (geminiApiKey) {
      populateModels().then(() => {
        if (geminiModelName && settingModel.querySelector(`option[value="${geminiModelName}"]`)) {
          settingModel.value = geminiModelName;
        }
      });
    }
  });

  settingsClose?.addEventListener('click', () => {
    settingsModal.style.display = 'none';
  });

  btnSaveSettings?.addEventListener('click', () => {
    const key = settingApiKey.value.trim();
    const vcKey = settingVcKey.value.trim();
    const model = settingModel.value;

    if (key) {
      localStorage.setItem(GEMINI_API_KEY_STORAGE, key);
      geminiApiKey = key;
      document.body.classList.add('has-api-key');
    } else {
      localStorage.removeItem(GEMINI_API_KEY_STORAGE);
      geminiApiKey = null;
      document.body.classList.remove('has-api-key');
    }

    if (vcKey) {
      localStorage.setItem(VC_API_KEY_STORAGE, vcKey);
      visualCrossingKey = vcKey;
    } else {
      localStorage.removeItem(VC_API_KEY_STORAGE);
      visualCrossingKey = null;
    }

    if (model) {
      localStorage.setItem('gemini_model', model);
      geminiModelName = model;
    }

    settingsModal.style.display = 'none';
  });

  btnClearCache?.addEventListener('click', () => {
    if (confirm('Clear all cached weather, AI suggestions, and packing lists? Your saved itineraries will be preserved.')) {
      const keysToRemove = [];
      for (let i = 0; i < localStorage.length; i++) {
        const key = localStorage.key(i);
        if (key.startsWith(WEATHER_CACHE_PREFIX) || key.startsWith(AI_CACHE_PREFIX)) {
          keysToRemove.push(key);
        }
      }
      keysToRemove.forEach(k => localStorage.removeItem(k));
      alert('Cache cleared.');
    }
  });

  aiModalClose?.addEventListener('click', () => {
    aiModal.style.display = 'none';
  });
  window.addEventListener('click', (event) => {
    if (event.target == aiModal) {
      aiModal.style.display = 'none';
    }
    if (event.target == settingsModal) {
      settingsModal.style.display = 'none';
    }
  });

  function escapeHtml(s) {
    const div = document.createElement('div');
    div.textContent = s;
    return div.innerHTML;
  }

  function getTextContent(htmlStr) {
    const div = document.createElement('div');
    div.innerHTML = htmlStr;
    return div.textContent || '';
  }

  function linkifyAndSanitize(str) {
    const div = document.createElement('div');
    div.innerHTML = str;
    // Convert newlines to <br> to ensure they are preserved as breaks
    div.innerHTML = (str || '').replace(/\n/g, '<br>');

    function walk(node) {
      if (node.nodeType === Node.TEXT_NODE) {
        return escapeHtml(node.textContent);
      }
      if (node.nodeType === Node.ELEMENT_NODE) {
        if (node.tagName === 'BR') return '<br>';
        if (node.tagName === 'A' && node.hasAttribute('href')) {
          const a = document.createElement('a');
          a.href = node.getAttribute('href');
          a.target = '_blank';
          a.rel = 'noopener noreferrer';
          a.innerHTML = Array.from(node.childNodes).map(walk).join('');
          return a.outerHTML;
        }
        return Array.from(node.childNodes).map(walk).join('');
      }
      return '';
    }
    return Array.from(div.childNodes).map(walk).join('');
  }

  function simplifyHtml(html) {
    const div = document.createElement('div');
    div.innerHTML = html;
    function walk(node) {
      if (node.nodeType === Node.TEXT_NODE) return node.textContent;
      if (node.nodeType === Node.ELEMENT_NODE) {
        if (node.tagName === 'A' && node.hasAttribute('href')) {
          const content = Array.from(node.childNodes).map(walk).join('');
          return `<a href="${node.getAttribute('href')}">${content}</a>`;
        }
        if (node.tagName === 'BR') return '\n';
        if (node.tagName === 'P' || node.tagName === 'DIV') return '\n' + Array.from(node.childNodes).map(walk).join('');
        return Array.from(node.childNodes).map(walk).join('');
      }
      return '';
    }
    return Array.from(div.childNodes).map(walk).join('').trim();
  }

  shareBtn?.addEventListener('click', async () => {
    const longUrl = window.location.href;
    const urlParams = new URLSearchParams(window.location.search);
    const isDocIdUrl = urlParams.has('docId');

    const originalContent = shareBtn.innerHTML;
    const originalTitle = shareBtn.title;

    let success = false;
    let urlToCopy = longUrl;

    if (!isDocIdUrl) {
      shareBtn.innerHTML = '...';
      shareBtn.disabled = true;
      shareBtn.title = 'Shortening...';

      try {
        const res = await fetch(`https://tinyurl.com/api-create.php?url=${encodeURIComponent(longUrl)}`);
        if (res.ok) {
          const shortUrl = await res.text();
          if (shortUrl.startsWith('http')) {
            urlToCopy = shortUrl;
          }
        }
      } catch (e) {
        console.error('Could not shorten URL, will copy long URL instead:', e);
      }
    }

    try {
      await navigator.clipboard.writeText(urlToCopy);
      success = true;
    } catch (e) {
      console.error('Failed to copy to clipboard:', e);
    }

    shareBtn.innerHTML = success ? '✓' : '✗';
    shareBtn.style.color = success ? 'var(--success)' : '#e88';
    shareBtn.title = success ? 'Copied!' : 'Failed to copy';

    setTimeout(() => {
      shareBtn.innerHTML = originalContent;
      shareBtn.disabled = false;
      shareBtn.style.color = '';
      shareBtn.title = originalTitle;
    }, 2000);
  });

  btnParse?.addEventListener('click', () => buildItinerary(csvInput.value));

  btnClearDoc?.addEventListener('click', resetApp);

  csvInput?.addEventListener('paste', (event) => {
    const html = (event.clipboardData || window.clipboardData).getData('text/html');
    if (!html || !html.includes('<table')) {
      return; // Let default paste happen
    }

    event.preventDefault();

    const div = document.createElement('div');
    div.innerHTML = html;
    const table = div.querySelector('table');
    if (!table) {
      const text = (event.clipboardData || window.clipboardData).getData('text/plain');
      csvInput.value = text;
      return;
    }

    const grid = [];
    const tableRows = table.querySelectorAll('tr');
    tableRows.forEach((row, rowIndex) => {
      if (!grid[rowIndex]) grid[rowIndex] = [];
      const cells = row.querySelectorAll('td, th');
      let colIndex = 0;
      cells.forEach(cell => {
        while (grid[rowIndex][colIndex] !== undefined) colIndex++;
        const htmlContent = simplifyHtml(cell.innerHTML);
        const colspan = parseInt(cell.getAttribute('colspan') || '1', 10);
        const rowspan = parseInt(cell.getAttribute('rowspan') || '1', 10);
        for (let r = 0; r < rowspan; r++) {
          const targetRow = rowIndex + r;
          if (!grid[targetRow]) grid[targetRow] = [];
          for (let c = 0; c < colspan; c++) {
            grid[targetRow][colIndex + c] = htmlContent;
          }
        }
      });
    });

    const maxCols = grid.reduce((max, row) => Math.max(max, row ? row.length : 0), 0);
    const rowsWithHtml = grid.map(row => {
      const newRow = [];
      for (let i = 0; i < maxCols; i++) newRow[i] = (row && row[i] !== undefined) ? row[i] : '';
      return newRow;
    });

    const toCsvValue = (val) => {
      let v = (val || '').toString();
      if (/[\t\n\r"]/.test(v)) {
        return '"' + v.replace(/"/g, '""') + '"';
      }
      return v;
    };

    const csvTextForUrl = rowsWithHtml.map(row => row.map(toCsvValue).join('\t')).join('\n');
    csvInput.value = csvTextForUrl;
    buildItineraryFromRows(rowsWithHtml, false, csvTextForUrl);
  });

  tabs.forEach(tab => {
    tab.addEventListener('click', () => {
      const id = tab.dataset.tab;
      tabs.forEach(t => t.classList.remove('active'));
      panels.forEach(p => {
        p.classList.toggle('active', p.id === 'panel-' + id);
      });
      tab.classList.add('active');
      showError('');
    });
  });

  // View toggles
  viewBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      const view = btn.dataset.view;
      viewBtns.forEach(b => b.classList.remove('active'));
      btn.classList.add('active');

      const searchWrapper = document.querySelector('.search-wrapper');
      if (searchWrapper) searchWrapper.style.display = view === 'list' ? '' : 'none';

      document.getElementById('itinerary-list').classList.remove('visible');
      document.getElementById('itinerary-map').classList.remove('visible');
      document.getElementById('itinerary-calendar').classList.remove('visible');
      document.getElementById('itinerary-packing').classList.remove('visible');

      if (view === 'map') {
        document.getElementById('itinerary-map').classList.add('visible');
        if (map) {
          map.invalidateSize();
          if (mapPoints.length > 0) {
            const bounds = L.latLngBounds(mapPoints);
            map.fitBounds(bounds, { padding: [50, 50], maxZoom: 12 });
          }
        }
      } else if (view === 'calendar') {
        document.getElementById('itinerary-calendar').classList.add('visible');
      } else if (view === 'packing') {
        document.getElementById('itinerary-packing').classList.add('visible');
        // Load persisted count
        const savedCount = localStorage.getItem(PACKING_COUNT_STORAGE);
        if (savedCount && packingCountInput) packingCountInput.value = savedCount;

        if (parsed) {
          // Reset to correct count logic based on active tab
          const activeTab = document.querySelector('.packing-tab.active');
          currentChecklistType = activeTab ? activeTab.dataset.type : 'packing';
          document.getElementById('packing-inputs-wrapper').style.display = currentChecklistType === 'packing' ? 'flex' : 'none';

          const count = currentChecklistType === 'packing' ? (packingCountInput?.value || 1) : 1;
          const key = getChecklistCacheKey(currentChecklistType, count);
          currentAiCacheKey = key;

          let contentFound = false;
          try {
            const cachedJson = localStorage.getItem(key);
            if (cachedJson) {
              const cached = JSON.parse(cachedJson);
              if (Date.now() - cached.timestamp < AI_CACHE_EXPIRY) {
                packingContent.innerHTML = renderMarkdown(cached.text, key);
                contentFound = true;
              }
            }
          } catch (e) { }

          if (!contentFound) {
            packingContent.innerHTML = '<p style="color:var(--text-muted)">Click Regenerate to create the list.</p>';
          }
        }
      } else {
        document.getElementById('itinerary-list').classList.add('visible');

        if (!hasScrolledToToday) {
          const activeDay = document.querySelector('.day-group.today-active');
          const activeCard = document.querySelector('.location-card.expanded');
          if (activeDay || activeCard) {
            setTimeout(() => {
              (activeDay || activeCard).scrollIntoView({ behavior: 'smooth', block: activeDay ? 'center' : 'start' });
            }, 100);
            hasScrolledToToday = true;
          }
        }
      }
    });
  });

  // Reset functionality
  h1.title = 'Clear data and start over';
  h1.style.cursor = 'pointer';
  h1?.addEventListener('click', resetApp);

  // Initial load
  (async () => {
    if (geminiApiKey) {
      document.body.classList.add('has-api-key');
      populateModels();
    }

    // Wait for Google Identity Services library to load before attempting silent sign-in
    const waitGsi = setInterval(() => {
      if (typeof google !== 'undefined' && google.accounts && google.accounts.id && google.accounts.oauth2) {
        clearInterval(waitGsi);
        initTokenClient();

        // Use One Tap with a notification listener to debug failures
        google.accounts.id.prompt((notification) => {
          if (notification.isSkippedMoment() || notification.isDismissedMoment()) {
            // If the automatic prompt is skipped (due to the 10-minute cooldown 
            // or blocked cookies) or dismissed, revert to the manual sign-in button.
            if (!accessToken) updateGoogleStatus(false);
          }
        });
      }
    }, 500);

    recentDocsSort?.addEventListener('change', renderRecentDocs);

    window.addEventListener('popstate', loadFromUrl);
    renderRecentDocs();
    loadFromUrl(); // Load itinerary data from URL if present
  })();

  // Offline support & Service Worker Registration
  const offlineIndicator = document.getElementById('offline-indicator');

  function updateOnlineStatus() {
    if (!navigator.onLine) {
      offlineIndicator.classList.add('visible');
    } else {
      offlineIndicator.classList.remove('visible');
    }
  }

  window.addEventListener('online', updateOnlineStatus);
  window.addEventListener('offline', updateOnlineStatus);
  updateOnlineStatus();

  if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('sw.js')
      .catch(err => console.log('Service Worker registration failed', err));

    // Automatically reload the page when a new service worker takes over
    let refreshing = false;
    navigator.serviceWorker.addEventListener('controllerchange', () => {
      if (!refreshing) {
        window.location.reload();
        refreshing = true;
      }
    });
  }
})();
