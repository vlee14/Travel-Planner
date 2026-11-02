// Loading a trip and drawing the list and calendar.
export function installItinerary(app) {
  app.showError = function showError(msg) {
    const text = msg || '';
    const inBuilder = document.getElementById('itinerary-builder')?.classList.contains('visible');
    app.errorMsg.textContent = text;
    app.errorMsg.classList.toggle('visible', !!text && !inBuilder);
    const builderError = document.getElementById('builder-error');
    if (builderError) {
      builderError.textContent = text;
      builderError.classList.toggle('visible', !!text && !!inBuilder);
    }
  };

  app.resetApp = function resetApp(skipHistory = false) {
    const skip = skipHistory === true;
    // Clear data model
    app.parsed = null;

    // Clear inputs
    app.csvInput.value = '';
    if (app.driveUrlInput) app.driveUrlInput.value = '';
    app.showError('');

    // Reset UI visibility
    app.itinerarySection.classList.remove('visible');
    app.inputSection.style.display = 'block';
    app.emptyState.style.display = 'block';
    app.h1.textContent = 'Travel Itinerary';
    app.shareBtn.style.display = 'none';
    if (app.printBtn) app.printBtn.style.display = 'none';
    if (app.syncBtn) app.syncBtn.style.display = 'none';
    if (app.tripProgressBar) app.tripProgressBar.style.display = 'none';
    const syncLabel = document.getElementById('sync-status-label');
    if (syncLabel) syncLabel.style.display = 'none';
    if (app.btnClearDoc) app.btnClearDoc.style.display = 'none';

    // Stop auto-sync
    if (app.autoSyncInterval) { clearInterval(app.autoSyncInterval); app.autoSyncInterval = null; }
    app.lastSyncedDocId = null;

    // Clear rendered content
    app.listContent.innerHTML = '';
    document.getElementById('itinerary-calendar').innerHTML = '';
    if (app.map) {
      app.resetMapData();
    }

    // Refresh the recent docs list on the home page
    app.renderRecentDocs();

    app.parkBuilderPanel();
    app.showBuildPanel();
    app.ensureBuilder();
    app.setBuilderMode('create');

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
  };

  app.buildItineraryFromRows = function buildItineraryFromRows(rows, fromUrl = false, csvTextForUrl = '') {
    app.showError('');
    if (!rows.length) {
      app.showError('No data found. Paste your sheet content (first row = locations, second = dates, first column = times).');
      return;
    }
    app.parsed = app.parseSheetRows(rows);
    const hasDay = app.parsed && app.parsed.locations.some((loc, i) =>
      i > 0 && (loc || '').toString().trim() && (app.parsed.dates[i] || '').toString().trim()
    );
    if (!app.parsed || (!app.parsed.rows.length && !hasDay)) {
      app.showError('Could not parse structure. Ensure first row = locations, second row = dates, first column = times.');
      return;
    }
    app.setBuilderMode('create');
    app.parkBuilderPanel();
    app.render();
    app.itinerarySection.classList.add('visible');
    app.emptyState.style.display = 'none';
    app.inputSection.style.display = 'none';
    app.shareBtn.style.display = 'flex';
    if (app.printBtn) app.printBtn.style.display = 'flex';

    // Switch to calendar view by default, or list view on small screens
    if (window.innerWidth < 768) {
      document.querySelector('.view-btn[data-view="list"]').click();
    } else {
      document.querySelector('.view-btn[data-view="calendar"]').click();
    }

    app.initMap();
    app.fetchAllWeather();
    app.renderTripProgress();

    if (!fromUrl && csvTextForUrl) {
      app.updateUrlWithData(csvTextForUrl);
    }

    // Update recent docs metadata with date range if it's a docId-based load
    const urlParams = new URLSearchParams(window.location.search);
    const docId = urlParams.get('docId');
    if (docId) {
      app.trackRecentDoc(docId);
      if (app.syncBtn && app.lastSyncedDocId === docId) app.syncBtn.style.display = 'flex';
    } else if (app.syncBtn) {
      app.syncBtn.style.display = 'none';
    }
  };

  app.buildItinerary = function buildItinerary(csvText, fromUrl = false) {
    const rows = app.parseCSV(csvText);
    app.buildItineraryFromRows(rows, fromUrl, csvText);
  };

  app.render = function render() {
    if (!app.parsed) return;
    app.resetMapData();
    app.hasScrolledToToday = false;
    const { locations, dates, rows, hotelCells } = app.parsed;

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
    app.listContent.innerHTML = '';
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
        ts: app.parseDateForSort((dates[i] || '').toString().trim())
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
    locationRuns.sort((a, b) => app.parseDateForSort(a.dateKeys[0]) - app.parseDateForSort(b.dateKeys[0]));

    const now = new Date();
    now.setHours(0, 0, 0, 0);
    const todayYmd = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;

    app.searchInput.value = ''; // Reset search

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
          if (e.time) sortTracker = app.parseTimeForSort(e.time);
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
      if (run.dateKeys.some(k => app.parseSmartDate(k) === todayYmd)) {
        isCurrent = true;
      }

      const startTs = app.parseDateForSort(run.dateKeys[0]);
      const endTs = app.parseDateForSort(run.dateKeys[run.dateKeys.length - 1]);
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
      summary.innerHTML = `<h3 class="location-card-name">${app.escapeHtml(run.locName)}</h3>`;

      const meta = document.createElement('div');
      meta.className = 'location-card-meta';

      const dateDiv = document.createElement('div');
      dateDiv.innerHTML = `<span class="date-range">${app.escapeHtml(dateRangeStr)}</span>`;
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
          if (e.time) sortTracker = app.parseTimeForSort(e.time);
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
        if (app.parseSmartDate(dateLabel) === todayYmd) {
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
        titleDiv.innerHTML = `<span>${app.escapeHtml(dateLabel)}</span>`;

        const copyBtn = document.createElement('button');
        copyBtn.className = 'btn-copy-day';
        copyBtn.title = 'Copy day plan';
        copyBtn.setAttribute('aria-label', 'Copy day plan');
        copyBtn.innerHTML = `<svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"></rect><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"></path></svg>`;
        copyBtn.onclick = async (e) => {
          e.stopPropagation();
          const lines = [`${dateLabel} — ${run.locName}`];
          if (hotelForDay) lines.push(`Hotel: ${hotelForDay}`);
          events.forEach(({ time, content: evtContent }) => {
            lines.push(`${time} - ${app.getTextContent(evtContent)}`);
          });
          notes.forEach(n => lines.push(`Note: ${app.getTextContent(n.content)}`));
          try {
            await navigator.clipboard.writeText(lines.join('\n'));
            copyBtn.classList.add('copied');
            setTimeout(() => copyBtn.classList.remove('copied'), 2000);
          } catch (err) {
            console.error('Copy failed:', err);
          }
        };
        titleDiv.appendChild(copyBtn);

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
            p.innerHTML = app.linkifyAndSanitize(n.content);
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
        suggestBtn.innerHTML = '<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 3l1.6 4.8L18.5 9.4 13.6 11 12 15.8 10.4 11 5.5 9.4l4.9-1.6L12 3z"></path><path d="M18 14l.6 1.8 1.8.6-1.8.6L18 19l-.6-1.8-1.8-.6 1.8-.6L18 14z"></path></svg>';
        suggestBtn.onclick = (e) => {
          e.stopPropagation();
          if (!app.currentAiKey()) {
            alert(app.aiKeyMessage());
            return;
          }
          const weather = titleDiv.dataset.weatherSummary;
          app.getAiSuggestions(run.locName, dateLabel, dailyPlan, fullLocationPlan, hotelInfo, weather);
        };
        titleDiv.appendChild(suggestBtn);

        events.forEach(({ time, content }) => {
          const plainContent = app.getTextContent(content);
          const checkKey = `${app.EVENT_CHECK_PREFIX}${encodeURIComponent(run.locName)}_${encodeURIComponent(dateLabel)}_${encodeURIComponent(time)}_${encodeURIComponent(plainContent)}`;
          const isChecked = app.db.get(checkKey) === 'true';

          const eventCard = document.createElement('div');
          eventCard.className = isChecked ? 'event-card checked' : 'event-card';

          const checkbox = document.createElement('input');
          checkbox.type = 'checkbox';
          checkbox.className = 'event-card-checkbox';
          checkbox.dataset.key = checkKey;
          checkbox.checked = isChecked;
          checkbox.title = 'Mark as done';
          checkbox.setAttribute('aria-label', 'Mark activity as done');
          checkbox.onchange = (e) => {
            e.stopPropagation();
            app.db.set(checkKey, checkbox.checked ? 'true' : 'false');
            eventCard.classList.toggle('checked', checkbox.checked);
          };

          const timeSpan = document.createElement('span');
          timeSpan.className = 'event-time';
          timeSpan.textContent = time;

          const contentSpan = document.createElement('span');
          contentSpan.className = 'event-content';
          contentSpan.innerHTML = app.linkifyAndSanitize(content);

          eventCard.appendChild(checkbox);
          eventCard.appendChild(timeSpan);
          eventCard.appendChild(contentSpan);
          dayDiv.appendChild(eventCard);
        });
        daysContainer.appendChild(dayDiv);
        daysInner.appendChild(dayDiv);
      });
      daysContainer.appendChild(daysInner);
      cardEl.appendChild(daysContainer);
      app.listContent.appendChild(cardEl);
    });
    app.renderCalendar();
  };

  app.renderCalendar = function renderCalendar() {
    const cal = document.getElementById('itinerary-calendar');
    cal.innerHTML = '';
    if (!app.parsed) return;
    const { locations, dates, rows } = app.parsed;

    // Backfill hotel logic for calendar (similar to render/list view)
    let effectiveHotel = [];
    if (app.parsed.hotelCells && app.parsed.locations.length) {
      const pad = Array(app.parsed.locations.length).fill('');
      const cells = app.parsed.hotelCells.length >= app.parsed.locations.length ? app.parsed.hotelCells : [...parsed.hotelCells, ...pad.slice(app.parsed.hotelCells.length)];
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

      const ymd = app.parseSmartDate(rawDate); // "YYYY-MM-DD"
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
            num.classList.add('month-start');
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
              pill.innerHTML = app.linkifyAndSanitize(timePart + ev.content);
              pill.title = app.getTextContent(timePart + ev.content);
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
  };

  app.getTripDateRange = function getTripDateRange() {
    if (!app.parsed) return null;
    const timestamps = [];
    for (let i = 1; i < app.parsed.dates.length; i++) {
      const d = (app.parsed.dates[i] || '').trim();
      const loc = (app.parsed.locations[i] || '').trim();
      if (d && loc) {
        const ts = app.parseDateForSort(d);
        if (ts !== Infinity) timestamps.push(ts);
      }
    }
    if (!timestamps.length) return null;
    timestamps.sort((a, b) => a - b);
    const startDate = new Date(timestamps[0]);
    startDate.setHours(0, 0, 0, 0);
    const endDate = new Date(timestamps[timestamps.length - 1]);
    endDate.setHours(0, 0, 0, 0);
    const totalDays = Math.round((endDate - startDate) / 86400000) + 1;
    return { startDate, endDate, totalDays };
  };

  app.renderTripProgress = function renderTripProgress() {
    if (!app.tripProgressBar || !app.tripProgressLabel || !app.tripProgressFill) return;
    const range = app.getTripDateRange();
    if (!range) {
      app.tripProgressBar.style.display = 'none';
      return;
    }
    const { startDate, endDate, totalDays } = range;
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const todayTs = today.getTime();
    let label;
    let fillPct;

    if (todayTs < startDate.getTime()) {
      const daysUntil = Math.round((startDate - today) / 86400000);
      label = daysUntil === 1 ? 'Trip starts tomorrow' : `Trip starts in ${daysUntil} days`;
      fillPct = 0;
    } else if (todayTs > endDate.getTime()) {
      label = 'Trip completed';
      fillPct = 100;
    } else {
      const currentDay = Math.round((todayTs - startDate.getTime()) / 86400000) + 1;
      label = `Day ${currentDay} of ${totalDays}`;
      fillPct = Math.min(100, (currentDay / totalDays) * 100);
    }

    app.tripProgressLabel.textContent = label;
    app.tripProgressFill.style.width = `${fillPct}%`;
    app.tripProgressBar.style.display = 'block';
  };

  app.updateUrlWithData = async function updateUrlWithData(csvText) {
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
      const title = (app.h1?.textContent || '').trim();
      if (title && title !== 'Travel Itinerary') url.searchParams.set('title', title);
      else url.searchParams.delete('title');
      window.history.pushState({}, '', url);
    } catch (e) {
      console.error('Failed to compress and update URL:', e);
    }
  };

  app.updateUrlWithDocId = function updateUrlWithDocId(fileId, gid = '0') {
    const url = new URL(window.location);
    url.searchParams.delete('data');
    url.searchParams.set('docId', fileId);
    if (gid !== '0') url.searchParams.set('gid', gid);
    else url.searchParams.delete('gid');
    window.history.pushState({}, '', url);
  };

  app.loadFromUrl = async function loadFromUrl() {
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
        app.buildItinerary(decompressed, true);
        const title = urlParams.get('title');
        if (title) app.h1.textContent = title;
      } catch (e) {
        console.error('Failed to decompress data from URL:', e);
        app.showError('Could not load itinerary from URL. The data might be corrupted.');
      }
    } else if (docId) {
      const metadata = JSON.parse(app.db.get(app.DOC_METADATA_KEY, '{}'));
      const key = `${docId}_${gid}`;
      const existingName = metadata[key]?.name;

      // Attempt to load from cache immediately for better UX
      const cached = app.db.get(app.DOC_CACHE_PREFIX + key);
      if (cached && !app.parsed) {
        app.trackRecentDoc(docId, existingName, gid); // Update access timestamp

        if (existingName) app.h1.textContent = existingName;

        if (cached.startsWith('[[')) {
          // Cached JSON grid from Sheets API
          app.buildItineraryFromRows(JSON.parse(cached), true);
        } else {
          // Legacy CSV cache
          app.buildItinerary(cached, true);
        }
        if (app.btnClearDoc) app.btnClearDoc.style.display = 'block';
        app.enableSheetSync(docId, gid);
      }

      if (!app.googleClientId) {
        if (!cached) app.showError('A Google Sheet was linked, but authentication is required. Please sign in.');
        return;
      }

      // If we don't have an access token yet, we stop here. 
      // The auth callbacks will trigger loadFromUrl() again once a token is acquired.
      if (!app.accessToken) return;

      app.fetchSheetData(docId, existingName || 'Linked Spreadsheet', gid, true);
    } else {
      app.resetApp(true);
    }
  };
}
