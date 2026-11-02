// Trip builder and editor.
export function installBuilder(app) {
  app.blankDay = function blankDay(partial = {}) {
    const activities = partial.activities && partial.activities.length
      ? partial.activities.map(a => ({ time: a.time || '', content: a.content || '' }))
      : [{ time: '09:00', content: '' }];
    return {
      date: partial.date || app.todayYmd(),
      location: partial.location || '',
      hotel: partial.hotel || '',
      activities,
      notes: partial.notes || ''
    };
  };

  app.placeholderDay = function placeholderDay(date) {
    return app.blankDay({
      date: date || '',
      location: '',
      hotel: '',
      activities: [{ time: '', content: '' }],
      notes: ''
    });
  };

  app.earliestBuilderDate = function earliestBuilderDate(days) {
    let best = '';
    (days || []).forEach(day => {
      if (day.date && (!best || day.date < best)) best = day.date;
    });
    return best;
  };

  app.tripDateSpan = function tripDateSpan(days) {
    const dates = (days || []).map(day => day.date).filter(Boolean).sort();
    if (!dates.length) return null;
    return { start: dates[0], end: dates[dates.length - 1] };
  };

  app.rangePick = { anchor: '', hover: '', year: 0, month: 0 };

  app.paintSchedule = function paintSchedule(days) {
    const list = days || app.collectBuilderDays();
    const span = app.tripDateSpan(list);
    const rangeEl = document.getElementById('schedule-range');
    const btn = document.getElementById('schedule-range-btn');
    const count = document.getElementById('schedule-count');
    const label = !span
      ? 'Add dates'
      : (span.start === span.end
        ? app.formatTripDate(span.start)
        : `${app.formatTripDate(span.start)} – ${app.formatTripDate(span.end)}`);
    if (rangeEl) rangeEl.textContent = label;
    if (btn) {
      btn.disabled = !list.length;
      btn.setAttribute('aria-label', span ? `Change trip dates, ${label}` : 'Set trip dates');
    }
    if (count) {
      const n = list.length;
      count.textContent = n ? (n === 1 ? '1 day' : `${n} days`) : '';
    }
  };

  app.dayHasContent = function dayHasContent(day) {
    if ((day.location || '').trim() || (day.hotel || '').trim() || (day.notes || '').trim()) return true;
    return (day.activities || []).some(act => (act.content || '').trim());
  };

  app.cloneBuilderDay = function cloneBuilderDay(day) {
    return {
      ...day,
      activities: (day.activities || []).map(act => ({ ...act }))
    };
  };

  app.daysForRange = function daysForRange(days, start, end) {
    if (end < start) { const swap = start; start = end; end = swap; }
    const used = new Set();
    const next = [];
    let cursor = start;
    while (cursor <= end && next.length <= 366) {
      const match = days.findIndex((day, index) => day.date === cursor && !used.has(index));
      if (match >= 0) {
        used.add(match);
        next.push(app.cloneBuilderDay(days[match]));
      } else {
        next.push(app.blankDay({ date: cursor, location: '' }));
      }
      cursor = app.addDaysYmd(cursor, 1);
    }
    return {
      next,
      removed: days.filter((_, index) => !used.has(index)),
      start,
      end
    };
  };

  app.confirmDiscardDays = function confirmDiscardDays(lost) {
    const modal = document.getElementById('builder-confirm-modal');
    const title = document.getElementById('builder-confirm-title');
    const copy = document.getElementById('builder-confirm-copy');
    const okBtn = document.getElementById('builder-confirm-ok');
    const cancelBtn = document.getElementById('builder-confirm-cancel');
    if (!modal || !title || !copy || !okBtn || !cancelBtn) return Promise.resolve(false);
    const names = lost.slice(0, 3).map(day => {
      const when = day.date ? app.formatTripDate(day.date) : 'An undated day';
      return day.location ? `${when} (${day.location})` : when;
    });
    const extra = lost.length > 3 ? `, and ${lost.length - 3} more` : '';
    title.textContent = lost.length === 1 ? 'Remove this day?' : `Remove ${lost.length} days?`;
    copy.textContent = `${names.join(', ')}${extra} ${lost.length === 1 ? 'has' : 'have'} plans that won't fit the new dates.`;
    modal.classList.add('open');
    modal.setAttribute('aria-hidden', 'false');
    okBtn.focus();
    return new Promise(resolve => {
      const finish = (value) => {
        modal.classList.remove('open');
        modal.setAttribute('aria-hidden', 'true');
        modal.removeEventListener('click', onBackdrop);
        document.removeEventListener('keydown', onKey);
        okBtn.onclick = null;
        cancelBtn.onclick = null;
        resolve(value);
      };
      const onBackdrop = (event) => {
        if (event.target === modal) finish(false);
      };
      const onKey = (event) => {
        if (event.key === 'Escape') finish(false);
      };
      okBtn.onclick = () => finish(true);
      cancelBtn.onclick = () => finish(false);
      modal.addEventListener('click', onBackdrop);
      document.addEventListener('keydown', onKey);
    });
  };

  app.tripRangeBounds = function tripRangeBounds() {
    const pick = app.rangePick;
    if (pick && pick.anchor) {
      const end = pick.hover || pick.anchor;
      return pick.anchor <= end ? { start: pick.anchor, end } : { start: end, end: pick.anchor };
    }
    return app.tripDateSpan(app.collectBuilderDays());
  };

  app.renderTripRange = function renderTripRange() {
    const grid = document.getElementById('schedule-grid');
    const label = document.getElementById('schedule-month');
    const hint = document.getElementById('schedule-hint');
    const pick = app.rangePick;
    if (!grid || !label || !pick) return;
    label.textContent = new Date(pick.year, pick.month, 1).toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
    if (hint) hint.textContent = pick.anchor ? 'Choose an end date' : 'Choose a start date';
    grid.innerHTML = '';
    const firstDow = new Date(pick.year, pick.month, 1).getDay();
    const daysInMonth = new Date(pick.year, pick.month + 1, 0).getDate();
    for (let i = 0; i < firstDow; i++) {
      const pad = document.createElement('span');
      pad.className = 'range-pad';
      grid.appendChild(pad);
    }
    const today = app.todayYmd();
    for (let day = 1; day <= daysInMonth; day++) {
      const ymd = `${pick.year}-${String(pick.month + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'range-day';
      btn.dataset.ymd = ymd;
      if (ymd === today) btn.classList.add('is-today');
      const mark = document.createElement('span');
      mark.textContent = String(day);
      btn.appendChild(mark);
      grid.appendChild(btn);
    }
    app.paintTripRangeSelection();
  };

  app.paintTripRangeSelection = function paintTripRangeSelection() {
    const bounds = app.tripRangeBounds();
    document.querySelectorAll('#schedule-grid [data-ymd]').forEach(btn => {
      const ymd = btn.dataset.ymd;
      const inRange = !!(bounds && ymd >= bounds.start && ymd <= bounds.end);
      const edge = !!(bounds && (ymd === bounds.start || ymd === bounds.end));
      const multi = !!(bounds && bounds.start !== bounds.end);
      btn.classList.toggle('in-range', !!(inRange && multi));
      btn.classList.toggle('is-edge', edge);
      btn.classList.toggle('is-start', !!(multi && ymd === bounds.start));
      btn.classList.toggle('is-end', !!(multi && ymd === bounds.end));
      btn.setAttribute('aria-pressed', edge ? 'true' : 'false');
    });
  };

  app.openTripRange = function openTripRange() {
    const pop = document.getElementById('schedule-popover');
    const btn = document.getElementById('schedule-range-btn');
    const card = document.getElementById('builder-schedule');
    if (!pop || !btn) return;
    const span = app.tripDateSpan(app.collectBuilderDays());
    const ymd = (span && span.start) || app.todayYmd();
    const [year, month] = ymd.split('-').map(Number);
    app.rangePick = { anchor: '', hover: '', year, month: month - 1 };
    pop.hidden = false;
    btn.setAttribute('aria-expanded', 'true');
    card?.classList.add('is-picking');
    app.renderTripRange();
  };

  app.closeTripRange = function closeTripRange() {
    const pop = document.getElementById('schedule-popover');
    const btn = document.getElementById('schedule-range-btn');
    const card = document.getElementById('builder-schedule');
    if (pop) pop.hidden = true;
    if (btn) btn.setAttribute('aria-expanded', 'false');
    card?.classList.remove('is-picking');
    if (app.rangePick) {
      app.rangePick.anchor = '';
      app.rangePick.hover = '';
    }
  };

  app.shiftTripRangeMonth = function shiftTripRangeMonth(delta) {
    const pick = app.rangePick;
    if (!pick) return;
    const next = new Date(pick.year, pick.month + delta, 1);
    pick.year = next.getFullYear();
    pick.month = next.getMonth();
    app.renderTripRange();
  };

  app.pickTripRangeDay = async function pickTripRangeDay(ymd) {
    const pick = app.rangePick;
    if (!pick || !ymd) return;
    if (!pick.anchor) {
      pick.anchor = ymd;
      pick.hover = '';
      app.renderTripRange();
      return;
    }
    let start = pick.anchor;
    let end = ymd;
    if (end < start) { const swap = start; start = end; end = swap; }
    const length = app.daysBetweenYmd(start, end) + 1;
    if (length > 366) {
      app.showError('Choose a range of a year or less.');
      return;
    }
    const days = app.collectBuilderDays();
    const plan = app.daysForRange(days, start, end);
    const lost = plan.removed.filter(app.dayHasContent);
    if (lost.length && !(await app.confirmDiscardDays(lost))) {
      pick.anchor = '';
      pick.hover = '';
      app.renderTripRange();
      return;
    }
    app.closeTripRange();
    app.showError('');
    app.commitBuilderDays(plan.next);
  };

  app.insertBuilderDay = function insertBuilderDay(days, index) {
    const next = days.map(day => ({ ...day }));
    if (index < 0) {
      const first = next[0];
      const earliest = app.earliestBuilderDate(next);
      const date = first && first.date
        ? app.addDaysYmd(first.date, -1)
        : (earliest ? app.addDaysYmd(earliest, -1) : app.todayYmd());
      next.unshift(app.blankDay({
        date,
        location: first ? (first.location || '') : ''
      }));
      return next;
    }
    const prev = next[index];
    const date = prev && prev.date
      ? app.addDaysYmd(prev.date, 1)
      : (next[index + 1] && next[index + 1].date ? next[index + 1].date : app.todayYmd());
    const collide = next.slice(index + 1).some(day => day.date && day.date <= date);
    if (collide) {
      for (let i = index + 1; i < next.length; i++) {
        if (next[i].date) next[i].date = app.addDaysYmd(next[i].date, 1);
      }
    }
    next.splice(index + 1, 0, app.blankDay({
      date,
      location: prev ? (prev.location || '') : ''
    }));
    return next;
  };

  app.shiftDatedDays = function shiftDatedDays(days, delta) {
    if (!delta) return days.map(day => ({ ...day }));
    return days.map(day => (
      day.date ? { ...day, date: app.addDaysYmd(day.date, delta) } : { ...day }
    ));
  };

  app.applyDayRemoval = function applyDayRemoval(days, index, mode) {
    if (mode === 'placeholder') {
      const next = days.map(day => ({ ...day }));
      next[index] = app.placeholderDay(days[index].date);
      return next;
    }
    const next = days.filter((_, i) => i !== index).map(day => ({ ...day }));
    for (let i = index; i < next.length; i++) {
      if (next[i].date) next[i].date = app.addDaysYmd(next[i].date, -1);
    }
    if (!next.length) next.push(app.blankDay());
    return next;
  };

  app.askRemoveDay = function askRemoveDay(day, index) {
    const modal = document.getElementById('builder-remove-modal');
    const title = document.getElementById('builder-remove-title');
    const copy = document.getElementById('builder-remove-copy');
    const placeholderBtn = document.getElementById('builder-remove-placeholder');
    const shiftBtn = document.getElementById('builder-remove-shift');
    const cancelBtn = document.getElementById('builder-remove-cancel');
    if (!modal || !title || !copy || !placeholderBtn || !shiftBtn || !cancelBtn) return Promise.resolve(null);
    const when = day.date ? app.formatTripDate(day.date) : 'No date';
    const where = day.location ? `, ${day.location}` : '';
    title.textContent = `Remove day ${index + 1}?`;
    copy.textContent = `${when}${where}. Keep this date empty, or delete it and move every later date back by one.`;
    modal.classList.add('open');
    modal.setAttribute('aria-hidden', 'false');
    placeholderBtn.focus();
    return new Promise(resolve => {
      const finish = (value) => {
        modal.classList.remove('open');
        modal.setAttribute('aria-hidden', 'true');
        modal.removeEventListener('click', onBackdrop);
        document.removeEventListener('keydown', onKey);
        placeholderBtn.onclick = null;
        shiftBtn.onclick = null;
        cancelBtn.onclick = null;
        resolve(value);
      };
      const onBackdrop = (event) => {
        if (event.target === modal) finish(null);
      };
      const onKey = (event) => {
        if (event.key === 'Escape') finish(null);
      };
      placeholderBtn.onclick = () => finish('placeholder');
      shiftBtn.onclick = () => finish('shift');
      cancelBtn.onclick = () => finish(null);
      modal.addEventListener('click', onBackdrop);
      document.addEventListener('keydown', onKey);
    });
  };

  app.commitBuilderDays = function commitBuilderDays(days) {
    app.builderReady = true;
    app.showError('');
    app.renderBuilderDays(days);
  };

  app.setBuilderMode = function setBuilderMode(mode) {
    app.builderEditing = mode === 'edit';
    const applyBtn = document.getElementById('btn-builder-apply');
    const cancelBtn = document.getElementById('btn-builder-cancel');
    const note = document.getElementById('builder-sheet-note');
    const heading = document.querySelector('.input-section h2');
    if (applyBtn) applyBtn.textContent = app.builderEditing ? 'Save changes' : 'Create itinerary';
    if (cancelBtn) cancelBtn.hidden = !app.builderEditing;
    if (note) note.hidden = !(app.builderEditing && app.lastSyncedDocId);
    if (heading) {
      const active = document.querySelector('.tab.active')?.dataset.tab;
      heading.textContent = app.builderEditing ? 'Edit trip' : (active === 'build' ? 'Build a trip' : 'Import');
    }
  };

  app.collectBuilderDays = function collectBuilderDays() {
    return Array.from(document.querySelectorAll('#builder-days .builder-day')).map(card => ({
      date: card.querySelector('[data-field="date"]').value,
      location: card.querySelector('[data-field="location"]').value.trim(),
      hotel: card.querySelector('[data-field="hotel"]').value.trim(),
      activities: Array.from(card.querySelectorAll('.builder-activity')).map(row => ({
        time: row.querySelector('[data-field="time"]').value,
        content: row.querySelector('[data-field="content"]').value.trim()
      })),
      notes: card.querySelector('[data-field="notes"]').value
    }));
  };

  app.renderBuilderDays = function renderBuilderDays(days) {
    const daysEl = document.getElementById('builder-days');
    if (!daysEl) return;
    daysEl.innerHTML = '';
    daysEl.appendChild(app.dayInsertButton(-1, 'Insert a day at the start'));
    days.forEach((day, index) => {
      const card = document.createElement('div');
      card.className = 'builder-day';
      card.dataset.index = String(index);

      const head = document.createElement('div');
      head.className = 'builder-day-head';
      const handle = document.createElement('button');
      handle.type = 'button';
      handle.className = 'day-drag';
      handle.setAttribute('aria-label', `Drag day ${index + 1} to reorder`);
      handle.innerHTML = '<svg width="14" height="18" viewBox="0 0 14 18" aria-hidden="true" fill="currentColor"><circle cx="4" cy="3.5" r="1.25"/><circle cx="10" cy="3.5" r="1.25"/><circle cx="4" cy="9" r="1.25"/><circle cx="10" cy="9" r="1.25"/><circle cx="4" cy="14.5" r="1.25"/><circle cx="10" cy="14.5" r="1.25"/></svg>';
      const title = document.createElement('div');
      title.className = 'builder-day-title';
      const kicker = document.createElement('span');
      kicker.className = 'day-kicker';
      kicker.textContent = `Day ${index + 1}`;
      const when = document.createElement('span');
      when.className = 'day-when';
      when.textContent = day.date ? app.formatTripDate(day.date) : 'No date';
      title.append(kicker, when);
      const tools = document.createElement('div');
      tools.className = 'builder-day-tools';
      const removeBtn = app.miniButton('', 'remove-day', false, true);
      removeBtn.classList.add('day-delete');
      removeBtn.setAttribute('aria-label', 'Remove day');
      removeBtn.innerHTML = '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M3 6h18"/><path d="M8 6V4h8v2"/><path d="M19 6l-1 14H6L5 6"/><path d="M10 11v6"/><path d="M14 11v6"/></svg>';
      tools.append(removeBtn);
      head.append(handle, title, tools);
      app.bindDayDrag(card, handle);

      const grid = document.createElement('div');
      grid.className = 'builder-grid';
      grid.append(
        app.builderField('Date', 'date', 'date', day.date || ''),
        app.builderField('Location', 'location', 'text', day.location || '', 'City or place'),
        app.builderField('Hotel', 'hotel', 'text', day.hotel || '', 'Optional stay', true)
      );

      const activities = document.createElement('div');
      activities.className = 'builder-activities';
      const list = day.activities && day.activities.length ? day.activities : [{ time: '', content: '' }];
      list.forEach((act, actIndex) => activities.appendChild(app.activityRow(act, actIndex, list.length === 1)));

      const addAct = app.miniButton('Add activity', 'add-activity');
      const notesField = app.builderField('Notes', 'notes', 'textarea', day.notes || '', 'Optional', true);

      card.append(head, grid, activities, addAct, notesField);
      daysEl.appendChild(card);
      const atEnd = index === days.length - 1;
      daysEl.appendChild(app.dayInsertButton(
        index,
        atEnd ? 'Add a day at the end' : `Insert a day after day ${index + 1}`
      ));
    });
    app.paintSchedule(days);
  };

  app.moveBuilderDay = function moveBuilderDay(from, to) {
    if (from === to || from < 0 || to < 0) return;
    const days = app.collectBuilderDays();
    if (to >= days.length) return;
    const [item] = days.splice(from, 1);
    days.splice(to, 0, item);
    app.commitBuilderDays(days);
    document.querySelector(`#builder-days .builder-day[data-index="${to}"] .day-drag`)?.focus();
  };

  app.bindDayDrag = function bindDayDrag(card, handle) {
    handle.addEventListener('pointerdown', (event) => {
      if (event.button !== 0) return;
      const daysEl = card.parentElement;
      if (!daysEl || daysEl.querySelectorAll('.builder-day').length < 2) return;
      event.preventDefault();
      try { handle.setPointerCapture(event.pointerId); } catch (err) { /* untrusted pointers still receive events on the handle */ }
      card.classList.add('is-dragging');
      const dayOrder = () => [...daysEl.querySelectorAll('.builder-day')].map(el => el.dataset.index).join(',');
      const inserts = [...daysEl.querySelectorAll('.day-insert')];
      inserts.forEach(el => { el.hidden = true; });
      const startOrder = dayOrder();

      const onMove = (ev) => {
        const siblings = [...daysEl.querySelectorAll('.builder-day')].filter(el => el !== card);
        const next = siblings.find(el => {
          const box = el.getBoundingClientRect();
          return ev.clientY < box.top + box.height / 2;
        });
        if (next) daysEl.insertBefore(card, next);
        else daysEl.appendChild(card);
      };
      const onUp = () => {
        handle.removeEventListener('pointermove', onMove);
        handle.removeEventListener('pointerup', onUp);
        handle.removeEventListener('pointercancel', onUp);
        card.classList.remove('is-dragging');
        const order = dayOrder();
        if (order === startOrder) {
          inserts.forEach(el => { el.hidden = false; });
          return;
        }
        const movedTo = [...daysEl.querySelectorAll('.builder-day')].indexOf(card);
        const days = app.collectBuilderDays();
        app.commitBuilderDays(days);
        document.querySelector(`#builder-days .builder-day[data-index="${movedTo}"] .day-drag`)?.focus();
      };
      handle.addEventListener('pointermove', onMove);
      handle.addEventListener('pointerup', onUp);
      handle.addEventListener('pointercancel', onUp);
    });

    handle.addEventListener('keydown', (event) => {
      if (event.key !== 'ArrowUp' && event.key !== 'ArrowDown') return;
      const daysEl = card.parentElement;
      if (!daysEl) return;
      const cards = [...daysEl.querySelectorAll('.builder-day')];
      const index = cards.indexOf(card);
      const next = event.key === 'ArrowUp' ? index - 1 : index + 1;
      if (next < 0 || next >= cards.length) return;
      event.preventDefault();
      app.moveBuilderDay(index, next);
    });
  };

  app.dayInsertButton = function dayInsertButton(index, label) {
    const insert = document.createElement('button');
    insert.type = 'button';
    insert.className = 'day-insert';
    insert.dataset.insertAfter = String(index);
    insert.setAttribute('aria-label', label);
    insert.innerHTML = '<svg width="22" height="22" viewBox="0 0 22 22" aria-hidden="true"><path d="M11 6.5v9M6.5 11h9" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round"/></svg>';
    return insert;
  };

  app.miniButton = function miniButton(label, action, disabled = false, danger = false) {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = danger ? 'builder-mini danger' : 'builder-mini';
    btn.dataset.action = action;
    btn.textContent = label;
    btn.disabled = disabled;
    return btn;
  };

  app.builderField = function builderField(labelText, field, type, value, placeholder, span) {
    const wrap = document.createElement('label');
    wrap.className = span ? 'builder-field builder-span' : 'builder-field';
    const label = document.createElement('span');
    label.className = 'builder-label';
    label.textContent = labelText;
    let input;
    if (type === 'textarea') {
      input = document.createElement('textarea');
      input.value = value || '';
    } else {
      input = document.createElement('input');
      input.type = type;
      input.value = value || '';
    }
    input.dataset.field = field;
    if (placeholder) input.placeholder = placeholder;
    if (type === 'text' || type === 'textarea') input.autocomplete = 'off';
    wrap.append(label, input);
    return wrap;
  };

  app.activityRow = function activityRow(act, actIndex, onlyOne) {
    const row = document.createElement('div');
    row.className = 'builder-activity';
    const time = document.createElement('input');
    time.type = 'time';
    time.dataset.field = 'time';
    time.value = act.time || '';
    time.setAttribute('aria-label', 'Activity time');
    const content = document.createElement('input');
    content.type = 'text';
    content.dataset.field = 'content';
    content.value = act.content || '';
    content.placeholder = 'Activity';
    content.autocomplete = 'off';
    content.setAttribute('aria-label', 'Activity');
    const remove = app.miniButton('Remove', 'remove-activity', onlyOne, true);
    remove.dataset.index = String(actIndex);
    remove.setAttribute('aria-label', 'Remove activity');
    row.append(time, content, remove);
    return row;
  };

  app.ensureBuilder = function ensureBuilder() {
    if (app.builderReady) return;
    app.builderReady = true;
    app.renderBuilderDays([app.blankDay()]);
    app.setBuilderMode('create');
  };

  app.showBuildPanel = function showBuildPanel() {
    app.tabs.forEach(t => t.classList.remove('active'));
    app.panels.forEach(p => p.classList.remove('active'));
    document.querySelector('.tab[data-tab="build"]').classList.add('active');
    document.getElementById('panel-build').classList.add('active');
  };

  app.parsedToDays = function parsedToDays(model) {
    const days = [];
    for (let i = 1; i < model.locations.length; i++) {
      const location = app.getTextContent(model.locations[i] || '').trim();
      const dateLabel = app.getTextContent(model.dates[i] || '').trim();
      if (!location && !dateLabel) continue;
      let lastTimed = -1;
      model.rows.forEach((row, idx) => {
        const content = app.getTextContent(row.cells[i - 1] || '').trim();
        if (content && (row.time || '').trim()) lastTimed = idx;
      });
      const activities = [];
      const noteLines = [];
      model.rows.forEach((row, idx) => {
        const content = app.getTextContent(row.cells[i - 1] || '').trim();
        if (!content) return;
        const time = (row.time || '').trim();
        if (idx <= lastTimed && time) activities.push({ time: app.timeToInputValue(time), content });
        else noteLines.push(content);
      });
      days.push({
        date: app.parseSmartDate(dateLabel) || '',
        location,
        hotel: model.hotelCells ? app.getTextContent(model.hotelCells[i] || '').trim() : '',
        activities: activities.length ? activities : [{ time: '', content: '' }],
        notes: noteLines.join('\n')
      });
    }
    return days.length ? days : [app.blankDay()];
  };

  app.daysToRows = function daysToRows(days) {
    const perDay = days.map(day => {
      const notes = (day.notes || '').split('\n').map(line => line.trim()).filter(Boolean);
      const activities = [];
      (day.activities || []).forEach(act => {
        const content = (act.content || '').trim();
        if (!content) return;
        if (!act.time) {
          notes.push(content);
          return;
        }
        activities.push({ time: app.formatTimeLabel(act.time), sort: app.timeToMinutes(act.time), content });
      });
      activities.sort((a, b) => a.sort - b.sort);
      return {
        location: day.location.trim(),
        date: app.formatTripDate(day.date),
        hotel: (day.hotel || '').trim(),
        activities,
        notes
      };
    });

    const timeRows = [];
    perDay.forEach((day, dayIndex) => {
      day.activities.forEach(act => {
        const open = timeRows.find(row => row.time === act.time && !row.cells[dayIndex]);
        if (open) {
          open.cells[dayIndex] = act.content;
          return;
        }
        const cells = Array(perDay.length).fill('');
        cells[dayIndex] = act.content;
        const row = { time: act.time, sort: act.sort, cells };
        const insertAt = timeRows.findIndex(existing => existing.sort > act.sort);
        if (insertAt === -1) timeRows.push(row);
        else timeRows.splice(insertAt, 0, row);
      });
    });

    const grid = [
      ['Time', ...perDay.map(day => day.location)],
      ['', ...perDay.map(day => day.date)],
      ['Hotel', ...perDay.map(day => day.hotel)],
      ...timeRows.map(row => [row.time, ...row.cells])
    ];
    const maxNotes = perDay.reduce((max, day) => Math.max(max, day.notes.length), 0);
    for (let n = 0; n < maxNotes; n++) {
      grid.push(['', ...perDay.map(day => day.notes[n] || '')]);
    }
    return grid;
  };

  app.parkBuilderPanel = function parkBuilderPanel() {
    const host = document.getElementById('itinerary-builder');
    if (host) host.classList.remove('visible');
    const panel = document.getElementById('panel-build');
    const error = document.getElementById('error-msg');
    if (panel && error && panel.parentElement !== error.parentElement) {
      error.parentElement.insertBefore(panel, error);
    }
    if (panel && !document.querySelector('.tab[data-tab="build"].active')) {
      panel.classList.remove('active');
    }
  };

  app.enterBuilderMode = function enterBuilderMode() {
    if (!app.parsed) return;
    const host = document.getElementById('itinerary-builder');
    const panel = document.getElementById('panel-build');
    host.appendChild(panel);
    panel.classList.add('active');
    host.classList.add('visible');
    const titleInput = document.getElementById('builder-title');
    if (titleInput) titleInput.value = app.h1.textContent === 'Travel Itinerary' ? '' : app.h1.textContent;
    app.renderBuilderDays(app.parsedToDays(app.parsed));
    app.builderReady = true;
    app.setBuilderMode('edit');
    app.showError('');
  };

  app.dayHasDetails = function dayHasDetails(day) {
    return !!(day.location || day.hotel || (day.notes || '').trim() || (day.activities || []).some(act => (act.content || '').trim()));
  };

  app.applyBuilder = function applyBuilder() {
    const all = app.collectBuilderDays();
    const incomplete = all.filter(day => app.dayHasDetails(day) && !(day.location && day.date));
    if (incomplete.length) {
      app.showError('Each day needs both a date and a location.');
      return;
    }
    const days = all.filter(day => day.location && day.date);
    if (!days.length) {
      app.showError('Add at least one day with a date and a location.');
      return;
    }
    const title = (document.getElementById('builder-title').value || '').trim() || 'Travel Itinerary';
    app.h1.textContent = title;
    app.detachSheetSync();
    const rows = app.daysToRows(days);
    app.buildItineraryFromRows(rows, false, app.rowsToCsv(rows));
  };

  app.cancelEditor = function cancelEditor() {
    app.parkBuilderPanel();
    app.setBuilderMode('create');
    app.showError('');
    const preferred = window.innerWidth < 768
      ? document.querySelector('.view-btn[data-view="list"]')
      : document.querySelector('.view-btn[data-view="calendar"]');
    preferred?.click();
  };
}
