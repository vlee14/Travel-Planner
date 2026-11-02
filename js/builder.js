// Trip builder and editor.
export function installBuilder(app) {
  app.blankDay = function blankDay(partial = {}) {
    const activities = partial.activities && partial.activities.length
      ? partial.activities.map(a => ({ time: a.time || '', content: a.content || '' }))
      : [{ time: '', content: '' }];
    return {
      date: partial.date || app.todayYmd(),
      location: partial.location || '',
      hotel: partial.hotel || '',
      activities,
      notes: partial.notes || ''
    };
  };

  app.placeholderDay = function placeholderDay(date, source = {}) {
    const day = app.blankDay({
      date: date || '',
      location: '',
      hotel: '',
      activities: [{ time: '', content: '' }],
      notes: ''
    });
    if (source.sourceCol != null && source.sourceCol !== '') day.sourceCol = Number(source.sourceCol);
    return day;
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

  app.askConfirm = function askConfirm(titleText, copyText, okText) {
    const modal = document.getElementById('builder-confirm-modal');
    const title = document.getElementById('builder-confirm-title');
    const copy = document.getElementById('builder-confirm-copy');
    const okBtn = document.getElementById('builder-confirm-ok');
    const cancelBtn = document.getElementById('builder-confirm-cancel');
    if (!modal || !title || !copy || !okBtn || !cancelBtn) return Promise.resolve(false);
    title.textContent = titleText;
    copy.textContent = copyText;
    okBtn.textContent = okText || 'Remove';
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
        okBtn.textContent = 'Remove';
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

  app.confirmDiscardDays = function confirmDiscardDays(lost) {
    const names = lost.slice(0, 3).map(day => {
      const when = day.date ? app.formatTripDate(day.date) : 'An undated day';
      return day.location ? `${when} (${day.location})` : when;
    });
    const extra = lost.length > 3 ? `, and ${lost.length - 3} more` : '';
    const titleText = lost.length === 1 ? 'Remove this day?' : `Remove ${lost.length} days?`;
    const copyText = `${names.join(', ')}${extra} ${lost.length === 1 ? 'has' : 'have'} plans that won't fit the new dates.`;
    return app.askConfirm(titleText, copyText, 'Remove');
  };

  app.tripRangeBounds = function tripRangeBounds() {
    const pick = app.rangePick;
    if (pick && pick.anchor) {
      const end = pick.hover || pick.anchor;
      return pick.anchor <= end ? { start: pick.anchor, end } : { start: end, end: pick.anchor };
    }
    if (pick && pick.mode === 'ai') {
      const btn = document.getElementById('ai-trip-range-btn');
      const start = btn?.dataset.start || '';
      const end = btn?.dataset.end || start;
      if (start && end) return { start, end };
      return null;
    }
    return app.tripDateSpan(app.collectBuilderDays());
  };

  app.renderTripRange = function renderTripRange() {
    const grid = document.getElementById('schedule-grid');
    const label = document.getElementById('schedule-month');
    const pick = app.rangePick;
    if (!grid || !label || !pick) return;
    const zoom = pick.zoom || 'day';
    grid.classList.toggle('is-months', zoom === 'month');
    grid.classList.toggle('is-years', zoom === 'year');
    const weekdays = document.querySelector('.schedule-weekdays');
    if (weekdays) weekdays.hidden = zoom !== 'day';
    if (zoom === 'month') app.renderTripMonths();
    else if (zoom === 'year') app.renderTripYears();
    else app.renderTripDays();
    const anchor = pick.target && pick.target.isConnected
      ? pick.target
      : document.getElementById('schedule-range-btn');
    if (anchor) app.placeSchedulePopover(anchor);
  };

  app.renderTripDays = function renderTripDays() {
    const grid = document.getElementById('schedule-grid');
    const label = document.getElementById('schedule-month');
    const hint = document.getElementById('schedule-hint');
    const pick = app.rangePick;
    if (!grid || !label || !pick) return;
    label.textContent = new Date(pick.year, pick.month, 1).toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
    label.disabled = false;
    label.setAttribute('aria-label', `Choose month and year, ${label.textContent}`);
    app.labelTripRangeNav('Previous month', 'Next month');
    if (hint) {
      hint.hidden = false;
      hint.textContent = pick.mode === 'day'
        ? 'Choose a date'
        : (pick.anchor ? 'Choose an end date' : 'Choose a start date');
    }
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

  app.renderTripMonths = function renderTripMonths() {
    const grid = document.getElementById('schedule-grid');
    const label = document.getElementById('schedule-month');
    const hint = document.getElementById('schedule-hint');
    const pick = app.rangePick;
    if (!grid || !label || !pick) return;
    label.textContent = String(pick.year);
    label.disabled = false;
    label.setAttribute('aria-label', `Choose year, ${pick.year}`);
    app.labelTripRangeNav('Previous year', 'Next year');
    if (hint) {
      hint.hidden = false;
      hint.textContent = 'Choose a month';
    }
    grid.innerHTML = '';
    const today = new Date();
    for (let month = 0; month < 12; month++) {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'range-month';
      btn.dataset.month = String(month);
      btn.textContent = new Date(pick.year, month, 1).toLocaleDateString('en-US', { month: 'short' });
      btn.setAttribute('aria-label', new Date(pick.year, month, 1).toLocaleDateString('en-US', { month: 'long', year: 'numeric' }));
      if (month === pick.month) btn.classList.add('is-selected');
      if (pick.year === today.getFullYear() && month === today.getMonth()) btn.classList.add('is-today');
      btn.setAttribute('aria-pressed', month === pick.month ? 'true' : 'false');
      grid.appendChild(btn);
    }
  };

  app.renderTripYears = function renderTripYears() {
    const grid = document.getElementById('schedule-grid');
    const label = document.getElementById('schedule-month');
    const hint = document.getElementById('schedule-hint');
    const pick = app.rangePick;
    if (!grid || !label || !pick) return;
    if (pick.yearPage == null) pick.yearPage = Math.floor(pick.year / 12) * 12;
    const start = pick.yearPage;
    label.textContent = `${start} – ${start + 11}`;
    label.disabled = true;
    label.setAttribute('aria-label', `Years ${start} to ${start + 11}`);
    app.labelTripRangeNav('Previous years', 'Next years');
    if (hint) {
      hint.hidden = false;
      hint.textContent = 'Choose a year';
    }
    grid.innerHTML = '';
    const todayYear = new Date().getFullYear();
    for (let offset = 0; offset < 12; offset++) {
      const year = start + offset;
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'range-year';
      btn.dataset.year = String(year);
      btn.textContent = String(year);
      if (year === pick.year) btn.classList.add('is-selected');
      if (year === todayYear) btn.classList.add('is-today');
      btn.setAttribute('aria-pressed', year === pick.year ? 'true' : 'false');
      grid.appendChild(btn);
    }
  };

  app.labelTripRangeNav = function labelTripRangeNav(previous, next) {
    document.getElementById('schedule-prev')?.setAttribute('aria-label', previous);
    document.getElementById('schedule-next')?.setAttribute('aria-label', next);
  };

  app.zoomTripRange = function zoomTripRange() {
    const pick = app.rangePick;
    if (!pick || pick.zoom === 'year') return;
    if (pick.zoom === 'month') {
      pick.zoom = 'year';
      pick.yearPage = Math.floor(pick.year / 12) * 12;
    } else {
      pick.zoom = 'month';
    }
    app.renderTripRange();
  };

  app.pickTripMonth = function pickTripMonth(month) {
    const pick = app.rangePick;
    if (!pick || month < 0 || month > 11) return;
    pick.month = month;
    pick.zoom = 'day';
    app.renderTripRange();
  };

  app.pickTripYear = function pickTripYear(year) {
    const pick = app.rangePick;
    if (!pick || !year) return;
    pick.year = year;
    pick.zoom = 'month';
    app.renderTripRange();
  };

  app.paintTripRangeSelection = function paintTripRangeSelection() {
    const pick = app.rangePick;
    const selected = pick && pick.mode === 'day' ? pick.anchor : '';
    const bounds = pick && pick.mode === 'day' ? null : app.tripRangeBounds();
    document.querySelectorAll('#schedule-grid [data-ymd]').forEach(btn => {
      const ymd = btn.dataset.ymd;
      const edge = selected
        ? ymd === selected
        : !!(bounds && (ymd === bounds.start || ymd === bounds.end));
      const multi = !!(bounds && bounds.start !== bounds.end);
      const inRange = !!(bounds && ymd >= bounds.start && ymd <= bounds.end);
      btn.classList.toggle('in-range', !!(inRange && multi));
      btn.classList.toggle('is-edge', edge);
      btn.classList.toggle('is-start', !!(multi && ymd === bounds.start));
      btn.classList.toggle('is-end', !!(multi && ymd === bounds.end));
      btn.setAttribute('aria-pressed', edge ? 'true' : 'false');
    });
  };

  app.placeSchedulePopover = function placeSchedulePopover(anchor) {
    const pop = document.getElementById('schedule-popover');
    const card = anchor?.closest('.ai-build-card, #builder-schedule') || document.getElementById('builder-schedule');
    if (!pop || !anchor || !card) return;
    const cardBox = card.getBoundingClientRect();
    const anchorBox = anchor.getBoundingClientRect();
    const margin = 12;
    const viewportCap = window.innerWidth - margin * 2;
    const cardCap = Math.max(160, cardBox.width - 8);
    const compact = cardCap > 26 * 16;
    const width = compact ? Math.min(22 * 16, cardCap, viewportCap) : Math.min(cardCap, viewportCap);
    let left = compact ? anchorBox.left : cardBox.left + 4;
    left = Math.max(margin, Math.min(left, window.innerWidth - margin - width));
    pop.style.width = `${width}px`;
    pop.style.left = `${left}px`;
    pop.style.maxHeight = `${window.innerHeight - margin * 2}px`;
    const height = pop.getBoundingClientRect().height;
    let top = anchorBox.bottom + 8;
    if (top + height > window.innerHeight - margin) {
      const above = anchorBox.top - 8 - height;
      top = above > margin ? above : Math.max(margin, window.innerHeight - margin - height);
    }
    pop.style.top = `${top}px`;
  };

  app.timeParts = function timeParts(hhmm) {
    const match = /^(\d{2}):(\d{2})$/.exec(hhmm || '');
    if (!match) return null;
    const hour = Number(match[1]);
    const minute = Number(match[2]);
    if (hour > 23 || minute > 59) return null;
    return {
      hour,
      minute,
      period: hour >= 12 ? 'pm' : 'am',
      hour12: hour % 12 || 12
    };
  };

  app.composeTime = function composeTime(hour12, minute, period) {
    let hour = hour12 % 12;
    if (period === 'pm') hour += 12;
    return `${String(hour).padStart(2, '0')}:${String(minute).padStart(2, '0')}`;
  };

  app.syncTimeTrigger = function syncTimeTrigger(btn, hhmm) {
    if (!btn) return;
    btn.value = hhmm || '';
    btn.textContent = hhmm ? app.formatTimeLabel(hhmm) : 'Add time';
    btn.classList.toggle('is-empty', !hhmm);
  };

  app.ensureTimePopover = function ensureTimePopover() {
    const hours = document.getElementById('time-hours');
    const minutes = document.getElementById('time-minutes');
    const period = document.getElementById('time-period');
    if (!hours || !minutes || !period || hours.childElementCount) return;
    app.fillTimeWheel(hours, [12, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11].map(hour => ({
      value: String(hour),
      label: String(hour)
    })), 3);
    app.fillTimeWheel(minutes, Array.from({ length: 60 }, (_, minute) => ({
      value: String(minute),
      label: String(minute).padStart(2, '0')
    })), 3);
    app.fillTimeWheel(period, [
      { value: 'am', label: 'AM' },
      { value: 'pm', label: 'PM' }
    ], 1);
  };

  app.fillTimeWheel = function fillTimeWheel(wheel, values, copies) {
    wheel.dataset.count = String(values.length);
    wheel.dataset.copies = String(copies);
    wheel.replaceChildren();
    for (let copy = 0; copy < copies; copy++) {
      values.forEach(item => {
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'time-wheel-item';
        btn.dataset.value = item.value;
        btn.textContent = item.label;
        btn.tabIndex = -1;
        btn.setAttribute('aria-label', item.label);
        btn.addEventListener('click', () => {
          const itemH = btn.offsetHeight || 1;
          const index = [...wheel.children].indexOf(btn);
          wheel.dataset.user = '1';
          wheel.scrollTo({ top: index * itemH, behavior: 'smooth' });
        });
        wheel.appendChild(btn);
      });
    }
    wheel.addEventListener('pointerdown', () => { wheel.dataset.user = '1'; });
    wheel.addEventListener('wheel', () => { wheel.dataset.user = '1'; }, { passive: true });
    wheel.addEventListener('scroll', () => app.onTimeWheelScroll(wheel), { passive: true });
    wheel.addEventListener('scrollend', () => app.settleTimeWheel(wheel));
  };

  app.wheelMetrics = function wheelMetrics(wheel) {
    const item = wheel?.firstElementChild;
    if (!item) return null;
    const itemH = item.offsetHeight || 1;
    const count = Number(wheel.dataset.count) || 1;
    const copies = Number(wheel.dataset.copies) || 1;
    let index = Math.round(wheel.scrollTop / itemH);
    const max = count * copies - 1;
    if (index < 0) index = 0;
    if (index > max) index = max;
    return { itemH, count, copies, index, local: ((index % count) + count) % count };
  };

  app.wheelValue = function wheelValue(wheel) {
    const metrics = app.wheelMetrics(wheel);
    if (!metrics) return '';
    return wheel.children[metrics.local]?.dataset.value || '';
  };

  app.scrollTimeWheel = function scrollTimeWheel(wheel, value) {
    if (!wheel?.firstElementChild) return;
    const count = Number(wheel.dataset.count) || 1;
    const copies = Number(wheel.dataset.copies) || 1;
    const local = [...wheel.children].slice(0, count).findIndex(el => el.dataset.value === value);
    if (local < 0) return;
    const itemH = wheel.firstElementChild.offsetHeight || 1;
    const index = Math.floor(copies / 2) * count + local;
    app.timeWheelLock = (app.timeWheelLock || 0) + 1;
    wheel.scrollTop = index * itemH;
    app.timeWheelLock--;
    clearTimeout(wheel._settle);
  };

  app.frameTimeWheels = function frameTimeWheels(hhmm) {
    const hours = document.getElementById('time-hours');
    const item = hours?.firstElementChild;
    if (!item || item.offsetHeight < 1) {
      const pick = app.timePick;
      if (!pick) return;
      pick.frameTries = (pick.frameTries || 0) + 1;
      if (pick.frameTries > 6) return;
      requestAnimationFrame(() => {
        if (app.timePick === pick) app.frameTimeWheels(hhmm);
      });
      return;
    }
    const parts = app.timeParts(hhmm) || { hour12: 9, minute: 0, period: 'am' };
    app.scrollTimeWheel(hours, String(parts.hour12));
    app.scrollTimeWheel(document.getElementById('time-minutes'), String(parts.minute));
    app.scrollTimeWheel(document.getElementById('time-period'), parts.period);
  };

  app.onTimeWheelScroll = function onTimeWheelScroll(wheel) {
    if (app.timeWheelLock || wheel.dataset.user !== '1') return;
    clearTimeout(wheel._settle);
    wheel._settle = setTimeout(() => app.settleTimeWheel(wheel), 90);
  };

  app.settleTimeWheel = function settleTimeWheel(wheel) {
    if (!wheel || app.timeWheelLock) return;
    const metrics = app.wheelMetrics(wheel);
    if (!metrics) return;
    const middle = Math.floor(metrics.copies / 2) * metrics.count + metrics.local;
    if (metrics.copies > 1 && metrics.index !== middle) {
      app.timeWheelLock = (app.timeWheelLock || 0) + 1;
      wheel.scrollTop = middle * metrics.itemH;
      app.timeWheelLock--;
    }
    if (wheel.dataset.user !== '1') return;
    wheel.dataset.user = '';
    app.commitTimeWheels();
  };

  app.commitTimeWheels = function commitTimeWheels() {
    const target = app.timePick?.target;
    if (!target) return;
    const hour = Number(app.wheelValue(document.getElementById('time-hours')));
    const minute = Number(app.wheelValue(document.getElementById('time-minutes')));
    const period = app.wheelValue(document.getElementById('time-period'));
    if (!hour || Number.isNaN(minute) || (period !== 'am' && period !== 'pm')) return;
    app.syncTimeTrigger(target, app.composeTime(hour, minute, period));
  };

  app.nudgeTimeWheel = function nudgeTimeWheel(key) {
    const ids = ['time-hours', 'time-minutes', 'time-period'];
    if (!app.timePick) return;
    if (app.timePick.column == null) app.timePick.column = 0;
    if (key === 'ArrowLeft') {
      app.timePick.column = (app.timePick.column + ids.length - 1) % ids.length;
      return;
    }
    if (key === 'ArrowRight') {
      app.timePick.column = (app.timePick.column + 1) % ids.length;
      return;
    }
    const wheel = document.getElementById(ids[app.timePick.column]);
    const metrics = app.wheelMetrics(wheel);
    if (!metrics) return;
    const delta = key === 'ArrowDown' ? 1 : -1;
    let next = metrics.local + delta;
    next = metrics.copies === 1
      ? Math.max(0, Math.min(metrics.count - 1, next))
      : (next + metrics.count) % metrics.count;
    const index = Math.floor(metrics.copies / 2) * metrics.count + next;
    wheel.dataset.user = '1';
    wheel.scrollTo({ top: index * metrics.itemH, behavior: 'smooth' });
  };

  app.placeTimePopover = function placeTimePopover(anchor) {
    const pop = document.getElementById('time-popover');
    if (!pop || !anchor) return;
    const box = anchor.getBoundingClientRect();
    const margin = 12;
    const card = anchor.closest('.ai-build-card, #builder-schedule') || document.getElementById('builder-schedule');
    const cardBox = card?.getBoundingClientRect();
    const cardCap = cardBox && cardBox.width > 160 ? cardBox.width - 8 : window.innerWidth - margin * 2;
    const compact = cardCap > 26 * 16;
    const font = parseFloat(getComputedStyle(document.documentElement).fontSize) || 17;
    const width = Math.round(compact
      ? Math.min(16.5 * font, cardCap, window.innerWidth - margin * 2)
      : Math.min(cardCap, window.innerWidth - margin * 2));
    pop.style.width = `${width}px`;
    const height = pop.offsetHeight;
    let left = compact ? box.left : (cardBox ? cardBox.left + 4 : box.left);
    left = Math.max(margin, Math.min(left, window.innerWidth - margin - width));
    let top = box.bottom + 8;
    if (top + height > window.innerHeight - margin) {
      const above = box.top - 8 - height;
      top = above > margin ? above : Math.max(margin, window.innerHeight - margin - height);
    }
    pop.style.left = `${left}px`;
    pop.style.top = `${top}px`;
  };

  app.openTimePicker = function openTimePicker(target) {
    const pop = document.getElementById('time-popover');
    if (!pop || !target) return;
    app.closeTripRange();
    app.ensureTimePopover();
    document.querySelectorAll('.time-trigger[aria-expanded="true"]').forEach(el => {
      if (el !== target) el.setAttribute('aria-expanded', 'false');
    });
    app.timePick = { target, column: 0, frameTries: 0 };
    pop.hidden = false;
    target.setAttribute('aria-expanded', 'true');
    app.placeTimePopover(target);
    app.frameTimeWheels(target.value);
    pop.focus({ preventScroll: true });
  };

  app.closeTimePicker = function closeTimePicker() {
    const pop = document.getElementById('time-popover');
    if (pop) pop.hidden = true;
    document.querySelectorAll('.time-trigger[aria-expanded="true"]').forEach(el => {
      el.setAttribute('aria-expanded', 'false');
    });
    app.timePick = null;
  };

  app.restoreSchedulePopover = function restoreSchedulePopover() {
    const pop = document.getElementById('schedule-popover');
    const home = document.getElementById('builder-schedule');
    if (!pop) return;
    pop.classList.remove('is-overlay');
    if (home && pop.parentElement !== home) home.appendChild(pop);
  };

  app.liftSchedulePopover = function liftSchedulePopover() {
    const pop = document.getElementById('schedule-popover');
    if (!pop) return;
    if (pop.parentElement !== document.body) document.body.appendChild(pop);
    pop.classList.add('is-overlay');
  };

  app.openTripRange = function openTripRange() {
    const pop = document.getElementById('schedule-popover');
    const btn = document.getElementById('schedule-range-btn');
    const card = document.getElementById('builder-schedule');
    if (!pop || !btn) return;
    app.restoreSchedulePopover();
    app.closeTimePicker();
    const span = app.tripDateSpan(app.collectBuilderDays());
    const ymd = (span && span.start) || app.todayYmd();
    const [year, month] = ymd.split('-').map(Number);
    app.rangePick = { mode: 'range', anchor: '', hover: '', year, month: month - 1, zoom: 'day', target: null };
    document.querySelectorAll('.day-date-btn[aria-expanded="true"]').forEach(el => {
      el.setAttribute('aria-expanded', 'false');
    });
    pop.hidden = false;
    btn.setAttribute('aria-expanded', 'true');
    card?.classList.add('is-picking');
    app.renderTripRange();
  };

  app.openDayDate = function openDayDate(btn) {
    const pop = document.getElementById('schedule-popover');
    if (!pop || !btn) return;
    app.restoreSchedulePopover();
    app.closeTimePicker();
    const ymd = btn.value || app.todayYmd();
    const [year, month] = ymd.split('-').map(Number);
    document.querySelectorAll('.day-date-btn[aria-expanded="true"]').forEach(el => {
      if (el !== btn) el.setAttribute('aria-expanded', 'false');
    });
    app.rangePick = { mode: 'day', anchor: btn.value || '', hover: '', year, month: month - 1, zoom: 'day', target: btn };
    pop.hidden = false;
    btn.setAttribute('aria-expanded', 'true');
    document.getElementById('schedule-range-btn')?.setAttribute('aria-expanded', 'false');
    document.getElementById('builder-schedule')?.classList.add('is-picking');
    app.renderTripRange();
  };

  app.closeTripRange = function closeTripRange() {
    const pop = document.getElementById('schedule-popover');
    const btn = document.getElementById('schedule-range-btn');
    const card = document.getElementById('builder-schedule');
    if (pop) pop.hidden = true;
    if (btn) btn.setAttribute('aria-expanded', 'false');
    document.getElementById('ai-trip-range-btn')?.setAttribute('aria-expanded', 'false');
    document.querySelectorAll('.day-date-btn[aria-expanded="true"]').forEach(el => {
      el.setAttribute('aria-expanded', 'false');
    });
    card?.classList.remove('is-picking');
    app.restoreSchedulePopover();
    if (app.rangePick) {
      app.rangePick.mode = '';
      app.rangePick.anchor = '';
      app.rangePick.hover = '';
      app.rangePick.target = null;
    }
  };

  app.openAiTripRange = function openAiTripRange() {
    const pop = document.getElementById('schedule-popover');
    const btn = document.getElementById('ai-trip-range-btn');
    if (!pop || !btn) return;
    const ymd = btn.dataset.start || app.todayYmd();
    const [year, month] = ymd.split('-').map(Number);
    document.querySelectorAll('.day-date-btn[aria-expanded="true"]').forEach(el => {
      el.setAttribute('aria-expanded', 'false');
    });
    document.getElementById('schedule-range-btn')?.setAttribute('aria-expanded', 'false');
    document.getElementById('builder-schedule')?.classList.remove('is-picking');
    app.rangePick = { mode: 'ai', anchor: '', hover: '', year, month: month - 1, zoom: 'day', target: btn };
    app.liftSchedulePopover();
    pop.hidden = false;
    btn.setAttribute('aria-expanded', 'true');
    app.renderTripRange();
  };

  app.shiftTripRangeMonth = function shiftTripRangeMonth(delta) {
    const pick = app.rangePick;
    if (!pick) return;
    if (pick.zoom === 'year') {
      pick.yearPage = (pick.yearPage ?? Math.floor(pick.year / 12) * 12) + delta * 12;
    } else if (pick.zoom === 'month') {
      pick.year += delta;
    } else {
      const next = new Date(pick.year, pick.month + delta, 1);
      pick.year = next.getFullYear();
      pick.month = next.getMonth();
    }
    app.renderTripRange();
  };

  app.pickTripRangeDay = async function pickTripRangeDay(ymd) {
    const pick = app.rangePick;
    if (!pick || !ymd) return;
    if (pick.mode === 'day') {
      const field = pick.target;
      if (field && field.isConnected) {
        field.value = ymd;
        field.textContent = app.formatTripDate(ymd);
        field.classList.remove('is-empty');
        const when = field.closest('.builder-day')?.querySelector('.day-when');
        if (when) when.textContent = app.formatTripDate(ymd);
        app.paintSchedule();
      }
      app.closeTripRange();
      return;
    }
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
    if (pick.mode === 'ai') {
      if (app.AI_TRIP_MAX_DAYS && length > app.AI_TRIP_MAX_DAYS) {
        if (app.setAiStatus) app.setAiStatus('ai-trip-status', app.aiTripLengthMessage(), true);
        const hint = document.getElementById('schedule-hint');
        if (hint) hint.textContent = 'Up to ' + app.AI_TRIP_MAX_DAYS + ' days';
        return;
      }
      app.setAiTripRange(start, end);
      if (app.setAiStatus) app.setAiStatus('ai-trip-status', '');
      app.closeTripRange();
      return;
    }
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
      next[index] = app.placeholderDay(days[index].date, days[index]);
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
    const linked = app.builderEditing && !!(app.lastSyncedDocId && app.linkedSheet);
    if (applyBtn) {
      applyBtn.hidden = linked;
      applyBtn.textContent = app.builderEditing ? 'Save changes' : 'Create itinerary';
    }
    const saveSheetBtn = document.getElementById('btn-builder-save-sheet');
    const forkBtn = document.getElementById('btn-builder-fork');
    if (saveSheetBtn) saveSheetBtn.hidden = !linked;
    if (forkBtn) forkBtn.hidden = !linked;
    if (cancelBtn) cancelBtn.hidden = !app.builderEditing;
    if (note) note.hidden = !linked;
    if (heading) app.syncInputHeading(document.querySelector('.tabs .tab.active')?.dataset.tab);
  };

  app.collectBuilderDays = function collectBuilderDays() {
    return Array.from(document.querySelectorAll('#builder-days .builder-day')).map(card => {
      const sourceCol = card.dataset.sourceCol;
      const day = {
        date: card.querySelector('[data-field="date"]').value,
        location: card.querySelector('[data-field="location"]').value.trim(),
        hotel: card.querySelector('[data-field="hotel"]').value.trim(),
        activities: Array.from(card.querySelectorAll('.builder-activity')).map(row => {
          const sourceRow = row.dataset.sourceRow;
          const act = {
            time: row.querySelector('[data-field="time"]').value,
            content: row.querySelector('[data-field="content"]').value.trim()
          };
          if (sourceRow != null && sourceRow !== '') act.sourceRow = Number(sourceRow);
          return act;
        }),
        notes: card.querySelector('[data-field="notes"]').value
      };
      if (sourceCol != null && sourceCol !== '') day.sourceCol = Number(sourceCol);
      return day;
    });
  };

  app.renderBuilderDays = function renderBuilderDays(days) {
    app.closeTimePicker();
    const daysEl = document.getElementById('builder-days');
    if (!daysEl) return;
    if (app.rangePick?.mode === 'day') app.closeTripRange();
    daysEl.innerHTML = '';
    daysEl.appendChild(app.dayInsertButton(-1, 'Insert a day at the start'));
    days.forEach((day, index) => {
      const card = document.createElement('div');
      card.className = 'builder-day';
      card.dataset.index = String(index);
      if (day.sourceCol != null && day.sourceCol !== '') card.dataset.sourceCol = String(day.sourceCol);

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
      const aiBtn = document.createElement('button');
      aiBtn.type = 'button';
      aiBtn.className = 'ai-choice';
      aiBtn.dataset.action = 'ai-day';
      aiBtn.title = 'Dream up this day';
      aiBtn.setAttribute('aria-label', 'Dream up this day');
      aiBtn.innerHTML = app.icon('sparkle');
      const removeBtn = app.miniButton('', 'remove-day', false, true);
      removeBtn.classList.add('day-delete');
      removeBtn.setAttribute('aria-label', 'Remove day');
      removeBtn.innerHTML = app.icon('trash');
      tools.append(aiBtn, removeBtn);
      head.append(handle, title, tools);
      app.bindDayDrag(card, handle);

      const grid = document.createElement('div');
      grid.className = 'builder-grid';
      grid.append(
        app.builderField('Date', 'date', 'date', day.date || ''),
        app.builderField('Location', 'location', 'text', day.location || '', 'City or place'),
        app.builderField('Hotel', 'hotel', 'text', day.hotel || '', 'Optional stay')
      );

      const activities = document.createElement('div');
      activities.className = 'builder-activities';
      const list = app.withTrailingActivity(day.activities);
      list.forEach((act, actIndex) => {
        const ghost = actIndex === list.length - 1 && !app.activityHasText(act.content);
        activities.appendChild(app.activityRow(act, actIndex, ghost));
      });

      const notesField = app.builderField('Notes', 'notes', 'textarea', day.notes || '', 'Optional', true);

      card.append(head, grid, activities, notesField);
      daysEl.appendChild(card);
      const atEnd = index === days.length - 1;
      daysEl.appendChild(app.dayInsertButton(
        index,
        atEnd ? 'Add a day at the end' : `Insert a day after day ${index + 1}`
      ));
    });
    daysEl.querySelectorAll('.builder-activity textarea').forEach(el => app.fitActivityField(el));
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

  app.builderDateField = function builderDateField(labelText, field, value) {
    const wrap = document.createElement('label');
    wrap.className = 'builder-field';
    const label = document.createElement('span');
    label.className = 'builder-label';
    label.textContent = labelText;
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'date-trigger day-date-btn';
    btn.dataset.field = field;
    btn.value = value || '';
    btn.textContent = value ? app.formatTripDate(value) : 'Add date';
    btn.classList.toggle('is-empty', !value);
    btn.setAttribute('aria-haspopup', 'dialog');
    btn.setAttribute('aria-expanded', 'false');
    btn.setAttribute('aria-controls', 'schedule-popover');
    wrap.append(label, btn);
    return wrap;
  };

  app.builderField = function builderField(labelText, field, type, value, placeholder, span) {
    if (type === 'date') return app.builderDateField(labelText, field, value);
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

  const linkIcon = '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M10 13a5 5 0 0 0 7.54.54l2.12-2.12a5 5 0 0 0-7.07-7.07L11.2 5.7"/><path d="M14 11a5 5 0 0 0-7.54-.54L4.34 12.6a5 5 0 0 0 7.07 7.07l1.39-1.39"/></svg>';
  let activityFieldSerial = 0;

  app.activityIconButton = function activityIconButton(label, className, svg) {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = `builder-mini activity-icon ${className}`;
    btn.setAttribute('aria-label', label);
    btn.title = label;
    btn.innerHTML = svg;
    return btn;
  };

  app.fitActivityField = function fitActivityField(el) {
    el.style.height = 'auto';
    el.style.height = `${el.scrollHeight}px`;
  };

  app.activityHasText = function activityHasText(value) {
    return !!String(value || '').trim();
  };

  app.withTrailingActivity = function withTrailingActivity(activities) {
    const list = activities && activities.length
      ? activities.map(act => ({ ...act }))
      : [{ time: '', content: '' }];
    while (
      list.length > 1 &&
      !app.activityHasText(list[list.length - 1].content) &&
      !app.activityHasText(list[list.length - 2].content)
    ) {
      list.pop();
    }
    if (app.activityHasText(list[list.length - 1].content)) list.push({ time: '', content: '' });
    return list;
  };

  app.syncTrailingActivity = function syncTrailingActivity(card) {
    const list = card.querySelector('.builder-activities');
    if (!list) return;
    const textOf = (row) => row.querySelector('[data-field="content"]')?.value || '';
    const rows = () => [...list.children].filter(el => el.classList.contains('builder-activity'));
    let current = rows();
    while (current.length > 1) {
      const tail = current[current.length - 1];
      const prev = current[current.length - 2];
      if (app.activityHasText(textOf(tail)) || app.activityHasText(textOf(prev))) break;
      if (tail.contains(document.activeElement)) break;
      tail.remove();
      current.pop();
    }
    const tail = current[current.length - 1];
    if (tail && app.activityHasText(textOf(tail))) {
      list.appendChild(app.activityRow({ time: '', content: '' }, current.length, true));
      current.push(list.lastElementChild);
    }
    current.forEach((row, index) => {
      const btn = row.querySelector('[data-action="remove-activity"]');
      if (!btn) return;
      btn.dataset.index = String(index);
      btn.disabled = index === current.length - 1 && !app.activityHasText(textOf(row));
    });
  };

  app.activityRow = function activityRow(act, actIndex, onlyOne) {
    const row = document.createElement('div');
    row.className = 'builder-activity';
    const content = document.createElement('textarea');
    content.dataset.field = 'content';
    content.rows = 1;
    content.value = act.content || '';
    content.placeholder = 'Activity';
    content.autocomplete = 'off';
    content.setAttribute('aria-label', 'Activity');
    content.addEventListener('input', () => {
      app.fitActivityField(content);
      syncLinkState();
      const card = row.closest('.builder-day');
      if (card) app.syncTrailingActivity(card);
    });

    const tools = document.createElement('div');
    tools.className = 'activity-tools';
    const time = document.createElement('button');
    time.type = 'button';
    time.className = 'time-trigger';
    time.dataset.field = 'time';
    time.setAttribute('aria-label', 'Activity time');
    time.setAttribute('aria-haspopup', 'dialog');
    time.setAttribute('aria-expanded', 'false');
    time.setAttribute('aria-controls', 'time-popover');
    app.syncTimeTrigger(time, act.time || '');
    if (act.sourceRow != null && act.sourceRow !== '') row.dataset.sourceRow = String(act.sourceRow);

    const boxId = `activity-linkbox-${++activityFieldSerial}`;
    const linkBtn = app.activityIconButton('Add link', 'activity-link', linkIcon);
    linkBtn.setAttribute('aria-expanded', 'false');
    linkBtn.setAttribute('aria-controls', boxId);
    const remove = app.activityIconButton('Remove activity', 'danger activity-delete', app.icon('trash'));
    remove.dataset.action = 'remove-activity';
    remove.dataset.index = String(actIndex);
    remove.disabled = onlyOne;
    tools.append(linkBtn, remove);

    const linkbox = document.createElement('div');
    linkbox.className = 'activity-linkbox';
    linkbox.id = boxId;
    linkbox.hidden = true;
    const labelInput = document.createElement('input');
    labelInput.type = 'text';
    labelInput.placeholder = 'Link text';
    labelInput.autocomplete = 'off';
    labelInput.setAttribute('aria-label', 'Link text');
    const urlInput = document.createElement('input');
    urlInput.type = 'url';
    urlInput.placeholder = 'https://';
    urlInput.inputMode = 'url';
    urlInput.autocomplete = 'off';
    urlInput.autocapitalize = 'off';
    urlInput.spellcheck = false;
    urlInput.setAttribute('aria-label', 'Link URL');
    const error = document.createElement('p');
    error.className = 'activity-link-error';
    error.hidden = true;
    const actions = document.createElement('div');
    actions.className = 'activity-link-actions';
    const cancelBtn = document.createElement('button');
    cancelBtn.type = 'button';
    cancelBtn.className = 'builder-mini';
    cancelBtn.textContent = 'Cancel';
    const addBtn = document.createElement('button');
    addBtn.type = 'button';
    addBtn.className = 'builder-mini';
    addBtn.dataset.linkAdd = 'true';
    addBtn.textContent = 'Add link';
    actions.append(cancelBtn, addBtn);
    linkbox.append(labelInput, urlInput, error, actions);

    let saved = { start: 0, end: 0 };
    let pendingSelection = null;
    const hasLink = () => /\]\(https?:\/\//i.test(content.value);
    function syncLinkState() {
      linkBtn.classList.toggle('is-on', !linkbox.hidden || hasLink());
    }
    function closeLinkBox() {
      linkbox.hidden = true;
      linkBtn.setAttribute('aria-expanded', 'false');
      error.hidden = true;
      urlInput.removeAttribute('aria-invalid');
      syncLinkState();
    }
    function openLinkBox() {
      saved = pendingSelection || { start: content.value.length, end: content.value.length };
      pendingSelection = null;
      const selected = content.value.slice(saved.start, saved.end).trim();
      labelInput.value = selected;
      urlInput.value = '';
      error.hidden = true;
      urlInput.removeAttribute('aria-invalid');
      linkbox.hidden = false;
      linkBtn.setAttribute('aria-expanded', 'true');
      syncLinkState();
      (selected ? urlInput : labelInput).focus();
    }
    function insertLink() {
      let url = urlInput.value.trim();
      if (url && !/^https?:\/\//i.test(url)) url = `https://${url}`;
      const href = app.safeHttpUrl(url);
      if (!href) {
        error.hidden = false;
        error.textContent = 'Enter a full web address, like https://example.com';
        urlInput.setAttribute('aria-invalid', 'true');
        urlInput.focus();
        return;
      }
      const start = Math.min(saved.start, content.value.length);
      const end = Math.max(start, Math.min(saved.end, content.value.length));
      const selected = content.value.slice(start, end).trim();
      const label = (labelInput.value.trim() || selected || href).replace(/[\[\]]/g, '').replace(/\s+/g, ' ');
      const before = content.value.slice(0, start);
      let lead = '';
      if (start === end && before && !/\s$/.test(before)) {
        lead = start === content.value.length ? '\n' : ' ';
      }
      const markdown = `${lead}[${label}](${href.replace(/[()]/g, match => (match === '(' ? '%28' : '%29'))})`;
      content.value = before + markdown + content.value.slice(end);
      const caret = start + markdown.length;
      closeLinkBox();
      app.fitActivityField(content);
      content.focus();
      content.setSelectionRange(caret, caret);
      const card = row.closest('.builder-day');
      if (card) app.syncTrailingActivity(card);
    }

    linkBtn.addEventListener('pointerdown', () => {
      pendingSelection = document.activeElement === content
        ? { start: content.selectionStart, end: content.selectionEnd }
        : null;
    });
    linkBtn.addEventListener('click', (event) => {
      event.stopPropagation();
      if (linkbox.hidden) openLinkBox();
      else closeLinkBox();
    });
    cancelBtn.addEventListener('click', (event) => {
      event.stopPropagation();
      closeLinkBox();
      content.focus();
    });
    addBtn.addEventListener('click', (event) => {
      event.stopPropagation();
      insertLink();
    });
    urlInput.addEventListener('keydown', (event) => {
      if (event.key === 'Enter') {
        event.preventDefault();
        insertLink();
      } else if (event.key === 'Escape') {
        event.preventDefault();
        closeLinkBox();
        content.focus();
      }
    });
    labelInput.addEventListener('keydown', (event) => {
      if (event.key === 'Enter') {
        event.preventDefault();
        urlInput.focus();
      } else if (event.key === 'Escape') {
        event.preventDefault();
        closeLinkBox();
        content.focus();
      }
    });
    urlInput.addEventListener('input', () => {
      error.hidden = true;
      urlInput.removeAttribute('aria-invalid');
    });

    const line = document.createElement('div');
    line.className = 'activity-line';
    line.append(time, content, tools);
    row.append(line, linkbox);
    queueMicrotask(() => app.fitActivityField(content));
    syncLinkState();
    return row;
  };

  app.ensureBuilder = function ensureBuilder() {
    if (app.builderReady) return;
    app.builderReady = true;
    app.renderBuilderDays([app.blankDay()]);
    app.setBuilderMode('create');
  };

  app.resetBuilder = function resetBuilder() {
    app.closeTripRange();
    if (app.setAiPanel) app.setAiPanel('');
    const titleInput = document.getElementById('builder-title');
    if (titleInput) titleInput.value = '';
    app.commitBuilderDays([app.blankDay()]);
    app.setBuilderMode('create');
  };

  app.inputHeading = function inputHeading(tabId) {
    if (tabId === 'build' && app.builderEditing) return 'Edit trip';
    if (tabId === 'build') return 'Plan your adventure';
    if (tabId === 'drive' || tabId === 'paste') return 'Import';
    return '';
  };

  app.syncInputHeading = function syncInputHeading(tabId) {
    const heading = document.getElementById('input-heading');
    if (!heading) return;
    const text = app.inputHeading(tabId);
    heading.textContent = text;
    heading.hidden = !text;
  };

  app.showInputTab = function showInputTab(id) {
    app.tabs.forEach(t => {
      const on = t.dataset.tab === id;
      t.classList.toggle('active', on);
      t.setAttribute('aria-selected', on ? 'true' : 'false');
    });
    app.panels.forEach(p => {
      p.classList.toggle('active', p.id === 'panel-' + id);
    });
    app.syncInputHeading(id);
    if (id === 'build') app.ensureBuilder();
  };

  app.parsedToDays = function parsedToDays(model) {
    const days = [];
    for (let i = 1; i < model.locations.length; i++) {
      const location = app.getTextContent(model.locations[i] || '').trim();
      const dateLabel = app.getTextContent(model.dates[i] || '').trim();
      if (!location && !dateLabel) continue;
      let lastTimed = -1;
      model.rows.forEach((row, idx) => {
        const content = app.activityEditorText(row.cells[i - 1] || '');
        if (content && (row.time || '').trim()) lastTimed = idx;
      });
      const activities = [];
      const noteLines = [];
      model.rows.forEach((row, idx) => {
        const content = app.activityEditorText(row.cells[i - 1] || '');
        if (!content) return;
        const time = (row.time || '').trim();
        if (idx <= lastTimed && time) {
          const activity = { time: app.timeToInputValue(time), content };
          if (row.gridRow != null) activity.sourceRow = row.gridRow;
          activities.push(activity);
        } else noteLines.push(content);
      });
      const day = {
        date: app.parseSmartDate(dateLabel) || '',
        location,
        hotel: model.hotelCells ? app.getTextContent(model.hotelCells[i] || '').trim() : '',
        activities: activities.length ? activities : [{ time: '', content: '' }],
        notes: noteLines.join('\n'),
        sourceCol: i
      };
      days.push(day);
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

  app.readBuilderDraft = function readBuilderDraft() {
    const all = app.collectBuilderDays();
    const incomplete = all.filter(day => app.dayHasDetails(day) && !(day.location && day.date));
    if (incomplete.length) {
      app.showError('Each day needs both a date and a location.');
      return null;
    }
    const days = all.filter(day => day.location && day.date);
    if (!days.length) {
      app.showError('Add at least one day with a date and a location.');
      return null;
    }
    const title = (document.getElementById('builder-title').value || '').trim() || 'Travel Itinerary';
    return { days, title };
  };

  app.setBuilderBusy = function setBuilderBusy(busy) {
    ['btn-builder-apply', 'btn-builder-save-sheet', 'btn-builder-fork', 'btn-builder-cancel'].forEach(id => {
      const btn = document.getElementById(id);
      if (btn) btn.disabled = busy;
    });
    const saveLabel = document.querySelector('#btn-builder-save-sheet .btn-label');
    if (saveLabel) saveLabel.textContent = busy ? 'Saving…' : 'Save to sheet';
  };

  app.saveBuilderFork = function saveBuilderFork(draft) {
    app.h1.textContent = draft.title;
    app.detachSheetSync();
    const rows = app.daysToRows(draft.days);
    app.buildItineraryFromRows(rows, false, app.rowsToCsv(rows));
  };

  app.applyBuilder = function applyBuilder() {
    const draft = app.readBuilderDraft();
    if (!draft) return;
    app.saveBuilderFork(draft);
  };

  app.saveBuilderToSheet = async function saveBuilderToSheet() {
    const draft = app.readBuilderDraft();
    if (!draft) return;
    if (!app.lastSyncedDocId || !app.linkedSheet) {
      app.saveBuilderFork(draft);
      return;
    }
    app.setBuilderBusy(true);
    app.showError('');
    try {
      app.h1.textContent = draft.title;
      const plan = await app.writeLinkedSheet(draft.days, draft.title);
      const rows = plan.grid;
      const fileId = app.linkedSheet.fileId;
      const gid = app.linkedSheet.gid || '0';
      app.db.set(app.DOC_CACHE_PREFIX + fileId + '_' + gid, JSON.stringify(rows));
      app.trackRecentDoc(fileId, draft.title, gid);
      app.enableSheetSync(fileId, gid);
      app.updateUrlWithDocId(fileId, gid);
      app.buildItineraryFromRows(rows, true);
    } catch (error) {
      console.error(error);
      app.showError(error.message || 'Could not save to Google Sheets.');
    } finally {
      app.setBuilderBusy(false);
    }
  };

  app.setAiPanel = function setAiPanel(which) {
    if (app.rangePick?.mode === 'ai') app.closeTripRange();
    const tripOpen = which === 'trip';
    const dayOpen = which === 'day';
    const tripForm = document.getElementById('ai-trip-form');
    const dayForm = document.getElementById('ai-day-form');
    const tripBtn = document.getElementById('btn-ai-trip');
    const modal = document.getElementById('ai-build-modal');
    const title = document.getElementById('ai-build-title');
    const open = tripOpen || dayOpen;
    if (!open) app.aiOpenDay = null;
    if (tripForm) tripForm.hidden = !tripOpen;
    if (dayForm) dayForm.hidden = !dayOpen;
    if (modal) {
      modal.classList.toggle('open', open);
      modal.setAttribute('aria-hidden', open ? 'false' : 'true');
    }
    if (title) title.textContent = dayOpen ? 'Dream up this day' : 'Dream up a trip';
    if (tripBtn) {
      tripBtn.classList.toggle('is-open', tripOpen);
      tripBtn.setAttribute('aria-expanded', tripOpen ? 'true' : 'false');
    }
    document.querySelectorAll('[data-action="ai-day"]').forEach(btn => {
      const card = btn.closest('.builder-day');
      const on = dayOpen && card && Number(card.dataset.index) === app.aiOpenDay;
      btn.classList.toggle('is-open', on);
      btn.setAttribute('aria-expanded', on ? 'true' : 'false');
    });
  };

  app.paintAiTripRange = function paintAiTripRange() {
    const btn = document.getElementById('ai-trip-range-btn');
    if (!btn) return;
    const start = btn.dataset.start || '';
    const end = btn.dataset.end || '';
    let label = 'Add dates';
    if (start && end && start !== end) label = `${app.formatTripDate(start)} – ${app.formatTripDate(end)}`;
    else if (start) label = app.formatTripDate(start);
    const text = document.getElementById('ai-trip-range');
    if (text) text.textContent = label;
    btn.classList.toggle('is-empty', !start);
    btn.setAttribute('aria-label', start ? `Change dates, ${label}` : 'Set trip dates');
  };

  app.setAiTripRange = function setAiTripRange(start, end) {
    const btn = document.getElementById('ai-trip-range-btn');
    if (!btn) return;
    if (start) btn.dataset.start = start;
    else delete btn.dataset.start;
    if (end) btn.dataset.end = end;
    else delete btn.dataset.end;
    app.paintAiTripRange();
  };

  app.prepareAiTripForm = function prepareAiTripForm() {
    const where = document.getElementById('ai-trip-where');
    const btn = document.getElementById('ai-trip-range-btn');
    const days = app.collectBuilderDays();
    const named = app.destinationNames(days);
    if (where && !where.value.trim() && named.length) where.value = named.join(', ');
    const span = (days.length > 1 || named.length) ? app.tripDateSpan(days) : null;
    if (btn && !btn.dataset.start && span) app.setAiTripRange(span.start, span.end);
    else app.paintAiTripRange();
    if (app.setAiStatus) app.setAiStatus('ai-trip-status', '');
  };

  app.prepareAiDayForm = function prepareAiDayForm(index) {
    const where = document.getElementById('ai-day-where');
    const dateEl = document.getElementById('ai-day-date');
    const day = Number.isInteger(index) ? app.collectBuilderDays()[index] : null;
    const when = day && day.date ? app.formatTripDate(day.date) : '';
    if (dateEl) {
      dateEl.textContent = when || 'No date';
      dateEl.classList.toggle('is-empty', !when);
    }
    if (where) where.value = day ? (day.location || '') : '';
    if (app.setAiStatus) app.setAiStatus('ai-day-status', '');
  };

  app.openAiTripForm = function openAiTripForm() {
    if (app.aiBuildBusy || !app.currentAiKey()) return;
    const form = document.getElementById('ai-trip-form');
    if (form && !form.hidden) {
      app.setAiPanel('');
      return;
    }
    app.ensureBuilder();
    app.closeTripRange();
    app.aiOpenDay = null;
    app.setAiPanel('trip');
    app.prepareAiTripForm();
    document.getElementById('ai-trip-where')?.focus();
  };

  app.openAiDayForm = function openAiDayForm(index) {
    if (app.aiBuildBusy || !app.currentAiKey() || !Number.isInteger(index)) return;
    const form = document.getElementById('ai-day-form');
    if (form && !form.hidden && app.aiOpenDay === index) {
      app.setAiPanel('');
      return;
    }
    app.ensureBuilder();
    app.closeTripRange();
    app.aiOpenDay = index;
    app.setAiPanel('day');
    app.prepareAiDayForm(index);
    const where = document.getElementById('ai-day-where');
    if (where && !where.value.trim()) where.focus();
    else document.getElementById('ai-day-prompt')?.focus();
  };

  app.applyGeneratedTrip = function applyGeneratedTrip(plan) {
    const current = app.collectBuilderDays();
    plan.days.forEach(day => {
      const prev = current.find(item => item.date === day.date && item.sourceCol != null);
      if (prev) day.sourceCol = prev.sourceCol;
    });
    const titleInput = document.getElementById('builder-title');
    if (titleInput && plan.title && !titleInput.value.trim()) titleInput.value = plan.title;
    app.commitBuilderDays(plan.days);
  };

  app.applyGeneratedDay = function applyGeneratedDay(index, planned) {
    const days = app.collectBuilderDays();
    let target = index;
    if (target == null || target < 0 || target >= days.length) {
      if (days.length === 1 && !app.dayHasContent(days[0])) target = 0;
      else target = -1;
    }
    if (days.some((day, i) => i !== target && day.date && day.date === planned.date)) {
      throw new Error('Another day already uses that date.');
    }
    if (target < 0) {
      const at = days.findIndex(day => day.date && planned.date && day.date > planned.date);
      if (at === -1) days.push(planned);
      else days.splice(at, 0, planned);
      target = days.indexOf(planned);
    } else {
      const prev = days[target];
      if (prev.sourceCol != null) planned.sourceCol = prev.sourceCol;
      if (!planned.hotel && prev.hotel) planned.hotel = prev.hotel;
      if (!planned.notes && prev.notes) planned.notes = prev.notes;
      days[target] = planned;
    }
    app.commitBuilderDays(days);
    return target;
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
