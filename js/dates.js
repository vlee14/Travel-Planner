// Date and time parsing for trip days.
export function installDates(app) {
  app.detectItineraryYear = function detectItineraryYear(dates) {
    const now = new Date();
    const currentYear = now.getFullYear();
    if (!dates || !dates.length) return currentYear;

    for (let i = 1; i < dates.length; i++) {
      const s = (dates[i] || '').toString();
      const m = s.match(/\b(202[0-9]|203[0-9])\b/);
      if (m) return parseInt(m[1], 10);
    }

    if (app.h1 && app.h1.textContent) {
      const m = app.h1.textContent.match(/\b(202[0-9]|203[0-9])\b/);
      if (m) return parseInt(m[1], 10);
    }

    for (let i = 1; i < dates.length; i++) {
      const s = (dates[i] || '').toString().trim();
      if (!s) continue;
      const testD = new Date(s);
      if (!isNaN(testD.getTime())) {
        testD.setFullYear(currentYear);
        if ((now.getTime() - testD.getTime()) > 30 * 86400000) {
          return currentYear + 1;
        }
      }
    }

    return currentYear;
  };

  app.parseDateForSort = function parseDateForSort(label, fallbackYear = null) {
    const s = (label || '').toString().trim();
    if (!s) return Infinity;
    const smartYmd = app.parseSmartDate(s, fallbackYear);
    if (smartYmd) {
      const parts = smartYmd.split('-').map(Number);
      return new Date(parts[0], parts[1] - 1, parts[2]).getTime();
    }
    return Infinity;
  };

  app.parseTimeForSort = function parseTimeForSort(timeStr) {
    const s = (timeStr || '').toString().trim();
    const m = s.match(/^(\d{1,2})(?::(\d{2}))?\s*(am|pm)?$/i);
    if (!m) return 0;
    let h = parseInt(m[1], 10);
    const min = m[2] ? parseInt(m[2], 10) : 0;
    const pm = (m[3] || '').toLowerCase() === 'pm';
    const am = (m[3] || '').toLowerCase() === 'am';
    if (pm && h < 12) h += 12;
    if (am && h === 12) h = 0;
    return h * 60 + min;
  };

  app.parseSmartDate = function parseSmartDate(str, fallbackYear = null) {
    const s = (str || '').toString().trim();
    if (!s) return null;
    // Date-only ISO strings are UTC midnight. Read the calendar day directly.
    if (/^\d{4}-\d{2}-\d{2}(?:[T\s].*)?$/.test(s)) return app.normalizeYmd(s) || null;

    const now = new Date();
    const yrMatch = s.match(/\b(202[0-9]|203[0-9])\b/);
    const explicitYear = yrMatch ? parseInt(yrMatch[1], 10) : null;
    const baseYear = explicitYear || fallbackYear || (app.parsed && app.parsed.dates ? app.detectItineraryYear(app.parsed.dates) : now.getFullYear());
    const months = {
      jan: 0, january: 0, feb: 1, february: 1, mar: 2, march: 2,
      apr: 3, april: 3, may: 4, jun: 5, june: 5, jul: 6, july: 6,
      aug: 7, august: 7, sep: 8, sept: 8, september: 8, oct: 9, october: 9,
      nov: 10, november: 10, dec: 11, december: 11
    };

    function asYmd(year, monthIndex, day) {
      const dt = new Date(year, monthIndex, day);
      if (dt.getFullYear() !== year || dt.getMonth() !== monthIndex || dt.getDate() !== day) return null;
      return `${year}-${String(monthIndex + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
    }

    const slashYear = s.match(/^(\d{4})[\/.](\d{1,2})[\/.](\d{1,2})$/);
    if (slashYear) {
      const year = Number(slashYear[1]);
      const useYear = year >= 2020 && year <= 2039 ? year : baseYear;
      return asYmd(useYear, Number(slashYear[2]) - 1, Number(slashYear[3]));
    }

    const body = s.replace(/^(?:sun|mon|tue|wed|thu|fri|sat)[a-z]*\s*,?\s+/i, '');
    const monthFirst = body.match(/\b([A-Za-z]{3,9})\.?\s+(\d{1,2})(?:st|nd|rd|th)?(?!\d)/i);
    if (monthFirst && months[monthFirst[1].toLowerCase()] != null) {
      return asYmd(baseYear, months[monthFirst[1].toLowerCase()], Number(monthFirst[2]));
    }
    const dayFirst = body.match(/\b(\d{1,2})(?:st|nd|rd|th)?\s+([A-Za-z]{3,9})\b/i);
    if (dayFirst && months[dayFirst[2].toLowerCase()] != null) {
      return asYmd(baseYear, months[dayFirst[2].toLowerCase()], Number(dayFirst[1]));
    }

    const numeric = body.match(/^(\d{1,2})[\/.\-](\d{1,2})(?:[\/.\-](\d{4}))?$/);
    if (!numeric) return null;
    return asYmd(baseYear, Number(numeric[1]) - 1, Number(numeric[2]));
  };

  app.todayYmd = function todayYmd() {
    const now = new Date();
    return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
  };

  app.addDaysYmd = function addDaysYmd(ymd, days) {
    const [y, m, d] = (ymd || '').split('-').map(Number);
    if (!y || !m || !d) return app.todayYmd();
    const dt = new Date(y, m - 1, d);
    dt.setDate(dt.getDate() + days);
    return `${dt.getFullYear()}-${String(dt.getMonth() + 1).padStart(2, '0')}-${String(dt.getDate()).padStart(2, '0')}`;
  };

  app.daysBetweenYmd = function daysBetweenYmd(fromYmd, toYmd) {
    const [y1, m1, d1] = (fromYmd || '').split('-').map(Number);
    const [y2, m2, d2] = (toYmd || '').split('-').map(Number);
    if (!y1 || !m1 || !d1 || !y2 || !m2 || !d2) return 0;
    const from = Date.UTC(y1, m1 - 1, d1);
    const to = Date.UTC(y2, m2 - 1, d2);
    return Math.round((to - from) / 86400000);
  };

  app.formatTripDate = function formatTripDate(ymd) {
    const [y, m, d] = (ymd || '').split('-').map(Number);
    if (!y || !m || !d) return '';
    return new Date(y, m - 1, d).toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
  };

  app.formatTimeLabel = function formatTimeLabel(hhmm) {
    const [hStr, mStr] = (hhmm || '').split(':');
    let h = parseInt(hStr, 10);
    if (Number.isNaN(h)) return '';
    const min = (mStr || '00').slice(0, 2);
    const ampm = h >= 12 ? 'PM' : 'AM';
    h = h % 12 || 12;
    return `${h}:${min} ${ampm}`;
  };

  app.timeToInputValue = function timeToInputValue(timeStr) {
    const s = (timeStr || '').trim();
    const match = s.match(/^(\d{1,2})(?::(\d{2}))?\s*(am|pm)?$/i);
    if (!match) return '';
    let h = parseInt(match[1], 10);
    const min = match[2] || '00';
    const ap = (match[3] || '').toLowerCase();
    if (ap === 'pm' && h < 12) h += 12;
    if (ap === 'am' && h === 12) h = 0;
    if (h > 23) return '';
    return `${String(h).padStart(2, '0')}:${min}`;
  };

  app.timeToMinutes = function timeToMinutes(hhmm) {
    const [hStr, mStr] = (hhmm || '').split(':');
    const h = parseInt(hStr, 10);
    const m = parseInt(mStr || '0', 10);
    if (Number.isNaN(h)) return 0;
    return h * 60 + (Number.isNaN(m) ? 0 : m);
  };

  app.normalizeYmd = function normalizeYmd(value) {
    const match = String(value || '').trim().match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (!match) return '';
    const y = Number(match[1]);
    const m = Number(match[2]);
    const d = Number(match[3]);
    const dt = new Date(y, m - 1, d);
    if (dt.getFullYear() !== y || dt.getMonth() !== m - 1 || dt.getDate() !== d) return '';
    return `${match[1]}-${match[2]}-${match[3]}`;
  };
}
