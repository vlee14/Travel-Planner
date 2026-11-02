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
    const now = new Date();

    const yrMatch = s.match(/\b(202[0-9]|203[0-9])\b/);
    const explicitYear = yrMatch ? parseInt(yrMatch[1], 10) : null;
    const baseYear = explicitYear || fallbackYear || (app.parsed && app.parsed.dates ? app.detectItineraryYear(app.parsed.dates) : now.getFullYear());

    let d = new Date(s);
    if (isNaN(d.getTime()) || !yrMatch) {
      const cleanStr = s.replace(/^[a-zA-Z]{3,}\s*,?\s*/, '');
      const testWithYear = new Date(`${cleanStr}, ${baseYear}`);
      if (!isNaN(testWithYear.getTime())) {
        d = testWithYear;
      } else if (!isNaN(d.getTime())) {
        d.setFullYear(baseYear);
      }
    } else if (explicitYear) {
      d.setFullYear(explicitYear);
    }

    if (isNaN(d.getTime())) return null;
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
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
