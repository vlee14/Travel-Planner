import assert from 'node:assert/strict';
import { describe, test } from 'node:test';
import { createApp } from './harness.js';

describe('dates', () => {
  test('normalizes a calendar date and rejects impossible ones', () => {
    const app = createApp();
    assert.equal(app.normalizeYmd('2026-11-01'), '2026-11-01');
    assert.equal(app.normalizeYmd('2026-11-01T10:00'), '2026-11-01');
    assert.equal(app.normalizeYmd('2026-02-31'), '');
    assert.equal(app.normalizeYmd('2026-13-01'), '');
    assert.equal(app.normalizeYmd('11/01/2026'), '');
  });

  test('adds days and counts the nights between two dates', () => {
    const app = createApp();
    assert.equal(app.addDaysYmd('2026-01-31', 1), '2026-02-01');
    assert.equal(app.addDaysYmd('2026-12-31', 1), '2027-01-01');
    assert.equal(app.addDaysYmd('2024-02-28', 1), '2024-02-29');
    assert.equal(app.addDaysYmd('nope', 1), app.todayYmd());
    assert.equal(app.daysBetweenYmd('2026-11-01', '2026-11-03'), 2);
    assert.equal(app.daysBetweenYmd('2026-11-03', '2026-11-01'), -2);
    assert.equal(app.daysBetweenYmd('', '2026-11-01'), 0);
  });

  test('formats a trip date and a 24-hour time', () => {
    const app = createApp();
    assert.equal(app.formatTripDate('2026-11-01'), 'Nov 1, 2026');
    assert.equal(app.formatTripDate(''), '');
    assert.equal(app.formatTimeLabel('00:00'), '12:00 AM');
    assert.equal(app.formatTimeLabel('09:05'), '9:05 AM');
    assert.equal(app.formatTimeLabel('12:00'), '12:00 PM');
    assert.equal(app.formatTimeLabel('13:05'), '1:05 PM');
    assert.equal(app.formatTimeLabel(''), '');
  });

  test('parses sheet times into editor values and minutes', () => {
    const app = createApp();
    assert.equal(app.timeToInputValue('9 AM'), '09:00');
    assert.equal(app.timeToInputValue('9:05'), '09:05');
    assert.equal(app.timeToInputValue('12:00 AM'), '00:00');
    assert.equal(app.timeToInputValue('12 PM'), '12:00');
    assert.equal(app.timeToInputValue('7:30 pm'), '19:30');
    assert.equal(app.timeToInputValue('24:00'), '');
    assert.equal(app.timeToInputValue('noon'), '');
    assert.equal(app.timeToMinutes('02:15'), 135);
    assert.equal(app.timeToMinutes(''), 0);
    assert.equal(app.parseTimeForSort('9:30 PM'), 21 * 60 + 30);
    assert.equal(app.parseTimeForSort('12:15 PM'), 12 * 60 + 15);
    assert.equal(app.parseTimeForSort('12:00 AM'), 0);
  });

  test('parses the date labels a sheet actually stores', () => {
    const app = createApp();
    assert.equal(app.parseSmartDate('Nov 1, 2026'), '2026-11-01');
    assert.equal(app.parseSmartDate('Saturday, Nov 1', 2026), '2026-11-01');
    assert.equal(app.parseSmartDate('November 1', 2026), '2026-11-01');
    assert.equal(app.parseSmartDate('Nov. 1, 2026'), '2026-11-01');
    assert.equal(app.parseSmartDate('1 Nov 2026'), '2026-11-01');
    assert.equal(app.parseSmartDate('November 1st, 2026'), '2026-11-01');
    assert.equal(app.parseSmartDate('11/1/2026'), '2026-11-01');
    assert.equal(app.parseSmartDate('11/1', 2026), '2026-11-01');
    assert.equal(app.parseSmartDate('2026/11/01'), '2026-11-01');
    assert.equal(app.parseSmartDate('Nov 1, 2019', 2026), '2026-11-01');
    assert.equal(app.parseSmartDate(''), null);
    assert.equal(app.parseSmartDate('   '), null);
    assert.equal(app.parseDateForSort('Nov 1, 2026'), new Date(2026, 10, 1).getTime());
    assert.equal(app.parseDateForSort(''), Infinity);
    assert.equal(app.parseDateForSort('   '), Infinity);
  });

  test('keeps an ISO date on that calendar day and rejects labels that are not dates', () => {
    const app = createApp();
    assert.equal(app.parseSmartDate('2026-11-01'), '2026-11-01');
    assert.equal(app.parseSmartDate('2026-11-01T00:00'), '2026-11-01');
    assert.equal(app.parseSmartDate('2026-02-31'), null);
    assert.equal(app.parseSmartDate('nope'), null);
    assert.equal(app.parseSmartDate('???'), null);
    assert.equal(app.parseSmartDate('TBD'), null);
    assert.equal(app.parseSmartDate('May'), null);
    assert.equal(app.parseSmartDate('2026'), null);
    assert.equal(app.parseDateForSort('nope'), Infinity);
  });

  test('finds an explicit year before guessing from today', () => {
    const app = createApp();
    const year = new Date().getFullYear();
    assert.equal(app.detectItineraryYear([]), year);
    assert.equal(app.detectItineraryYear(['', 'Nov 1, 2026']), 2026);
    app.h1 = { textContent: 'Japan 2028' };
    assert.equal(app.detectItineraryYear(['', 'Nov 1']), 2028);

    const undated = createApp();
    const jan1 = new Date(year, 0, 1);
    const rolled = (Date.now() - jan1.getTime()) > 30 * 86400000;
    assert.equal(undated.detectItineraryYear(['', 'January 1']), rolled ? year + 1 : year);
  });
});
