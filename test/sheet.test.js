import assert from 'node:assert/strict';
import { describe, test } from 'node:test';
import { createApp } from './harness.js';

const GRID = [
  ['Time', 'Kyoto', 'Osaka'],
  ['', 'Nov 1, 2026', 'Nov 2, 2026'],
  ['Hotel', 'Ryokan', 'Hotel B'],
  ['9:00 AM', 'Temple', 'Castle'],
  ['', 'Bring cash', '']
];

describe('sheet', () => {
  test('parses comma and tab grids, including quotes and newlines', () => {
    const app = createApp();
    assert.deepEqual(app.parseCSV(''), []);
    assert.deepEqual(
      app.parseCSV('Kyoto,Osaka\n"Nov 1, 2026","Nov 2, 2026"\r\n'),
      [
        ['Kyoto', 'Osaka'],
        ['Nov 1, 2026', 'Nov 2, 2026']
      ]
    );
    assert.deepEqual(app.parseCSV('  Kyoto  \tOsaka'), [['Kyoto', 'Osaka']]);
    assert.deepEqual(app.parseCSV('"say ""hi"""\t"line\n2"'), [['say "hi"', 'line\n2']]);
  });

  test('writes a grid as tab-separated text and reads it back', () => {
    const app = createApp();
    const rows = [
      ['Time', 'Kyoto'],
      ['', 'Nov 1, 2026'],
      ['9:00 AM', 'say "hi"']
    ];
    assert.deepEqual(app.parseCSV(app.rowsToCsv(rows)), rows);
  });

  test('recognizes a hotel row and keeps it out of the activities', () => {
    const app = createApp();
    assert.equal(app.looksLikeHotelHeader('Hotel'), true);
    assert.equal(app.looksLikeHotelHeader('HOTEL'), true);
    assert.equal(app.looksLikeHotelHeader('The hotel'), true);
    assert.equal(app.looksLikeHotelHeader('Hotels'), false);
    assert.equal(app.looksLikeHotelHeader(''), false);

    const model = app.parseSheetRows([
      ['Time', 'Kyoto', '<b>Osaka</b>'],
      ['', 'Nov 1, 2026', 'Nov 2, 2026'],
      ['Hotel', 'Ryokan'],
      ['9:00 AM', 'Temple &amp; garden', 'Castle']
    ]);
    assert.equal(model.locations[2], 'Osaka');
    assert.equal(model.hotelGridRow, 2);
    assert.deepEqual(model.hotelCells, ['Hotel', 'Ryokan', '']);
    assert.equal(model.rows.length, 1);
    assert.equal(model.rows[0].time, '9:00 AM');
    assert.equal(model.rows[0].gridRow, 3);
    assert.deepEqual(model.rows[0].cells, ['Temple &amp; garden', 'Castle']);
    assert.equal(app.parsedToGrid(model)[3][1], 'Temple & garden');
    assert.equal(app.parseSheetRows([]), null);
  });

  test('lines a day column up with its location and date', () => {
    const app = createApp();
    const model = app.parseSheetRows([
      ['Time', 'Kyoto', '', 'Osaka'],
      ['', 'Nov 1, 2026', '', 'Nov 2, 2026'],
      ['9:00 AM', 'Temple', 'Ignored', 'Castle']
    ]);
    assert.deepEqual(model.locations, ['Time', 'Kyoto', 'Location 2', 'Osaka']);
    app.parsed = model;
    assert.equal(app.findDayColumn('Kyoto', 'Nov 1, 2026'), 1);
    assert.equal(app.findDayColumn('Osaka', 'Nov 2, 2026'), 3);
    assert.equal(app.findDayColumn('Kyoto', 'Nov 2, 2026'), -1);
    assert.equal(model.rows[0].cells[0], 'Temple');
    assert.equal(model.rows[0].cells[2], 'Castle');
  });

  test('a one-cell edit leaves the neighboring day and its note alone', () => {
    const app = createApp();
    const days = app.parsedToDays(app.parseSheetRows(GRID));
    assert.equal(days[0].activities[0].content, 'Temple');
    assert.equal(days[0].notes, 'Bring cash');
    days[0].activities[0].content = 'Fushimi Inari';

    const plan = app.planLinkedSheetEdits(GRID, days, []);
    assert.deepEqual(plan.cells, [{ row: 3, col: 1, value: 'Fushimi Inari' }]);
    assert.equal(plan.grid[3][1], 'Fushimi Inari');
    assert.equal(plan.grid[3][2], 'Castle');
    assert.equal(plan.grid[4][1], 'Bring cash');
    assert.equal(plan.grid[0][2], 'Osaka');
    assert.deepEqual(plan.requests, []);
  });

  test('keeps a uniform hotel merge and splits it when one stay changes', () => {
    const app = createApp();
    const grid = [
      ['Time', 'Kyoto', 'Kyoto'],
      ['', 'Nov 1, 2026', 'Nov 2, 2026'],
      ['Hotel', 'Ryokan', 'Ryokan'],
      ['9:00 AM', 'Temple', 'Castle']
    ];
    const merges = [{
      startRowIndex: 2,
      endRowIndex: 3,
      startColumnIndex: 1,
      endColumnIndex: 3
    }];
    const days = app.parsedToDays(app.parseSheetRows(grid));
    const same = app.planLinkedSheetEdits(grid, days, merges);
    assert.equal(same.cells.length, 0);
    assert.equal(same.requests.length, 0);
    assert.equal(same.merges.length, 1);

    days[1].hotel = 'Other';
    const changed = app.planLinkedSheetEdits(grid, days, merges);
    assert.equal(changed.requests.some(request => request.unmergeCells), true);
    assert.equal(changed.grid[2][1], 'Ryokan');
    assert.equal(changed.grid[2][2], 'Other');
    assert.equal(changed.grid[3][1], 'Temple');
    assert.equal(changed.grid[3][2], 'Castle');
  });
});
