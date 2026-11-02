import assert from 'node:assert/strict';
import { describe, test } from 'node:test';
import { createApp } from './harness.js';

function day(partial) {
  return {
    date: '',
    location: '',
    hotel: '',
    activities: [{ time: '', content: '' }],
    notes: '',
    ...partial
  };
}

describe('builder', () => {
  test('builds a blank day and a placeholder that remembers its column', () => {
    const app = createApp();
    const blank = app.blankDay({ date: '2026-11-01', location: 'Kyoto', hotel: 'Ryokan' });
    assert.equal(blank.date, '2026-11-01');
    assert.equal(blank.location, 'Kyoto');
    assert.equal(blank.hotel, 'Ryokan');
    assert.deepEqual(blank.activities, [{ time: '', content: '' }]);

    const placeholder = app.placeholderDay('2026-11-02', { sourceCol: '3', location: 'Osaka' });
    assert.equal(placeholder.date, '2026-11-02');
    assert.equal(placeholder.location, '');
    assert.equal(placeholder.sourceCol, 3);
  });

  test('knows when a day has content and clones activities apart from the original', () => {
    const app = createApp();
    const empty = app.blankDay({ date: '2026-11-01' });
    assert.equal(app.dayHasContent(empty), false);
    assert.equal(app.dayHasDetails(empty), false);
    assert.equal(app.dayHasContent({ ...empty, notes: 'Cash' }), true);

    const original = app.blankDay({
      date: '2026-11-01',
      activities: [{ time: '09:00', content: 'Temple' }]
    });
    const copy = app.cloneBuilderDay(original);
    copy.activities[0].content = 'Castle';
    assert.equal(original.activities[0].content, 'Temple');
  });

  test('keeps one empty activity ready for the next line', () => {
    const app = createApp();
    assert.equal(app.activityHasText('  '), false);
    assert.deepEqual(
      app.withTrailingActivity([{ time: '09:00', content: 'Temple' }]),
      [{ time: '09:00', content: 'Temple' }, { time: '', content: '' }]
    );
    assert.deepEqual(
      app.withTrailingActivity([{ time: '', content: '' }, { time: '', content: '' }]),
      [{ time: '', content: '' }]
    );
    assert.deepEqual(
      app.withTrailingActivity([{ time: '', content: 'Note' }, { time: '', content: '' }]),
      [{ time: '', content: 'Note' }, { time: '', content: '' }]
    );
  });

  test('fills gaps in a date range and shifts the days around an insert or removal', () => {
    const app = createApp();
    const days = [
      day({ date: '2026-11-01', location: 'Kyoto' }),
      day({ date: '2026-11-03', location: 'Nara' }),
      day({ date: '2026-12-01', location: 'Tokyo' })
    ];
    const range = app.daysForRange(days, '2026-11-03', '2026-11-01');
    assert.deepEqual(range.next.map(item => item.date), ['2026-11-01', '2026-11-02', '2026-11-03']);
    assert.equal(range.next[0].location, 'Kyoto');
    assert.equal(range.next[1].location, '');
    assert.equal(range.next[2].location, 'Nara');
    assert.equal(range.removed.length, 1);
    assert.equal(range.removed[0].location, 'Tokyo');
    assert.deepEqual(app.tripDateSpan(days), { start: '2026-11-01', end: '2026-12-01' });
    assert.equal(app.earliestBuilderDate(days), '2026-11-01');

    const inserted = app.insertBuilderDay([
      day({ date: '2026-11-01', location: 'Kyoto' }),
      day({ date: '2026-11-02', location: 'Osaka' })
    ], 0);
    assert.deepEqual(inserted.map(item => [item.date, item.location]), [
      ['2026-11-01', 'Kyoto'],
      ['2026-11-02', 'Kyoto'],
      ['2026-11-03', 'Osaka']
    ]);

    const shifted = app.shiftDatedDays([
      day({ date: '2026-11-01', location: 'Kyoto' }),
      day({ location: 'Open' })
    ], 2);
    assert.equal(shifted[0].date, '2026-11-03');
    assert.equal(shifted[1].date, '');

    const removed = app.applyDayRemoval([
      day({ date: '2026-11-01', location: 'Kyoto' }),
      day({ date: '2026-11-02', location: 'Osaka' }),
      day({ date: '2026-11-03', location: 'Nara' })
    ], 0, 'shift');
    assert.deepEqual(removed.map(item => [item.date, item.location]), [
      ['2026-11-01', 'Osaka'],
      ['2026-11-02', 'Nara']
    ]);

    const cleared = app.applyDayRemoval([
      day({ date: '2026-11-01', location: 'Kyoto', hotel: 'Ryokan', notes: 'Cash', sourceCol: 1 })
    ], 0, 'placeholder');
    assert.equal(cleared[0].date, '2026-11-01');
    assert.equal(cleared[0].location, '');
    assert.equal(cleared[0].hotel, '');
    assert.equal(cleared[0].notes, '');
    assert.equal(cleared[0].sourceCol, 1);
  });

  test('composes a wheel time back into 24-hour form', () => {
    const app = createApp();
    assert.deepEqual(app.timeParts('13:05'), { hour: 13, minute: 5, period: 'pm', hour12: 1 });
    assert.deepEqual(app.timeParts('00:00'), { hour: 0, minute: 0, period: 'am', hour12: 12 });
    assert.equal(app.timeParts('9:00'), null);
    assert.equal(app.timeParts('24:00'), null);
    assert.equal(app.composeTime(12, 0, 'am'), '00:00');
    assert.equal(app.composeTime(12, 30, 'pm'), '12:30');
    assert.equal(app.composeTime(1, 5, 'pm'), '13:05');
  });

  test('converts a sheet into days and back without moving activities', () => {
    const app = createApp();
    const rows = [
      ['Time', 'Kyoto', 'Osaka'],
      ['', 'Nov 1, 2026', 'Nov 2, 2026'],
      ['Hotel', 'Ryokan', 'Hotel B'],
      ['9:00 AM', 'Temple', 'Castle'],
      ['2:00 PM', '', 'Food'],
      ['', 'Bring cash', '']
    ];
    const days = app.parsedToDays(app.parseSheetRows(rows));
    assert.deepEqual(days.map(item => ({
      date: item.date,
      location: item.location,
      hotel: item.hotel,
      notes: item.notes,
      sourceCol: item.sourceCol,
      activities: item.activities.map(act => ({ time: act.time, content: act.content, sourceRow: act.sourceRow }))
    })), [
      {
        date: '2026-11-01',
        location: 'Kyoto',
        hotel: 'Ryokan',
        notes: 'Bring cash',
        sourceCol: 1,
        activities: [{ time: '09:00', content: 'Temple', sourceRow: 3 }]
      },
      {
        date: '2026-11-02',
        location: 'Osaka',
        hotel: 'Hotel B',
        notes: '',
        sourceCol: 2,
        activities: [
          { time: '09:00', content: 'Castle', sourceRow: 3 },
          { time: '14:00', content: 'Food', sourceRow: 4 }
        ]
      }
    ]);

    const grid = app.daysToRows([
      day({
        date: '2026-11-01',
        location: 'Kyoto',
        hotel: 'Ryokan',
        activities: [{ time: '15:00', content: 'Late' }],
        notes: 'Bring cash'
      }),
      day({
        date: '2026-11-02',
        location: 'Osaka',
        activities: [
          { time: '15:00', content: 'Also' },
          { time: '', content: 'Wander' },
          { time: '08:00', content: 'Early' }
        ]
      })
    ]);
    assert.deepEqual(grid, [
      ['Time', 'Kyoto', 'Osaka'],
      ['', 'Nov 1, 2026', 'Nov 2, 2026'],
      ['Hotel', 'Ryokan', ''],
      ['8:00 AM', '', 'Early'],
      ['3:00 PM', 'Late', 'Also'],
      ['', 'Bring cash', 'Wander']
    ]);
  });

  test('an empty model still opens as one blank day', () => {
    const app = createApp();
    const days = app.parsedToDays({ locations: ['Time'], dates: [''], rows: [] });
    assert.equal(days.length, 1);
    assert.equal(days[0].date, app.todayYmd());
    assert.equal(days[0].location, '');
    assert.deepEqual(days[0].activities, [{ time: '', content: '' }]);
  });
});
