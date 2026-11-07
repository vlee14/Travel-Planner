import assert from 'node:assert/strict';
import { describe, test } from 'node:test';
import { createApp } from './harness.js';
import { installAi } from '../js/ai.js';

function createAiApp() {
  const app = createApp();
  installAi(app);
  return app;
}

describe('ai trip constraints', () => {
  test('formats stays, bookings, and wishlist into the prompt', () => {
    const app = createAiApp();
    const text = app.formatTripConstraintsPrompt({
      hotels: [{ hotel: 'Park Hyatt Tokyo', start: '2026-03-12', end: '2026-03-14' }],
      booked: [{ date: '2026-03-13', time: '19:00', content: 'Sushi Saito' }],
      wishlist: [{ content: 'teamLab Planets' }, { content: 'Meiji Shrine' }]
    });
    assert.match(text, /Locked stays/);
    assert.match(text, /Park Hyatt Tokyo: 2026-03-12 through 2026-03-14/);
    assert.match(text, /Already booked/);
    assert.match(text, /2026-03-13 19:00 Sushi Saito/);
    assert.match(text, /Want to see/);
    assert.match(text, /teamLab Planets/);
    assert.match(text, /Meiji Shrine/);
  });

  test('includes constraints in the trip prompt', () => {
    const app = createAiApp();
    const prompt = app.tripPrompt({
      where: 'Tokyo',
      start: '2026-03-12',
      end: '2026-03-14',
      interests: 'food',
      pace: 'balanced',
      constraints: {
        hotels: [{ hotel: 'Park Hyatt Tokyo', start: '2026-03-12', end: '2026-03-13' }],
        booked: [{ date: '2026-03-13', time: '19:00', content: 'Sushi Saito' }],
        wishlist: [{ content: 'teamLab Planets' }]
      }
    });
    assert.match(prompt, /Park Hyatt Tokyo/);
    assert.match(prompt, /Sushi Saito/);
    assert.match(prompt, /teamLab Planets/);
    assert.match(prompt, /2026-03-12, 2026-03-13, 2026-03-14/);
  });

  test('forces locked hotels and booked activities onto generated days', () => {
    const app = createAiApp();
    const plan = app.generatedTripDays({
      title: 'Tokyo',
      days: [
        {
          date: '2026-03-12',
          location: 'Shinjuku',
          hotel: 'Somewhere else',
          activities: [{ time: '10:00', content: 'Walk around' }]
        },
        {
          date: '2026-03-13',
          location: 'Shibuya',
          hotel: '',
          activities: [{ time: '11:00', content: 'Shopping' }]
        },
        {
          date: '2026-03-14',
          location: 'Shibuya',
          hotel: '',
          activities: [{ time: '10:00', content: 'Coffee' }]
        }
      ]
    }, '2026-03-12', '2026-03-14', 'Tokyo', 'balanced', {
      hotels: [{ hotel: 'Park Hyatt Tokyo', start: '2026-03-12', end: '2026-03-13' }],
      booked: [{ date: '2026-03-13', time: '19:00', content: 'Sushi Saito' }],
      wishlist: [{ content: 'teamLab Planets' }]
    });

    assert.equal(plan.days[0].hotel, 'Park Hyatt Tokyo');
    assert.equal(plan.days[1].hotel, 'Park Hyatt Tokyo');
    assert.equal(plan.days[2].hotel, '');
    assert.deepEqual(
      plan.days[1].activities.find(act => act.content === 'Sushi Saito'),
      { time: '19:00', content: 'Sushi Saito' }
    );
    assert.equal(plan.days[1].activities[0].time <= plan.days[1].activities[1].time, true);
  });

  test('applyTripConstraints updates hotel nights and booking times in place', () => {
    const app = createAiApp();
    const days = [
      { date: '2026-04-01', location: 'Kyoto', hotel: '', activities: [{ time: '09:00', content: 'Temple' }] },
      { date: '2026-04-02', location: 'Kyoto', hotel: 'Inn', activities: [{ time: '12:00', content: 'Lunch' }] }
    ];
    app.applyTripConstraints(days, {
      hotels: [{ hotel: 'Hiiragiya', start: '2026-04-01', end: '2026-04-02' }],
      booked: [{ date: '2026-04-02', time: '18:30', content: 'Kaiseki dinner' }]
    });
    assert.equal(days[0].hotel, 'Hiiragiya');
    assert.equal(days[1].hotel, 'Hiiragiya');
    assert.deepEqual(days[1].activities, [
      { time: '12:00', content: 'Lunch' },
      { time: '18:30', content: 'Kaiseki dinner' }
    ]);
  });
});
