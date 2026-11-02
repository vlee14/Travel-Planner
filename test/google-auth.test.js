import assert from 'node:assert/strict';
import { describe, test } from 'node:test';
import { installDrive } from '../js/drive.js';

function memoryDb() {
  const store = new Map();
  return {
    get(key, fallback = null) {
      return store.has(key) ? store.get(key) : fallback;
    },
    set(key, value) {
      store.set(key, String(value));
      return true;
    },
    remove(key) {
      store.delete(key);
      return true;
    }
  };
}

function createDriveApp() {
  const app = { db: memoryDb(), googleClientId: 'test-client' };
  installDrive(app);
  return app;
}

function fakeGoogle(app, respond) {
  const requests = [];
  let revoked = 0;
  globalThis.google = {
    accounts: {
      oauth2: {
        initTokenClient(config) {
          app._googleConfig = config;
          return {
            requestAccessToken(options) {
              requests.push(options);
              queueMicrotask(() => respond(config, options, requests.length));
            }
          };
        },
        revoke() {
          revoked += 1;
        }
      }
    }
  };
  return {
    requests,
    revoked: () => revoked
  };
}

describe('google sign-in', () => {
  test('reuses a stored token until it is near expiry', () => {
    const app = createDriveApp();
    app.rememberGoogleToken('cached-token', 3600);
    assert.equal(app.googleTokenFresh(), true);
    assert.equal(app.restoreGoogleSession(), true);
    assert.equal(app.accessToken, 'cached-token');
    assert.equal(app.db.get(app.GOOGLE_SIGNED_IN_KEY), 'true');

    app.accessTokenExpiresAt = Date.now() + 30 * 1000;
    assert.equal(app.googleTokenFresh(), false);
    app.db.set(app.GOOGLE_TOKEN_KEY, JSON.stringify({
      access_token: 'stale',
      expires_at: Date.now() - 1000
    }));
    assert.equal(app.restoreGoogleSession(), false);
    assert.equal(app.accessToken, null);
    assert.equal(app.db.get(app.GOOGLE_TOKEN_KEY), null);
  });

  test('returns the cached token without opening Google', async () => {
    const app = createDriveApp();
    const google = fakeGoogle(app, () => {});
    app.rememberGoogleToken('cached-token', 3600);
    const token = await app.ensureSheetsAccess({ interactive: true });
    assert.equal(token, 'cached-token');
    assert.equal(google.requests.length, 0);
    delete globalThis.google;
  });

  test('one click shares a single Google request', async () => {
    const app = createDriveApp();
    const google = fakeGoogle(app, (config) => {
      config.callback({ access_token: 'new-token', expires_in: 3600 });
    });
    const first = app.ensureSheetsAccess({ interactive: true });
    const second = app.ensureSheetsAccess({ interactive: true });
    const tokens = await Promise.all([first, second]);
    assert.deepEqual(tokens, ['new-token', 'new-token']);
    assert.equal(google.requests.length, 1);
    assert.equal(google.requests[0].prompt, '');
    assert.equal(app.db.get(app.GOOGLE_TOKEN_KEY) != null, true);
    delete globalThis.google;
  });

  test('a silent refresh does not show Google, and an interactive click does', async () => {
    const app = createDriveApp();
    const prompts = [];
    const google = fakeGoogle(app, (config, options) => {
      prompts.push(options.prompt);
      if (options.prompt === 'none') config.callback({ error: 'immediate_failed' });
      else config.callback({ access_token: 'clicked', expires_in: 1800 });
    });
    app.db.set(app.GOOGLE_SIGNED_IN_KEY, 'true');
    await assert.rejects(app.ensureSheetsAccess({ interactive: false }), (error) => {
      assert.equal(error.silent, true);
      return true;
    });
    const token = await app.ensureSheetsAccess({ interactive: true });
    assert.equal(token, 'clicked');
    assert.deepEqual(prompts, ['none', '']);
    assert.equal(google.requests.length, 2);
    delete globalThis.google;
  });

  test('sign-out drops the saved token and ignores a late Google callback', async () => {
    const app = createDriveApp();
    fakeGoogle(app, (config) => {
      config.callback({ access_token: 'late', expires_in: 3600 });
    });
    app.rememberGoogleToken('cached-token', 3600);
    const pending = app.ensureSheetsAccess({ interactive: true, force: true });
    app.signOutGoogle();
    await assert.rejects(pending);
    assert.equal(app.accessToken, null);
    assert.equal(app.db.get(app.GOOGLE_TOKEN_KEY), null);
    assert.equal(app.db.get(app.GOOGLE_SIGNED_IN_KEY), null);
    delete globalThis.google;
  });

  test('a rejected token is refreshed once, then replaced when Google repeats it', async () => {
    const app = createDriveApp();
    const prompts = [];
    fakeGoogle(app, (config, options) => {
      prompts.push(options.prompt);
      const token = options.prompt === 'consent' ? 'consented' : 'rejected';
      config.callback({ access_token: token, expires_in: 3600 });
    });
    app.accessToken = 'rejected';
    app.accessTokenExpiresAt = Date.now() + 3600 * 1000;
    app.db.set(app.GOOGLE_SIGNED_IN_KEY, 'true');
    const token = await app.recoverGoogleAccess(true);
    assert.equal(token, 'consented');
    assert.deepEqual(prompts, ['', 'consent']);
    delete globalThis.google;
  });
});
