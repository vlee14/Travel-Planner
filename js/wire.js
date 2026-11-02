// Startup, settings, and event wiring.
export function installWire(app) {
app.csvInput = document.getElementById('csv-input');
app.btnParse = document.getElementById('btn-parse');
app.errorMsg = document.getElementById('error-msg');
app.itinerarySection = document.getElementById('itinerary-section');
app.inputSection = document.querySelector('.input-section');
app.emptyState = document.getElementById('empty-state');
app.listContent = document.getElementById('list-content');
app.tabs = document.querySelectorAll('.tab');
app.panels = document.querySelectorAll('.input-panel');
app.viewBtns = document.querySelectorAll('.view-btn[data-view]');
app.homeBtn = document.getElementById('home-btn');
app.h1 = document.querySelector('.header-row h1');
app.shareBtn = document.getElementById('share-btn');
app.printBtn = document.getElementById('print-btn');
app.syncBtn = document.getElementById('sync-btn');
app.searchInput = document.getElementById('search-input');
app.settingsBtn = document.getElementById('settings-btn');
app.settingsModal = document.getElementById('settings-modal');
app.settingsClose = document.getElementById('settings-close');
app.settingProvider = document.getElementById('setting-provider');
app.settingApiKey = document.getElementById('setting-api-key');
app.settingApiKeyLabel = document.getElementById('setting-api-key-label');
app.settingApiKeyLink = document.getElementById('setting-api-key-link');
app.settingKeyNote = document.getElementById('setting-key-note');
app.settingVcKey = document.getElementById('setting-vc-key');
app.settingModel = document.getElementById('setting-model');
app.btnGoogleSignIn = document.getElementById('btn-google-signin');
app.btnDriveAuth = document.getElementById('btn-drive-auth');
app.googleStatus = document.getElementById('google-status');
app.btnDrivePicker = document.getElementById('btn-drive-picker');
app.recentDocsSort = document.getElementById('recent-docs-sort');
app.btnClearDoc = document.getElementById('btn-clear-doc');
app.driveUrlInput = document.getElementById('drive-url-input');
app.btnDriveUrlFetch = document.getElementById('btn-drive-url-fetch');
app.driveManualInput = document.getElementById('google-auth-container');
app.btnSaveSettings = document.getElementById('btn-save-settings');
app.btnClearCache = document.getElementById('btn-clear-cache');
app.aiModal = document.getElementById('ai-modal');
app.aiModalClose = document.getElementById('ai-modal-close');
app.aiModalTitle = document.getElementById('ai-modal-title');
app.aiModalBody = document.getElementById('ai-modal-body');
app.aiFollowUpInput = document.getElementById('ai-follow-up-input');
app.aiFollowUpBtn = document.getElementById('ai-follow-up-btn');
app.packingCountInput = document.getElementById('packing-count-input');
app.btnGenerateChecklist = document.getElementById('btn-generate-checklist');
app.packingContent = document.getElementById('packing-content');
app.packingTabs = document.querySelectorAll('.packing-tab');
app.tripProgressBar = document.getElementById('trip-progress-bar');
app.tripProgressLabel = document.getElementById('trip-progress-label');
app.tripProgressFill = document.getElementById('trip-progress-fill');
app.mapRouteBtn = document.getElementById('map-route-btn');

app.parsed = null; // { locations, dates, rows: [ { time, cells[] } ] }
app.builderEditing = false;
app.builderReady = false;
app.map = null;
app.mapLayerGroup = null;
app.mapTileLayer = null;
app.polyline = null;
app.mapPoints = []; // Array of {lat, lon}
app.mapMarkers = {}; // Object to track unique locations
app.hasScrolledToToday = false
app.currentTempUnit = 'F'; // 'F' or 'C'
app.autoSyncInterval = null; // timer handle for auto-sync
app.lastSyncedDocId = null; // track which doc is auto-synced
app.lastSyncedGid = '0';
app.lastSyncedTime = null; // timestamp of last successful sync
app.GEMINI_API_KEY_STORAGE = 'gemini_api_key';
app.AI_PROVIDER_STORAGE = 'ai_provider';
app.VC_API_KEY_STORAGE = 'visual_crossing_key';
app.WEATHER_CACHE_PREFIX = 'weather_cache_v1_';
app.WEATHER_CACHE_EXPIRY = 1000 * 60 * 60 * 6; // 6 hours
app.AI_CACHE_PREFIX = 'ai_chat_cache_v1_';
app.AI_CACHE_EXPIRY = 1000 * 60 * 60 * 3; // 3 hours
app.PACKING_COUNT_STORAGE = 'packing_travelers_count';
app.DOC_CACHE_PREFIX = 'doc_cache_v1_';
app.DOC_METADATA_KEY = 'doc_metadata_v1';
app.EVENT_CHECK_PREFIX = 'event_chk_v1_';
app.TEMP_UNIT_STORAGE = 'temp_unit';

app.db = {
  get(key, defaultValue = null) {
    try {
      const val = localStorage.getItem(key);
      return val !== null ? val : defaultValue;
    } catch (e) {
      console.warn(`Error reading from localStorage key "${key}":`, e);
      return defaultValue;
    }
  },
  set(key, value) {
    try {
      localStorage.setItem(key, value);
      return true;
    } catch (e) {
      console.warn(`Error writing to localStorage key "${key}":`, e);
      if (e.name === 'QuotaExceededError' || e.code === 22) {
        app.db.evictOldCache();
        try {
          localStorage.setItem(key, value);
          return true;
        } catch (retryError) {
          console.error('Eviction failed to free enough space:', retryError);
        }
      }
      return false;
    }
  },
  remove(key) {
    try {
      localStorage.removeItem(key);
      return true;
    } catch (e) {
      console.warn(`Error removing localStorage key "${key}":`, e);
      return false;
    }
  },
  evictOldCache() {
    try {
      const keys = [];
      for (let i = 0; i < localStorage.length; i++) {
        const key = localStorage.key(i);
        if (key && (key.startsWith(app.WEATHER_CACHE_PREFIX) || key.startsWith(app.AI_CACHE_PREFIX))) {
          let timestamp = 0;
          try {
            const data = JSON.parse(localStorage.getItem(key));
            if (data && data.timestamp) timestamp = data.timestamp;
          } catch (err) {}
          keys.push({ key, timestamp });
        }
      }
      keys.sort((a, b) => a.timestamp - b.timestamp);
      const toRemove = keys.slice(0, Math.ceil(keys.length * 0.3));
      toRemove.forEach(item => {
        localStorage.removeItem(item.key);
      });
      console.info(`Evicted ${toRemove.length} old cache entries due to quota limits.`);
    } catch (e) {
      console.error('Failed to evict old cache:', e);
    }
  }
};

// Initialize settings immediately from storage to prevent empty fields on quick click
app.visualCrossingKey = app.db.get(app.VC_API_KEY_STORAGE);
// Hardcoded default Client ID for easier setup
app.googleClientId = '663334830027-1182pamonrcagb0sgvg46jserh7v5jqg.apps.googleusercontent.com';
app.tokenClient = null;
app.accessToken = null;
const savedProvider = app.db.get(app.AI_PROVIDER_STORAGE, 'google');
app.aiProvider = app.AI_PROVIDERS[savedProvider] ? savedProvider : 'google';
app.aiKeys = {};
app.aiModels = {};
Object.keys(app.AI_PROVIDERS).forEach(id => {
  const meta = app.AI_PROVIDERS[id];
  app.aiKeys[id] = app.db.get(meta.storageKey) || '';
  app.aiModels[id] = app.db.get(meta.modelKey, meta.defaultModel);
});
app.geminiApiKey = app.aiKeys.google || null;
app.geminiModelName = app.aiModels.google;
app.aiConversationHistory = [];
app.suggestionTarget = null;
app.currentAiCacheKey = null;
app.currentChecklistType = 'packing'; // 'packing' or 'apps'
app.currentTempUnit = app.db.get(app.TEMP_UNIT_STORAGE, 'F');

app.showAiProvider = function showAiProvider(provider) {
  const meta = app.AI_PROVIDERS[provider];
  if (!meta) return;
  const keys = app.aiDraft ? app.aiDraft.keys : app.aiKeys;
  const models = app.aiDraft ? app.aiDraft.models : app.aiModels;
  if (app.settingProvider) app.settingProvider.value = provider;
  document.querySelectorAll('.provider-choice').forEach(btn => {
    const selected = btn.dataset.provider === provider;
    btn.classList.toggle('active', selected);
    btn.setAttribute('aria-checked', selected ? 'true' : 'false');
  });
  if (app.settingApiKeyLabel) app.settingApiKeyLabel.textContent = meta.keyLabel;
  if (app.settingApiKeyLink) app.settingApiKeyLink.href = meta.href;
  if (app.settingKeyNote) app.settingKeyNote.textContent = meta.note;
  if (app.settingApiKey) app.settingApiKey.value = keys[provider] || '';
  const model = models[provider] || meta.defaultModel;
  app.setModelOptions([{ id: model, label: model }], model);
};

app.captureAiDraft = function captureAiDraft() {
  if (!app.aiDraft || !app.settingProvider) return;
  const provider = app.settingProvider.value;
  app.aiDraft.provider = provider;
  app.aiDraft.keys[provider] = app.settingApiKey ? app.settingApiKey.value.trim() : '';
  if (app.settingModel && app.settingModel.value) app.aiDraft.models[provider] = app.settingModel.value;
};

// Pre-populate setting inputs so they are ready before the modal is even opened
app.showAiProvider(app.aiProvider);
if (app.settingVcKey) app.settingVcKey.value = app.visualCrossingKey || '';

// Theme Logic. Stored value is light, dark, or system.
app.themePreference = app.db.get('theme', 'light');
if (app.themePreference !== 'light' && app.themePreference !== 'dark' && app.themePreference !== 'system') {
  app.themePreference = 'light';
}
app.currentTheme = 'light';
app.htmlEl = document.documentElement;
app.systemThemeQuery = window.matchMedia('(prefers-color-scheme: dark)');

// Initialize theme and temp unit UI
app.updateTheme(app.themePreference);
app.updateTempUnitUI();

// Prevent layout shift when switching between short and long pages (scrollbar jump)
document.documentElement.style.overflowY = 'scroll';

document.querySelectorAll('.theme-choice').forEach(btn => {
  btn.addEventListener('click', () => app.updateTheme(btn.dataset.themeChoice));
});

app.systemThemeQuery.addEventListener('change', () => {
  if (app.themePreference === 'system') app.updateTheme('system');
});

document.querySelectorAll('.temp-unit-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    app.currentTempUnit = app.currentTempUnit === 'F' ? 'C' : 'F';
    app.db.set(app.TEMP_UNIT_STORAGE, app.currentTempUnit);
    app.updateTempUnitUI();
    app.rerenderAllWeatherBadges();
  });
});

// Update all clocks every minute
setInterval(() => {
  document.querySelectorAll('.local-time-display').forEach(app.updateSingleClock);
  app.updateSyncStatusLabel();
}, 60000);

// Event delegation for checkbox state storage
document.addEventListener('change', (e) => {
  if (e.target.matches('input[type="checkbox"][data-key]')) {
    app.db.set(e.target.dataset.key, e.target.checked);
    const card = e.target.closest('.event-card');
    if (card) card.classList.toggle('checked', e.target.checked);
  }
});

app.btnGenerateChecklist?.addEventListener('click', () => app.generateChecklist(true));

app.packingTabs.forEach(tab => {
  tab.addEventListener('click', () => {
    app.packingTabs.forEach(t => t.classList.remove('active'));
    tab.classList.add('active');
    app.currentChecklistType = tab.dataset.type;

    // Hide travelers input for apps
    document.getElementById('packing-inputs-wrapper').style.display = app.currentChecklistType === 'packing' ? 'flex' : 'none';

    // Attempt to load existing content
    if (app.parsed) {
      const count = app.currentChecklistType === 'packing' ? (app.packingCountInput?.value || 1) : 1;
      const key = app.getChecklistCacheKey(app.currentChecklistType, count);
      app.currentAiCacheKey = key;
      let loaded = false;
      try {
        const cachedJson = app.db.get(key);
        if (cachedJson) {
          const cached = JSON.parse(cachedJson);
          if (Date.now() - cached.timestamp < app.AI_CACHE_EXPIRY) {
            app.packingContent.innerHTML = app.renderMarkdown(cached.text, key);
            loaded = true;
          }
        }
      } catch (e) { }

      if (!loaded) {
        app.packingContent.innerHTML = '<p class="muted">Regenerate to create this list.</p>';
      }
    }
  });
});

// Search Logic
app.searchInput?.addEventListener('input', (e) => {
  const term = e.target.value.toLowerCase();
  const cards = document.querySelectorAll('.location-card');

  cards.forEach(card => {
    const days = card.querySelectorAll('.day-group');
    let hasVisibleDay = false;

    days.forEach(day => {
      const text = day.textContent.toLowerCase();
      if (text.includes(term)) {
        day.classList.remove('hidden');
        hasVisibleDay = true;
      } else {
        day.classList.add('hidden');
      }
    });

    // If card matches title, show all days? Or just filter days?
    // Let's hide card if no days match and card title doesn't match
    const titleMatch = card.querySelector('.location-card-name').textContent.toLowerCase().includes(term);

    if (titleMatch) {
      if (term) card.classList.remove('collapsed');
      // If title matches, show matches inside, but ensure card is visible
      // Optional: unhide all days if title matches? Let's keep granular filtering.
      card.classList.remove('hidden');
      if (!hasVisibleDay && term.length > 0) {
        // If title matches but no days, show card but maybe days are hidden. 
        // Let's unhide all days if title matches to avoid empty card
        days.forEach(d => d.classList.remove('hidden'));
      }
    } else {
      if (hasVisibleDay) {
        card.classList.remove('hidden');
        if (term) card.classList.remove('collapsed');
      }
      else card.classList.add('hidden');
    }
  });
});

app.aiFollowUpBtn?.addEventListener('click', app.handleFollowUp);
app.aiFollowUpInput?.addEventListener('keydown', (e) => {
  if (e.key === 'Enter' && !e.shiftKey) {
    e.preventDefault();
    app.handleFollowUp();
  }
});

app.handleGoogleSignIn = () => {
  if (app.accessToken && app.accessTokenExpiresAt > Date.now()) {
    app.signOutGoogle();
    return;
  }
  if (!app.googleClientId) {
    alert('Google Client ID is not configured.');
    return;
  }

  app.initTokenClient();
  if (!app.tokenClient) {
    alert('Google Auth library is still loading. Please wait a few seconds and try again.');
    return;
  }
  app.updateGoogleStatus(false, true);
  app.ensureSheetsAccess({ interactive: true }).then(() => {
    if (new URLSearchParams(window.location.search).get('docId')) app.loadFromUrl();
  }).catch((error) => {
    app.updateGoogleStatus(!!app.googleTokenFresh());
    if (error && error.blocked && app.googleStatus) {
      app.googleStatus.textContent = error.message;
      app.googleStatus.style.color = 'var(--danger)';
    }
  });
};

app.btnGoogleSignIn?.addEventListener('click', app.handleGoogleSignIn);
app.btnDriveAuth?.addEventListener('click', app.handleGoogleSignIn);

app.btnDrivePicker?.addEventListener('click', async () => {
  try {
    await app.ensureSheetsAccess({ interactive: true });
  } catch (error) {
    return;
  }
  if (typeof gapi === 'undefined') return;
  gapi.load('picker', { callback: app.createPicker });
});

app.btnDriveUrlFetch?.addEventListener('click', async () => {
  const input = app.driveUrlInput.value.trim();
  if (!input) return;

  // Extract File ID (supports /d/ID or /d/e/ID or just the ID)
  const fileIdMatch = input.match(/\/d\/(?:e\/)?([a-zA-Z0-9-_]+)/);
  const gidMatch = input.match(/[#&]gid=([0-9]+)/);

  const fileId = fileIdMatch ? fileIdMatch[1] : input;
  const gid = gidMatch ? gidMatch[1] : '0';

  // Normalize UI to show the document URL instead of internal export links or just the ID
  app.driveUrlInput.value = `https://docs.google.com/spreadsheets/d/${fileId}/edit${gid !== '0' ? '#gid=' + gid : ''}`;

  try {
    await app.ensureSheetsAccess({ interactive: true });
  } catch (error) {
    app.showError(error.message || 'Google sign-in failed. Try again.');
    return;
  }
  app.fetchSheetData(fileId, 'Manual Entry', gid);
});

app.refreshMenu = function refreshMenu(select) {
  if (!select) return;
  const menu = select.closest('.menu');
  const label = menu && menu.querySelector('.menu-value');
  if (!label) return;
  const selected = select.options[select.selectedIndex];
  label.textContent = selected ? selected.textContent : '';
};

app.closeMenu = function closeMenu() {
  if (app.menuList) app.menuList.remove();
  app.menuList = null;
  app.menuOpenSelect = null;
  document.querySelectorAll('.menu-button[aria-expanded="true"]').forEach(btn => {
    btn.setAttribute('aria-expanded', 'false');
  });
};

app.openMenu = function openMenu(select, button) {
  app.closeMenu();
  if (!select || !select.options.length) return;
  const list = document.createElement('ul');
  list.className = 'menu-list';
  list.setAttribute('role', 'listbox');
  const current = select.value;
  Array.from(select.options).forEach(option => {
    const item = document.createElement('li');
    const choice = document.createElement('button');
    choice.type = 'button';
    choice.setAttribute('role', 'option');
    choice.setAttribute('aria-selected', option.value === current ? 'true' : 'false');
    const text = document.createElement('span');
    text.textContent = option.textContent;
    choice.appendChild(text);
    choice.addEventListener('click', () => {
      if (select.value !== option.value) {
        select.value = option.value;
        select.dispatchEvent(new Event('change', { bubbles: true }));
      }
      app.refreshMenu(select);
      app.closeMenu();
      button.focus();
    });
    item.appendChild(choice);
    list.appendChild(item);
  });
  document.body.appendChild(list);
  const rect = button.getBoundingClientRect();
  const spaceBelow = window.innerHeight - rect.bottom;
  const spaceAbove = rect.top;
  const upward = spaceBelow < 220 && spaceAbove > spaceBelow;
  const maxHeight = Math.max(140, Math.min(320, (upward ? spaceAbove : spaceBelow) - 16));
  list.style.left = Math.max(8, rect.left) + 'px';
  list.style.width = rect.width + 'px';
  list.style.maxHeight = maxHeight + 'px';
  if (upward) list.style.top = Math.max(8, rect.top - Math.min(list.scrollHeight, maxHeight) - 6) + 'px';
  else list.style.top = (rect.bottom + 6) + 'px';
  const selected = list.querySelector('[aria-selected="true"]');
  if (selected) list.scrollTop = Math.max(0, selected.offsetTop - 8);
  button.setAttribute('aria-expanded', 'true');
  app.menuList = list;
  app.menuOpenSelect = select;
};

document.querySelectorAll('.provider-choice').forEach(btn => {
  btn.addEventListener('click', () => {
    if (!app.settingProvider || app.settingProvider.value === btn.dataset.provider) return;
    app.settingProvider.value = btn.dataset.provider;
    app.settingProvider.dispatchEvent(new Event('change', { bubbles: true }));
  });
});

document.querySelectorAll('.menu').forEach(menu => {
  const button = menu.querySelector('.menu-button');
  const select = menu.querySelector('select');
  if (!button || !select) return;
  button.addEventListener('click', (event) => {
    event.stopPropagation();
    if (app.menuOpenSelect === select) app.closeMenu();
    else app.openMenu(select, button);
  });
});

document.addEventListener('click', (event) => {
  if (!app.menuList) return;
  if (app.menuList.contains(event.target)) return;
  app.closeMenu();
});

document.addEventListener('keydown', (event) => {
  if (event.key === 'Escape' && app.menuList) {
    event.preventDefault();
    event.stopPropagation();
    const select = app.menuOpenSelect;
    const button = select && select.closest('.menu') && select.closest('.menu').querySelector('.menu-button');
    app.closeMenu();
    if (button) button.focus();
  }
});

window.addEventListener('resize', () => app.closeMenu());
app.settingsModal?.addEventListener('scroll', () => app.closeMenu());
document.querySelector('.settings-modal-content')?.addEventListener('scroll', () => app.closeMenu());
app.refreshMenu(app.settingModel);

app.settingProvider?.addEventListener('change', () => {
  if (!app.aiDraft) return;
  const previous = app.aiDraft.provider;
  app.aiDraft.keys[previous] = app.settingApiKey.value.trim();
  if (app.settingModel.value) app.aiDraft.models[previous] = app.settingModel.value;
  app.aiDraft.provider = app.settingProvider.value;
  app.showAiProvider(app.aiDraft.provider);
  const key = (app.aiDraft.keys[app.aiDraft.provider] || '').trim();
  if (key) app.populateModels(key);
});

app.settingApiKey?.addEventListener('change', () => {
  const key = app.settingApiKey.value.trim();
  if (app.aiDraft) app.aiDraft.keys[app.settingProvider.value] = key;
  if (key) app.populateModels(key);
});

app.settingsBtn?.addEventListener('click', () => {
  app.settingsModal.style.display = 'flex';
  app.aiDraft = {
    provider: app.aiProvider,
    keys: { ...app.aiKeys },
    models: { ...app.aiModels }
  };
  app.showAiProvider(app.aiProvider);
  app.settingVcKey.value = app.visualCrossingKey || '';
  const key = (app.aiDraft.keys[app.aiProvider] || '').trim();
  if (key) app.populateModels(key);
});

app.settingsClose?.addEventListener('click', () => {
  app.aiDraft = null;
  app.settingsModal.style.display = 'none';
});

app.btnSaveSettings?.addEventListener('click', () => {
  const vcKey = app.settingVcKey.value.trim();
  app.captureAiDraft();
  const draft = app.aiDraft || {
    provider: app.settingProvider ? app.settingProvider.value : app.aiProvider,
    keys: { ...app.aiKeys, [app.aiProvider]: app.settingApiKey.value.trim() },
    models: { ...app.aiModels }
  };
  if (app.settingModel && app.settingModel.value) draft.models[draft.provider] = app.settingModel.value;

  app.aiProvider = app.AI_PROVIDERS[draft.provider] ? draft.provider : 'google';
  app.db.set(app.AI_PROVIDER_STORAGE, app.aiProvider);
  Object.keys(app.AI_PROVIDERS).forEach(id => {
    const meta = app.AI_PROVIDERS[id];
    const key = (draft.keys[id] || '').trim();
    app.aiKeys[id] = key;
    if (key) app.db.set(meta.storageKey, key);
    else app.db.remove(meta.storageKey);
    const model = (draft.models[id] || '').trim() || meta.defaultModel;
    app.aiModels[id] = model;
    app.db.set(meta.modelKey, model);
  });
  app.geminiApiKey = app.aiKeys.google || null;
  app.geminiModelName = app.aiModels.google;
  app.syncAiAvailability();

  if (vcKey) {
    app.db.set(app.VC_API_KEY_STORAGE, vcKey);
    app.visualCrossingKey = vcKey;
  } else {
    app.db.remove(app.VC_API_KEY_STORAGE);
    app.visualCrossingKey = null;
  }

  app.aiDraft = null;
  app.settingsModal.style.display = 'none';
});

app.btnClearCache?.addEventListener('click', () => {
  if (confirm('Clear all cached weather, AI suggestions, and packing lists? Your saved itineraries will be preserved.')) {
    const keysToRemove = [];
    for (let i = 0; i < localStorage.length; i++) {
      const key = localStorage.key(i);
      if (key && (key.startsWith(app.WEATHER_CACHE_PREFIX) || key.startsWith(app.AI_CACHE_PREFIX))) {
        keysToRemove.push(key);
      }
    }
    keysToRemove.forEach(k => app.db.remove(k));
    alert('Cache cleared.');
  }
});

app.aiModalClose?.addEventListener('click', () => {
  app.aiModal.style.display = 'none';
});
window.addEventListener('click', (event) => {
  if (event.target == app.aiModal) {
    app.aiModal.style.display = 'none';
  }
  if (event.target == app.settingsModal) {
    app.aiDraft = null;
    app.settingsModal.style.display = 'none';
  }
});

app.printBtn?.addEventListener('click', () => {
  const collapsed = [];
  document.querySelectorAll('.location-card.collapsed').forEach(card => {
    collapsed.push(card);
    card.classList.remove('collapsed');
  });
  const restore = () => {
    collapsed.forEach(card => card.classList.add('collapsed'));
    window.removeEventListener('afterprint', restore);
  };
  window.addEventListener('afterprint', restore);
  window.print();
});

app.handleSyncClick = async () => {
  if (!app.lastSyncedDocId) return;
  const btns = document.querySelectorAll('.sync-action-btn');
  btns.forEach(btn => { btn.classList.add('syncing'); btn.disabled = true; });
  try {
    await app.ensureSheetsAccess({ interactive: true });
    await app.fetchSheetData(app.lastSyncedDocId, app.h1.textContent, app.lastSyncedGid, true, true);
  } catch (error) {
    if (error && !error.cancelled && !error.silent) app.showError(error.message || 'Google sign-in failed. Try again.');
  } finally {
    btns.forEach(btn => { btn.classList.remove('syncing'); btn.disabled = false; });
  }
};

document.querySelectorAll('.sync-action-btn').forEach(btn => {
  btn.addEventListener('click', app.handleSyncClick);
});

app.mapRouteBtn?.addEventListener('click', () => {
  const url = app.buildGoogleMapsRouteUrl();
  if (url) window.open(url, '_blank', 'noopener,noreferrer');
});

app.shareBtn?.addEventListener('click', async () => {
  const longUrl = window.location.href;
  const urlParams = new URLSearchParams(window.location.search);
  const isDocIdUrl = urlParams.has('docId');

  const originalContent = app.shareBtn.innerHTML;
  const originalTitle = app.shareBtn.title;

  let success = false;
  let urlToCopy = longUrl;

  if (!isDocIdUrl) {
    app.shareBtn.innerHTML = '...';
    app.shareBtn.disabled = true;
    app.shareBtn.title = 'Shortening...';

    try {
      const res = await fetch(`https://tinyurl.com/api-create.php?url=${encodeURIComponent(longUrl)}`);
      if (res.ok) {
        const shortUrl = await res.text();
        if (shortUrl.startsWith('http')) {
          urlToCopy = shortUrl;
        }
      }
    } catch (e) {
      console.error('Could not shorten URL, will copy long URL instead:', e);
    }
  }

  try {
    await navigator.clipboard.writeText(urlToCopy);
    success = true;
  } catch (e) {
    console.error('Failed to copy to clipboard:', e);
  }

  app.shareBtn.innerHTML = success ? '✓' : '✗';
  app.shareBtn.style.color = success ? 'var(--success)' : '#e88';
  app.shareBtn.title = success ? 'Copied!' : 'Failed to copy';

  setTimeout(() => {
    app.shareBtn.innerHTML = originalContent;
    app.shareBtn.disabled = false;
    app.shareBtn.style.color = '';
    app.shareBtn.title = originalTitle;
  }, 2000);
});

app.btnParse?.addEventListener('click', () => app.buildItinerary(app.csvInput.value));

app.btnClearDoc?.addEventListener('click', app.resetApp);

app.csvInput?.addEventListener('paste', (event) => {
  const html = (event.clipboardData || window.clipboardData).getData('text/html');
  if (!html || !html.includes('<table')) {
    return; // Let default paste happen
  }

  event.preventDefault();

  const div = document.createElement('div');
  div.innerHTML = html;
  const table = div.querySelector('table');
  if (!table) {
    const text = (event.clipboardData || window.clipboardData).getData('text/plain');
    app.csvInput.value = text;
    return;
  }

  const grid = [];
  const tableRows = table.querySelectorAll('tr');
  tableRows.forEach((row, rowIndex) => {
    if (!grid[rowIndex]) grid[rowIndex] = [];
    const cells = row.querySelectorAll('td, th');
    let colIndex = 0;
    cells.forEach(cell => {
      while (grid[rowIndex][colIndex] !== undefined) colIndex++;
      const htmlContent = app.simplifyHtml(cell.innerHTML);
      const colspan = parseInt(cell.getAttribute('colspan') || '1', 10);
      const rowspan = parseInt(cell.getAttribute('rowspan') || '1', 10);
      for (let r = 0; r < rowspan; r++) {
        const targetRow = rowIndex + r;
        if (!grid[targetRow]) grid[targetRow] = [];
        for (let c = 0; c < colspan; c++) {
          grid[targetRow][colIndex + c] = htmlContent;
        }
      }
    });
  });

  const maxCols = grid.reduce((max, row) => Math.max(max, row ? row.length : 0), 0);
  const rowsWithHtml = grid.map(row => {
    const newRow = [];
    for (let i = 0; i < maxCols; i++) newRow[i] = (row && row[i] !== undefined) ? row[i] : '';
    return newRow;
  });

  const toCsvValue = (val) => {
    let v = (val || '').toString();
    if (/[\t\n\r"]/.test(v)) {
      return '"' + v.replace(/"/g, '""') + '"';
    }
    return v;
  };

  const csvTextForUrl = rowsWithHtml.map(row => row.map(toCsvValue).join('\t')).join('\n');
  app.csvInput.value = csvTextForUrl;
  app.buildItineraryFromRows(rowsWithHtml, false, csvTextForUrl);
});

document.getElementById('builder-days')?.addEventListener('click', async (e) => {
  const dateBtn = e.target.closest('.day-date-btn');
  if (dateBtn) {
    e.preventDefault();
    const pop = document.getElementById('schedule-popover');
    if (pop && !pop.hidden && app.rangePick?.target === dateBtn) app.closeTripRange();
    else app.openDayDate(dateBtn);
    return;
  }
  const insertBtn = e.target.closest('[data-insert-after]');
  if (insertBtn) {
    const after = Number(insertBtn.dataset.insertAfter);
    const inserted = app.insertBuilderDay(app.collectBuilderDays(), after);
    app.commitBuilderDays(inserted);
    document.querySelector(`#builder-days .builder-day[data-index="${after + 1}"] [data-field="location"]`)?.focus();
    return;
  }
  const btn = e.target.closest('[data-action]');
  if (!btn || btn.disabled) return;
  const card = btn.closest('.builder-day');
  if (!card) return;
  const index = Number(card.dataset.index);
  const action = btn.dataset.action;
  if (action === 'ai-day') {
    app.openAiDayForm(index);
    return;
  }
  const days = app.collectBuilderDays();
  if (action === 'remove-day') {
    if (index === days.length - 1) {
      app.commitBuilderDays(app.applyDayRemoval(days, index, 'shift'));
      return;
    }
    const choice = await app.askRemoveDay(days[index], index);
    if (!choice) return;
    app.commitBuilderDays(app.applyDayRemoval(days, index, choice));
    return;
  } else if (action === 'move-up' && index > 0) {
    const [item] = days.splice(index, 1);
    days.splice(index - 1, 0, item);
  } else if (action === 'move-down' && index < days.length - 1) {
    const [item] = days.splice(index, 1);
    days.splice(index + 1, 0, item);
  } else if (action === 'remove-activity') {
    days[index].activities.splice(Number(btn.dataset.index), 1);
    if (!days[index].activities.length) days[index].activities.push({ time: '', content: '' });
  } else {
    return;
  }
  app.renderBuilderDays(days);
});

document.getElementById('schedule-range-btn')?.addEventListener('click', () => {
  const pop = document.getElementById('schedule-popover');
  if (pop && !pop.hidden && app.rangePick?.mode !== 'day') app.closeTripRange();
  else app.openTripRange();
});

document.getElementById('ai-trip-range-btn')?.addEventListener('click', () => {
  const pop = document.getElementById('schedule-popover');
  if (pop && !pop.hidden && app.rangePick?.mode === 'ai') app.closeTripRange();
  else app.openAiTripRange();
});

document.getElementById('schedule-prev')?.addEventListener('click', () => app.shiftTripRangeMonth(-1));
document.getElementById('schedule-next')?.addEventListener('click', () => app.shiftTripRangeMonth(1));
document.getElementById('schedule-month')?.addEventListener('click', (event) => {
  event.stopPropagation();
  app.zoomTripRange();
});

document.getElementById('schedule-grid')?.addEventListener('click', (event) => {
  const month = event.target.closest('[data-month]');
  if (month && app.rangePick?.zoom === 'month') {
    event.stopPropagation();
    app.pickTripMonth(Number(month.dataset.month));
    return;
  }
  const year = event.target.closest('[data-year]');
  if (year && app.rangePick?.zoom === 'year') {
    event.stopPropagation();
    app.pickTripYear(Number(year.dataset.year));
    return;
  }
  const day = event.target.closest('[data-ymd]');
  if (!day) return;
  event.stopPropagation();
  app.pickTripRangeDay(day.dataset.ymd);
});

document.getElementById('schedule-grid')?.addEventListener('pointerover', (event) => {
  const day = event.target.closest('[data-ymd]');
  if (!day || app.rangePick?.mode === 'day' || !app.rangePick?.anchor || app.rangePick.hover === day.dataset.ymd) return;
  app.rangePick.hover = day.dataset.ymd;
  app.paintTripRangeSelection();
  const hint = document.getElementById('schedule-hint');
  if (hint) hint.textContent = 'Choose an end date';
});

document.addEventListener('click', (event) => {
  const timePop = document.getElementById('time-popover');
  const timeTrigger = event.target.closest('.time-trigger');
  if (timeTrigger) {
    if (timePop && !timePop.hidden && app.timePick?.target === timeTrigger) app.closeTimePicker();
    else app.openTimePicker(timeTrigger);
  } else if (timePop && !timePop.hidden && !timePop.contains(event.target)) {
    app.closeTimePicker();
  }

  const pop = document.getElementById('schedule-popover');
  const btn = document.getElementById('schedule-range-btn');
  if (!pop || pop.hidden) return;
  if (pop.contains(event.target) || btn?.contains(event.target) || event.target.closest('.day-date-btn') || event.target.closest('#ai-trip-range-btn')) return;
  app.closeTripRange();
});

document.addEventListener('keydown', (event) => {
  const timePop = document.getElementById('time-popover');
  const confirm = document.getElementById('builder-confirm-modal');
  if (event.key !== 'Escape' || confirm?.classList.contains('open')) return;
  if (timePop && !timePop.hidden) {
    app.closeTimePicker();
    return;
  }
  const pop = document.getElementById('schedule-popover');
  if (!pop || pop.hidden) return;
  if (app.rangePick?.zoom === 'year') {
    app.rangePick.zoom = 'month';
    app.renderTripRange();
    return;
  }
  if (app.rangePick?.zoom === 'month') {
    app.rangePick.zoom = 'day';
    app.renderTripRange();
    return;
  }
  const keepDialog = app.rangePick?.mode === 'ai';
  app.closeTripRange();
  if (keepDialog) event.stopImmediatePropagation();
});

function placeOpenSchedule() {
  const pop = document.getElementById('schedule-popover');
  if (!pop || pop.hidden || !app.rangePick) return;
  const anchor = app.rangePick.target && app.rangePick.target.isConnected
    ? app.rangePick.target
    : document.getElementById('schedule-range-btn');
  if (anchor) app.placeSchedulePopover(anchor);
}

function placeOpenTime() {
  const pop = document.getElementById('time-popover');
  if (!pop || pop.hidden || !app.timePick?.target) return;
  app.placeTimePopover(app.timePick.target);
}

window.addEventListener('scroll', placeOpenSchedule, true);
window.addEventListener('resize', placeOpenSchedule);
window.addEventListener('scroll', placeOpenTime, true);
window.addEventListener('resize', placeOpenTime);

document.addEventListener('keydown', (event) => {
  const pop = document.getElementById('time-popover');
  if (!pop || pop.hidden) return;
  if (!['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(event.key)) return;
  if (event.target.closest('input, textarea')) return;
  event.preventDefault();
  app.nudgeTimeWheel(event.key);
});

document.getElementById('ai-build-close')?.addEventListener('click', () => {
  if (!app.aiBuildBusy) app.setAiPanel('');
});
document.getElementById('ai-build-modal')?.addEventListener('click', (event) => {
  if (event.target.id === 'ai-build-modal' && !app.aiBuildBusy) app.setAiPanel('');
});
document.addEventListener('keydown', (event) => {
  const modal = document.getElementById('ai-build-modal');
  const confirm = document.getElementById('builder-confirm-modal');
  if (event.key !== 'Escape' || !modal?.classList.contains('open') || confirm?.classList.contains('open') || app.aiBuildBusy) return;
  app.setAiPanel('');
});

document.getElementById('btn-ai-trip')?.addEventListener('click', () => app.openAiTripForm());
document.getElementById('ai-trip-form')?.addEventListener('submit', (event) => {
  event.preventDefault();
  app.submitAiTrip();
});
document.getElementById('ai-day-form')?.addEventListener('submit', (event) => {
  event.preventDefault();
  app.submitAiDay();
});
document.getElementById('ai-trip-form')?.addEventListener('input', () => app.setAiStatus('ai-trip-status', ''));
document.getElementById('ai-day-form')?.addEventListener('input', () => app.setAiStatus('ai-day-status', ''));
document.querySelectorAll('.ai-pace').forEach(group => {
  group.addEventListener('click', (event) => {
    const paceBtn = event.target.closest('[data-pace]');
    if (!paceBtn || !group.contains(paceBtn)) return;
    group.querySelectorAll('[data-pace]').forEach(el => {
      const on = el === paceBtn;
      el.classList.toggle('active', on);
      el.setAttribute('aria-checked', on ? 'true' : 'false');
    });
    const form = group.closest('form');
    if (form?.id === 'ai-trip-form') app.setAiStatus('ai-trip-status', '');
    if (form?.id === 'ai-day-form') app.setAiStatus('ai-day-status', '');
  });
});

document.getElementById('btn-builder-apply')?.addEventListener('click', app.applyBuilder);
document.getElementById('btn-builder-save-sheet')?.addEventListener('click', app.saveBuilderToSheet);
document.getElementById('btn-builder-fork')?.addEventListener('click', app.applyBuilder);
document.getElementById('btn-builder-cancel')?.addEventListener('click', app.cancelEditor);

app.tabs.forEach(tab => {
  tab.addEventListener('click', () => {
    app.showError('');
    app.showInputTab(tab.dataset.tab);
  });
});

// View toggles
app.viewBtns.forEach(btn => {
  btn.addEventListener('click', () => {
    const view = btn.dataset.view;
    app.viewBtns.forEach(b => b.classList.remove('active'));
    btn.classList.add('active');

    const searchWrapper = document.querySelector('.search-wrapper');
    if (searchWrapper) searchWrapper.style.display = view === 'list' ? '' : 'none';

    document.getElementById('itinerary-list').classList.remove('visible');
    document.getElementById('itinerary-map').classList.remove('visible');
    document.getElementById('itinerary-calendar').classList.remove('visible');
    document.getElementById('itinerary-packing').classList.remove('visible');
    if (view !== 'builder') {
      app.parkBuilderPanel();
      if (app.builderEditing) app.setBuilderMode('create');
    }

    if (view === 'builder') {
      app.enterBuilderMode();
    } else if (view === 'map') {
      document.getElementById('itinerary-map').classList.add('visible');
      if (app.map) {
        app.map.invalidateSize();
        if (app.mapPoints.length > 0) {
          const bounds = L.latLngBounds(app.mapPoints);
          app.map.fitBounds(bounds, { padding: [50, 50], maxZoom: 12 });
        }
      }
    } else if (view === 'calendar') {
      document.getElementById('itinerary-calendar').classList.add('visible');
    } else if (view === 'packing') {
      document.getElementById('itinerary-packing').classList.add('visible');
      // Load persisted count
      const savedCount = app.db.get(app.PACKING_COUNT_STORAGE);
      if (savedCount && app.packingCountInput) app.packingCountInput.value = savedCount;

      if (app.parsed) {
        // Reset to correct count logic based on active tab
        const activeTab = document.querySelector('.packing-tab.active');
        app.currentChecklistType = activeTab ? activeTab.dataset.type : 'packing';
        document.getElementById('packing-inputs-wrapper').style.display = app.currentChecklistType === 'packing' ? 'flex' : 'none';

        const count = app.currentChecklistType === 'packing' ? (app.packingCountInput?.value || 1) : 1;
        const key = app.getChecklistCacheKey(app.currentChecklistType, count);
        app.currentAiCacheKey = key;

        let contentFound = false;
        try {
          const cachedJson = app.db.get(key);
          if (cachedJson) {
            const cached = JSON.parse(cachedJson);
            if (Date.now() - cached.timestamp < app.AI_CACHE_EXPIRY) {
              app.packingContent.innerHTML = app.renderMarkdown(cached.text, key);
              contentFound = true;
            }
          }
        } catch (e) { }

        if (!contentFound) {
          app.packingContent.innerHTML = '<p class="muted">Regenerate to create this list.</p>';
        }
      }
    } else if (view === 'list') {
      document.getElementById('itinerary-list').classList.add('visible');

      if (!app.hasScrolledToToday) {
        const activeDay = document.querySelector('.day-group.today-active');
        const activeCard = document.querySelector('.location-card.expanded');
        if (activeDay || activeCard) {
          setTimeout(() => {
            (activeDay || activeCard).scrollIntoView({ behavior: 'smooth', block: activeDay ? 'center' : 'start' });
          }, 100);
          app.hasScrolledToToday = true;
        }
      }
    }
  });
});

// Reset functionality
app.h1.title = 'Clear data and start over';
app.h1.style.cursor = 'pointer';
app.h1?.addEventListener('click', app.resetApp);

// Initial load
(async () => {
  app.syncAiAvailability();
  if (app.currentAiKey()) app.populateModels();

  // Reuse a token from this browser before asking Google. A silent refresh
  // never opens a popup; the sign-in button is what may show Google's window.
  if (app.restoreGoogleSession()) app.updateGoogleStatus(true);

  let gsiWaits = 0;
  const waitGsi = setInterval(() => {
    gsiWaits += 1;
    const ready = typeof google !== 'undefined' && google.accounts && google.accounts.oauth2;
    if (!ready && gsiWaits < 50) return;
    clearInterval(waitGsi);
    if (!ready) return;
    app.initTokenClient();
    if (app.googleTokenFresh() || app.db.get(app.GOOGLE_SIGNED_IN_KEY) !== 'true') return;
    app.updateGoogleStatus(false, true);
    app.ensureSheetsAccess({ interactive: false }).then(() => {
      if (new URLSearchParams(window.location.search).get('docId')) app.loadFromUrl();
    }).catch(() => {
      if (!app.googleTokenFresh()) app.updateGoogleStatus(false);
    });
  }, 200);

  app.recentDocsSort?.addEventListener('change', app.renderRecentDocs);

  window.addEventListener('popstate', app.loadFromUrl);
  window.addEventListener('storage', app.onGoogleAuthStorage);
  app.renderRecentDocs();
  app.loadFromUrl(); // Load itinerary data from URL if present
})();

// Offline support & Service Worker Registration
app.offlineIndicator = document.getElementById('offline-indicator');

window.addEventListener('online', app.updateOnlineStatus);
window.addEventListener('offline', app.updateOnlineStatus);
app.updateOnlineStatus();

if ('serviceWorker' in navigator) {
  navigator.serviceWorker.register('sw.js', { updateViaCache: 'none' })
    .catch(err => console.log('Service Worker registration failed', err));

  // Automatically reload the page when a new service worker takes over
  let refreshing = false;
  navigator.serviceWorker.addEventListener('controllerchange', () => {
    if (!refreshing) {
      window.location.reload();
      refreshing = true;
    }
  });
}
}
