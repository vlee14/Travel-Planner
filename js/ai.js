// Suggestions and checklists through Google, Grok, or OpenRouter.
export function installAi(app) {
  app.AI_PROVIDERS = {
    google: {
      label: 'Google',
      keyLabel: 'Google AI',
      href: 'https://aistudio.google.com/app/apikey',
      missing: 'Please set your Google AI API key in Settings first.',
      note: 'The key stays in this browser and is sent only to Google.',
      storageKey: 'gemini_api_key',
      modelKey: 'gemini_model',
      defaultModel: 'gemini-flash-latest'
    },
    grok: {
      label: 'Grok',
      keyLabel: 'xAI',
      href: 'https://console.x.ai/',
      missing: 'Please set your xAI API key in Settings first.',
      note: 'The key stays in this browser and is sent only to xAI.',
      storageKey: 'grok_api_key',
      modelKey: 'grok_model',
      defaultModel: 'grok-3'
    },
    openrouter: {
      label: 'OpenRouter',
      keyLabel: 'OpenRouter',
      href: 'https://openrouter.ai/keys',
      missing: 'Please set your OpenRouter API key in Settings first.',
      note: 'The key stays in this browser and is sent only to OpenRouter.',
      storageKey: 'openrouter_api_key',
      modelKey: 'openrouter_model',
      defaultModel: 'openrouter/auto'
    }
  };

  app.currentAiKey = function currentAiKey() {
    return ((app.aiKeys && app.aiKeys[app.aiProvider]) || '').trim();
  };

  app.currentAiModel = function currentAiModel() {
    const meta = app.AI_PROVIDERS[app.aiProvider];
    return ((app.aiModels && app.aiModels[app.aiProvider]) || (meta && meta.defaultModel) || '').trim();
  };

  app.aiKeyMessage = function aiKeyMessage() {
    const meta = app.AI_PROVIDERS[app.aiProvider];
    return meta ? meta.missing : 'Please set an API key in Settings first.';
  };

  app.syncAiAvailability = function syncAiAvailability() {
    document.body.classList.toggle('has-api-key', Boolean(app.currentAiKey()));
  };

  app.geminiChatModelName = function geminiChatModelName(model) {
    const name = (model.name || '').replace(/^models\//, '');
    if (!name.includes('gemini')) return null;
    const methods = model.supportedGenerationMethods || model.supportedActions || [];
    if (methods.length > 0 && !methods.includes('generateContent')) return null;
    if (/embed|image|imagen|tts|live|aqa|vision|transcribe|banana|omni/i.test(name)) return null;
    return name;
  };

  app.grokChatModelId = function grokChatModelId(modelId) {
    const name = (modelId || '').trim();
    if (!name || /embed|image|vision|tts|video|flux/i.test(name)) return null;
    return name;
  };

  app.openRouterChatModelId = function openRouterChatModelId(model) {
    const name = (model.id || '').trim();
    if (!name || /embed/i.test(name)) return null;
    const outputs = model.architecture && model.architecture.output_modalities;
    if (outputs && outputs.length > 0 && !outputs.includes('text')) return null;
    const modality = (model.architecture && model.architecture.modality) || '';
    if (modality && !/->text|text\+/.test(modality) && !modality.endsWith('text')) return null;
    return name;
  };

  app.chooseProviderModel = function chooseProviderModel(provider, names, saved) {
    const keep = (saved || '').trim();
    const fallback = app.AI_PROVIDERS[provider].defaultModel;
    if (keep && (names.includes(keep) || names.length === 0)) return keep;
    if (provider === 'google') {
      if (names.includes('gemini-flash-latest')) return 'gemini-flash-latest';
      const flash = names.find(name => name.includes('flash'));
      if (flash) return flash;
    }
    if (provider === 'grok') {
      if (fallback && names.includes(fallback)) return fallback;
      const named = names.filter(name => /^grok([.-]|$)/i.test(name));
      if (named.length) return named[named.length - 1];
    }
    if (fallback && names.includes(fallback)) return fallback;
    return names[0] || fallback;
  };

  app.geminiThinkingConfig = function geminiThinkingConfig(model) {
    const name = (model || '').toLowerCase();
    if (!name.includes('gemini')) return null;
    if (name.includes('gemini-3') || name.endsWith('-latest')) return { thinkingLevel: 'low' };
    if (name.includes('gemini-2.5')) return { thinkingBudget: 0 };
    return null;
  };

  app.readApiError = function readApiError(data, status, label) {
    const record = data && typeof data === 'object' ? data : {};
    const error = record.error;
    let message = '';
    if (typeof error === 'string') message = error;
    else if (error && typeof error === 'object' && typeof error.message === 'string') message = error.message;
    if (status === 401 || status === 403) return label + ' rejected the API key.';
    const clean = message.replace(/\s+/g, ' ').trim().slice(0, 240);
    return clean || (label + ' returned ' + status + '.');
  };

  app.scrubSecret = function scrubSecret(text, secret) {
    if (!secret || secret.length < 6 || !text.includes(secret)) return text;
    return text.split(secret).join('…');
  };

  app.requestAi = async function requestAi(label, url, init, secret, timeoutMs) {
    try {
      const response = await fetch(url, { ...init, signal: AbortSignal.timeout(timeoutMs) });
      const text = await response.text();
      let data = null;
      if (text) {
        try {
          data = JSON.parse(text);
        } catch (e) {
          data = { error: { message: text.slice(0, 240) } };
        }
      }
      if (!response.ok) {
        const error = new Error(app.scrubSecret(app.readApiError(data, response.status, label), secret));
        error.status = response.status;
        throw error;
      }
      return data || {};
    } catch (err) {
      if (err && err.status) throw err;
      if (err && (err.name === 'TimeoutError' || err.name === 'AbortError')) {
        throw new Error('Request timed out waiting for ' + label + '.');
      }
      throw new Error('Could not reach ' + label + '.');
    }
  };

  app.historyToChatMessages = function historyToChatMessages(history) {
    return (history || []).map(msg => ({
      role: msg.role === 'model' || msg.role === 'assistant' ? 'assistant' : 'user',
      content: (msg.parts || []).map(part => part.text || '').join('')
    })).filter(msg => msg.content.trim());
  };

  app.geminiReplyText = function geminiReplyText(data) {
    const block = data.promptFeedback && data.promptFeedback.blockReason;
    if (block) throw new Error('Google blocked the request (' + block + ').');
    const candidate = data.candidates && data.candidates[0];
    const parts = (candidate && candidate.content && candidate.content.parts) || [];
    const text = parts
      .filter(part => part && !part.thought && typeof part.text === 'string')
      .map(part => part.text)
      .join('');
    if (text.trim()) return text;
    if (candidate && candidate.finishReason && candidate.finishReason !== 'STOP') {
      return 'Could not generate suggestions. Reason: ' + candidate.finishReason;
    }
    throw new Error('Google returned an empty reply.');
  };

  app.chatReplyText = function chatReplyText(data, label) {
    const choice = data.choices && data.choices[0];
    const content = choice && choice.message && choice.message.content;
    if (typeof content === 'string' && content.trim()) return content;
    if (Array.isArray(content)) {
      const text = content.map(part => {
        if (typeof part === 'string') return part;
        if (!part || part.thought) return '';
        const kind = String(part.type || '').toLowerCase();
        if (kind === 'reasoning' || kind === 'thinking' || kind === 'thought') return '';
        return part.text || '';
      }).join('');
      if (text.trim()) return text;
    }
    throw new Error(label + ' returned an empty reply.');
  };

  app.completeGoogle = async function completeGoogle(key, model, history, maxTokens) {
    const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(model) + ':generateContent';
    const send = thinking => {
      const generationConfig = {
        temperature: 0.7,
        topK: 1,
        topP: 1,
        maxOutputTokens: maxTokens
      };
      if (thinking) {
        const config = app.geminiThinkingConfig(model);
        if (config) generationConfig.thinkingConfig = config;
      }
      return app.requestAi('Google', url, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'x-goog-api-key': key
        },
        body: JSON.stringify({ contents: history, generationConfig })
      }, key, 120000);
    };
    try {
      return app.geminiReplyText(await send(true));
    } catch (err) {
      if (err.status === 400 && /thinking/i.test(err.message)) {
        return app.geminiReplyText(await send(false));
      }
      throw err;
    }
  };

  app.completeGrok = async function completeGrok(key, model, history, maxTokens) {
    const send = function(useSearch) {
      const body = {
        model,
        messages: app.historyToChatMessages(history),
        temperature: 0.7,
        stream: false,
        max_completion_tokens: maxTokens
      };
      if (useSearch) body.search_parameters = { mode: 'off' };
      return app.requestAi('Grok', 'https://api.x.ai/v1/chat/completions', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          Authorization: 'Bearer ' + key
        },
        body: JSON.stringify(body)
      }, key, 120000);
    };
    try {
      return app.chatReplyText(await send(true), 'Grok');
    } catch (err) {
      if (err.status === 400 && /search/i.test(err.message)) {
        return app.chatReplyText(await send(false), 'Grok');
      }
      throw err;
    }
  };

  app.openRouterHeaders = function openRouterHeaders(key) {
    const headers = {
      Authorization: 'Bearer ' + key,
      'X-Title': 'Travel Planner'
    };
    if (location.origin) headers['HTTP-Referer'] = location.origin;
    return headers;
  };

  app.completeOpenRouter = function completeOpenRouter(key, model, history, maxTokens) {
    return app.requestAi('OpenRouter', 'https://openrouter.ai/api/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        ...app.openRouterHeaders(key)
      },
      body: JSON.stringify({
        model,
        messages: app.historyToChatMessages(history),
        temperature: 0.7,
        stream: false,
        max_tokens: maxTokens
      })
    }, key, 120000).then(data => app.chatReplyText(data, 'OpenRouter'));
  };

  app.completeAi = function completeAi(history, maxTokens) {
    const provider = app.aiProvider;
    const key = app.currentAiKey();
    const meta = app.AI_PROVIDERS[provider];
    if (!key) throw new Error(meta ? meta.missing : 'Please set an API key in Settings first.');
    const model = app.currentAiModel();
    if (provider === 'google') return app.completeGoogle(key, model, history, maxTokens);
    if (provider === 'grok') return app.completeGrok(key, model, history, maxTokens);
    return app.completeOpenRouter(key, model, history, maxTokens);
  };

  app.listGoogleModels = async function listGoogleModels(key) {
    const found = [];
    let pageToken = '';
    for (let page = 0; page < 10; page++) {
      const url = new URL('https://generativelanguage.googleapis.com/v1beta/models');
      url.searchParams.set('pageSize', '100');
      if (pageToken) url.searchParams.set('pageToken', pageToken);
      const data = await app.requestAi('Google', url.toString(), {
        method: 'GET',
        headers: { 'x-goog-api-key': key }
      }, key, 15000);
      (data.models || []).forEach(model => {
        const name = app.geminiChatModelName(model);
        if (!name) return;
        found.push({
          id: name,
          label: model.displayName ? model.displayName + ' (' + name + ')' : name
        });
      });
      pageToken = data.nextPageToken || '';
      if (!pageToken) break;
    }
    const seen = new Set();
    return found.filter(model => {
      if (seen.has(model.id)) return false;
      seen.add(model.id);
      return true;
    }).sort((a, b) => b.id.localeCompare(a.id));
  };

  app.listGrokModels = async function listGrokModels(key) {
    const data = await app.requestAi('Grok', 'https://api.x.ai/v1/models', {
      method: 'GET',
      headers: { Authorization: 'Bearer ' + key }
    }, key, 15000);
    const rows = data.data || data.models || [];
    const seen = new Set();
    return rows.map(row => {
      const id = app.grokChatModelId(row.id || row.name || '');
      return id ? { id, label: id } : null;
    }).filter(model => {
      if (!model || seen.has(model.id)) return false;
      seen.add(model.id);
      return true;
    }).sort((a, b) => a.id.localeCompare(b.id));
  };

  app.listOpenRouterModels = async function listOpenRouterModels(key) {
    const data = await app.requestAi('OpenRouter', 'https://openrouter.ai/api/v1/models', {
      method: 'GET',
      headers: app.openRouterHeaders(key)
    }, key, 15000);
    const seen = new Set();
    return (data.data || []).map(row => {
      const id = app.openRouterChatModelId(row);
      if (!id) return null;
      return { id, label: row.name ? row.name + ' (' + id + ')' : id };
    }).filter(model => {
      if (!model || seen.has(model.id)) return false;
      seen.add(model.id);
      return true;
    }).sort((a, b) => a.id.localeCompare(b.id));
  };

  app.setModelOptions = function setModelOptions(models, selected) {
    if (!app.settingModel) return;
    const select = app.settingModel;
    select.innerHTML = '';
    const seen = new Set();
    models.forEach(model => {
      if (!model || !model.id || seen.has(model.id)) return;
      seen.add(model.id);
      const opt = document.createElement('option');
      opt.value = model.id;
      opt.textContent = model.label || model.id;
      select.appendChild(opt);
    });
    const value = selected || app.AI_PROVIDERS[app.aiProvider].defaultModel;
    if (value && !seen.has(value)) {
      const opt = document.createElement('option');
      opt.value = value;
      opt.textContent = value;
      select.insertBefore(opt, select.firstChild);
    }
    if (value) select.value = value;
    if (app.refreshMenu) app.refreshMenu(select);
  };

  app.parseSuggestionBlocks = function parseSuggestionBlocks(markdown) {
    const text = markdown || '';
    if (!/^###\s+/m.test(text)) return null;
    const parts = text.split(/^###\s+/m);
    const blocks = parts.slice(1).map(block => {
      const lines = block.replace(/\s+$/, '').split('\n');
      const title = lines[0].replace(/[*_`]/g, '').trim();
      const body = lines.slice(1).join('\n').trim();
      const timeMatch = body.match(/^\s*`?Time:\s*([0-9]{1,2}(?::[0-9]{2})?\s*(?:AM|PM)?)`?/im);
      return {
        title,
        body: timeMatch ? body.replace(timeMatch[0], '').trim() : body,
        time: timeMatch ? timeMatch[1].trim() : ''
      };
    }).filter(block => block.title);
    if (!blocks.length) return null;
    return { preamble: parts[0].trim(), blocks };
  };

  app.renderModelMessage = function renderModelMessage(text) {
    const div = document.createElement('div');
    div.className = 'ai-message model';
    const blocks = app.suggestionTarget ? app.parseSuggestionBlocks(text) : null;
    if (!blocks) {
      div.innerHTML = app.renderMarkdown(text);
      return div;
    }
    if (blocks.preamble) {
      const pre = document.createElement('div');
      pre.innerHTML = app.renderMarkdown(blocks.preamble);
      div.appendChild(pre);
    }
    blocks.blocks.forEach(block => {
      const card = document.createElement('div');
      card.className = 'suggestion-card';
      const heading = document.createElement('h3');
      heading.textContent = block.title;
      card.appendChild(heading);
      if (block.body) {
        const body = document.createElement('div');
        body.innerHTML = app.renderMarkdown(block.body);
        card.appendChild(body);
      }
      const actions = document.createElement('div');
      actions.className = 'suggestion-actions';
      const time = document.createElement('input');
      time.type = 'time';
      time.value = app.timeToInputValue(block.time);
      time.setAttribute('aria-label', `Time for ${block.title}`);
      const add = document.createElement('button');
      add.type = 'button';
      add.className = 'btn btn-primary suggestion-add';
      add.textContent = 'Add to schedule';
      const note = document.createElement('span');
      note.className = 'suggestion-note';
      add.addEventListener('click', () => {
        const err = app.addSuggestionToDay(app.suggestionTarget.location, app.suggestionTarget.date, time.value, block.title);
        note.textContent = err || '';
        if (!err) {
          add.textContent = 'Added';
          add.disabled = true;
          time.disabled = true;
        }
      });
      actions.append(time, add, note);
      card.appendChild(actions);
      div.appendChild(card);
    });
    return div;
  };

  app.addSuggestionToDay = function addSuggestionToDay(location, dateLabel, timeValue, title) {
    const col = app.findDayColumn(location, dateLabel);
    if (col < 1) return 'That day is no longer on the schedule.';
    if (!timeValue) return 'Choose a time first.';
    const cells = Array(Math.max(app.parsed.locations.length - 1, 0)).fill('');
    cells[col - 1] = title;
    let insertAt = app.parsed.rows.length;
    while (insertAt > 0 && !(app.parsed.rows[insertAt - 1].time || '').trim()) insertAt--;
    app.parsed.rows.splice(insertAt, 0, { time: app.formatTimeLabel(timeValue), cells });
    app.detachSheetSync();
    app.render();
    app.fetchAllWeather();
    app.updateUrlWithData(app.rowsToCsv(app.parsedToGrid(app.parsed)));
    const host = document.getElementById('itinerary-builder');
    if (host && host.classList.contains('visible')) {
      const days = app.collectBuilderDays();
      const ymd = app.parseSmartDate(dateLabel) || '';
      const day = days.find(item => item.location === location && item.date === ymd);
      if (day) {
        day.activities = day.activities.filter(act => act.time || act.content);
        day.activities.push({ time: timeValue, content: title });
        app.renderBuilderDays(days);
      }
    }
    return '';
  };

  app.getAiSuggestions = async function getAiSuggestions(location, date, dailyPlan, fullLocationPlan, hotelInfo, weather) {
    if (!app.currentAiKey()) {
      alert(app.aiKeyMessage());
      return;
    }

    app.suggestionTarget = { location, date };
    app.currentAiCacheKey = `${app.AI_CACHE_PREFIX}${encodeURIComponent(location)}_${date}`;

    try {
      const cachedJson = app.db.get(app.currentAiCacheKey);
      if (cachedJson) {
        const cached = JSON.parse(cachedJson);
        if (Date.now() - cached.timestamp < app.AI_CACHE_EXPIRY) {
          app.aiConversationHistory = cached.history;
          app.aiModal.style.display = 'flex';
          app.aiModalTitle.textContent = `AI Suggestions for ${date}`;
          app.aiModalBody.innerHTML = '';

          app.aiConversationHistory.forEach((msg, index) => {
            if (index === 0) return; // Skip initial system prompt

            if (msg.role === 'model') {
              app.aiModalBody.appendChild(app.renderModelMessage(msg.parts[0].text));
            } else {
              const div = document.createElement('div');
              div.className = 'ai-message user';
              div.textContent = msg.parts[0].text;
              app.aiModalBody.appendChild(div);
            }
          });
          app.aiModalBody.scrollTop = 0;
          return;
        }
      }
    } catch (e) { console.warn('Cache read error', e); }

    app.aiModal.style.display = 'flex';
    app.aiModalTitle.textContent = `AI Suggestions for ${date}`;
    app.aiModalBody.innerHTML = ''; // Clear previous conversation
    app.aiConversationHistory = []; // Reset history

    const prompt = `You are an expert local travel guide and itinerary planner.
  I am planning a trip to **${location}**.
  ${hotelInfo ? `Accommodation: ${hotelInfo}` : ''}
  ${weather ? `Expected Weather: ${weather}` : ''}

  **Context: My full itinerary for ${location} is:**
  ${fullLocationPlan}

  **Focus Day: ${date}**
  My current plan for this specific day is:
  ${dailyPlan || "No specific activities planned yet."}

  **Task:**
  Please provide **3 specific, actionable, and distinct suggestions** to enhance or fill my plan for **${date}**.
  ${weather ? `Take the expected weather (${weather}) into account (e.g. prefer indoor activities if raining).` : ''}
  Consider logistical efficiency (places near my current activities), distinct experiences I might be missing, or highly-rated dining/relaxation spots nearby.

  **Constraints:**
  - Do NOT suggest attractions I am already visiting on this day or other days.
  - Ensure suggestions are geographically feasible.
  - Keep descriptions concise.

  **Format:**
  Use Markdown. For each suggestion, use an H3 (###) for the place name. On the next line write \`Time: 2:00 PM\` with a suggested start time that fits around the existing plan. Then a brief description and a specific reason why it fits this itinerary. Include a Markdown link to search for the location on Google Maps (e.g. Open in Maps).`;

    app.aiConversationHistory.push({
      role: 'user',
      parts: [{ text: prompt }]
    });

    await app.callAi();
  };

  app.callAi = async function callAi() {
    const loadingDiv = document.createElement('div');
    loadingDiv.className = 'ai-message loading';
    loadingDiv.textContent = 'Generating...';
    app.aiModalBody.appendChild(loadingDiv);
    app.aiModalBody.scrollTop = app.aiModalBody.scrollHeight;

    app.aiFollowUpInput.disabled = true;
    app.aiFollowUpBtn.disabled = true;

    try {
      const responseText = await app.completeAi(app.aiConversationHistory, 800);
      loadingDiv.remove();

      app.aiConversationHistory.push({ role: 'model', parts: [{ text: responseText }] });

      if (app.currentAiCacheKey) {
        app.db.set(app.currentAiCacheKey, JSON.stringify({
          timestamp: Date.now(),
          history: app.aiConversationHistory
        }));
      }

      app.aiModalBody.appendChild(app.renderModelMessage(responseText));
    } catch (error) {
      loadingDiv.remove();
      console.error('AI call failed:', error);
      const errorDiv = document.createElement('div');
      errorDiv.className = 'ai-message model';
      const title = document.createElement('p');
      title.style.color = 'var(--danger)';
      const strong = document.createElement('strong');
      strong.textContent = 'Error: ';
      title.append(strong, document.createTextNode(error.message || 'The request failed.'));
      const hint = document.createElement('p');
      hint.textContent = 'Check the API key for this provider and your network connection.';
      errorDiv.append(title, hint);
      app.aiModalBody.appendChild(errorDiv);
    } finally {
      app.aiFollowUpInput.disabled = false;
      app.aiFollowUpBtn.disabled = false;
      app.aiFollowUpInput.focus();
    }
  };

  app.getChecklistCacheKey = function getChecklistCacheKey(type, count) {
    if (!app.parsed) return null;
    const locs = [...new Set(app.parsed.locations.slice(1).filter(l => l && l.trim()))].join(', ');
    const validDates = app.parsed.dates.slice(1).filter(d => d && d.trim());
    const start = validDates.length > 0 ? validDates[0] : '';
    return `${app.AI_CACHE_PREFIX}${type}_${encodeURIComponent(locs)}_${start}_${count}`;
  };

  app.generateChecklist = async function generateChecklist(ignoreCache = false) {
    if (!app.parsed) {
      app.packingContent.innerHTML = '<p>Please build an itinerary first.</p>';
      return;
    }

    // Use 1 for apps to normalize cache key, as app suggestions rarely depend on traveler count
    const count = app.currentChecklistType === 'packing' ? (app.packingCountInput.value || 1) : 1;
    if (app.currentChecklistType === 'packing') {
      app.db.set(app.PACKING_COUNT_STORAGE, count);
    }

    app.packingContent.innerHTML = '<p>Generating checklist... <span class="ai-message loading">✨</span></p>';

    const locs = app.parsed ? [...new Set(app.parsed.locations.slice(1).filter(l => l && l.trim()))].join(', ') : '';
    const validDates = app.parsed ? app.parsed.dates.slice(1).filter(d => d && d.trim()) : [];
    const start = validDates.length > 0 ? validDates[0] : '';
    const end = validDates.length > 0 ? validDates[validDates.length - 1] : '';

    let prompt = '';
    if (app.currentChecklistType === 'packing') {
      const weatherSummaries = [];
      document.querySelectorAll('.location-card').forEach(card => {
        const l = card.dataset.location;
        const wInline = card.querySelector('.weather-summary-inline');
        if (l && wInline) weatherSummaries.push(`${l}: ${wInline.textContent}`);
      });
      const weatherContext = weatherSummaries.join('; ');
      const allContent = app.parsed ? app.parsed.rows.map(r => r.cells.join(' ')).join(' ').substring(0, 3000) : '';

      prompt = `Act as a travel expert. Create a comprehensive packing checklist for ${count} person(s) for a trip to ${locs}.\nDates: ${start} to ${end}.\n\nContext:\n- Weather Forecast: ${weatherContext || 'Not available'}\n- Activities overview: ${allContent}\n\nFormat the response as a Markdown checklist categorized by Essentials, Clothing, Toiletries, Electronics, and Misc. Include quantity suggestions based on the trip duration.`;
    } else {
      prompt = `Act as a travel expert. Create a checklist of essential and useful mobile apps to download for a trip to ${locs}.\nInclude categories like Navigation, Transport, Language, Finance/Payment, and Local Guides. Explain briefly why each app is useful for this specific destination.\nFormat as a Markdown checklist.`;
    }

    app.currentAiCacheKey = app.getChecklistCacheKey(app.currentChecklistType, count);

    // Check Cache
    if (!ignoreCache) {
      try {
        const cachedJson = app.db.get(app.currentAiCacheKey);
        if (cachedJson) {
          const cached = JSON.parse(cachedJson);
          if (Date.now() - cached.timestamp < app.AI_CACHE_EXPIRY) {
            app.packingContent.innerHTML = app.renderMarkdown(cached.text, app.currentAiCacheKey);
            return;
          }
        }
      } catch (e) { }
    }

    if (!app.currentAiKey()) {
      app.packingContent.innerHTML = '';
      const message = document.createElement('p');
      message.className = 'error-msg visible';
      message.textContent = app.aiKeyMessage();
      app.packingContent.appendChild(message);
      return;
    }

    try {
      const text = await app.completeAi([{ role: 'user', parts: [{ text: prompt }] }], 2000);
      app.packingContent.innerHTML = app.renderMarkdown(text, app.currentAiCacheKey);
      app.db.set(app.currentAiCacheKey, JSON.stringify({
        timestamp: Date.now(),
        text: text,
        history: []
      }));
    } catch (e) {
      console.error(e);
      app.packingContent.innerHTML = '';
      const message = document.createElement('p');
      message.className = 'error-msg visible';
      message.textContent = 'Error generating list: ' + (e.message || 'The request failed.');
      app.packingContent.appendChild(message);
    }
  };

  app.handleFollowUp = async function handleFollowUp() {
    const prompt = app.aiFollowUpInput.value.trim();
    if (!prompt) return;

    app.aiFollowUpInput.value = '';

    // Render user message
    const messageDiv = document.createElement('div');
    messageDiv.className = 'ai-message user';
    messageDiv.textContent = prompt;
    app.aiModalBody.appendChild(messageDiv);
    app.aiModalBody.scrollTop = app.aiModalBody.scrollHeight;

    // Add to history
    app.aiConversationHistory.push({
      role: 'user',
      parts: [{ text: prompt }]
    });

    await app.callAi();
  };

  app.populateModels = async function populateModels(keyOverride) {
    const draft = app.aiDraft;
    const provider = draft ? draft.provider : app.aiProvider;
    const meta = app.AI_PROVIDERS[provider];
    if (!meta) return;
    const key = (keyOverride != null ? keyOverride : ((draft ? draft.keys[provider] : app.aiKeys[provider]) || '')).trim();
    if (!key) return;
    const requestId = (app.modelListRequest = (app.modelListRequest || 0) + 1);
    try {
      const listed = provider === 'google'
        ? await app.listGoogleModels(key)
        : provider === 'grok'
          ? await app.listGrokModels(key)
          : await app.listOpenRouterModels(key);
      if (requestId !== app.modelListRequest) return;
      const names = listed.map(model => model.id);
      const stored = app.db.get(meta.modelKey) || '';
      const preferred = ((draft && draft.models[provider]) || stored || (app.aiModels && app.aiModels[provider]) || '').trim();
      let selected = preferred;
      if (!(preferred && (names.includes(preferred) || names.length === 0))) {
        selected = stored && preferred === stored
          ? stored
          : app.chooseProviderModel(provider, names, '');
      }
      const visible = app.aiDraft;
      const visibleProvider = visible ? visible.provider : app.aiProvider;
      if (visible && visible.provider === provider && app.settingModel && names.includes(app.settingModel.value)) {
        selected = app.settingModel.value;
      }
      if (visibleProvider !== provider) return;
      if (visible) visible.models[provider] = selected;
      else {
        app.aiModels[provider] = selected;
        if (provider === 'google') app.geminiModelName = app.aiModels.google;
      }
      const options = listed.slice();
      if (selected && !names.includes(selected)) options.unshift({ id: selected, label: selected });
      if (!options.length) options.push({ id: selected || meta.defaultModel, label: selected || meta.defaultModel });
      app.setModelOptions(options, selected);
    } catch (e) {
      console.error('Failed to list models', e);
    }
  };
}
