// Forecasts and location clocks.
export function installWeather(app) {
  app.formatTemp = function formatTemp(fahrenheit) {
    const f = Math.round(fahrenheit);
    if (app.currentTempUnit === 'C') {
      return `${Math.round((f - 32) * 5 / 9)}°C`;
    }
    return `${f}°F`;
  };

  app.formatTempPair = function formatTempPair(tempmax, tempmin) {
    const hi = Math.round(tempmax);
    const lo = Math.round(tempmin);
    if (hi === lo) return app.formatTemp(hi);
    if (app.currentTempUnit === 'C') {
      return `${app.formatTemp(hi)} / ${app.formatTemp(lo)}`;
    }
    return `${hi}°F / ${lo}°F`;
  };

  app.getWmoInfo = function getWmoInfo(code) {
    const map = {
      0: { icon: 'clear-day', label: 'Clear sky' },
      1: { icon: 'partly-cloudy-day', label: 'Mainly clear' },
      2: { icon: 'partly-cloudy-day', label: 'Partly cloudy' },
      3: { icon: 'cloudy', label: 'Overcast' },
      45: { icon: 'fog', label: 'Fog' },
      48: { icon: 'fog', label: 'Depositing rime fog' },
      51: { icon: 'rain', label: 'Light drizzle' },
      53: { icon: 'rain', label: 'Moderate drizzle' },
      55: { icon: 'rain', label: 'Dense drizzle' },
      56: { icon: 'sleet', label: 'Light freezing drizzle' },
      57: { icon: 'sleet', label: 'Dense freezing drizzle' },
      61: { icon: 'rain', label: 'Slight rain' },
      63: { icon: 'rain', label: 'Moderate rain' },
      65: { icon: 'rain', label: 'Heavy rain' },
      66: { icon: 'sleet', label: 'Light freezing rain' },
      67: { icon: 'sleet', label: 'Heavy freezing rain' },
      71: { icon: 'snow', label: 'Slight snow fall' },
      73: { icon: 'snow', label: 'Moderate snow fall' },
      75: { icon: 'snow', label: 'Heavy snow fall' },
      77: { icon: 'snow', label: 'Snow grains' },
      80: { icon: 'showers-day', label: 'Slight rain showers' },
      81: { icon: 'showers-day', label: 'Moderate rain showers' },
      82: { icon: 'showers-day', label: 'Violent rain showers' },
      85: { icon: 'snow-showers-day', label: 'Slight snow showers' },
      86: { icon: 'snow-showers-day', label: 'Heavy snow showers' },
      95: { icon: 'thunder-rain', label: 'Thunderstorm' },
      96: { icon: 'thunder-showers-day', label: 'Thunderstorm with slight hail' },
      99: { icon: 'thunder-showers-day', label: 'Thunderstorm with heavy hail' }
    };
    return map[code] || { icon: 'partly-cloudy-day', label: 'Unknown' };
  };

  app.getCoordinates = async function getCoordinates(location) {
    try {
      // Use Open-Meteo Geocoding API to avoid Nominatim's strict Referer/User-Agent policies
      const url = `https://geocoding-api.open-meteo.com/v1/search?name=${encodeURIComponent(location)}&count=1&language=en&format=json`;
      const res = await fetch(url);
      if (!res.ok) return null;
      const data = await res.json();
      if (data && data.results && data.results.length > 0) {
        return { lat: data.results[0].latitude, lon: data.results[0].longitude };
      }
    } catch (e) {
      console.error('Geocoding error:', e);
    }
    return null;
  };

  app.fetchCardWeatherData = async function fetchCardWeatherData(card) {
    const loc = card.dataset.location;
    const start = app.parseSmartDate(card.dataset.start);
    const end = app.parseSmartDate(card.dataset.end);
    if (!loc || !start) return;
    const skipMap = card.dataset.bookend === 'true';

    const source = app.visualCrossingKey ? 'vc' : 'om';
    const cacheKey = `${app.WEATHER_CACHE_PREFIX}${source}_${encodeURIComponent(loc)}_${start}_${end || ''}`;

    try {
      const cachedJson = app.db.get(cacheKey);
      if (cachedJson) {
        const cached = JSON.parse(cachedJson);
        if (Date.now() - cached.timestamp < app.WEATHER_CACHE_EXPIRY) {
          if (cached.coords) app.pinForecastLocation(loc, cached.coords.lat, cached.coords.lon, skipMap);
          return cached.data;
        }
      }
    } catch (e) {
      console.warn('Cache read error', e);
    }

    // Try Visual Crossing first if key exists
    if (app.visualCrossingKey) {
      try {
        const endDateStr = end || start;
        const url = `https://weather.visualcrossing.com/VisualCrossingWebServices/rest/services/timeline/${encodeURIComponent(loc)}/${start}/${endDateStr}?unitGroup=us&key=${app.visualCrossingKey}&include=days&contentType=json`;

        const res = await fetch(url);
        if (!res.ok) throw new Error(`Visual Crossing error: ${res.status}`);

        const data = await res.json();
        if (data) {
          let coords = null;
          if (data.latitude && data.longitude) {
            coords = { lat: data.latitude, lon: data.longitude };
            app.pinForecastLocation(loc, data.latitude, data.longitude, skipMap);
          }
          const timezone = data.timezone;
          const days = (data.days || []).map(d => ({
            datetime: d.datetime,
            tempmax: d.tempmax,
            tempmin: d.tempmin,
            icon: d.icon,
            label: d.conditions,
            isHistorical: false // VC handles mixed forecast/history automatically
          }));
          const result = { days, timezone };
          app.db.set(cacheKey, JSON.stringify({ timestamp: Date.now(), data: result, coords }));
          return result;
        }
      } catch (e) {
        console.warn('Visual Crossing fetch failed, falling back to Open-Meteo.', e);
      }
    }

    const coords = await app.getCoordinates(loc);
    if (!coords) return;

    app.pinForecastLocation(loc, coords.lat, coords.lon, skipMap);

    // Generate dates to fetch
    const daysToFetch = [];
    const startParts = start.split('-').map(Number);
    const startDate = new Date(startParts[0], startParts[1] - 1, startParts[2]);
    const endDate = end ? new Date(end.split('-')[0], end.split('-')[1] - 1, end.split('-')[2]) : new Date(startDate);

    let curr = new Date(startDate);
    while (curr <= endDate) {
      const y = curr.getFullYear();
      const m = String(curr.getMonth() + 1).padStart(2, '0');
      const d = String(curr.getDate()).padStart(2, '0');
      daysToFetch.push(`${y}-${m}-${d}`);
      curr.setDate(curr.getDate() + 1);
    }

    const now = new Date();
    now.setHours(0, 0, 0, 0);

    const days = [];
    let timezone = null;

    // Fetch per day
    for (const dateStr of daysToFetch) {
      const p = dateStr.split('-').map(Number);
      const dObj = new Date(p[0], p[1] - 1, p[2]);
      const diffDays = (dObj - now) / (1000 * 60 * 60 * 24);
      let isDayHistorical = diffDays < 0 || diffDays > 14;
      let dayData = null;

      // Try Forecast
      if (!isDayHistorical) {
        try {
          const url = `https://api.open-meteo.com/v1/forecast?latitude=${coords.lat}&longitude=${coords.lon}&daily=weathercode,temperature_2m_max,temperature_2m_min&timezone=auto&temperature_unit=fahrenheit&start_date=${dateStr}&end_date=${dateStr}`;
          const res = await fetch(url);
          if (res.ok) {
            const json = await res.json();
            if (json.timezone) timezone = json.timezone; // Capture timezone
            if (json.daily && json.daily.time && json.daily.time.length > 0) {
              dayData = {
                datetime: dateStr,
                tempmax: json.daily.temperature_2m_max[0],
                tempmin: json.daily.temperature_2m_min[0],
                ...getWmoInfo(json.daily.weathercode[0]),
                isHistorical: false
              };
            }
          }
        } catch (e) {
          // Ignore, fallback to historical
        }
      }

      // Fallback to Historical
      if (!dayData) {
        try {
          const prevDate = new Date(p[0] - 1, p[1] - 1, p[2]);
          const py = prevDate.getFullYear();
          const pm = String(prevDate.getMonth() + 1).padStart(2, '0');
          const pd = String(prevDate.getDate()).padStart(2, '0');
          const prevDateStr = `${py}-${pm}-${pd}`;

          const url = `https://archive-api.open-meteo.com/v1/archive?latitude=${coords.lat}&longitude=${coords.lon}&daily=weathercode,temperature_2m_max,temperature_2m_min&timezone=auto&temperature_unit=fahrenheit&start_date=${prevDateStr}&end_date=${prevDateStr}`;
          const res = await fetch(url);
          if (res.ok) {
            const json = await res.json();
            if (json.timezone) timezone = json.timezone;
            if (json.daily && json.daily.time && json.daily.time.length > 0) {
              dayData = {
                datetime: dateStr,
                tempmax: json.daily.temperature_2m_max[0],
                tempmin: json.daily.temperature_2m_min[0],
                ...getWmoInfo(json.daily.weathercode[0]),
                isHistorical: true
              };
            }
          }
        } catch (e) {
          console.error('Historical fetch failed', e);
        }
      }

      if (dayData) {
        days.push(dayData);
      }
    }

    const result = { days, timezone };
    app.db.set(cacheKey, JSON.stringify({ timestamp: Date.now(), data: result, coords }));
    return result;
  };

  app.renderCardWeather = function renderCardWeather(card, { days, timezone }) {
    try {
      if (days && days.length > 0) {
        const maxHigh = Math.round(Math.max(...days.map(d => d.tempmax)));
        const minHigh = Math.round(Math.min(...days.map(d => d.tempmax)));

        const uniqueIcons = [...new Set(days.map(d => d.icon))];
        const iconsHtml = uniqueIcons.map(icon =>
          `<img src="https://raw.githubusercontent.com/visualcrossing/WeatherIcons/refs/heads/main/PNG/2nd%20Set%20-%20Color/${icon}.png" alt="${icon}" title="${icon}" class="weather-icon" style="width: 1.4em; height: 1.4em; vertical-align: middle;">`
        ).join('');

        const dateRangeEl = card.querySelector('.date-range');
        if (dateRangeEl) {
          let sumEl = card.querySelector('.weather-summary-inline');
          if (!sumEl) {
            sumEl = document.createElement('span');
            sumEl.className = 'weather-summary-inline';
            dateRangeEl.parentNode.insertBefore(sumEl, dateRangeEl.nextSibling);
          }
          const tempStr = minHigh === maxHigh ? app.formatTemp(maxHigh) : `${app.formatTemp(minHigh)}–${app.formatTemp(maxHigh)}`;
          sumEl.innerHTML = `${iconsHtml} <span class="weather-temp">${tempStr}</span>`;

          if (days.some(d => d.isHistorical)) {
            const meta = card.querySelector('.location-card-meta');
            if (meta && !meta.querySelector('.historical-note')) {
              const note = document.createElement('div');
              note.className = 'historical-note';
              note.textContent = 'Includes historical data (last year)';
              meta.appendChild(note);
            }
          }
        }
      }

      if (timezone) {
        const meta = card.querySelector('.location-card-meta');
        if (meta && !meta.querySelector('.local-time-display')) {
          const clockEl = document.createElement('div');
          clockEl.className = 'local-time-display';
          clockEl.innerHTML = `Local <span class="clock-val">--:--</span>`;
          const firstHotel = meta.querySelector('.hotel');
          if (firstHotel) {
            meta.insertBefore(clockEl, firstHotel);
          } else {
            meta.appendChild(clockEl);
          }

          clockEl.dataset.tz = timezone;
          app.updateSingleClock(clockEl);
        }
      }

      const groups = card.querySelectorAll('.day-group-title');
      groups.forEach(title => {
        const smartDate = app.parseSmartDate(title.dataset.date);
        const tDate = smartDate;
        const dayData = days.find(d => d.datetime === tDate);
        if (dayData) {
          title.dataset.weatherSummary = `${app.formatTempPair(dayData.tempmax, dayData.tempmin)}, ${dayData.label}`;
          let badge = title.querySelector('.weather-badge');
          if (!badge) {
            badge = document.createElement('span');
            badge.className = 'weather-badge';
            const firstButton = title.querySelector('.day-notes-btn, .btn-ai-suggest, .btn-copy-day');
            title.insertBefore(badge, firstButton || null);
          }
          const iconUrl = `https://raw.githubusercontent.com/visualcrossing/WeatherIcons/refs/heads/main/PNG/2nd%20Set%20-%20Color/${dayData.icon}.png`;
          badge.innerHTML = `<img src="${iconUrl}" alt="${dayData.icon}" class="weather-icon" style="width: 1.4em; height: 1.4em; margin-right: 0.3em;"> ${app.formatTempPair(dayData.tempmax, dayData.tempmin)} ${dayData.label}`;
        }
      });

      if (days && days.length > 0) {
        days.forEach(dayData => {
          const calCell = document.querySelector(`.calendar-day[data-date="${dayData.datetime}"]`);
          if (calCell) {
            let wDiv = calCell.querySelector('.calendar-weather');
            if (!wDiv) {
              wDiv = document.createElement('div');
              wDiv.className = 'calendar-weather';
              const hotel = calCell.querySelector('.hotel-pill');
              if (hotel) {
                calCell.insertBefore(wDiv, hotel);
              } else {
                const firstEvent = calCell.querySelector('.calendar-event:not(.location-pill)');
                if (firstEvent) calCell.insertBefore(wDiv, firstEvent);
                else calCell.appendChild(wDiv);
              }
            }

            if (!wDiv.innerHTML) {
              const iconUrl = `https://raw.githubusercontent.com/visualcrossing/WeatherIcons/refs/heads/main/PNG/2nd%20Set%20-%20Color/${dayData.icon}.png`;
              const histIcon = dayData.isHistorical ? '<span class="hist-mark" title="Historical average">Past</span>' : '';
              wDiv.title = dayData.label;
              wDiv.innerHTML = `
            <img src="${iconUrl}" class="weather-icon" style="width:1.2em; height:1.2em;">
            <span>${app.formatTempPair(dayData.tempmax, dayData.tempmin)}${histIcon}</span>
          `;
              if (dayData.isHistorical) wDiv.classList.add('historical');
            }
          }
        });
      }

    } catch (e) {
      console.error('Error processing weather data:', e);
    }
  };

  app.updateSingleClock = function updateSingleClock(el) {
    try {
      const tz = el.dataset.tz;
      if (!tz) return;
      const now = new Date();
      const options = { hour: 'numeric', minute: '2-digit', timeZone: tz, hour12: true };
      // Format: "2:45 PM"
      const timeStr = new Intl.DateTimeFormat('en-US', options).format(now);
      const val = el.querySelector('.clock-val');
      if (val) val.textContent = timeStr;
    } catch (e) { console.error(e); }
  };

  app.fetchAllWeather = async function fetchAllWeather() {
    const cards = document.querySelectorAll('.location-card');
    const results = [];
    for (const card of cards) {
      const data = await app.fetchCardWeatherData(card);
      results.push({ card, data });
    }
    for (const { card, data } of results) {
      if (data) app.renderCardWeather(card, data);
    }
  };

  app.rerenderAllWeatherBadges = async function rerenderAllWeatherBadges() {
    const cards = document.querySelectorAll('.location-card');
    for (const card of cards) {
      const data = await app.fetchCardWeatherData(card);
      if (data) app.renderCardWeather(card, data);
    }
  };
}
