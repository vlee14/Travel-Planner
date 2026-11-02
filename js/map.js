// Leaflet map and route.
export function installMap(app) {
  app.routeColor = function routeColor() {
    return app.currentTheme === 'dark' ? '#0A84FF' : '#007AFF';
  };

  app.initMap = function initMap() {
    if (app.map) return;
    // Initialize map
    app.map = L.map('itinerary-map').setView([20, 0], 2);

    const isLight = app.currentTheme === 'light';
    const url = isLight ? 'https://{s}.basemaps.cartocdn.com/rastertiles/voyager/{z}/{x}/{y}{r}.png' : 'https://{s}.basemaps.cartocdn.com/dark_all/{z}/{x}/{y}{r}.png';

    app.mapTileLayer = L.tileLayer(url, {
      attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors &copy; <a href="https://carto.com/attributions">CARTO</a>',
      maxZoom: 18
    }).addTo(app.map);
    app.mapLayerGroup = L.layerGroup().addTo(app.map);
  };

  app.resetMapData = function resetMapData() {
    if (!app.map) return;
    app.mapLayerGroup.clearLayers();
    app.mapPoints.length = 0;
    for (const key in app.mapMarkers) delete app.mapMarkers[key];
    app.polyline = L.polyline([], { color: app.routeColor(), weight: 4, opacity: 0.9 }).addTo(app.mapLayerGroup);
    if (app.mapRouteBtn) app.mapRouteBtn.style.display = 'none';
  };

  app.pinForecastLocation = function pinForecastLocation(name, lat, lon, skip) {
    if (skip) return;
    app.addMapLocation(name, lat, lon);
  };

  app.addMapLocation = function addMapLocation(name, lat, lon) {
    if (!app.map) return;

    // Add marker if this specific location hasn't been marked yet
    if (!app.mapMarkers[name]) {
      const marker = L.marker([lat, lon]).addTo(app.mapLayerGroup);
      marker.bindPopup(`<b>${app.escapeHtml(name)}</b>`);
      app.mapMarkers[name] = marker;
    }

    // Add to polyline path (allows duplicate locations to show route)
    app.mapPoints.push([lat, lon]);
    if (app.polyline) {
      app.polyline.setLatLngs(app.mapPoints);
    }

    // Fit bounds if we have points
    if (app.mapPoints.length > 0) {
      const bounds = L.latLngBounds(app.mapPoints);
      // Add some padding
      app.map.fitBounds(bounds, { padding: [50, 50], maxZoom: 12 });
    }
    app.updateMapRouteButton();
  };

  app.updateMapRouteButton = function updateMapRouteButton() {
    if (!app.mapRouteBtn) return;
    const stops = app.getOrderedMapStops();
    app.mapRouteBtn.style.display = stops.length >= 2 ? 'flex' : 'none';
  };

  app.getOrderedMapStops = function getOrderedMapStops() {
    const stops = [];
    const seen = new Set();
    if (app.parsed) {
      for (let i = 1; i < app.parsed.locations.length; i++) {
        const loc = (app.parsed.locations[i] || '').trim();
        const date = (app.parsed.dates[i] || '').trim();
        if (!loc || !date || seen.has(loc)) continue;
        const marker = app.mapMarkers[loc];
        if (marker) {
          const { lat, lng } = marker.getLatLng();
          stops.push({ lat, lng, loc });
          seen.add(loc);
        }
      }
    }
    if (stops.length < 2 && app.mapPoints.length >= 2) {
      const fallback = [];
      const seenCoords = new Set();
      app.mapPoints.forEach(([lat, lon]) => {
        const key = `${lat.toFixed(4)},${lon.toFixed(4)}`;
        if (!seenCoords.has(key)) {
          fallback.push({ lat, lng: lon });
          seenCoords.add(key);
        }
      });
      return fallback;
    }
    return stops;
  };

  app.buildGoogleMapsRouteUrl = function buildGoogleMapsRouteUrl() {
    const stops = app.getOrderedMapStops();
    if (stops.length < 2) return null;
    const fmt = (s) => `${s.lat},${s.lng}`;
    const origin = encodeURIComponent(fmt(stops[0]));
    const destination = encodeURIComponent(fmt(stops[stops.length - 1]));
    let url = `https://www.google.com/maps/dir/?api=1&origin=${origin}&destination=${destination}&travelmode=driving`;
    if (stops.length > 2) {
      const waypoints = stops.slice(1, -1).map(fmt).join('|');
      url += `&waypoints=${encodeURIComponent(waypoints)}`;
    }
    return url;
  };
}
