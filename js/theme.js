// Appearance and temperature units.
export function installTheme(app) {
  app.resolvedTheme = function resolvedTheme(preference) {
    if (preference === 'system') return app.systemThemeQuery.matches ? 'dark' : 'light';
    return preference === 'dark' ? 'dark' : 'light';
  };

  app.updateTheme = function updateTheme(preference) {
    app.themePreference = preference;
    app.currentTheme = app.resolvedTheme(preference);
    app.htmlEl.setAttribute('data-theme', app.currentTheme);
    app.db.set('theme', preference);
    document.querySelectorAll('.theme-choice').forEach(btn => {
      const selected = btn.dataset.themeChoice === preference;
      btn.classList.toggle('active', selected);
      btn.setAttribute('aria-checked', selected ? 'true' : 'false');
    });
    if (app.mapTileLayer) {
      const isLight = app.currentTheme === 'light';
      const url = isLight
        ? 'https://{s}.basemaps.cartocdn.com/rastertiles/voyager/{z}/{x}/{y}{r}.png'
        : 'https://{s}.basemaps.cartocdn.com/dark_all/{z}/{x}/{y}{r}.png';
      app.mapTileLayer.setUrl(url);
    }
    if (app.polyline) app.polyline.setStyle({ color: app.routeColor() });
  };

  app.updateTempUnitUI = function updateTempUnitUI() {
    document.querySelectorAll('.temp-unit-label').forEach(el => {
      el.textContent = app.currentTempUnit === 'C' ? 'Celsius (°C)' : 'Fahrenheit (°F)';
    });
  };
}
