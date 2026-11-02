// Offline status.
export function installBoot(app) {
  app.updateOnlineStatus = function updateOnlineStatus() {
    if (!navigator.onLine) {
      app.offlineIndicator.classList.add('visible');
    } else {
      app.offlineIndicator.classList.remove('visible');
    }
  };
}
