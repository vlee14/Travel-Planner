import './dom.js';
import { installHtml } from '../js/html.js';
import { installDates } from '../js/dates.js';
import { installSheet } from '../js/sheet.js';
import { installBuilder } from '../js/builder.js';

export function createApp() {
  const app = {};
  installHtml(app);
  installDates(app);
  installSheet(app);
  installBuilder(app);
  return app;
}
