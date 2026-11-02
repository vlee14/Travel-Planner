# AGENTS.md

Guidance for coding agents working in this repository. Read this before changing behavior.

## What this is

A static single-page travel itinerary. There is no bundler, framework, or package manager. `index.html`, `styles.css`, and `js/*.js` are the app. `serve.py` is a local static server that stamps a cache name into the service worker. GitHub Pages hosts production; Jekyll front matter in `sw.js` supplies the cache revision there.

Run locally with `python serve.py` (default `http://localhost:8000`). Open the app over HTTP so the service worker can register. A file URL will not.

## Architecture

One shared mutable object, `app`, is the whole runtime. `js/app.js` creates it and calls each module’s `installX(app)` in dependency order. Modules attach functions and constants onto `app`. They do not export a class, and they do not import each other.

```js
export function installDates(app) {
  app.parseSmartDate = function parseSmartDate(str, fallbackYear = null) {
    // ...
  };
}
```

`js/wire.js` runs last. It queries the DOM, holds mutable state, wires listeners, and boots the page. Feature modules stay free of startup order except for what `app.js` already sequences.

Install order in `js/app.js` is the dependency order. `html` and `dates` come before `sheet`. `sheet` comes before `itinerary`, `builder`, and `drive`. `ai` comes before `wire`, because `wire` reads `app.AI_PROVIDERS` while initializing settings. Keep a new module in that list, and add its file to `STATIC_ASSETS` in `sw.js`.

Name functions when assigning them (`app.render = function render()`). That keeps stack traces readable.

## Module map

| File | Owns |
| --- | --- |
| `js/html.js` | Escape, sanitize, markdown, linkify, sheet HTML to editor text |
| `js/dates.js` | Trip dates and times. Canonical date is `YYYY-MM-DD` |
| `js/sheet.js` | CSV/TSV parse, sheet rows to `app.parsed`, linked-sheet edit plan |
| `js/theme.js` | Light, dark, and system theme; temperature unit UI |
| `js/weather.js` | Geocoding, forecast and historical weather, local clocks |
| `js/map.js` | Leaflet map, stop order, Google Maps route URL |
| `js/itinerary.js` | Load, render list and calendar, trip progress, share URLs |
| `js/sync.js` | Linked-sheet auto-sync (5 minutes) and the sync label |
| `js/builder.js` | Day editor, drag reorder, AI form chrome, day ↔ grid conversion |
| `js/ai.js` | Providers, chat, suggestions, checklists, trip and day generation |
| `js/drive.js` | Google sign-in, Picker, Sheets read, write-back |
| `js/boot.js` | Online and offline indicator |
| `js/wire.js` | DOM refs, `app.db`, settings, listeners, first load |
| `sw.js` | Precache, network-first navigations, stale-while-revalidate assets |
| `serve.py` | Local HTTP. Strips the Jekyll front matter from `sw.js` and rewrites `CACHE_NAME` from the git SHA (`-dirty` when the worktree is dirty) |
| `styles.css` | All presentation. Tokens on `:root` and `[data-theme="dark"]` |

## Data model

Three representations of the same trip. Convert at the boundaries. Do not invent a fourth.

**Sheet grid.** Row 0 is locations (`locations[0]` is the time header). Row 1 is dates. A row whose first cell contains the word hotel is the stay row. Later rows are times in column 0 and activities in the day columns. Untimed rows after the last timed activity become day notes in the builder. Merged cells are filled from the top-left value on read.

**`app.parsed`.** Output of `parseSheetRows`. Shape: `{ locations, dates, rows: [{ time, cells, gridRow }], hotelCells, hotelGridRow }`. Day columns start at index 1. List, calendar, and builder save treat a day as valid only when both location and date are non-empty. `parsedToDays` still keeps a column when only one of the two is set. `cells[i]` lines up with location column `i + 1`.

**Builder days.** One object per day:

```js
{ date, location, hotel, activities: [{ time, content, sourceRow? }], notes, sourceCol? }
```

`date` is `YYYY-MM-DD`. Activity `time` in the editor is `HH:MM` (24-hour). `sourceCol` and `sourceRow` point back at the linked grid so `planLinkedSheetEdits` can patch cells instead of rewriting the sheet. `parsedToDays` and `daysToRows` are the converters. `blankDay` and `placeholderDay` are the only ways to create an empty day.

In the list, empty hotel cells backfill forward while the location stays the same. A stay’s “through” date is the day before the hotel value changes. The calendar copies the previous hotel into any empty cell, including across a location change, and it does not show a through date.

## Design patterns

**Module installation.** Add behavior by extending `app` inside `installX`. Call other features as `app.otherFn()`, including functions defined in a different file. A new feature is a new `installX` plus a line in `js/app.js`, unless it is a small helper that clearly belongs in an existing module.

**DOM bag.** `wire.js` caches elements on `app` (`app.listContent`, `app.syncBtn`). Later code uses those refs. Query the DOM again only for nodes created during render.

**Render by replacement.** List, calendar, builder days, and recent docs rebuild their containers (`innerHTML = ''`, then append). There is no virtual DOM and no fine-grained patch of those views. After a data change, call the existing render (`app.render`, `app.renderBuilderDays`, `app.renderRecentDocs`) rather than mutating a single card in place.

**Visibility by class.** Major sections use `.visible` and `.active`. Tabs use `data-tab`. Itinerary views use `data-view` (`list`, `calendar`, `map`, `packing`, `builder`). Follow that instead of toggling `display` on new top-level regions. Inline `style.display` is reserved for header actions and one-off chrome that is absent until a trip exists.

**Sheet as source of truth when linked.** A Drive trip stores the grid on `app.linkedSheet` (`fileId`, `gid`, `sheetId`, `sheetName`, `title`, `grid`, `merges`). Edits go through `planLinkedSheetEdits`, then `writeLinkedSheet`. The planner diffs cells and preserves merges. A Hotel row is added on write only when some day has a stay. A paste or builder-only trip has no `linkedSheet`. Sharing uses a compressed `data` query param instead of `docId`.

**URL as navigation state.** `?docId=&gid=` is a Google Sheet. `?data=` is deflate-compressed, URL-safe base64 CSV, with optional `title`. `loadFromUrl` also accepts legacy gzip. `resetApp` strips both. `history.pushState` updates the URL; `popstate` reloads.

**Cache-aside in localStorage.** All persistence goes through `app.db.get` / `set` / `remove`. Keys are versioned prefixes (`weather_cache_v1_`, `ai_chat_cache_v1_`, `doc_cache_v1_`, `event_chk_v1_`). Weather expires in 6 hours, AI chats in 3 hours. On `QuotaExceededError`, `evictOldCache` drops the oldest 30% of weather and AI entries. Bump the prefix when a cached payload shape changes. API keys, theme, temp unit, and recent-doc metadata live in the same store and must not be evicted.

**Provider table.** AI vendors are data in `app.AI_PROVIDERS` (Google, Grok, OpenRouter): labels, key URL, storage keys, default model. `completeAi` dispatches. Adding a provider means a table entry plus a `completeX` and `listX` pair, not a new call site in the UI.

**Draft then commit.** Settings edits sit in `app.aiDraft` until Save. Builder edits sit in the DOM until `collectBuilderDays` / `applyBuilder`. Generated trips land in the builder for editing before any sheet write.

**Optimistic UI with a silent refresh.** `enableSheetSync` refetches every 5 minutes with `silent = true`. Failures fall back to the doc cache and leave the current view in place.

**Service worker split.** Navigations are network-first. Static assets are stale-while-revalidate. Cache writes are same-origin only. The fetch handler returns immediately for `accounts.google.com`, `apis.google.com`, Gemini `generateContent`, Visual Crossing, Open-Meteo, and TinyURL. Sheet reads, xAI, and OpenRouter are not on that skip list. The app caches sheet, weather, and AI responses itself when it needs freshness rules.

## UI and styling

iOS-like, system font, one stylesheet. Color, radius, and type come from custom properties on `:root` and `[data-theme="dark"]`. Use the tokens (`--bg`, `--surface`, `--text`, `--text-muted`, `--accent`, `--accent-soft`, `--success`, `--danger`, `--fill`, `--border`, `--radius`). Add a token when a new color is shared. Hard-coded colors belong only inside a token definition.

`html[data-theme]` is set by `updateTheme`. Preference is `light`, `dark`, or `system`, stored as `theme`.

Break the layout around 768px. Below that, a newly loaded trip opens on List. At 768 and above it opens on Calendar.

Interactive controls need an accessible name (`aria-label` on icon buttons). Tabs and segmented controls use `role="tab"` or `role="radio"` with `aria-checked`. Dialogs use `role="dialog"` and `aria-modal`.

User and model content that becomes HTML goes through `escapeHtml`, `linkifyAndSanitize`, or `renderMarkdown`. `safeHttpUrl` allows only `http:` and `https:`. Activity cells round-trip through `activityEditorText` (HTML or markdown links to editor text) and `linkifyPlainText` (editor text back to safe HTML).

Copy stays short and plain. Match the voice already on screen: “Plan your adventure”, “Dream up a trip”, “Synced just now”.

## External services

- **Google Identity Services and Picker** for Drive. Client id is the public OAuth client in `wire.js`. Tokens stay in memory (`app.accessToken`).
- **Sheets API v4** for read (`includeGridData`) and batch update. Scope failures use `isSheetsScopeError` and `sheetsPermissionMessage`.
- **AI** via the provider table. Keys stay in `localStorage` and are sent only to that vendor. `scrubSecret` strips the key from errors. `requestAi` aborts on timeout.
- **Weather.** Open-Meteo for coordinates, forecast, and archive (past dates, and anything beyond 14 days). Visual Crossing when `app.visualCrossingKey` is set, for the whole range in one call; a failure falls back to Open-Meteo. Icons are the Visual Crossing PNG set on GitHub.
- **Leaflet** from unpkg, with CARTO basemap tiles, and **marked** from jsDelivr. Both libraries are in the service worker precache list. Map tiles and weather icons are not. Route and place links open Google Maps.
- **TinyURL** only for shortening the share link.

Do not put API keys, tokens, or sheet contents in source. `.env` is gitignored and unused by the app. Keys are entered in Settings.

## Change rules

- Preserve the sheet layout: row 1 locations, row 2 dates, column 1 times, a Hotel row when stays exist.
- Keep day column index 1-based alignment between `locations`, `dates`, `hotelCells`, and `rows[].cells`.
- Parse dates with `parseSmartDate` / `formatTripDate`. Parse times with `timeToInputValue`, `formatTimeLabel`, and `timeToMinutes`. Leave year detection in `detectItineraryYear`.
- Write linked sheets through `planLinkedSheetEdits`. It is the merge-aware diff.
- Register new static files in `sw.js` `STATIC_ASSETS`.
- Extend `app.db` key prefixes with a new version suffix when the JSON shape changes, so old entries miss the cache and expire naturally.
- Guard DOM writes with the existing null checks (`if (app.syncBtn)`). Header controls are hidden until a trip is showing.
- Keep `serve.py` a static file server. Application behavior belongs in the browser.
- Verify UI changes in a browser against the flow you touched: Builder, Drive, Paste, list, calendar, map, packing, and a reload that hits `loadFromUrl`.

## Testing

There is no test runner. Check behavior by running `python serve.py` and exercising the path. Useful cases: paste a small grid, generate a trip in the builder, edit a day, reload a `?data=` link, and toggle theme and offline. For a linked sheet, confirm a one-cell edit does not wipe neighboring columns or merges.
