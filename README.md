# Travel Itinerary

A static single-page site that turns a Google Sheet travel itinerary into a clean web view. Host it on **GitHub Pages** for free.

## Sheet layout

Your Google Sheet should follow this structure:

|       | **Location A** | **Location B** | **Location C** |
|-------|----------------|---------------|----------------|
|       | *Date 1*       | *Date 2*      | *Date 3*       |
| 9:00  | Breakfast      |               |                |
| 10:00 | Museum        | Travel        | Check-in       |
| 14:00 | Park          | Lunch         |                |

- **Row 1:** Location names (column headers)
- **Row 2:** Dates for each location
- **Column 1:** Times
- **Rest:** Activity/notes for each time and location

## How to use

1. **Paste CSV**  
   In Google Sheets, select your itinerary range and copy (Ctrl+C / Cmd+C). On the site, open the “Paste CSV” tab, paste into the text area, and click **Build itinerary**.

2. **Google Sheet URL**  
   Publish the sheet as CSV: **File → Share → Publish to web** → choose the sheet → **Comma-separated values (.csv)** → Publish. Paste the export URL into the “Google Sheet URL” tab and click **Load from sheet**.  
   (If the URL doesn’t load due to CORS, use “Paste CSV” instead.)

3. **Views**  
   Use **Grid** (table) or **List** (cards by date) to read the itinerary.

## Host on GitHub Pages

1. Create a new repo (e.g. `travel-itinerary`).
2. Upload `index.html` to the root of the repo (or put it in a folder and set that as the site source).
3. In the repo: **Settings → Pages**.
4. Under **Source**, choose **Deploy from a branch**.
5. Branch: **main** (or **master**), folder: **/ (root)** (or the folder that contains `index.html`).
6. Save. The site will be at `https://<username>.github.io/<repo>/` (or `https://<username>.github.io/<repo>/` with a subfolder if you used one).

No build step or server required—just static HTML/CSS/JS.
