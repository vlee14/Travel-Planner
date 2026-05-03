# Travel Itinerary

A static single-page web app that transforms your Google Sheet travel itinerary into a rich, interactive experience. Host it on **GitHub Pages** for free.

## Features

- **Multiple Views**: Switch between a sequential **List**, a monthly **Calendar**, an interactive **Map**, and an AI-powered **Packing Checklist**.
- **Google Sheets Integration**: Directly link your private Google Sheets via Google Drive picker or URL.
- **Smart Weather**: Get live forecasts and historical weather data for your destinations.
- **AI-Powered Suggestions**: Get personalized activity ideas and travel checklists using Google Gemini.
- **Offline First**: Works offline as a PWA with local caching.
- **Easy Sharing**: Shorten your itinerary URL for easy sharing with travel companions.

## Sheet layout

Your Google Sheet should follow this structure for best results:

|       | **Location A** | **Location B** | **Location C** |
|-------|----------------|---------------|----------------|
|       | *Date 1*       | *Date 2*      | *Date 3*       |
| 9:00  | Breakfast      |               |                |
| 10:00 | Museum        | Travel        | Check-in       |
| 14:00 | Park          | Lunch         |                |

- **Row 1:** Location names (column headers)
- **Row 2:** Dates for each location
- **Column 1:** Times
- **Rest:** Activity/notes for each time and location.
- **Accommodation:** Add a row starting with "Hotel" in the first column to track your stays.

## How to use

1. **Connect a Sheet**  
   Open the **Google Drive** tab. Sign in to browse your spreadsheets or paste a specific Google Sheet URL. The app will automatically parse the content.
   
2. **Paste Content**  
   Alternatively, copy your itinerary range in Google Sheets or Excel and paste it directly into the **Paste CSV** tab.

3. **Configure AI & Weather (Optional)**  
   Click the **Settings** icon to add your **Google Gemini API Key** for suggestions and a **Visual Crossing API Key** for enhanced weather data.

## Host on GitHub Pages

1. Create a new repo (e.g. `travel-itinerary`).
2. Upload the project files to the root of the repo.
3. In the repo: **Settings → Pages**.
4. Under **Source**, choose **Deploy from a branch**.
5. Branch: **main** (or **master**), folder: **/ (root)** (or the folder that contains `index.html`).
6. Save. The site will be at `https://<username>.github.io/<repo>/` (or `https://<username>.github.io/<repo>/` with a subfolder if you used one).

No build step or server required—just static HTML/CSS/JS.
