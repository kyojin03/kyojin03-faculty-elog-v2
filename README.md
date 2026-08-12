# Faculty eLog V2 | GitHub Pages + Google Apps Script

Production frontend for the Good Samaritan Colleges Faculty Laboratory eLog System.

## Architecture

```text
GitHub Pages
    -> static HTML, CSS, and JavaScript
    -> fetch() requests
    -> existing Google Apps Script Web App /exec
    -> existing Google Sheets workbook
```

Google Sheets remains the source of truth. Google Apps Script remains the only backend. No additional database or hosting service is used.

## Project structure

```text
index.html
assets/
  images/
    good-samaritan-colleges-logo.png
css/
  styles.css
js/
  config.js
  api.js
  core.js
  logbook.js
  reports.js
  inventory.js
Code.gs
appsscript.json
.nojekyll
```

The approved user interface is in `index.html`, `css/styles.css`, and the feature modules. `js/api.js` is the HTTP integration layer. `Code.gs` is the JSON API deployed from the existing spreadsheet-bound Apps Script project.

## HTTP API

The configured API URL is stored once in `js/config.js`.

| Method | Action | Purpose |
|---|---|---|
| GET | `initialData` | Initial rooms, settings, inventories, forms, and Logbook data |
| GET | `readOnlyData` | Refresh rooms, settings, inventories, and forms |
| GET | `logbook` | Filtered Logbook records |
| GET | `reports` | Monthly read-only report data |
| POST | `submitLog` | Submit the only website write operation |

Examples:

```text
GET /exec?action=initialData
GET /exec?action=readOnlyData
GET /exec?action=logbook&search=microscope&department=COLLEGE%20OF%20NURSING
GET /exec?action=reports&month=2026-08
POST /exec?action=submitLog
```

The frontend uses `application/x-www-form-urlencoded` for `submitLog`. This is compatible with Apps Script Web Apps and avoids a cross-origin preflight request.

## Deploy the backend

1. Open the existing Faculty eLog Google Sheet.
2. Open **Extensions -> Apps Script**.
3. Replace the existing `Code.gs` with the supplied `Code.gs`.
4. In **Project Settings**, enable **Show "appsscript.json" manifest file in editor** and replace it with the supplied manifest if appropriate for the project.
5. Save the Apps Script project.
6. Open **Deploy -> Manage deployments**.
7. Edit the existing Web App deployment and select **New version**.
8. Set **Execute as** to the deploying account.
9. Set **Who has access** to **Anyone**. This is required for a browser-hosted GitHub Pages frontend to call the endpoint without an Apps Script sign-in redirect.
10. Deploy and retain the existing `/exec` URL.

After deployment, verify that this URL returns JSON rather than HTML:

```text
https://script.google.com/macros/s/AKfycby3R8jwOH3UrwidrSTva7zfUQw5uL6RkrDad0Q6CGxHx1B7ylh9chwR8rC8LJtx8svC/exec?action=initialData
```

## Deploy the frontend to GitHub Pages

1. Commit `index.html`, `.nojekyll`, `css/`, and `js/` to the GitHub repository branch used for Pages.
2. Confirm `js/config.js` contains the existing deployed Apps Script `/exec` URL.
3. In the GitHub repository, open **Settings -> Pages**.
4. Under **Build and deployment**, choose **Deploy from a branch**.
5. Select the frontend branch and `/ (root)`, then save.
6. Open the published GitHub Pages URL in a private browser window and confirm the Logbook loads.
7. Perform one controlled Logbook test submission, verify the new row in the Google Sheet, and remove that test row directly in Google Sheets if it should not remain in the official record.

Do not publish `Code.gs` as a browser script. It is copied only into the existing Apps Script backend project.

## Data ownership rule

Only Logbook accepts website input through the `submitLog` action. Reports, Laboratory Supplies, Glassware, Equipment, Specialized Equipment, and Laboratory Forms are read-only. Their information is maintained directly in Google Sheets.

`Rooms` controls the room dropdown and room-name display. Active `Activity Type` rows in `Settings` control the activity selector. Existing Logbook faculty names and departments provide form suggestions and filters.
