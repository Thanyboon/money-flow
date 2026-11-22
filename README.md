# MONEY FLOW

Personal finance PWA using GitHub Pages + Google Apps Script + Google Sheets.

## 1. Google Sheets
Create a spreadsheet named `MONEY FLOW DATABASE`. Copy its ID from the URL.

## 2. Apps Script
Open Extensions > Apps Script (or script.google.com), replace Code.gs with `gas/Code.gs`.
In Project Settings > Script Properties add:
- `SPREADSHEET_ID` = spreadsheet ID
- `APP_PASSWORD` = a long unique password (do not reuse your Google password)

Save, run `initialize_` once from the editor and approve permissions. Then Deploy > New deployment > Web app. Execute as Me. Access: Anyone (the app password is the application gate; anyone who obtains the URL can reach the login endpoint). Copy the `/exec` URL.

## 3. GitHub Pages
Upload all files in this folder to the root of a GitHub repository. Enable Settings > Pages > Deploy from branch > main > /(root). Open the published HTTPS URL.

## 4. Install
On iPad Safari: Share > Add to Home Screen.

## Important security and feature notes
- This is a single-owner starter implementation. Protect the deployment URL and use a unique strong app password. For stronger privacy, restrict the Apps Script deployment to your Google account and ensure all devices use that account, subject to Google's deployment access behavior.
- Session tokens are stored in Apps Script Cache for up to 6 hours and in sessionStorage on the device. Cache eviction may require signing in again.
- The current split-bill implementation creates equal splits. Custom per-item/percentage splits, payment settlement tracking, push notifications, multi-user accounts and robust offline write queues are not implemented in this initial release.
- Offline mode serves cached app shell and last cached data snapshot. Writes require connectivity.
- Apps Script/Sheets are not a transactional database. Back up the spreadsheet regularly.
