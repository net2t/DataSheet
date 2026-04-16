# DataSheet (NeoBrutalism)

This project is a small **single-page web app** (HTML + JS) that displays your trademark/application tracking data in a NeoBrutalism UI.

It supports:
- List view + Card view
- Compact mode
- Status/Sub-status dependency
- Duplicate + TM match indicator
- Attach file metadata per record
- Log history + per-record popup
- Pagination (**100 rows per page**)
- Load data from **Google Sheets** (read-only)

## 1) Run locally

Option A (Python):

```bash
python -m http.server 8000
```

Then open:

```text
http://localhost:8000
```

## 2) Link / Load from Google Sheets (Read-only)

The app loads data from Google Sheets using Google’s `gviz` endpoint:

- It works only if your sheet is **published to the web** OR **shared publicly**.
- This approach is **read-only** (it will not write back to the sheet).

### Step A: Get Sheet ID

Your URL looks like:

```text
https://docs.google.com/spreadsheets/d/SHEET_ID/edit?gid=GID
```

Copy the `SHEET_ID` part.

### Step B: Get GID (tab id)

In the same URL, `gid=XXXX` is your tab’s **GID**.

### Step C: Make the sheet accessible

Choose one:

1. **Publish to web** (recommended)
   - In Google Sheets:
     - File -> Share -> Publish to web
     - Publish the sheet

2. **Share publicly**
   - Share -> General access -> Anyone with the link -> Viewer

### Step D: Load in the app

In the app UI:
- Enter `Google Sheet ID`
- Enter `Sheet GID`
- Click **Load Sheet**
- Optionally click the save (disk) icon to remember the config in your browser.

## 3) Sheet column mapping

The loader expects these column headers in your Google Sheet:

- `DATE TIME`
- `CASE NO`
- `APP NAME`
- `TM NO`
- `CLASS`
- `APPLICATION STATUS`
- `APPLICATION SUB STATUS`
- `NOTES`

If headers differ, rename them in the sheet to match.

## 4) Pagination

Pagination is built-in:
- 100 records per page
- Use the **left/right arrows** in the Pagination bar.

## 5) Field search

Use:
- Search text box
- Search field dropdown (CASE NO / TM NO / APP NAME / etc.)

## 6) Writing back to Google Sheets (Not implemented)

This project currently **does not write** updates back to Google Sheets.

To enable saving edits back to a Google Sheet you need one of these:
- A backend (Node/Python) using Google Sheets API + OAuth/service account
- OR a Google Apps Script Web App endpoint (easier) that accepts POST requests

If you want, tell me which method you prefer and I can implement it.
