# Consular Request Checker

A small static HTML/CSS/JavaScript project that checks a request ID + phone number
against an Excel spreadsheet (`database.xlsx`).

## Included

- `index.html` - web form
- `styles.css` - simple responsive styling
- `app.js` - form logic + built-in XLSX reader (no external libraries/CDNs)
- `database.xlsx` - sample Excel database
- `README.md` - setup instructions
- `run_offline.bat` - starts a local Python web server for offline testing

## Important

This is a static-site solution. The browser downloads `database.xlsx` and reads it.
Therefore, the spreadsheet is publicly accessible to anyone who can access your
website. Do NOT put sensitive/private personal information in a publicly hosted
Excel file.

For a real consular production system, use a server-side database/API instead of
exposing an Excel file to the browser.

## Offline testing

### Windows

**Recommended:** double-click `run_offline.bat`, then open:

http://localhost:8000

If Python is installed, the batch file starts a local web server in this folder.

**Alternative:** double-click `index.html`. If the browser blocks automatic access
to the Excel file, the page will show an Offline Testing file picker. Select
`database.xlsx` and then run the check.

### Any OS

From this folder run:

`python -m http.server 8000`

Then open:

http://localhost:8000

Opening `index.html` directly is supported for offline testing, but browsers do
not normally allow JavaScript to fetch a neighboring file in `file://` mode.
Use the database file picker if you open the page this way.

## Edit the database

Open `database.xlsx` in Excel or LibreOffice.

The first worksheet must have these column headers:

- request_id
- phone_number
- name
- request_type
- request_status
- notes

The checker only displays records whose `request_status` is `Completed` and whose
request ID + phone number match.

## Change the appointment link

Open `app.js` and change:

`appointmentUrl: "https://example.com/book-appointment"`

to your real booking URL.

## Hosting

Because this project is static, it can be hosted on many free static hosting
services. Upload all files while keeping `database.xlsx` in the same directory
as `index.html`.

## Matching behavior

- Request ID: ignores leading/trailing spaces and letter case.
- Phone number: compares digits only, so spaces, parentheses and hyphens do not matter.
- Status: `Completed` is matched case-insensitively.
