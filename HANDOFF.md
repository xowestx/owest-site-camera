# O West Site Camera: project handoff

Paste or upload this file into a new Claude chat to continue the work.
Also upload `Code.gs` and `index.html` (or connect GitHub and point Claude to the repo).

## What it is
Engineers on site open a website on their phone, pick a unit, take photos, and fill in the work done and manpower (contractor · item · trade · count) plus internal notes. Everything is filed by unit in Google Drive and logged in a Google Sheet. Reports are PDFs made automatically. It doesn't need Claude, ClickUp or SharePoint access to run.

## Live pieces
| Piece | Where |
|---|---|
| Website (GitHub Pages) | https://xowestx.github.io/owest-site-camera/ |
| Hidden extra reports | https://xowestx.github.io/owest-site-camera/#extra, or press and hold the O WEST title for 1.5 s. Code is set in Apps Script (setPasswordNow). |
| Code repo | https://github.com/xowestx/owest-site-camera (`index.html` = site, `backend/Code.gs` = server) |
| Google Sheet (database) | "O West Site Camera", ID `1kWngNEurLts1djgYRbjmoUFhyCHm6i1vqpx5--D1Ngc`, owner account x.owest.x |
| Apps Script (backend API) | Sheet → Extensions → Apps Script. Web App: Execute as Me, access Anyone |
| API URL (in index.html) | https://script.google.com/macros/s/AKfycbyPqz639Y6R1--7qazMyHZ0-LpQFSX7e_daagYkQgmgUK2TDYilBjrbOwUdojpsIB3_qQ/exec |
| Drive photos | `O West Site Photos/<Unit>/<YYYY-MM-DD>/…jpg`, plus a "<Unit> - Site log" Google Doc per unit |
| Drive reports | `O West Reports/Daily | Weekly | Manpower/<date>/` |
| ClickUp source | List `1201250000001676`, status "site execution". Token stored in Script Properties (CLICKUP_TOKEN), never in chat |

ClickUp custom field IDs: Request Type `8f185b3c-db17-4dac-b577-d377a949db9f` (gives the items per unit), Contractor `8d4a9069-1ade-4b4e-9702-ef05ff3ecf97`, Description `af390993-91be-4db5-b8d9-c4a2dca60ce8`, Client Name `75747da4-250c-42e7-8d65-98d9d503060f`, Email `ecdc4bdb-8444-4074-b32c-cd4c36274f25`.

## Weekly client report (redesigned 12 Nov 2026)
One A4 page per unit: header with date, unit and client; stage tracker; 2 site photos; progress ring with On track / Behind plan badge; elapsed vs planned days; payment date, mobilization, contract start/finish, expected finish; scope, quotation ref and amount; latest updates (AI); project manager.
Data is read live from the unit's ClickUp tasks when the PDF is made (`cuReportInfo_`): Actual POC %, Payment Date, Site In (contract start), Site Due Date (contract finish and expected finish), Request Type, Quotation No., Amount, PM, task status (stage) and Construction Status = Completed (Quality Inspection). Mobilization = days from Payment Date to Site In. "On track" means actual % is at most 10 points under planned % (ON_TRACK_SLACK).
The WhatsApp button flow in ClickUp is unchanged.

## Sheet tabs
- **Units**: Unit, Active, Project, Client name, Client email, Source (ClickUp or Manual). 15 Authority units are Manual; ClickUp rows are rebuilt by sync.
- **Scope**: Unit → Item (the products each villa bought). These are the photo chips.
- **Unit Contractors**: Unit → Contractor (from ClickUp). The manpower dropdown shows only these, plus "Other contractor…".
- **Contractors**: Contractor, Active, Same as (an alias maps to one name). Active = No hides a name.
- **Trades**, **Engineers**, **Settings** (report emails, hours, weekly day Thursday, "Send weekly to clients" = No).
- **Photos**, **Visits** (work done, manpower total, internal notes), **Manpower** (Visit ID, Date, Unit, Project, Contractor, Trade, Count, By, Status, Item), **Reports**.

## Automations (Apps Script triggers)
- `dailyJob` daily at Settings hour (19:00): a PDF per engineer and a manpower PDF, emailed to "Daily report to".
- `weeklyJob` Thursday 17:00: a weekly PDF per unit (internal; sent to clients only if the setting is Yes).
- `syncJob` every 30 min: ClickUp → Units / Scope / Unit Contractors.
- The Sheet menu "O West Camera" has: make/send reports, delete selected photos/visits, set password, connect ClickUp, sync, apply schedule, refresh lists.

## Rules the owner asked for
- Very simple and very fast; small photos (≤1280 px, about 450 KB); no web fonts.
- At least 1 photo, any number. Photos are never cropped in PDFs.
- Internal notes are saved but **never shown in reports**.
- Daily reports are open to anyone. Manpower and weekly reports are hidden ("Extra reports") and code-protected; only the owner uses them.
- Name placeholder: "e.g. Mohamed Emira".
- Data is structured so questions like "unit progress in the last 2 weeks, manpower by contractor and item" can be answered from the Visits and Manpower tabs.

## How to update
- **Website**: edit `index.html`, then commit and push to GitHub. It's live in about 1 minute; no other step.
- **Backend**: paste the new `Code.gs` in Apps Script, then Ctrl+S. Then Deploy → Manage deployments → ✏️ → New version → Deploy. The URL stays the same.

## Status as of 10 Nov 2026
Done: site live, ClickUp connected (168 units synced), hidden extra reports, contractor list per unit.

Still to check:
1. Run `setPasswordNow` once if it hasn't been run yet.
2. Deploy a new version of Apps Script if the latest Code.gs isn't deployed.
3. Record one full test visit, then run "Make today's daily reports (no email)" and check the PDF.
4. Delete the old test rows from 5 Nov.
5. Add engineer emails in Engineers.
6. Decide whether "Real Estate" in Contractors is a real contractor.

Later: custom report buttons and templates the owner will describe.
