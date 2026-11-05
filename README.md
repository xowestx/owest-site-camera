# O West Site Camera

Phone page for site engineers: pick a unit, take photos, done.

- **Page**: `index.html`, served by GitHub Pages.
- **Backend**: `backend/Code.gs`, a Google Apps Script bound to the "O West Site Camera" Google Sheet.
  Deployed as a Web app (Execute as: Me, Who has access: Anyone). Its `/exec` URL goes in `API_URL` in `index.html`.
- **Units**: first tab of the sheet (Unit | Active | Project). Set Active to `No` to hide a unit.
- **Photos**: `O West Site Photos / <Unit> / <YYYY-MM-DD> / <Unit>_<date>_<time>_<Category>_<NN>.jpg` in the sheet owner's Drive. Every upload is logged in the sheet's `Log` tab.
