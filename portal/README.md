# O-Plus Client Portal

The agreed record between O West Extra Works and each client. Clients use it to approve scope items, date changes and material selections (with sample photos), and to send concerns that the PM answers. It replaces WhatsApp as the reference. WhatsApp stays for quick messages.

It is **separate from O West Site Camera**. It has its own Apps Script and its own Google Sheet, and it only *reads* the Site Camera sheet (Units, Engineers, Photos, Reports, ClickUp Tasks). Nothing in `index.html` or `backend/Code.gs` at the repo root is used or changed.

| Piece | Where |
|---|---|
| Client page | `https://xowestx.github.io/owest-site-camera/portal/#t=<token>`. Each unit has its own private link, sent by email or WhatsApp. |
| Team page (PMs) | `https://xowestx.github.io/owest-site-camera/portal/pm.html` |
| Backend | `portal/backend/Code.gs`, an Apps Script bound to a new Google Sheet "O-Plus Client Portal" |
| Photos | Drive folder `O West Client Portal/<unit>/`. Files stay private and are served only through the backend. |

## How it works

- **Client access:** a private link per unit, with no account and no password. To approve or reject an item, the client enters a 6-digit code that we email to them. The decision is recorded with their name, the date and time, and the email the code went to.
- **Versions:** a rejected item is revised and sent again as version 2. The old version stays in the record, marked "Replaced".
- **The record:** the "Portal Log" tab gets a row for every send, approval, rejection, concern and reply. It is protected, so only the sheet owner can edit it.
- **Snag list (handover):** from Quality Inspection on, the client gets a Snags tab. The client or the PM adds a snag with a room and photos. The PM marks it fixed with an after photo. The client then confirms it is fixed (closed) or sends it back with a note. Emails are grouped: at most one per unit every 30 minutes, so a walk-through of 20 snags is not 20 emails.
- **Arabic:** the client page has an English / عربي button in the top bar. It switches the page to Arabic, right to left, and remembers the choice on that phone. What the PM types (titles, details, replies) shows as written. Emails and the team page stay in English.
- **Variation orders:** a fourth item type for extra work outside the contract. Each new one is numbered per unit (VO-01, VO-02 …), and the client approves or rejects it with the email code like any other item. A revision keeps its number. No prices: the client page says the cost is agreed with the PM separately. Clients can also tap **Request extra work** on Approvals; it arrives as a concern titled "Extra work request", and the PM answers with a variation order.
- **Handover certificate:** when every snag is closed, the PM taps **Issue handover certificate** on the unit page (handover date, what was handed over such as keys and meter readings, notes). The client reviews it in the portal (approved items, snag summary) and signs with the 6-digit email code, or taps "Something is missing" to send it back. On signing, a branded PDF is saved in the unit's Drive folder and emailed to the client and the team. A signed certificate is final. Records are in the "Portal Handover" tab.
- **Overview (management dashboard):** on the team page, tap **Overview**. It shows units on site, late units (past the handover date), units 15+ points behind plan (planned % runs in a straight line from Site in to Site due), units with no Site Camera visit in 7 days, manpower for the last 14 days, top contractors by worker-days, a PM table and a unit list, with filters. ClickUp is read by an hourly job into the "Portal Dashboard" tab, so the page opens fast. Use **O-Plus Portal → Refresh dashboard now** to read it right away. No amounts or prices are shown.
- **Home page data:** progress %, stage, handover date and PM come from the unit's ClickUp tasks (cached for 3 hours). The newest site photos and the latest weekly report come from Site Camera.
- **Emails:** the client is emailed when they are invited, when an item is sent and when the PM replies. The team ("Team emails" in Settings) is emailed on every approval, rejection, concern and client reply.

## Setup (once, about 10 minutes)

Use the same Google account that owns the Site Camera sheet (x.owest.x), so the portal can read it and its photos.

1. Create a new Google Sheet named **O-Plus Client Portal**.
2. In that sheet, open Extensions → Apps Script. Paste `portal/backend/Code.gs` and save.
3. Reload the sheet. Run the menu **O-Plus Portal → Set up (run once)** and allow the permissions.
4. **O-Plus Portal → Set team password**. PMs sign in with this password.
5. **O-Plus Portal → Connect ClickUp**. Paste a ClickUp API token. This is only for progress, handover date and PM.
6. In the **Settings** tab, fill **Team emails** (comma separated).
7. In Apps Script, go to Deploy → New deployment → Web app. Set Execute as **Me** and Who has access **Anyone**. Copy the `/exec` URL.
8. Put that URL in `portal/portal.js` (`API_URL`), then commit. GitHub Pages updates in about a minute.

## Daily use (PM)

1. Open the team page and sign in. Choose your name once on each phone.
2. Open a unit and check the client email. Tap **Send invitation**, or **Copy client link** to paste it on WhatsApp.
3. Use **Send new item for approval**: pick the type, title, details ("Label: value" lines show as a table) and sample photos.
4. **Needs your action** lists rejected items to revise and concerns to answer, oldest first.
5. **Snags:** on the unit page, use **Add snag** at inspection. Open a snag and use **Mark as fixed** with at least one after photo. Units with open snags show in **Needs your action**.
6. **Handover:** once all snags are closed, use **Issue handover certificate** at the bottom of the unit page.
7. If a link leaks, use **Make a new link**. The old link stops working at once.
