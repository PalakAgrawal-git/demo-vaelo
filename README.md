# Vaelo — internal tools (prototype)

Role-based login in front of the Vaelo content pipeline, plus scoped client
portals, a finance/billing panel, and optional Google Sheets sync.

## Files
- `vaelo-login.html` — the app: login + routing to the four roles.
- `vaelo-content.html` — the content pipeline (embedded, unchanged).
- `vaelo-ops.html` — the ops hub (client pipeline, production, tasks, approvals).
- `vaelo-apps-script.gs` — Google Apps Script backend for Sheet sync.
- `VAELO-SYNC-SETUP.md` — how to wire up the Google Sheet sync.

## Roles
- **Team** — opens the content pipeline with edit access.
- **Dhruv (lead)** — pipeline + internal-approval queue + team activity.
- **Client staff** — their brand's pipeline (read-only) + approvals.
- **Client head** — same as staff + expenses/billing.

## Google Sheet sync
The sync URL is **not** committed. On the login screen use **Connect / change
Google Sheet** to paste your Apps Script `/exec` URL (stored per-device in the
browser). For local convenience you can instead create a git-ignored
`vaelo-config.js` that sets `window.VAELO_SYNC_URL`.

## Note
This is a **prototype**. Credentials live in the page source — it separates
views but is not real security. Do not host publicly with real passwords or
client data. See `vaelo-systems-dev-brief.md` (kept out of this repo) for the
path to a proper backend.
