# Daily Tracker PWA

Plain HTML/CSS/JS PWA (`index.html`, `styles.css`, `app.js`), no build step, no dependencies. State lives in the browser's localStorage on the user's phone.

Deployed via **GitHub Pages** from the `main` branch of `github.com/ab7888/daily-tracker`. The app on the user's phone loads from that hosted URL (not local files), so changes only reach the phone after being pushed to `main` — GitHub Pages then redeploys automatically within a minute or two.

## Workflow

After making and verifying a change (per the `run`/`verify` skill), commit and push to `main` without asking first — the user has asked for this to happen automatically going forward. Use a clear commit message describing the change. No need to check in before pushing on this repo.
