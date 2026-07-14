# Daily Tracker (PWA)

A personal daily routine / habit / score tracker, ported from your `DailyTracker_v6` Excel workbook into a small website you can install to your iPhone home screen. No backend, no login — everything is saved in Safari's local storage on your phone.

## What's included

- `index.html`, `styles.css`, `app.js` — the app (plain HTML/CSS/JS, no build step, no dependencies)
- `manifest.json`, `sw.js`, `icons/` — makes it installable ("Add to Home Screen") and usable offline
- Tabs: **Today** (routine tasks, meals log, timing tracker, gym schedule, side missions, score calendar), **Monthly** (gym/sauna/red-light/HBOT session tracker), **Archive** (auto-saved month summaries), **Notes** (your reference notes from the spreadsheet, editable)

## What's different from the spreadsheet (on purpose)

- **Daily reset is automatic.** At midnight, today's score is saved to the calendar and all checkboxes reset — no more manual "+ Add to Calendar" button or clearing ticks by hand.
- **Monthly archiving is automatic**, same idea as the VBA macro, just triggered by the browser instead of `Workbook_Open`.
- **Totals are calculated dynamically** from whatever tasks/points you have configured, rather than a few hardcoded numbers in the original formulas (which had drifted out of sync with the actual task list — e.g. the meals max was hardcoded to 24 pts but the four meal rows add up to 32).
- **Last Coffee is now included in your daily total.** In the spreadsheet it was scored but never added into the subtotal — looked like an oversight, so I folded it in.
- Every task list (routine, meals, side missions) is editable in the app itself: tap ✎ to rename/re-point a task, 🗑 to delete, or use the "+ Add" row to create new ones.

If you'd rather any of these matched the original behaviour exactly, tell me and I'll adjust — it's a small change in `app.js`.

## Run it locally first (in VS Code)

1. Open the `daily-tracker-app` folder in VS Code.
2. Install the **Live Server** extension (or any static server) and click "Go Live" — or run `python3 -m http.server 8080` from a terminal in this folder.
3. Visit the local address in a browser and click around to sanity-check it before deploying.

## Put it online with GitHub Pages (free, gives you a permanent HTTPS URL)

This step matters: Safari's "Add to Home Screen" and offline support work far more reliably from a real HTTPS URL than from a file on your computer.

1. Create a new **public** repo on GitHub, e.g. `daily-tracker`.
2. From this folder:
   ```
   git init
   git add .
   git commit -m "Initial daily tracker PWA"
   git branch -M main
   git remote add origin https://github.com/<your-username>/daily-tracker.git
   git push -u origin main
   ```
3. On GitHub: **Settings → Pages → Source → Deploy from a branch → `main` / root → Save**.
4. After a minute, your site is live at `https://<your-username>.github.io/daily-tracker/`.

## Add to your iPhone home screen

1. Open the GitHub Pages URL in **Safari** on your iPhone (must be Safari, not Chrome, for this to work).
2. Tap the **Share** icon → **Add to Home Screen** → **Add**.
3. Launch it from the home screen icon — it opens full-screen, no browser chrome, and works offline after the first load.

## A note on your data

Everything lives in Safari's local storage on that one phone. That's the simplest, zero-cost setup, and what you asked for — but it also means: if you delete the app data, clear Safari storage, or switch phones, your history goes with it. If you want a cheap safety net later, the easiest addition is an "export to JSON" button that saves a backup file you can re-import — say the word and I'll add it.

## Making future changes

Everything is plain HTML/CSS/JS in three files (`index.html`, `styles.css`, `app.js`) — no framework, no build step. Edit in VS Code, refresh the browser to see changes, then `git push` when you're happy and GitHub Pages updates automatically within a minute or two.
