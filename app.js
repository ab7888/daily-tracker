/* =========================================================================
   Daily Tracker — vanilla JS PWA
   Ported from "DailyTracker_v6" Excel workbook.
   All data lives in localStorage on this device (phone-only, per your choice).
   ========================================================================= */

const STORAGE_KEY = "dailyTracker.state.v1";

/* --------------------------- Custom icon set ----------------------------
   Small hand-drawn line icons (24x24, currentColor stroke) standing in for
   the emoji this UI used to lean on — keeps section chrome legible without
   relying on the OS emoji font.
   ------------------------------------------------------------------------- */
const ICON_SVG = {
  medal: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="8" r="5"/><path d="M9 12.5 7 20l5-3 5 3-2-7.5"/></svg>',
  clipboard: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><rect x="5" y="4" width="14" height="17" rx="2"/><rect x="9" y="2.5" width="6" height="3" rx="1"/><path d="M8.5 11l1.5 1.5L13 9"/><line x1="8.5" y1="15" x2="15.5" y2="15"/></svg>',
  salad: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><path d="M4 11a8 8 0 0 0 16 0"/><line x1="4" y1="11" x2="20" y2="11"/><line x1="9" y1="8" x2="9" y2="5"/><line x1="12" y1="8" x2="12" y2="4"/><line x1="15" y1="8" x2="15" y2="5"/></svg>',
  clock: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="8"/><path d="M12 8v4l3 2"/></svg>',
  moon: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><path d="M18 12.5A7.5 7.5 0 1 1 11.5 5a6 6 0 0 0 6.5 7.5Z"/></svg>',
  fork: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><path d="M7 2v6M10 2v6M13 2v6"/><path d="M7 8a3 3 0 0 0 6 0"/><line x1="10" y1="8" x2="10" y2="22"/></svg>',
  dumbbell: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><rect x="2" y="9" width="3" height="6" rx="1"/><rect x="19" y="9" width="3" height="6" rx="1"/><line x1="5" y1="12" x2="19" y2="12"/><rect x="6" y="7" width="2" height="10" rx="1"/><rect x="16" y="7" width="2" height="10" rx="1"/></svg>',
  coffee: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><path d="M5 9h11v6a4 4 0 0 1-4 4H9a4 4 0 0 1-4-4V9Z"/><path d="M16 10h1.5a2.5 2.5 0 0 1 0 5H16"/><path d="M8 5c0 1-1 1-1 2M11 5c0 1-1 1-1 2"/></svg>',
  noPhone: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><rect x="7" y="2" width="10" height="20" rx="2"/><line x1="11" y1="18" x2="13" y2="18"/><line x1="3" y1="3" x2="21" y2="21"/></svg>',
  target: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="8"/><circle cx="12" cy="12" r="4.5"/><circle cx="12" cy="12" r="1"/></svg>',
  calendar: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="5" width="18" height="16" rx="2"/><line x1="3" y1="10" x2="21" y2="10"/><line x1="8" y1="3" x2="8" y2="7"/><line x1="16" y1="3" x2="16" y2="7"/></svg>',
  archive: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="5" rx="1.5"/><path d="M4 9v9a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9"/><line x1="10" y1="13" x2="14" y2="13"/></svg>',
  book: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><path d="M4 5a2 2 0 0 1 2-2h6v18H6a2 2 0 0 1-2-2V5Z"/><path d="M20 5a2 2 0 0 0-2-2h-6v18h6a2 2 0 0 0 2-2V5Z"/><line x1="12" y1="3" x2="12" y2="21"/></svg>',
  save: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><path d="M5 3h11l3 3v13a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2Z"/><path d="M8 3v6h7V3"/><path d="M7 21v-7h10v7"/></svg>',
  note: '<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><rect x="4" y="3" width="12" height="15" rx="1.5"/><line x1="7" y1="7" x2="13" y2="7"/><line x1="7" y1="10.5" x2="13" y2="10.5"/><line x1="7" y1="14" x2="10.5" y2="14"/></svg>',
  sauna: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><rect x="10" y="3" width="4" height="12" rx="2"/><circle cx="12" cy="18" r="3"/><path d="M5 9c1 1 1 2 0 3M19 9c-1 1-1 2 0 3"/></svg>',
  sun: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="4"/><line x1="12" y1="2" x2="12" y2="5"/><line x1="12" y1="19" x2="12" y2="22"/><line x1="2" y1="12" x2="5" y2="12"/><line x1="19" y1="12" x2="22" y2="12"/><line x1="4.5" y1="4.5" x2="6.5" y2="6.5"/><line x1="17.5" y1="17.5" x2="19.5" y2="19.5"/></svg>',
  wind: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><path d="M3 8h12a3 3 0 1 0-3-3"/><path d="M3 16h15a3 3 0 1 1-3 3"/></svg>',
  trash: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><path d="M4 7h16"/><path d="M9 7V4.5a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1V7"/><path d="M6 7l1 13a1.5 1.5 0 0 0 1.5 1.5h7a1.5 1.5 0 0 0 1.5-1.5l1-13"/><line x1="10" y1="11" x2="10" y2="17"/><line x1="14" y1="11" x2="14" y2="17"/></svg>',
  upload: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3v12"/><path d="M7 8l5-5 5 5"/><path d="M4 15v3a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-3"/></svg>',
  droplet: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3c3 4 6 8 6 11a6 6 0 0 1-12 0c0-3 3-7 6-11Z"/></svg>',
  stretch: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="5" r="2"/><path d="M12 8v6"/><path d="M7 11l5-2 5 2"/><path d="M9 20l3-6 3 6"/></svg>',
  foot: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><ellipse cx="11" cy="14" rx="4" ry="7"/><circle cx="9" cy="4" r="1.2"/><circle cx="12" cy="3" r="1.2"/><circle cx="15" cy="4.5" r="1.2"/></svg>',
  footsteps: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><ellipse cx="7" cy="8" rx="2" ry="3.2" transform="rotate(-15 7 8)"/><ellipse cx="16" cy="15" rx="2" ry="3.2" transform="rotate(15 16 15)"/></svg>',
  headphones: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><path d="M4 14v-2a8 8 0 0 1 16 0v2"/><rect x="2.5" y="14" width="4" height="6" rx="1.5"/><rect x="17.5" y="14" width="4" height="6" rx="1.5"/></svg>',
  lotus: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="6" r="2"/><path d="M12 8v3"/><path d="M6 19c0-3 3-5 6-5s6 2 6 5"/><path d="M9 14l-3 3M15 14l3 3"/></svg>',
  shower: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><path d="M7 4a5 5 0 0 1 10 0"/><line x1="12" y1="4" x2="12" y2="2"/><rect x="5" y="4" width="14" height="4" rx="2"/><line x1="8" y1="12" x2="8" y2="14"/><line x1="12" y1="12" x2="12" y2="14"/><line x1="16" y1="12" x2="16" y2="14"/><line x1="8" y1="17" x2="8" y2="19"/><line x1="12" y1="17" x2="12" y2="19"/><line x1="16" y1="17" x2="16" y2="19"/></svg>',
  pill: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><path d="M6.5 17.5a5 5 0 0 1 0-7l4-4a5 5 0 0 1 7 7l-4 4a5 5 0 0 1-7 0Z"/><line x1="10" y1="10" x2="14" y2="14"/></svg>',
  sparkle: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3v4M12 17v4M3 12h4M17 12h4M6 6l2.5 2.5M15.5 15.5 18 18M18 6l-2.5 2.5M8.5 15.5 6 18"/></svg>',
  brain: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><circle cx="9" cy="10" r="4"/><circle cx="14" cy="8" r="5"/><circle cx="17" cy="12.5" r="3"/><circle cx="6.5" cy="19" r="1.2"/><circle cx="4" cy="22" r="0.8"/></svg>',
  flame: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3c2 3-1 4-1 7a3 3 0 0 0 6 0c0-1-.5-2-1-3 1 .5 2 2 2 4a6 6 0 0 1-12 0c0-4 3-5 3-8 1 1 2 1 3 0Z"/></svg>',
  pot: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><path d="M4 11h16"/><path d="M5 11v6a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2v-6"/><path d="M9 11V8a3 3 0 0 1 6 0v3"/><line x1="12" y1="5" x2="12" y2="3"/></svg>',
  inhaler: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><rect x="8" y="8" width="8" height="12" rx="2"/><path d="M10 8V5a2 2 0 0 1 2-2h0a2 2 0 0 1 2 2v3"/><line x1="10" y1="20" x2="10" y2="22"/><line x1="14" y1="20" x2="14" y2="22"/></svg>',
  plate: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><circle cx="12" cy="12" r="5"/></svg>',
  undo: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><path d="M4 11h11a5 5 0 0 1 0 10h-2"/><path d="M9 6 4 11l5 5"/></svg>',
  pin: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><path d="M12 21s-6.5-6.1-6.5-11A6.5 6.5 0 0 1 18.5 10c0 4.9-6.5 11-6.5 11Z"/><circle cx="12" cy="10" r="2.2"/></svg>',
  split: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><path d="M12 4v6"/><path d="M12 10 4 20"/><path d="M12 10l8 10"/></svg>',
  checkCircle: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M8 12.5l2.5 2.5L16 9.5"/></svg>',
  listChecks: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><path d="M4 6.5 5.5 8 8 5"/><line x1="11" y1="6.5" x2="20" y2="6.5"/><path d="M4 12.5 5.5 14 8 11"/><line x1="11" y1="12.5" x2="20" y2="12.5"/><path d="M4 18.5 5.5 20 8 17"/><line x1="11" y1="18.5" x2="20" y2="18.5"/></svg>',
  sunrise: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><path d="M5 18a7 7 0 0 1 14 0"/><line x1="2" y1="18" x2="22" y2="18"/><line x1="12" y1="2" x2="12" y2="6"/><line x1="4.5" y1="9" x2="7" y2="10.5"/><line x1="19.5" y1="9" x2="17" y2="10.5"/></svg>',
  cloud: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><path d="M7 18a4.5 4.5 0 0 1-.5-8.97A5.5 5.5 0 0 1 17 8.5a4 4 0 0 1-.7 7.94"/><path d="M16.3 16.44H7"/></svg>',
  dice: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="18" height="18" rx="3"/><circle cx="8" cy="8" r="1.2"/><circle cx="16" cy="8" r="1.2"/><circle cx="8" cy="16" r="1.2"/><circle cx="16" cy="16" r="1.2"/><circle cx="12" cy="12" r="1.2"/></svg>',
  folder: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><path d="M3 7a2 2 0 0 1 2-2h4l2 2.5h8a2 2 0 0 1 2 2V18a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V7Z"/></svg>',
  newspaper: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="5" width="14" height="15" rx="1.5"/><line x1="6" y1="9" x2="14" y2="9"/><line x1="6" y1="12.5" x2="14" y2="12.5"/><line x1="6" y1="16" x2="10" y2="16"/><path d="M17 8h2a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H7"/></svg>',
  train: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><rect x="5" y="3" width="14" height="13" rx="4"/><line x1="5" y1="10" x2="19" y2="10"/><circle cx="8.5" cy="13.5" r="0.8"/><circle cx="15.5" cy="13.5" r="0.8"/><path d="M8 16l-2 5M16 16l2 5"/></svg>',
  rain: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><path d="M7 15a4.5 4.5 0 0 1-.5-8.97A5.5 5.5 0 0 1 17 5.5a4 4 0 0 1-.7 7.94"/><line x1="8" y1="18" x2="8" y2="21"/><line x1="12" y1="18" x2="12" y2="21"/><line x1="16" y1="18" x2="16" y2="21"/></svg>',
  snow: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><path d="M7 14a4.5 4.5 0 0 1-.5-8.97A5.5 5.5 0 0 1 17 4.5a4 4 0 0 1-.7 7.94"/><path d="M8 18v3M6.5 19.5h3M12 18v3M10.5 19.5h3M16 18v3M14.5 19.5h3"/></svg>',
  storm: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><path d="M7 13a4.5 4.5 0 0 1-.5-8.97A5.5 5.5 0 0 1 17 3.5a4 4 0 0 1-.7 7.94"/><path d="M13 14l-4 6h3l-1 5 5-7h-3l1-4Z"/></svg>',
  refresh: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><path d="M4 12a8 8 0 0 1 14-5.3M20 12a8 8 0 0 1-14 5.3"/><path d="M18 3v4h-4M6 21v-4h4"/></svg>',
  heart: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><path d="M12 20s-7-4.35-9.5-8.8C.8 7.9 2.4 4.5 6 4.1c2-.2 3.6.8 6 3.1 2.4-2.3 4-3.3 6-3.1 3.6.4 5.2 3.8 3.5 7.1C19 15.65 12 20 12 20Z"/><path d="M6 11.5h2.5l1.5-2.5 2 5 1.5-2.5H16"/></svg>'
};

function injectIcons() {
  $$("[data-icon]").forEach((el) => {
    const svg = ICON_SVG[el.dataset.icon];
    if (svg) el.innerHTML = svg;
  });
}

// Inline icon markup for templates built via innerHTML at render time
// (as opposed to the static [data-icon] placeholders injectIcons fills once).
function iconTag(name, extraClass) {
  return `<span class="icon icon-sm${extraClass ? " " + extraClass : ""}">${ICON_SVG[name] || ""}</span>`;
}

const GYM_DAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];
const MONTH_NAMES = ["January","February","March","April","May","June","July","August","September","October","November","December"];

/* ------------------------- New training routine -------------------------- */
// Sourced from New-Training-Routine-v2.xlsx. This is the single source of
// truth: the Training tab AND the Gym Schedule card both read from it, so
// there's nowhere else that needs updating if the program changes again.

const TRAINING_SESSIONS = {
  pushA: {
    key: "pushA",
    day: 0, // Monday
    shortTitle: "Push A",
    title: "Push A — HSPU Strength",
    focus: "Primary: anterior delts · triceps · upper traps · serratus",
    exercises: [
      { section: "SKILL PRIMER", name: "Wall Handstand Hold", sets: 3, reps: "30–45s", weight: "Bodyweight", notes: "Extend hold duration weekly · Dial in straight body line before pressing" },
      { section: null, name: "Pike Push-Up (slow eccentric)", sets: 3, reps: "8–10", weight: "Bodyweight", notes: "Add deficit once 10 reps easy · 3s down — warms up the exact HSPU pattern" },
      { section: "MAIN STRENGTH", name: "Strict Barbell Overhead Press", sets: 4, reps: "5–8", weight: "42.5kg", notes: "+2.5kg when you hit 4×8 · No leg drive. Core lift — log every set" },
      { section: null, name: "Wall HSPU (or deficit)", sets: 4, reps: "5–8", weight: "Bodyweight / deficit", notes: "Add 2–3cm deficit when 4×8 clean · Work toward freestanding. Head to floor, full lockout" },
      { section: null, name: "Weighted Dips", sets: 3, reps: "8–10", weight: "10–15kg added", notes: "+2.5kg when 3×10 achieved · Leans forward slightly — chest and tricep pressing" },
      { section: "ACCESSORY", name: "Overhead Tricep Extension (EZ bar)", sets: 3, reps: "10–12", weight: "22.5kg", notes: "+2.5kg when 3×12 · Long-head tricep — HSPU lockout strength" },
      { section: null, name: "Lateral Raise (3s eccentric)", sets: 3, reps: "12–15", weight: "10kg DBs", notes: "+2kg when 3×15 · Medial delt endurance for handstand holds" },
      { section: null, name: "Serratus Punch / Cable Push-Through", sets: 3, reps: "12–15", weight: "12kg", notes: "+2.5kg when 3×15 · Scapular upward rotation — critical for HSPU ceiling" },
      { section: null, name: "Wrist Prep", sets: 2, reps: "15 each", weight: "5kg", notes: "Maintain · Flexion + extension — non-negotiable for handstand volume" }
    ]
  },
  pullA: {
    key: "pullA",
    day: 1, // Tuesday
    shortTitle: "Pull A",
    title: "Pull A — Weighted Pull-Up & OAP Progression",
    focus: "Primary: lats · biceps · scapular depressors · unilateral pulling strength · core-to-lever carryover",
    exercises: [
      { section: "MAIN LIFT", name: "Weighted Pull-Up (pronated, sh-width)", sets: 4, reps: "5–8", weight: "+20kg", notes: "Progress: +2.5kg at 4×6 · Target: +30kg" },
      { section: "OAP PROGRESSION", name: "Archer Pull-Up", sets: 4, reps: "5–6 ea. side", weight: "Bodyweight", notes: "Extend non-working arm straighter each week" },
      { section: null, name: "One-Arm Negative (OAN)", sets: 3, reps: "3–5 ea. side", weight: "Bodyweight", notes: "5s controlled lower per rep" },
      { section: null, name: "One-Arm Dead Hang", sets: 3, reps: "10–20s ea.", weight: "Bodyweight", notes: "Add load when 20s comfortable" },
      { section: "SUPPLEMENTARY", name: "Straight-Arm Lat Pulldown", sets: 4, reps: "10–12", weight: "50kg", notes: "Progress: +2.5kg at 4×12" },
      { section: null, name: "Barbell Pendlay Row", sets: 4, reps: "6–8", weight: "60kg", notes: "Progress: +2.5kg at 4×8" },
      { section: null, name: "Seated Cable Row (wide grip)", sets: 3, reps: "10", weight: "62.5kg", notes: "Progress: +2.5kg at 3×10" },
      { section: "ACCESSORY", name: "Bicep Curl (drop set)", sets: 2, reps: "10 → failure", weight: "14kg → 8kg", notes: "" },
      { section: null, name: "Cable Face Pull", sets: 3, reps: "15", weight: "17.5kg", notes: "" },
      { section: null, name: "Scapular Pull-Ups", sets: 3, reps: "10–12", weight: "Bodyweight", notes: "" },
      { section: "CORE FINISHER", name: "Toes-to-Bar", sets: 3, reps: "8–12", weight: "Bodyweight", notes: "Strict — no kipping · Feeds directly into OAP and front lever compression" },
      { section: null, name: "Dragon Flag (negative)", sets: 3, reps: "4–6", weight: "Bodyweight", notes: "5s controlled lower · Build to full dragon flag before adding load" },
      { section: null, name: "L-Sit Hold (parallel bars or floor)", sets: 3, reps: "10–20s", weight: "Bodyweight", notes: "Rest 45s between · Progress: tuck → one leg extended → full L-sit" }
    ]
  },
  legsA: {
    key: "legsA",
    day: 2, // Wednesday
    shortTitle: "Legs A",
    title: "Legs A — Squat & Deadlift",
    focus: "Primary: quads · glutes · hamstrings · spinal erectors",
    exercises: [
      { section: "MAIN LIFTS", name: "ATG Barbell Squat", sets: 4, reps: "5–8", weight: "70kg", notes: "+2.5kg when 4×8 achieved · Full depth, heels flat. Drive numbers up each session" },
      { section: null, name: "Barbell Deadlift (conventional)", sets: 4, reps: "4–6", weight: "80kg", notes: "+2.5kg when 4×6 achieved · Brace hard. Your DL is ahead of squat — keep pushing both" },
      { section: "SUPPLEMENTARY", name: "Leg Press (feet low, narrow)", sets: 3, reps: "10–12", weight: "120kg", notes: "+5kg when 3×12 · Quad isolation after squats" },
      { section: null, name: "Romanian Deadlift", sets: 3, reps: "10–12", weight: "62.5kg", notes: "+2.5kg when 3×12 · Hamstring stretch is the goal — not a max effort lift" },
      { section: null, name: "Leg Extension (3s eccentric)", sets: 3, reps: "12–15", weight: "45kg stack", notes: "+5kg when 3×15 · VMO strength. 3-second lower" },
      { section: "CALF & CORE", name: "Standing Calf Raise", sets: 4, reps: "8–12", weight: "90kg", notes: "+5kg when 4×12 · Full stretch at bottom — not partial reps" },
      { section: null, name: "Ab Wheel Rollout", sets: 3, reps: "8–10", weight: "Bodyweight", notes: "Elevate feet when 3×10 easy · Anti-extension core. Direct HSPU and handstand carryover" },
      { section: null, name: "Hanging Leg Raise", sets: 3, reps: "10–12", weight: "Bodyweight", notes: "Add ankle weight when 3×12 easy · Straight legs. Hip flexor compression for lever work" }
    ]
  },
  pushB: {
    key: "pushB",
    day: 3, // Thursday
    shortTitle: "Push B",
    title: "Push B — Handstand Skill & Shoulder Accessory",
    focus: "Primary: handstand balance · scapular control · shoulder endurance · anti-extension core",
    exercises: [
      { section: "HANDSTAND SKILL", name: "Freestanding Handstand Attempts", sets: "5–8", reps: "Max hold", weight: "Bodyweight", notes: "Log best hold each session · Target: 20s → 30s → 60s" },
      { section: null, name: "Wall Handstand (chest to wall)", sets: 3, reps: "45–60s", weight: "Bodyweight", notes: "Better body line than back-to-wall" },
      { section: null, name: "Handstand Shoulder Taps", sets: 3, reps: "8–12 taps", weight: "Bodyweight", notes: "Against wall · Builds unilateral shoulder stability" },
      { section: null, name: "Pike Compression Hold", sets: 3, reps: "20–30s", weight: "Bodyweight / 5kg on legs", notes: "" },
      { section: "CORE (handstand-specific)", name: "Ab Wheel Rollout", sets: 4, reps: "8–10", weight: "Bodyweight", notes: "Anti-extension core — the exact demand of a held handstand · Elevate feet when 3×10 easy" },
      { section: null, name: "Hollow Body Hold", sets: 3, reps: "30–45s", weight: "Bodyweight", notes: "Mimics the handstand body line exactly · Extend duration before adding load" },
      { section: null, name: "Straddle Press Negatives", sets: 3, reps: "3–5", weight: "Bodyweight", notes: "Lower slowly from handstand to straddle seated position · Feeds press-to-handstand long term" },
      { section: "SHOULDER ACCESSORY", name: "Incline DB Press", sets: 3, reps: "10–12", weight: "20kg DBs", notes: "" },
      { section: null, name: "Cable Pull-Over (straight arm)", sets: 3, reps: "12", weight: "32.5kg", notes: "" },
      { section: null, name: "Rear Delt Fly (pause at top)", sets: 3, reps: "12–15", weight: "10kg DBs", notes: "" },
      { section: null, name: "Band Pull-Aparts", sets: 3, reps: "20", weight: "Medium band", notes: "" },
      { section: null, name: "Rotator Cuff External Rotation", sets: 2, reps: "15 each", weight: "5kg", notes: "" }
    ]
  },
  legsB: {
    key: "legsB",
    day: 4, // Friday
    shortTitle: "Legs B",
    title: "Legs B — Single Leg",
    focus: "Primary: glutes · hamstrings · unilateral strength · hip extension",
    exercises: [
      { section: "MAIN LIFTS", name: "Bulgarian Split Squat (BB)", sets: 4, reps: "8–10 ea.", weight: "30–35kg total", notes: "+2.5kg when 4×10 each side · Front foot further = more glute. Most important single-leg builder" },
      { section: null, name: "Single-Leg Deadlift (DB)", sets: 3, reps: "8–10 ea.", weight: "22.5–25kg", notes: "+2.5kg when 3×10 each side · Lower spinal load than barbell DL. Balance is the challenge — go slow" },
      { section: "POSTERIOR CHAIN", name: "Barbell Hip Thrust", sets: 4, reps: "8–10", weight: "65kg", notes: "+2.5kg when 4×10 · Full hip extension, posterior pelvic tilt at top" },
      { section: null, name: "Seated Leg Curl", sets: 4, reps: "10–12", weight: "37.5kg", notes: "+2.5kg when 4×12 · Isolated hamstring. Slow eccentric" },
      { section: null, name: "Cable Pull-Through", sets: 3, reps: "12–15", weight: "32.5kg", notes: "+2.5kg when 3×15 · Hip hinge pattern. Drive hips forward — not a back exercise" },
      { section: null, name: "Nordic Curl (or GHR if available)", sets: 3, reps: "4–6", weight: "Bodyweight", notes: "Add reps weekly — very hard · Best hamstring strength exercise. Use pad for knees" },
      { section: "CALF & CORE", name: "Seated Calf Raise", sets: 4, reps: "12–15", weight: "55kg", notes: "+5kg when 4×15 · Soleus-focused. Complement to standing raises on Legs A" },
      { section: null, name: "Hollow Body Hold", sets: 3, reps: "30–45s", weight: "Bodyweight", notes: "Extend to 60s then add light plate · Full body tension — feeds handstand and HSPU body line" },
      { section: null, name: "Dragon Flag Negatives", sets: 3, reps: "4–6", weight: "Bodyweight", notes: "Slow the eccentric — target 5s down · Core and hip flexor strength. Hard — don't rush progression" }
    ]
  }
};

// Mon..Sun -> session key, or null for a rest day.
const WEEKLY_SESSION_KEYS = ["pushA", "pullA", "legsA", "pushB", "legsB", null, null];

// 0 = Mon .. 4 = Fri. Pancake note: "Best slotted after Legs A, Legs B, and
// Push B (Thursday skill session)".
const PANCAKE_SUGGESTED_DAYS = [2, 3, 4];

const REHAB_EXERCISES = [
  { section: "SPINAL & LATERAL CHAIN", name: "Jefferson Curl", sets: 3, reps: 8, weight: "5–7.5kg", notes: "Roll down vertebra by vertebra · Do NOT rush progression — this is rehab, not strength work" },
  { section: null, name: "Kettlebell Side Bend", sets: 2, reps: "12 ea.", weight: "6kg", notes: "+1kg when 12 reps feel easy each side · Keep hips square, slow and controlled" },
  { section: "HIP & GLUTE STABILITY", name: "Single Leg Bridge (SL)", sets: 3, reps: "8 ea.", weight: "Bodyweight", notes: "Add 2s pause at top when easy · Full extension at top, don't let hips rotate" },
  { section: null, name: "Inward Foot Rotations", sets: 3, reps: "10–15 ea.", weight: "Bodyweight", notes: "Increase range of motion over weeks, not reps · Slow and deliberate" },
  { section: "LOWER LEG & ANKLE", name: "Single Leg Calf Raise", sets: 3, reps: "12–15 ea.", weight: "Bodyweight", notes: "Add load once 15 reps per leg is easy · Full range — deep stretch at bottom, full rise at top" }
];

const PANCAKE_EXERCISES = [
  { section: "ACTIVE WARM-UP", name: "Hip CARs", sets: 2, reps: "5 ea. dir.", weight: "Bodyweight", notes: "Full hip circles, slow and controlled — primes the joint before loading" },
  { section: null, name: "Seated Straddle — active lifts", sets: 3, reps: "10 ea. leg", weight: "Bodyweight", notes: "Activates the hip flexors needed for pancake" },
  { section: null, name: "Frog Stretch", sets: 3, reps: "30s", weight: "Bodyweight", notes: "Inner groin and adductor opener — essential prep before straddle work" },
  { section: "PANCAKE SPECIFIC", name: "Straddle Compression (active)", sets: 4, reps: "10–15", weight: "Bodyweight", notes: "Fold forward actively using hip flexors — don't just hang. Chest to floor is the goal" },
  { section: null, name: "Pancake Hold (passive assisted)", sets: 3, reps: "30–60s", weight: "Light plate on back", notes: "Gravity-assisted fold. Relax into it — don't force" },
  { section: null, name: "Straddle Wall Slides", sets: 3, reps: 10, weight: "Bodyweight", notes: "Walk hands up the wall as you fold — opens chest + adductors" },
  { section: null, name: "Weighted Straddle Sit (isometric)", sets: 3, reps: "45–60s", weight: "5kg plate on legs", notes: "Builds active flexibility, not just passive" },
  { section: "ADDUCTOR & HAMSTRING LOADING", name: "Copenhagen Plank", sets: 3, reps: "20–30s ea.", weight: "Bodyweight", notes: "Weakness here limits straddle width and pancake depth" },
  { section: null, name: "Seated Pike Compression", sets: 3, reps: "10–15", weight: "Bodyweight", notes: "Fold forward with straight legs — same movement as pancake but legs together" },
  { section: null, name: "Jefferson Curl (mobility version)", sets: 2, reps: 6, weight: "2.5–5kg", notes: "Keep very light — this is mobility, not the rehab strength version" },
  { section: "END OF SESSION", name: "Long Pancake Hold", sets: 1, reps: "2–5 min", weight: "Bodyweight / light plate", notes: "Fold and relax completely. Use a timer — this is where the real gains happen" }
];

function sessionForDay(dayIdx) {
  const key = state.weeklySessionKeys[dayIdx];
  return key ? TRAINING_SESSIONS[key] : null;
}

function exerciseId(sessionKey, idx) {
  return `${sessionKey}-${idx}`;
}

// The full exercise list for a session — the fixed program exercises plus
// any the user has added via the Training tab's "Add Exercise" box. Each
// entry is { ex, id, custom, section, showSection } so callers don't care
// where it came from. state.training.done and state.trainingActuals are
// both keyed by `id`. Custom exercises are spliced in next to the last
// exercise sharing their chosen section heading (so an "added" MAIN
// STRENGTH lift sits with the other MAIN STRENGTH lifts, not at the end);
// a custom exercise with no section, or one whose section no longer
// exists, falls to the bottom.
function sessionExerciseEntries(session) {
  const entries = [];
  let cur = null;
  session.exercises.forEach((ex, i) => {
    if (ex.section) cur = ex.section;
    entries.push({ ex, id: exerciseId(session.key, i), custom: false, section: cur });
  });
  (state.customExercises[session.key] || []).forEach((ex) => {
    const wanted = ex.section || null;
    let insertAt = -1;
    for (let i = 0; i < entries.length; i++) {
      if ((entries[i].section || null) === wanted) insertAt = i;
    }
    const entry = { ex, id: ex.id, custom: true, section: wanted };
    if (insertAt === -1) entries.push(entry);
    else entries.splice(insertAt + 1, 0, entry);
  });
  // Header shows only where the section changes from the row above.
  let prev = null;
  entries.forEach((e, i) => {
    e.showSection = (i === 0 || (e.section || null) !== (prev || null)) ? (e.section || "") : "";
    prev = e.section;
  });
  return entries;
}

// Distinct section headings in a session, in order — the options offered
// by the "Add Exercise" section picker.
function sessionSections(session) {
  const seen = [];
  session.exercises.forEach((ex) => {
    if (ex.section && !seen.includes(ex.section)) seen.push(ex.section);
  });
  return seen;
}

// Look up an exercise definition from an id, whether it's a built-in
// (`legsA-3`) or a custom one (`legsA-c<uid>`). Used when archiving a logged
// session into trainingHistory so the report can show what was prescribed.
function resolveExercise(sessionKey, id) {
  const session = TRAINING_SESSIONS[sessionKey];
  const m = /-(\d+)$/.exec(id);
  if (session && m) return session.exercises[Number(m[1])] || null;
  return (state.customExercises[sessionKey] || []).find((e) => e.id === id) || null;
}

// Every logged rep/set/weight for one exercise, newest first — powers the
// history view opened by tapping an exercise. Reads straight off
// trainingHistory (already the source for the Monthly Training Report)
// instead of keeping a second, per-exercise log that could drift out of
// sync with it — one history, two views onto it. Includes today's
// not-yet-archived entry, if there is one, marked in-progress.
function exerciseHistoryFor(id) {
  const out = [];
  (state.trainingHistory || []).forEach((s) => {
    (s.lines || []).forEach((ln) => {
      if (ln.id === id) out.push({ date: s.date, actual: ln.actual, target: ln.target });
    });
  });
  const inProgress = state.trainingActuals[id];
  if (inProgress && inProgress.trim()) {
    out.push({ date: state.currentDate, actual: inProgress.trim(), target: null, inProgress: true });
  }
  out.sort((a, b) => b.date.localeCompare(a.date));
  return out;
}

// Once every exercise in *today's* scheduled session is logged, auto-check
// the Gym Schedule box for today so the daily score reflects it without an
// extra manual tap. One-directional — unticking a single exercise afterward
// won't claw the bonus back. Shared by the checkbox toggle and the Log
// button, since both can be what finally completes the session.
function maybeAutoCheckGymBonus() {
  const todayIdx = todayDayIdx();
  const session = sessionForDay(trainingViewDay);
  if (session && trainingViewDay === todayIdx) {
    const allDone = sessionExerciseEntries(session).every((e) => state.training.done[e.id]);
    if (allDone) state.gymSchedule[todayIdx].done = true;
  }
}

let uidCounter = 1;
function uid() { return "id" + (uidCounter++) + "_" + Math.random().toString(36).slice(2, 7); }

function defaultState() {
  const today = todayISO();
  const d = new Date();
  const sideMissions = [
    ["Shave", 5, ""],
    ["Make bed", 5, ""]
  ].map(([label, points, icon]) => ({ id: uid(), label, points, done: false, time: "", icon }));
  return {
    version: 2,
    currentDate: today,
    routineTasks: [
      ["Face Routine (AM)", 5, "droplet"],
      ["Stretch", 8, "stretch"],
      ["Ankle Training", 5, "foot"],
      ["Walk / Steps", 10, "footsteps"],
      ["Audiobook", 4, "headphones"],
      ["Meditate", 12, "lotus"],
      ["Shower", 5, "shower"],
      ["Supplements", 5, "pill"],
      ["Clean", 5, "sparkle"],
      ["QL Recovery", 5, "footsteps"],
      ["Relapse", -80, "noPhone"],
      ["Mental dump + Controlled silence", 10, "brain"],
      ["Gym Session", 15, "dumbbell"],
      ["Face Routine (PM)", 5, "droplet"],
      ["Language learning", 10, "book"],
      ["Respondent surveys", 5, "brain"],
      ["Face your fears", 10, "flame"],
      ["Evening Wind Down", 5, "moon"],
      ["Meal prep", 8, "pot"],
      ["Asthma Inhaler", 6, "inhaler"],
      ["Water Floss", 2, ""]
    ].map(([label, points, icon]) => ({ id: uid(), label, points, done: false, time: "", icon })),
    meals: [
      ["Healthy Meal 1", 8, "plate"],
      ["Healthy Meal 2", 8, "plate"],
      ["Healthy Meal 3", 8, "plate"],
      ["No Junk Food?", 8, "salad"]
    ].map(([label, points, icon]) => ({ id: uid(), label, points, done: false, time: "", icon })),
    timing: { sleep: "", lastMeal: "", gymFinish: "", lastCoffee: "", noScroll: false },
    gymSchedule: GYM_DAYS.map((day) => ({ day, done: false })),
    weeklySessionKeys: WEEKLY_SESSION_KEYS.slice(),
    sideMissions,
    eliminateToday: { date: today, ids: pickRandomSubset(sideMissions, Math.min(5, sideMissions.length)) },
    projects: [],
    trainingActuals: {},
    trainingLog: {},
    trainingHistory: [],
    customExercises: {},
    training: { done: {}, rehabDone: {}, pancakeDone: {} },
    dailyLog: {},
    dayPlans: {},
    tomorrowPlan: { order: [], times: {} },
    rightNow: [],
    monthly: {
      year: d.getFullYear(),
      month: d.getMonth() + 1, // 1-12
      days: {}
    },
    archive: []
  };
}

function pickRandomSubset(list, n) {
  const ids = list.map((x) => x.id);
  for (let i = ids.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [ids[i], ids[j]] = [ids[j], ids[i]];
  }
  return ids.slice(0, n);
}

function ensureEliminateToday() {
  const today = todayISO();
  if (!state.eliminateToday || state.eliminateToday.date !== today) {
    state.eliminateToday = { date: today, ids: pickRandomSubset(state.sideMissions, Math.min(5, state.sideMissions.length)) };
  }
}

// Maps an already-saved task's exact old (emoji-prefixed) label to its new
// icon-based label/icon — lets loadState() upgrade a device's saved tasks
// in place without touching anything the user has renamed or added.
const ROUTINE_ICON_MIGRATIONS = {
  "🧴 Face Routine (AM)": { label: "Face Routine (AM)", icon: "droplet" },
  "🧘 Stretch": { label: "Stretch", icon: "stretch" },
  "🦶 Ankle Training": { label: "Ankle Training", icon: "foot" },
  "🚶 Walk / Steps": { label: "Walk / Steps", icon: "footsteps" },
  "📚 Audiobook": { label: "Audiobook", icon: "headphones" },
  "🧘 Meditate": { label: "Meditate", icon: "lotus" },
  "🚿 Shower": { label: "Shower", icon: "shower" },
  "💊 Supplements": { label: "Supplements", icon: "pill" },
  "🧹 Clean": { label: "Clean", icon: "sparkle" },
  "🚶 QL Recovery": { label: "QL Recovery", icon: "footsteps" },
  "📵 Relapse": { label: "Relapse", icon: "noPhone" },
  "🧠 Mental dump + Controlled silence": { label: "Mental dump + Controlled silence", icon: "brain" },
  "💪 Gym Session": { label: "Gym Session", icon: "dumbbell" },
  "🧴 Face Routine (PM)": { label: "Face Routine (PM)", icon: "droplet" },
  "📚 Language learning": { label: "Language learning", icon: "book" },
  "🧠 Respondent surveys": { label: "Respondent surveys", icon: "brain" },
  "😤 Face your fears": { label: "Face your fears", icon: "flame" },
  "🌙 Evening Wind Down": { label: "Evening Wind Down", icon: "moon" },
  "👨‍🍳 Meal prep": { label: "Meal prep", icon: "pot" },
  "🧴 Asthma Inhaler": { label: "Asthma Inhaler", icon: "inhaler" }
};

const MEALS_ICON_MIGRATIONS = {
  "🥩 Healthy Meal 1": { label: "Healthy Meal 1", icon: "plate" },
  "🥦 Healthy Meal 2": { label: "Healthy Meal 2", icon: "plate" },
  "🫠 Healthy Meal 3": { label: "Healthy Meal 3", icon: "plate" },
  "🥦 No Junk Food?": { label: "No Junk Food?", icon: "salad" }
};

function migrateTaskIcons(list, map) {
  if (!Array.isArray(list)) return;
  list.forEach((t) => {
    const m = map[t.label];
    if (m && !t.icon) {
      t.label = m.label;
      t.icon = m.icon;
    }
    if (t.time === undefined) t.time = "";
    if (t.icon === undefined) t.icon = "";
  });
}

function todayISO(date) {
  const d = date || new Date();
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

function todayDayIdx() {
  const dow = new Date().getDay(); // 0=Sun..6=Sat
  return dow === 0 ? 6 : dow - 1; // convert to 0=Mon..6=Sun
}

function formatDateShort(iso) {
  const d = new Date(iso + "T00:00:00");
  return d.toLocaleDateString(undefined, { weekday: "short", day: "2-digit", month: "short" });
}

// Forward-compat merge/migration for a raw state blob, regardless of
// whether it came from localStorage (loadState) or the cloud
// (pullOrSeedCloudState) — one shared implementation so the two paths can't
// drift out of sync with each other again (they did, once: a migration only
// got added to loadState(), and cloud-pulling old-shaped data crashed).
function normalizeStateBlob(raw) {
  const base = defaultState();
  // tomorrowPlan used to be a flat array of ad-hoc {id,label,points}
  // entries; it's now { order, times } keyed to real routineTasks ids.
  // Old-shaped data doesn't map onto the new format, so just start fresh.
  const tomorrowPlan = (raw.tomorrowPlan && !Array.isArray(raw.tomorrowPlan))
    ? { order: Array.isArray(raw.tomorrowPlan.order) ? raw.tomorrowPlan.order : [], times: raw.tomorrowPlan.times || {} }
    : { order: [], times: {} };
  const merged = Object.assign({}, base, raw, {
    timing: Object.assign({}, base.timing, raw.timing),
    monthly: Object.assign({}, base.monthly, raw.monthly),
    training: Object.assign({}, base.training, raw.training),
    customExercises: Object.assign({}, base.customExercises, raw.customExercises),
    trainingHistory: Array.isArray(raw.trainingHistory) ? raw.trainingHistory : [],
    dayPlans: Object.assign({}, base.dayPlans, raw.dayPlans),
    tomorrowPlan
  });
  // Gym schedule day *labels* always come fresh from the current program —
  // only the per-day "done" flags are worth keeping from a save.
  if (Array.isArray(raw.gymSchedule)) {
    merged.gymSchedule = GYM_DAYS.map((day, i) => ({
      day,
      done: !!(raw.gymSchedule[i] && raw.gymSchedule[i].done)
    }));
  }
  migrateTaskIcons(merged.routineTasks, ROUTINE_ICON_MIGRATIONS);
  migrateTaskIcons(merged.meals, MEALS_ICON_MIGRATIONS);
  migrateTaskIcons(merged.sideMissions, {});
  return merged;
}

function loadState() {
  const raw = localStorage.getItem(STORAGE_KEY);
  if (!raw) return defaultState();
  try {
    return normalizeStateBlob(JSON.parse(raw));
  } catch (e) {
    console.error("Failed to parse saved state, starting fresh.", e);
    return defaultState();
  }
}

function saveState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  pushStateToCloud();
}

let state = loadState();

/* ------------------------------- Cloud Sync -------------------------------
   Optional: syncs `state` to Supabase when logged in, so it carries over
   between devices. Works fully offline / logged-out too — everything above
   this point already round-trips through localStorage on its own.
   ------------------------------------------------------------------------- */

const SUPABASE_URL = "https://ogakjdqgwwtgwfvetgnq.supabase.co";
const SUPABASE_ANON_KEY = "sb_publishable_J--TUohNkNVSsmKbBsxiyg_7Jj0X7IX";

const supabaseClient = (window.supabase && window.supabase.createClient)
  ? window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY)
  : null;

let currentUser = null;
let cloudPushTimer = null;
let syncStatus = "offline"; // offline | idle | syncing | synced | error

function setSyncStatus(s) {
  syncStatus = s;
  renderAccount();
}

async function initAuth() {
  if (!supabaseClient) return;
  setSyncStatus("idle");
  const { data } = await supabaseClient.auth.getSession();
  currentUser = data.session ? data.session.user : null;
  renderAccount();
  if (currentUser) await pullOrSeedCloudState();

  supabaseClient.auth.onAuthStateChange(async (event, session) => {
    currentUser = session ? session.user : null;
    renderAccount();
    if (event === "SIGNED_IN") await pullOrSeedCloudState();
  });
}

async function pullOrSeedCloudState() {
  if (!supabaseClient || !currentUser) return;
  setSyncStatus("syncing");
  try {
    const { data, error } = await supabaseClient
      .from("daily_tracker_state")
      .select("state")
      .eq("user_id", currentUser.id)
      .maybeSingle();
    if (error) throw error;
    if (data && data.state) {
      state = normalizeStateBlob(data.state);
      localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
      renderAll();
    } else {
      // No cloud copy yet for this account — seed it with what's on this device.
      await pushStateToCloud(true);
    }
    setSyncStatus("synced");
  } catch (err) {
    console.error("Cloud pull failed", err);
    setSyncStatus("error");
  }
}

// Debounced by default (rapid clicks shouldn't each fire a network request);
// pass immediate=true for the one-off "seed the cloud" push right after login,
// or to flush a pending change right before the page backgrounds/closes.
function pushStateToCloud(immediate) {
  if (!supabaseClient || !currentUser) return;
  if (cloudPushTimer) { clearTimeout(cloudPushTimer); cloudPushTimer = null; }
  const doPush = async () => {
    setSyncStatus("syncing");
    try {
      const { error } = await supabaseClient
        .from("daily_tracker_state")
        .upsert({ user_id: currentUser.id, state, updated_at: new Date().toISOString() });
      if (error) throw error;
      setSyncStatus("synced");
    } catch (err) {
      console.error("Cloud push failed", err);
      setSyncStatus("error");
    }
  };
  if (immediate) return doPush();
  cloudPushTimer = setTimeout(doPush, 800);
}

// A debounced push scheduled just before the tab is backgrounded or closed
// (phone locked, app switched, browser tab closed) would otherwise get
// silently killed before its timer ever fires — the change stays on that
// device's localStorage but never reaches the cloud, so other devices never
// see it. Flushing immediately on these signals closes that gap.
document.addEventListener("visibilitychange", () => {
  if (document.hidden) pushStateToCloud(true);
});
window.addEventListener("pagehide", () => pushStateToCloud(true));

async function signUpAccount(email, password) {
  return supabaseClient.auth.signUp({ email, password });
}

async function logInAccount(email, password) {
  return supabaseClient.auth.signInWithPassword({ email, password });
}

async function logOutAccount() {
  await supabaseClient.auth.signOut();
  currentUser = null;
  setSyncStatus("offline");
}

function renderAccount() {
  const loggedOutEl = $("#account-logged-out");
  const loggedInEl = $("#account-logged-in");
  if (!loggedOutEl || !loggedInEl) return;
  if (currentUser) {
    loggedOutEl.style.display = "none";
    loggedInEl.style.display = "block";
    $("#account-email-display").textContent = currentUser.email;
    const statusText = {
      idle: "Not synced yet", syncing: "Syncing…", synced: "Synced",
      error: "Sync error — will retry on your next change", offline: "Offline"
    }[syncStatus] || "";
    $("#account-sync-status").textContent = statusText;
  } else {
    loggedOutEl.style.display = "block";
    loggedInEl.style.display = "none";
  }
}

// Which day's session is currently being viewed in the Training tab.
// Not persisted — always opens on today when the app is reloaded.
let trainingViewDay = todayDayIdx();

/* -------------------------------- Undo ----------------------------------- */
// In-memory only (cleared on reload) — a snapshot of `state` is pushed right
// before each user-initiated mutation, so Undo just pops the last one back.

let undoStack = [];
const UNDO_LIMIT = 30;

function pushUndo() {
  undoStack.push(JSON.stringify(state));
  if (undoStack.length > UNDO_LIMIT) undoStack.shift();
  updateUndoButton();
}

function undoLastAction() {
  if (undoStack.length === 0) return;
  state = JSON.parse(undoStack.pop());
  renderAll();
  updateUndoButton();
}

function updateUndoButton() {
  const btn = $("#undo-btn");
  if (btn) btn.disabled = undoStack.length === 0;
}

/* ---------------------------- Scoring logic ---------------------------- */

function sumTasks(list) {
  const earned = list.reduce((s, t) => s + (t.done ? t.points : 0), 0);
  const max = list.reduce((s, t) => s + (t.points > 0 ? t.points : 0), 0);
  return { earned, max };
}

// Convert "HH:MM" to a decimal hour, rolling times before 6am to the next day
// (so 23:30 and 00:45 both compare correctly against a ~22:00 bedtime).
function normalizedHour(hhmm) {
  if (!hhmm) return null;
  const [h, m] = hhmm.split(":").map(Number);
  let hour = h + m / 60;
  if (hour < 6) hour += 24;
  return hour;
}

function sleepScore(sleepHHMM) {
  const h = normalizedHour(sleepHHMM);
  if (h === null) return null;
  if (h <= 22) return 20;
  if (h <= 22.25) return 15;
  if (h <= 22.5) return 10;
  if (h <= 22.75) return 5;
  return 0;
}

function gapHours(sleepHHMM, otherHHMM) {
  const s = normalizedHour(sleepHHMM);
  const o = normalizedHour(otherHHMM);
  if (s === null || o === null) return null;
  return s - o;
}

function mealOrGymScore(sleepHHMM, otherHHMM) {
  const gap = gapHours(sleepHHMM, otherHHMM);
  if (gap === null) return null;
  if (gap >= 4) return 20;
  if (gap >= 3) return 15;
  if (gap >= 2) return 10;
  if (gap >= 1) return 5;
  return 0;
}

function coffeeScore(sleepHHMM, coffeeHHMM) {
  const gap = gapHours(sleepHHMM, coffeeHHMM);
  if (gap === null) return null;
  if (gap >= 9) return 10;
  if (gap >= 8) return 7;
  if (gap >= 7) return 4;
  if (gap >= 6) return 1;
  if (gap >= 5) return -2;
  if (gap >= 4) return -5;
  if (gap >= 3) return -7;
  return -10;
}

function computeTiming(t) {
  const sleep = sleepScore(t.sleep);
  const meal = mealOrGymScore(t.sleep, t.lastMeal);
  const gym = mealOrGymScore(t.sleep, t.gymFinish);
  const coffee = coffeeScore(t.sleep, t.lastCoffee);
  const noScroll = t.noScroll ? 20 : 0;
  const earned = (sleep || 0) + (meal || 0) + (gym || 0) + (coffee || 0) + noScroll;
  const max = 20 + 20 + 20 + 10 + 20; // sleep + meal + gym + coffee + no-scroll
  return { sleep, meal, gym, coffee, noScroll, earned, max };
}

function gymBonus(state) {
  const idx = todayDayIdx();
  const today = state.gymSchedule[idx];
  return { earned: today && today.done ? 10 : 0, max: 10, today };
}

function computeTotals(state) {
  const routine = sumTasks(state.routineTasks);
  const meals = sumTasks(state.meals);
  const timing = computeTiming(state.timing);
  const bonus = gymBonus(state);
  const side = sumTasks(state.sideMissions);
  const earned = routine.earned + meals.earned + timing.earned + bonus.earned + side.earned;
  const max = routine.max + meals.max + timing.max + bonus.max + side.max;
  return { routine, meals, timing, bonus, side, earned, max };
}

/* ------------------------------ Rollover -------------------------------- */

function monthKey(year, month) {
  return `${year}-${String(month).padStart(2, "0")}`;
}

function checkRollover() {
  const today = todayISO();
  if (state.currentDate === today) return;

  // Log the day that just ended.
  const totals = computeTotals(state);
  state.dailyLog[state.currentDate] = { earned: totals.earned, max: totals.max };

  const prev = new Date(state.currentDate + "T00:00:00");
  const now = new Date(today + "T00:00:00");
  const monthChanged = prev.getFullYear() !== now.getFullYear() || prev.getMonth() !== now.getMonth();

  if (monthChanged) {
    archiveMonth(prev.getFullYear(), prev.getMonth() + 1);
    state.monthly = { year: now.getFullYear(), month: now.getMonth() + 1, days: {} };
  }

  // Reset the daily checklist for the new day. Target-completion times reset
  // too — they're a plan for *that* day, not a permanent setting.
  state.routineTasks.forEach((t) => { t.done = false; t.time = ""; });
  state.meals.forEach((t) => { t.done = false; t.time = ""; });
  // Ticked Side Missions are done — they don't come back tomorrow. Anything
  // left unticked carries over (just with its time cleared, same as above).
  state.sideMissions = state.sideMissions.filter((t) => !t.done);
  state.sideMissions.forEach((t) => { t.time = ""; });
  state.gymSchedule.forEach((g) => (g.done = false));
  state.timing = { sleep: "", lastMeal: "", gymFinish: "", lastCoffee: "", noScroll: false };
  state.training = { done: {}, rehabDone: {}, pancakeDone: {} };

  // Archive any logged "actual" sets/reps/weight into per-session history —
  // grouped by session key (parsed off the front of each exercise id, e.g.
  // "legsA-3" -> "legsA") so it's there to check next time that session
  // comes around, regardless of which day it lands on. Only the most recent
  // logged session of each type is kept.
  const actualsBySession = {};
  Object.entries(state.trainingActuals).forEach(([id, value]) => {
    if (!value || !value.trim()) return;
    const sessionKey = id.slice(0, id.lastIndexOf("-"));
    if (!actualsBySession[sessionKey]) actualsBySession[sessionKey] = {};
    actualsBySession[sessionKey][id] = value.trim();
  });
  Object.entries(actualsBySession).forEach(([sessionKey, entries]) => {
    state.trainingLog[sessionKey] = { date: state.currentDate, entries };
    // Also append to the permanent per-day history that the Monthly
    // Training Report reads from. trainingLog only keeps the *latest*
    // session of each type; trainingHistory keeps every logged day.
    const session = TRAINING_SESSIONS[sessionKey];
    const lines = Object.entries(entries).map(([id, actual]) => {
      const ex = resolveExercise(sessionKey, id);
      return {
        id,
        name: ex ? ex.name : id,
        target: ex ? `${ex.sets} × ${ex.reps} · ${ex.weight}` : "",
        actual
      };
    });
    state.trainingHistory.push({
      date: state.currentDate,
      sessionKey,
      title: session ? session.title : sessionKey,
      lines
    });
  });
  // Guard against unbounded growth (roughly two years of daily sessions).
  if (state.trainingHistory.length > 800) {
    state.trainingHistory = state.trainingHistory.slice(-800);
  }
  state.trainingActuals = {};

  // Apply whatever order/times were set on the "Plan for Tomorrow" tab, now
  // that tomorrow has arrived, then clear the plan for the next cycle.
  state.routineTasks = tomorrowPlanDisplayOrder();
  Object.keys(state.tomorrowPlan.times).forEach((id) => {
    const task = state.routineTasks.find((t) => t.id === id);
    if (task) task.time = state.tomorrowPlan.times[id];
  });
  state.tomorrowPlan = { order: [], times: {} };

  state.currentDate = today;
  trainingViewDay = todayDayIdx();
  saveState();
}

// A day's resting-heart-rate reading is a plain number, not an achievement
// toggle, so it never contributes to totalPts — this just averages whatever
// got logged. Shared by the live Monthly view and the archive snapshot so
// the two can't disagree on what counts as "logged".
function averageRHR(days) {
  let sum = 0, count = 0;
  Object.values(days).forEach((d) => {
    const v = Number(d.rhr);
    if (d.rhr !== "" && d.rhr != null && !Number.isNaN(v)) { sum += v; count++; }
  });
  return count ? { avg: Math.round(sum / count), count } : null;
}

function archiveMonth(year, month) {
  const days = state.monthly.days || {};
  let gym = 0, sauna = 0, redlight = 0, hbot = 0, totalPts = 0, bestDay = 0;
  Object.values(days).forEach((d) => {
    if (d.gym) gym++;
    if (d.sauna) sauna++;
    if (d.redlight) redlight++;
    if (d.hbot) hbot++;
    const pts = (d.gym ? 10 : 0) + (d.sauna ? 5 : 0) + (d.redlight ? 5 : 0) + (d.hbot ? 8 : 0);
    totalPts += pts;
    if (pts > bestDay) bestDay = pts;
  });
  const label = `${MONTH_NAMES[month - 1]} ${year}`;
  // avoid duplicate archive entries
  if (state.archive.some((a) => a.label === label)) return;
  if (Object.keys(days).length === 0) return; // nothing tracked, skip
  const rhr = averageRHR(days);
  state.archive.unshift({ label, gym, sauna, redlight, hbot, totalPts, bestDay, avgRHR: rhr ? rhr.avg : null });
}

/* -------------------------------- Render -------------------------------- */

const $ = (sel, root) => (root || document).querySelector(sel);
const $$ = (sel, root) => Array.from((root || document).querySelectorAll(sel));

function esc(str) {
  const div = document.createElement("div");
  div.textContent = str;
  return div.innerHTML;
}

function renderAll() {
  checkRollover();
  ensureEliminateToday();
  renderHeader();
  renderTaskSection("routine-list", state.routineTasks, "routine", true);
  applyTimeShiftHighlight();
  renderTaskSection("meals-list", state.meals, "meals", false);
  renderTaskSection("side-list", state.sideMissions, "side", true);
  renderEliminateToday();
  renderTiming();
  renderGymSchedule();
  renderCalendar();
  renderMonthly();
  renderArchive();
  renderTraining();
  renderSaveScore();
  renderRightNow();
  renderTomorrowPlan();
  renderProjects();
  updateUndoButton();
  saveState();
}

function renderHeader() {
  const totals = computeTotals(state);
  const d = new Date();
  const dayName = d.toLocaleDateString(undefined, { weekday: "long" });
  const dateStr = d.toLocaleDateString(undefined, { day: "2-digit", month: "short", year: "numeric" });
  $("#today-label").textContent = `${dayName} ${dateStr}`;
  $("#score-label-text").textContent = `${totals.earned} / ${totals.max} pts`;
  const pct = totals.max > 0 ? Math.max(0, Math.min(100, (totals.earned / totals.max) * 100)) : 0;
  $("#score-bar-fill").style.width = pct + "%";
}

function renderTaskSection(listId, tasks, kind, reorderable) {
  const el = $("#" + listId);
  el.innerHTML = "";
  tasks.forEach((t) => {
    const row = document.createElement("div");
    row.className = "task-row" + (t.done ? " done" : "");
    row.dataset.id = t.id;
    row.dataset.kind = kind;
    row.innerHTML = `
      ${reorderable ? `<span class="drag-handle" aria-label="drag to reorder"><svg viewBox="0 0 16 16" width="14" height="14"><circle cx="4" cy="3" r="1.4"/><circle cx="4" cy="8" r="1.4"/><circle cx="4" cy="13" r="1.4"/><circle cx="10" cy="3" r="1.4"/><circle cx="10" cy="8" r="1.4"/><circle cx="10" cy="13" r="1.4"/></svg></span>` : ""}
      <button class="check-btn" data-kind="${kind}" data-id="${t.id}" aria-label="toggle">${t.done ? "✓" : "✗"}</button>
      ${t.icon ? iconTag(t.icon) : ""}
      <span class="task-label" data-kind="${kind}" data-id="${t.id}">${esc(t.label)}</span>
      <input type="time" class="task-time" data-kind="${kind}" data-id="${t.id}" value="${t.time || ""}" title="Desired completion time" />
      ${kind === "routine" ? `<span class="time-bump-group"><button class="time-bump-btn" data-bump-minutes="5">+5</button><button class="time-bump-btn" data-bump-minutes="20">+20</button></span>` : ""}
      <span class="task-points ${t.points < 0 ? "neg" : ""}" data-kind="${kind}" data-id="${t.id}">${t.points > 0 ? "+" : ""}${t.points}</span>
      <button class="icon-btn edit-btn" data-kind="${kind}" data-id="${t.id}" aria-label="edit">✎</button>
      <button class="icon-btn del-btn" data-kind="${kind}" data-id="${t.id}" aria-label="delete">${iconTag("trash")}</button>
    `;
    el.appendChild(row);
  });
  if (reorderable) attachDragHandlers(el, getList(kind), swapListItems);
}

/* ---------------------------- Drag to reorder ---------------------------- */
// Pointer Events (not the HTML5 drag-and-drop API) so this works with touch
// on iOS Safari, not just mouse on desktop.

function swapListItems(list, idA, idB) {
  const iA = list.findIndex((x) => x.id === idA);
  const iB = list.findIndex((x) => x.id === idB);
  if (iA === -1 || iB === -1) return;
  [list[iA], list[iB]] = [list[iB], list[iA]];
}

// `list` + `swapFn` are generic so this same drag machinery works for both
// arrays of task objects (swapListItems) and plain id arrays like
// tomorrowPlan.order (swapIdArray).
function attachDragHandlers(containerEl, list, swapFn) {
  const rows = Array.from(containerEl.querySelectorAll(".task-row"));
  rows.forEach((row) => {
    const handle = row.querySelector(".drag-handle");
    if (!handle) return;

    handle.addEventListener("pointerdown", (e) => {
      e.preventDefault();
      pushUndo();
      const dragEl = row;
      let startY = e.clientY;

      dragEl.classList.add("dragging");
      try { dragEl.setPointerCapture(e.pointerId); } catch (err) {}

      const onMove = (ev) => {
        // Self-heal: if this row got detached from the page (e.g. a render
        // happened for some other reason mid-drag), stop listening instead
        // of operating on a stale node — this is what was causing the list
        // to occasionally snap back to its old order.
        if (!dragEl.isConnected) {
          onUp(ev);
          return;
        }
        const deltaY = ev.clientY - startY;
        dragEl.style.transform = `translateY(${deltaY}px)`;

        const dragRect = dragEl.getBoundingClientRect();
        const dragCenter = dragRect.top + dragRect.height / 2;

        // Only ever compare against the row immediately adjacent to dragEl in
        // the direction of travel — not every row in the list. Checking the
        // whole list (in top-to-bottom order) meant that once dragEl had
        // passed a row, that row was still first in iteration order and
        // trivially satisfied "dragCenter is past it" on every subsequent
        // tick, re-triggering a swap *with the row already passed*. The DOM
        // move was a no-op (they were already adjacent) but the array swap
        // still ran, silently flipping the order back and forth on every
        // pointermove — so wherever the drag happened to end, the array
        // could easily land back at its original order.
        if (deltaY > 0) {
          const next = dragEl.nextElementSibling;
          if (next) {
            const nextRect = next.getBoundingClientRect();
            const nextCenter = nextRect.top + nextRect.height / 2;
            if (dragCenter > nextCenter) {
              containerEl.insertBefore(dragEl, next.nextSibling);
              swapFn(list, dragEl.dataset.id, next.dataset.id);
              startY = ev.clientY;
              dragEl.style.transform = "translateY(0px)";
            }
          }
        } else if (deltaY < 0) {
          const prev = dragEl.previousElementSibling;
          if (prev) {
            const prevRect = prev.getBoundingClientRect();
            const prevCenter = prevRect.top + prevRect.height / 2;
            if (dragCenter < prevCenter) {
              containerEl.insertBefore(dragEl, prev);
              swapFn(list, dragEl.dataset.id, prev.dataset.id);
              startY = ev.clientY;
              dragEl.style.transform = "translateY(0px)";
            }
          }
        }
      };

      const onUp = (ev) => {
        try { dragEl.releasePointerCapture(e.pointerId); } catch (err) {}
        dragEl.classList.remove("dragging");
        dragEl.style.transform = "";
        document.removeEventListener("pointermove", onMove);
        document.removeEventListener("pointerup", onUp);
        document.removeEventListener("pointercancel", onUp);
        saveState();
      };

      document.addEventListener("pointermove", onMove);
      document.addEventListener("pointerup", onUp);
      document.addEventListener("pointercancel", onUp);
    });
  });
}

function renderTiming() {
  const t = state.timing;
  $("#sleep-time").value = t.sleep;
  $("#last-meal-time").value = t.lastMeal;
  $("#gym-finish-time").value = t.gymFinish;
  $("#last-coffee-time").value = t.lastCoffee;
  $("#no-scroll-check").checked = t.noScroll;
  const c = computeTiming(t);
  $("#sleep-score").textContent = c.sleep === null ? "–" : c.sleep;
  $("#meal-score").textContent = c.meal === null ? "–" : c.meal;
  $("#gym-score").textContent = c.gym === null ? "–" : c.gym;
  $("#coffee-score").textContent = c.coffee === null ? "–" : c.coffee;
}

const SESSION_SELECT_OPTIONS = [
  { key: "", label: "Rest" },
  { key: "pushA", label: "Push A" },
  { key: "pullA", label: "Pull A" },
  { key: "legsA", label: "Legs A" },
  { key: "pushB", label: "Push B" },
  { key: "legsB", label: "Legs B" }
];

function renderGymSchedule() {
  const todayIdx = todayDayIdx();
  const el = $("#gym-schedule-list");
  el.innerHTML = "";
  state.gymSchedule.forEach((g, i) => {
    const currentKey = state.weeklySessionKeys[i] || "";
    const row = document.createElement("div");
    row.className = "gym-row" + (i === todayIdx ? " today" : "");
    row.innerHTML = `
      <span class="gym-day">${g.day}${i === todayIdx ? " ← today" : ""}</span>
      <select class="gym-type-select" data-day-idx="${i}" aria-label="${g.day} session">
        ${SESSION_SELECT_OPTIONS.map((o) => `<option value="${o.key}" ${o.key === currentKey ? "selected" : ""}>${o.label}</option>`).join("")}
      </select>
      <button class="check-btn small" data-idx="${i}" id="gym-done-${i}">${g.done ? "✓" : "✗"}</button>
    `;
    el.appendChild(row);
  });
  const bonus = gymBonus(state);
  $("#gym-bonus-label").textContent = `Schedule bonus: ${bonus.earned} / 10 pts · full session on the Training tab`;
}

function renderCalendar() {
  const el = $("#calendar-grid");
  el.innerHTML = "";
  const now = new Date();
  const year = now.getFullYear(), month = now.getMonth();
  $("#calendar-month-label").textContent = `${MONTH_NAMES[month]} ${year}`;
  const firstDow = (new Date(year, month, 1).getDay() + 6) % 7; // Mon=0
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  ["Mon","Tue","Wed","Thu","Fri","Sat","Sun"].forEach((d) => {
    const h = document.createElement("div");
    h.className = "cal-head";
    h.textContent = d;
    el.appendChild(h);
  });
  for (let i = 0; i < firstDow; i++) {
    el.appendChild(document.createElement("div"));
  }
  const totals = computeTotals(state);
  for (let day = 1; day <= daysInMonth; day++) {
    const iso = todayISO(new Date(year, month, day));
    const cell = document.createElement("div");
    cell.className = "cal-cell";
    cell.dataset.date = iso;
    let entry = state.dailyLog[iso];
    if (iso === todayISO()) entry = { earned: totals.earned, max: totals.max };
    let ptsHtml = "";
    const titleParts = [];
    if (entry && entry.max > 0) {
      const pct = Math.max(0, Math.min(1, entry.earned / entry.max));
      cell.style.background = heatColor(pct);
      ptsHtml = `<span class="cal-cell-pts">${entry.earned}</span>`;
      titleParts.push(`${entry.earned} / ${entry.max} pts`);
    }
    const plan = state.dayPlans[iso];
    const planMark = plan ? `<span class="cal-cell-plan">${ICON_SVG.note}</span>` : "";
    if (plan) titleParts.push(`Plan: ${plan}`);
    cell.title = titleParts.join(" — ");
    if (iso === todayISO()) cell.classList.add("today");
    cell.innerHTML = `<span class="cal-cell-day">${day}</span>${ptsHtml}${planMark}`;
    el.appendChild(cell);
  }
}

function heatColor(pct) {
  // 0 -> muted red, 1 -> green, matching the app's dark theme
  const r = Math.round(180 - 100 * pct);
  const g = Math.round(60 + 130 * pct);
  const b = 60;
  return `rgba(${r},${g},${b},0.55)`;
}

function renderMonthly() {
  const m = state.monthly;
  $("#monthly-label").textContent = `${MONTH_NAMES[m.month - 1]} ${m.year}`;
  const el = $("#monthly-list");
  el.innerHTML = "";
  const daysInMonth = new Date(m.year, m.month, 0).getDate();
  let totalPts = 0;
  const header = document.createElement("div");
  header.className = "monthly-row monthly-header";
  header.innerHTML = `<span>DATE</span><span>${iconTag("dumbbell")} GYM</span><span>${iconTag("sauna")} SAUNA</span><span>${iconTag("sun")} RED LIGHT</span><span>${iconTag("wind")} HBOT</span><span>${iconTag("heart")} RHR</span><span>PTS</span>`;
  el.appendChild(header);
  for (let day = 1; day <= daysInMonth; day++) {
    const key = String(day);
    const d = m.days[key] || { gym: false, sauna: false, redlight: false, hbot: false, rhr: "" };
    const pts = (d.gym ? 10 : 0) + (d.sauna ? 5 : 0) + (d.redlight ? 5 : 0) + (d.hbot ? 8 : 0);
    totalPts += pts;
    const dateObj = new Date(m.year, m.month - 1, day);
    const label = dateObj.toLocaleDateString(undefined, { weekday: "short", day: "2-digit", month: "short" });
    const row = document.createElement("div");
    row.className = "monthly-row";
    row.innerHTML = `
      <span>${label}</span>
      <button class="check-btn small" data-field="gym" data-day="${day}">${d.gym ? "✓" : "✗"}</button>
      <button class="check-btn small" data-field="sauna" data-day="${day}">${d.sauna ? "✓" : "✗"}</button>
      <button class="check-btn small" data-field="redlight" data-day="${day}">${d.redlight ? "✓" : "✗"}</button>
      <button class="check-btn small" data-field="hbot" data-day="${day}">${d.hbot ? "✓" : "✗"}</button>
      <input type="number" inputmode="numeric" class="rhr-input" data-day="${day}" value="${esc(d.rhr || "")}" placeholder="–" min="30" max="220" />
      <span>${pts}</span>
    `;
    el.appendChild(row);
  }
  const rhr = averageRHR(m.days);
  $("#monthly-total").textContent = `Total: ${totalPts} pts${rhr ? ` · Avg RHR: ${rhr.avg} bpm (${rhr.count} day${rhr.count === 1 ? "" : "s"})` : ""}`;
}

function renderArchive() {
  const el = $("#archive-list");
  el.innerHTML = "";
  if (state.archive.length === 0) {
    el.innerHTML = `<p class="muted">No archived months yet — the current month archives automatically once a new month begins.</p>`;
    return;
  }
  state.archive.forEach((a) => {
    const row = document.createElement("div");
    row.className = "archive-row";
    row.innerHTML = `
      <strong>${esc(a.label)}</strong>
      <span>${iconTag("dumbbell")} ${a.gym} &nbsp; ${iconTag("sauna")} ${a.sauna} &nbsp; ${iconTag("sun")} ${a.redlight} &nbsp; ${iconTag("wind")} ${a.hbot}${a.avgRHR != null ? ` &nbsp; ${iconTag("heart")} ${a.avgRHR} bpm avg` : ""}</span>
      <span>Total: ${a.totalPts} pts &middot; Best day: ${a.bestDay} pts</span>
    `;
    el.appendChild(row);
  });
}

/* --------------------------- Save / Backdate Score ------------------------ */
// Keeps the Save/Backdate Score inputs in sync: live totals while the
// selected date is today (read-only, since that's always the current
// running score), editable and untouched otherwise so a backdate-in-
// progress doesn't get clobbered by unrelated renders.
function applySaveScoreToday(dateInput, earnedInput, maxInput) {
  const isToday = dateInput.value === todayISO();
  earnedInput.readOnly = isToday;
  maxInput.readOnly = isToday;
  if (isToday) {
    const totals = computeTotals(state);
    earnedInput.value = totals.earned;
    maxInput.value = totals.max;
  }
  return isToday;
}

function renderSaveScore() {
  const dateInput = $("#save-score-date");
  if (!dateInput) return;
  if (!dateInput.value) dateInput.value = todayISO();
  applySaveScoreToday(dateInput, $("#save-score-earned"), $("#save-score-max"));
}

/* -------------------------------- Right Now -------------------------------- */
// A separate, unscored scratch checklist — for whatever's on your plate right
// now, independent of the scored Daily Routine Tasks.

function renderRightNow() {
  const el = $("#right-now-list");
  if (!el) return;
  el.innerHTML = "";
  state.rightNow.forEach((item) => {
    const row = document.createElement("div");
    row.className = "task-row" + (item.done ? " done" : "");
    row.innerHTML = `
      <button class="check-btn" data-right-now-id="${item.id}" aria-label="toggle">${item.done ? "✓" : "✗"}</button>
      <span class="task-label">${esc(item.label)}</span>
      <button class="icon-btn del-btn" data-right-now-del="${item.id}" aria-label="delete">${iconTag("trash")}</button>
    `;
    el.appendChild(row);
  });
}

/* ----------------------------- Plan for Tomorrow --------------------------- */
// A live mirror of the Daily Routine Tasks — same tasks, but here you can
// only reorder them and set a planned time. Nothing here can be checked off
// or deleted; do that from the real list. Applied to routineTasks the moment
// the date actually rolls over (checkRollover), then cleared for next time.

// The order rows should display in: tomorrowPlan.order first (for ids it
// knows about), then any routineTasks not yet covered by the plan (new since
// last visit), in their existing relative order.
function tomorrowPlanDisplayOrder() {
  const byId = new Map(state.routineTasks.map((t) => [t.id, t]));
  const ordered = [];
  const order = (state.tomorrowPlan && Array.isArray(state.tomorrowPlan.order)) ? state.tomorrowPlan.order : [];
  order.forEach((id) => {
    if (byId.has(id)) {
      ordered.push(byId.get(id));
      byId.delete(id);
    }
  });
  state.routineTasks.forEach((t) => {
    if (byId.has(t.id)) ordered.push(t);
  });
  return ordered;
}

function swapIdArray(arr, idA, idB) {
  const iA = arr.indexOf(idA);
  const iB = arr.indexOf(idB);
  if (iA === -1 || iB === -1) return;
  [arr[iA], arr[iB]] = [arr[iB], arr[iA]];
}

function renderTomorrowPlan() {
  const el = $("#tomorrow-plan-list");
  if (!el) return;
  // Keep tomorrowPlan.order in sync with the live task list — idempotent
  // when nothing's changed, and self-heals if tasks were added/removed.
  const items = tomorrowPlanDisplayOrder();
  state.tomorrowPlan.order = items.map((t) => t.id);
  el.innerHTML = "";
  if (items.length === 0) {
    el.innerHTML = `<p class="muted">Add some Daily Routine Tasks and they'll show up here to plan.</p>`;
    return;
  }
  items.forEach((t) => {
    const row = document.createElement("div");
    row.className = "task-row";
    row.dataset.id = t.id;
    row.innerHTML = `
      <span class="drag-handle" aria-label="drag to reorder"><svg viewBox="0 0 16 16" width="14" height="14"><circle cx="4" cy="3" r="1.4"/><circle cx="4" cy="8" r="1.4"/><circle cx="4" cy="13" r="1.4"/><circle cx="10" cy="3" r="1.4"/><circle cx="10" cy="8" r="1.4"/><circle cx="10" cy="13" r="1.4"/></svg></span>
      ${t.icon ? iconTag(t.icon) : ""}
      <span class="task-label">${esc(t.label)}</span>
      <input type="time" class="plan-time" data-id="${t.id}" value="${state.tomorrowPlan.times[t.id] || ""}" title="Planned time for tomorrow" />
      <span class="time-bump-group">
        <button class="time-bump-btn" data-bump-minutes="5">+5</button>
        <button class="time-bump-btn" data-bump-minutes="20">+20</button>
      </span>
      <span class="task-points ${t.points < 0 ? "neg" : ""}">${t.points > 0 ? "+" : ""}${t.points}</span>
    `;
    el.appendChild(row);
  });
  attachDragHandlers(el, state.tomorrowPlan.order, swapIdArray);
}

/* ----------------------------- Eliminate Today ----------------------------- */
// 5 Side Missions picked at random each day (ensureEliminateToday, called
// from renderAll) — a focus subset of the full backlog. Checking one off
// here toggles the same underlying mission, so it counts the same as
// completing it from the full Side Missions list.

function renderEliminateToday() {
  const el = $("#eliminate-today-list");
  if (!el) return;
  el.innerHTML = "";
  const items = state.eliminateToday.ids
    .map((id) => state.sideMissions.find((m) => m.id === id))
    .filter(Boolean);
  if (items.length === 0) {
    el.innerHTML = `<p class="muted">Add some Side Missions and today's picks will show up here.</p>`;
    return;
  }
  items.forEach((t) => {
    const row = document.createElement("div");
    row.className = "task-row" + (t.done ? " done" : "");
    row.innerHTML = `
      <button class="check-btn" data-kind="side" data-id="${t.id}" aria-label="toggle">${t.done ? "✓" : "✗"}</button>
      ${t.icon ? iconTag(t.icon) : ""}
      <span class="task-label">${esc(t.label)}</span>
      <span class="task-points ${t.points < 0 ? "neg" : ""}">${t.points > 0 ? "+" : ""}${t.points}</span>
    `;
    el.appendChild(row);
  });
}

/* -------------------------------- Projects --------------------------------- */
// Each project is a freeform notes page. The note textarea saves via its own
// debounced "input" listener (see event wiring) rather than going through
// renderAll()/renderProjects() on every keystroke — rebuilding the list DOM
// mid-typing would blow away focus and cursor position.

function renderProjects() {
  const el = $("#projects-list");
  if (!el) return;
  el.innerHTML = "";
  if (state.projects.length === 0) {
    el.innerHTML = `<div class="card"><p class="muted">No projects yet — tap "+ Add New" above to start one.</p></div>`;
    return;
  }
  state.projects.forEach((p) => {
    const card = document.createElement("div");
    card.className = "card";
    card.innerHTML = `
      <div class="project-header">
        <h2 class="project-title">${esc(p.title)}</h2>
        <div class="project-actions">
          <button class="icon-btn" data-project-rename="${p.id}" aria-label="rename">&#9998;</button>
          <button class="icon-btn del-btn" data-project-del="${p.id}" aria-label="delete">${iconTag("trash")}</button>
        </div>
      </div>
      <textarea class="project-note" data-project-id="${p.id}" rows="12" placeholder="Start typing…">${esc(p.content)}</textarea>
    `;
    el.appendChild(card);
  });
}

/* ------------------------------- News Board --------------------------------- */
// Live external data (London weather + a few transport lines) — deliberately
// kept out of `state` since it's not the user's data and shouldn't sync or
// persist; just refetched on boot, on a timer, and on manual refresh.

const WEATHER_CODES = {
  0: { label: "Clear sky", icon: "sun" },
  1: { label: "Mainly clear", icon: "sun" },
  2: { label: "Partly cloudy", icon: "cloud" },
  3: { label: "Overcast", icon: "cloud" },
  45: { label: "Fog", icon: "cloud" },
  48: { label: "Depositing rime fog", icon: "cloud" },
  51: { label: "Light drizzle", icon: "rain" },
  53: { label: "Drizzle", icon: "rain" },
  55: { label: "Dense drizzle", icon: "rain" },
  56: { label: "Freezing drizzle", icon: "rain" },
  57: { label: "Dense freezing drizzle", icon: "rain" },
  61: { label: "Slight rain", icon: "rain" },
  63: { label: "Rain", icon: "rain" },
  65: { label: "Heavy rain", icon: "rain" },
  66: { label: "Freezing rain", icon: "rain" },
  67: { label: "Heavy freezing rain", icon: "rain" },
  71: { label: "Slight snow", icon: "snow" },
  73: { label: "Snow", icon: "snow" },
  75: { label: "Heavy snow", icon: "snow" },
  77: { label: "Snow grains", icon: "snow" },
  80: { label: "Slight rain showers", icon: "rain" },
  81: { label: "Rain showers", icon: "rain" },
  82: { label: "Violent rain showers", icon: "rain" },
  85: { label: "Slight snow showers", icon: "snow" },
  86: { label: "Heavy snow showers", icon: "snow" },
  95: { label: "Thunderstorm", icon: "storm" },
  96: { label: "Thunderstorm, slight hail", icon: "storm" },
  99: { label: "Thunderstorm, heavy hail", icon: "storm" }
};
function weatherInfo(code) {
  return WEATHER_CODES[code] || { label: "Unknown", icon: "cloud" };
}

let newsWeather = null;
let newsLines = null;
let newsError = null;
let newsLastFetched = null;
let newsLoading = false;

async function fetchNews() {
  newsLoading = true;
  newsError = null;
  renderNews();
  try {
    const [weatherRes, linesRes] = await Promise.all([
      fetch("https://api.open-meteo.com/v1/forecast?latitude=51.5074&longitude=-0.1278&daily=weathercode,temperature_2m_max,temperature_2m_min&timezone=Europe%2FLondon&forecast_days=2"),
      fetch("https://api.tfl.gov.uk/Line/elizabeth,metropolitan,dlr,c2c/Status")
    ]);
    if (!weatherRes.ok || !linesRes.ok) throw new Error("Request failed");
    const weatherData = await weatherRes.json();
    newsLines = await linesRes.json();
    newsWeather = weatherData.daily;
    newsLastFetched = new Date();
  } catch (err) {
    console.error("News fetch failed", err);
    newsError = "Couldn't load news right now — check your connection and try again.";
  }
  newsLoading = false;
  renderNews();
}

function renderNews() {
  const weatherEl = $("#news-weather");
  const linesEl = $("#news-lines");
  const updatedEl = $("#news-updated");
  if (!weatherEl || !linesEl) return;

  if (newsError) {
    weatherEl.innerHTML = `<p class="muted">${esc(newsError)}</p>`;
    linesEl.innerHTML = "";
  } else if (!newsWeather) {
    weatherEl.innerHTML = `<p class="muted">${newsLoading ? "Loading…" : ""}</p>`;
  } else {
    const dayLabels = ["Today", "Tomorrow"];
    weatherEl.innerHTML = newsWeather.time.map((date, i) => {
      const info = weatherInfo(newsWeather.weathercode[i]);
      const max = Math.round(newsWeather.temperature_2m_max[i]);
      const min = Math.round(newsWeather.temperature_2m_min[i]);
      return `
        <div class="weather-row">
          ${iconTag(info.icon)}
          <span class="weather-day">${dayLabels[i] || date}</span>
          <span class="weather-desc">${esc(info.label)}</span>
          <span class="weather-temp">${max}&deg; / ${min}&deg;</span>
        </div>
      `;
    }).join("");
  }

  if (newsLines && !newsError) {
    linesEl.innerHTML = newsLines.map((line) => {
      const status = line.lineStatuses && line.lineStatuses[0];
      const desc = status ? status.statusSeverityDescription : "Unknown";
      const good = desc === "Good Service";
      return `
        <div class="line-row">
          <span class="line-name">${esc(line.name)}</span>
          <span class="line-status ${good ? "good" : "disrupted"}">${esc(desc)}</span>
        </div>
      `;
    }).join("");
  } else if (!newsError && newsLoading) {
    linesEl.innerHTML = `<p class="muted">Loading…</p>`;
  }

  if (updatedEl) {
    updatedEl.textContent = newsLastFetched
      ? `Updated ${newsLastFetched.toLocaleTimeString(undefined, { hour: "2-digit", minute: "2-digit" })}`
      : "";
  }
}

/* ------------------------------- Training -------------------------------- */

// actualId/lastActual are only passed for the main weekly-session exercises
// (not Rehab/Pancake) — that's the "refer back next leg day" use case.
function exerciseRowHtml(ex, checked, dataAttrs, actualId, lastActual, removeId, sectionLabel) {
  const section = sectionLabel === undefined ? ex.section : sectionLabel;
  return `
    <div class="ex-row">
      ${section ? `<div class="ex-section">${esc(section)}</div>` : ""}
      <div class="ex-line">
        <button class="check-btn small" ${dataAttrs}>${checked ? "✓" : "✗"}</button>
        <div class="ex-info">
          <div class="ex-name-row">
            <div class="ex-name">${esc(ex.name)}${removeId ? ` <span class="ex-custom-tag">added</span>` : ""}</div>
            ${actualId ? `<button class="ex-history-btn" data-history-id="${actualId}" aria-label="View history" title="View logged history">${iconTag("clock")}</button>` : ""}
          </div>
          <div class="ex-meta">${esc(String(ex.sets))} × ${esc(String(ex.reps))} &middot; ${esc(ex.weight)}</div>
          ${ex.notes ? `<div class="ex-notes">${esc(ex.notes)}</div>` : ""}
          ${actualId ? `
            <div class="ex-actual-row">
              <input type="text" class="ex-actual" data-actual-id="${actualId}" value="${esc(state.trainingActuals[actualId] || "")}" placeholder="Log actual sets/reps/weight…" />
              <button class="ex-log-btn" data-actual-id="${actualId}">${iconTag("checkCircle")} Log</button>
            </div>
            ${lastActual ? `<div class="ex-last">Last time: ${esc(lastActual)}</div>` : ""}
          ` : ""}
        </div>
        ${removeId ? `<button class="ex-remove-btn" data-remove-custom-id="${esc(removeId)}" aria-label="Remove exercise" title="Remove">${iconTag("trash")}</button>` : ""}
      </div>
    </div>
  `;
}

function openExerciseHistory(id) {
  const sessionKey = id.slice(0, id.lastIndexOf("-"));
  const ex = resolveExercise(sessionKey, id);
  const history = exerciseHistoryFor(id);
  $("#history-modal-title").textContent = ex ? ex.name : "Exercise History";
  const body = $("#history-modal-body");
  if (history.length === 0) {
    body.innerHTML = `<p class="hint">No logged history yet — hit Log on a set to start tracking this one over time.</p>`;
  } else {
    body.innerHTML = history.map((h) => `
      <div class="history-row">
        <span class="history-date">${esc(formatDateShort(h.date))}${h.inProgress ? " · today" : ""}</span>
        <span class="history-actual">${esc(h.actual)}</span>
      </div>
    `).join("");
  }
  $("#history-modal").hidden = false;
}

function renderTraining() {
  const pillsEl = $("#training-day-pills");
  if (!pillsEl) return; // tab not in the DOM yet on first paint edge case

  const todayIdx = todayDayIdx();
  pillsEl.innerHTML = "";
  GYM_DAYS.forEach((day, i) => {
    const session = sessionForDay(i);
    const btn = document.createElement("button");
    btn.className = "day-pill" + (i === trainingViewDay ? " active" : "") + (i === todayIdx ? " today" : "");
    btn.dataset.day = i;
    btn.innerHTML = `<span>${day}</span><small>${session ? session.shortTitle : "Rest"}</small>`;
    pillsEl.appendChild(btn);
  });

  const session = sessionForDay(trainingViewDay);
  const sessionEl = $("#training-session");
  const addCard = $("#training-add-card");
  if (!session) {
    sessionEl.innerHTML = `
      <h2>${iconTag("moon")} Rest Day</h2>
      <p class="hint">No lifting session scheduled${trainingViewDay === todayIdx ? " today" : ""}. Good day for the Pancake Program below, or just recover.</p>
    `;
    if (addCard) addCard.hidden = true;
  } else {
    const isToday = trainingViewDay === todayIdx;
    const entries = sessionExerciseEntries(session);
    const doneCount = entries.filter((e) => state.training.done[e.id]).length;
    const lastEntry = state.trainingLog[session.key];
    sessionEl.innerHTML = `
      <h2>${isToday ? iconTag("pin") + " Today: " : ""}${esc(session.title)}</h2>
      <p class="hint">${esc(session.focus)}</p>
      <p class="hint">${doneCount} / ${entries.length} logged${lastEntry ? ` &middot; Last logged ${esc(formatDateShort(lastEntry.date))}` : ""}</p>
      <div id="training-exercise-list"></div>
      <button class="add-task-btn" id="training-complete-all" style="margin-top:10px;">${iconTag("checkCircle")} Mark all complete</button>
    `;
    const listEl = $("#training-exercise-list", sessionEl);
    entries.forEach((e) => {
      const checked = !!state.training.done[e.id];
      const lastActual = lastEntry && lastEntry.entries[e.id];
      listEl.insertAdjacentHTML("beforeend", exerciseRowHtml(e.ex, checked, `data-training-id="${e.id}"`, e.id, lastActual, e.custom ? e.id : null, e.showSection));
    });
    if (addCard) {
      addCard.hidden = false;
      const hintEl = $("#training-add-hint", addCard);
      if (hintEl) hintEl.textContent = `Adds to ${session.shortTitle} permanently — it comes back every ${session.shortTitle} day.`;
      const secSel = $("#cx-section");
      if (secSel) {
        secSel.innerHTML = `<option value="">— No section —</option>` +
          sessionSections(session).map((s) => `<option value="${esc(s)}">${esc(s)}</option>`).join("");
      }
    }
  }

  // Rehab & Maintenance (always available, daily).
  const rehabEl = $("#rehab-list");
  rehabEl.innerHTML = "";
  REHAB_EXERCISES.forEach((ex, i) => {
    const checked = !!state.training.rehabDone[i];
    rehabEl.insertAdjacentHTML("beforeend", exerciseRowHtml(ex, checked, `data-rehab-idx="${i}"`));
  });

  // Pancake Program (always available; flagged as suggested on certain days).
  const pancakeEl = $("#pancake-list");
  const suggested = PANCAKE_SUGGESTED_DAYS.includes(trainingViewDay);
  $("#pancake-suggested-badge").style.display = suggested ? "inline-block" : "none";
  pancakeEl.innerHTML = "";
  PANCAKE_EXERCISES.forEach((ex, i) => {
    const checked = !!state.training.pancakeDone[i];
    pancakeEl.insertAdjacentHTML("beforeend", exerciseRowHtml(ex, checked, `data-pancake-idx="${i}"`));
  });
}

// Plain-text summary of every logged training session in the current
// calendar month, built from state.trainingHistory (archived on each day
// rollover) plus whatever's logged-but-not-yet-archived for today. Meant
// to be copied straight into a message to Claude.
function buildTrainingReport() {
  const now = new Date();
  const y = now.getFullYear();
  const m = now.getMonth() + 1;
  const monthPrefix = `${y}-${String(m).padStart(2, "0")}`;

  const sessions = (state.trainingHistory || [])
    .filter((s) => typeof s.date === "string" && s.date.startsWith(monthPrefix))
    .map((s) => ({ date: s.date, title: s.title, lines: s.lines || [] }));

  // Today's session — logged but still sitting in trainingActuals until the
  // next rollover. Group by session key the same way checkRollover() does.
  const todayBySession = {};
  Object.entries(state.trainingActuals).forEach(([id, value]) => {
    if (!value || !value.trim()) return;
    const key = id.slice(0, id.lastIndexOf("-"));
    if (!todayBySession[key]) todayBySession[key] = [];
    const ex = resolveExercise(key, id);
    todayBySession[key].push({
      name: ex ? ex.name : id,
      target: ex ? `${ex.sets} × ${ex.reps} · ${ex.weight}` : "",
      actual: value.trim()
    });
  });
  Object.entries(todayBySession).forEach(([key, lines]) => {
    const session = TRAINING_SESSIONS[key];
    sessions.push({ date: state.currentDate, title: (session ? session.title : key) + " (in progress)", lines });
  });

  sessions.sort((a, b) => a.date.localeCompare(b.date));

  const out = [];
  out.push(`TRAINING LOG — ${MONTH_NAMES[m - 1]} ${y}`);
  out.push("");

  if (sessions.length === 0) {
    out.push("No sessions logged yet this month.");
    return out.join("\n");
  }

  sessions.forEach((s) => {
    out.push(`=== ${formatDateShort(s.date)} · ${s.title} ===`);
    if (s.lines.length === 0) {
      out.push("  (no exercises logged)");
    } else {
      s.lines.forEach((ln) => {
        out.push(`  ${ln.name}${ln.target ? `  [target ${ln.target}]` : ""}`);
        out.push(`    did: ${ln.actual}`);
      });
    }
    out.push("");
  });

  out.push(`${sessions.length} session${sessions.length === 1 ? "" : "s"} logged this month.`);
  return out.join("\n");
}

/* ------------------------------ Event wiring ----------------------------- */

function getList(kind) {
  if (kind === "routine") return state.routineTasks;
  if (kind === "meals") return state.meals;
  if (kind === "side") return state.sideMissions;
  return null;
}

// Adds `minutes` to a "HH:MM" string (wrapping past midnight), defaulting to
// the current wall-clock time when empty. Pure — shared by the single-input
// bump (bumpTimeInput) and the cascade bump (cascadeBumpTime) below.
function addMinutesToTimeValue(value, minutes) {
  let base;
  if (value) {
    const [h, m] = value.split(":").map(Number);
    base = h * 60 + m;
  } else {
    const now = new Date();
    base = now.getHours() * 60 + now.getMinutes();
  }
  const total = ((base + minutes) % 1440 + 1440) % 1440;
  const hh = String(Math.floor(total / 60)).padStart(2, "0");
  const mm = String(total % 60).padStart(2, "0");
  return `${hh}:${mm}`;
}

// Bumps a time <input>'s value directly, then fires a real "change" event so
// the existing per-field handler persists it — same code path as picking a
// time by hand, no separate state-mutation logic needed.
function bumpTimeInput(inputEl, minutes) {
  inputEl.value = addMinutesToTimeValue(inputEl.value, minutes);
  inputEl.dispatchEvent(new Event("change", { bubbles: true }));
}

/* ------------------------- Cascade time-shift (Daily Routine Tasks) -------
   Tap a task's label to anchor it — that row and everything below it (in
   current order) highlight. The anchor's own +5/+20 buttons then shift the
   whole highlighted range by that amount instead of just themselves, so
   "everything from here on is running late" is a two-tap fix rather than
   bumping each task individually. Only rows that already have a time set
   get shifted (nothing fabricated for blank ones) — the anchor itself is the
   one exception, matching the single-row bump's "default to now" behavior.
   Selection is pure UI state, not app data — not persisted, not synced.
   ------------------------------------------------------------------------- */

let timeShiftAnchorId = null;

function applyTimeShiftHighlight() {
  const rows = $$("#routine-list .task-row");
  let foundAnchor = false;
  rows.forEach((row) => {
    const isAnchor = row.dataset.id === timeShiftAnchorId;
    if (isAnchor) foundAnchor = true;
    row.classList.toggle("shift-anchor-row", isAnchor);
    row.classList.toggle("shift-cascade", foundAnchor && !isAnchor);
    const label = row.querySelector(".task-label");
    if (label) label.classList.toggle("shift-anchor-label", isAnchor);
  });
  if (!foundAnchor) timeShiftAnchorId = null;
}

function cascadeBumpTime(anchorId, minutes) {
  const idx = state.routineTasks.findIndex((t) => t.id === anchorId);
  if (idx === -1) return;
  pushUndo();
  for (let i = idx; i < state.routineTasks.length; i++) {
    const task = state.routineTasks[i];
    if (i === idx || task.time) {
      task.time = addMinutesToTimeValue(task.time, minutes);
    }
  }
  renderAll();
}

document.addEventListener("click", (e) => {
  const t = e.target.closest("button") || e.target;

  if (t.id === "undo-btn") {
    undoLastAction();
    return;
  }

  if (t.classList.contains("check-btn") && t.dataset.kind) {
    pushUndo();
    const list = getList(t.dataset.kind);
    const item = list.find((x) => x.id === t.dataset.id);
    if (item) item.done = !item.done;
    renderAll();
    return;
  }

  if (t.classList.contains("del-btn") && t.dataset.kind) {
    if (!confirm("Delete this task?")) return;
    pushUndo();
    const list = getList(t.dataset.kind);
    const idx = list.findIndex((x) => x.id === t.dataset.id);
    if (idx > -1) list.splice(idx, 1);
    renderAll();
    return;
  }

  if (t.classList.contains("edit-btn")) {
    const list = getList(t.dataset.kind);
    const item = list.find((x) => x.id === t.dataset.id);
    if (!item) return;
    const newLabel = prompt("Task name:", item.label);
    if (newLabel === null) return;
    const newPoints = prompt("Points (use a negative number for a penalty):", item.points);
    if (newPoints === null) return;
    const n = Number(newPoints);
    pushUndo();
    item.label = newLabel.trim() || item.label;
    if (!Number.isNaN(n)) item.points = n;
    renderAll();
    return;
  }

  if (t.id && t.id.startsWith("gym-done-")) {
    pushUndo();
    const idx = Number(t.dataset.idx);
    state.gymSchedule[idx].done = !state.gymSchedule[idx].done;
    renderAll();
    return;
  }

  if (t.dataset && t.dataset.field && t.dataset.day) {
    pushUndo();
    const day = t.dataset.day;
    const field = t.dataset.field;
    if (!state.monthly.days[day]) state.monthly.days[day] = { gym: false, sauna: false, redlight: false, hbot: false };
    state.monthly.days[day][field] = !state.monthly.days[day][field];
    renderAll();
    return;
  }

  if (t.classList.contains("add-task-btn") && t.dataset.kind) {
    const kind = t.dataset.kind;
    const list = getList(kind);
    const labelInput = $(`#add-${kind}-label`);
    const timeInput = $(`#add-${kind}-time`);
    const pointsInput = $(`#add-${kind}-points`);
    const label = labelInput.value.trim();
    const points = Number(pointsInput.value);
    if (!label || Number.isNaN(points)) return;
    pushUndo();
    list.push({ id: uid(), label, points, done: false, time: timeInput ? timeInput.value : "", icon: "" });
    labelInput.value = "";
    pointsInput.value = "5";
    if (timeInput) timeInput.value = "";
    renderAll();
    return;
  }

  if (t.id === "add-side-bulk-btn") {
    const textarea = $("#add-side-bulk");
    const lines = textarea.value.split("\n").map((l) => l.trim()).filter(Boolean);
    if (lines.length === 0) return;
    pushUndo();
    lines.forEach((label) => {
      state.sideMissions.push({ id: uid(), label, points: 5, done: false, time: "", icon: "" });
    });
    textarea.value = "";
    renderAll();
    return;
  }

  if (t.classList.contains("tab-btn")) {
    switchTab(t.dataset.tab);
    return;
  }

  if (t.classList.contains("day-pill") || t.closest(".day-pill")) {
    const pill = t.classList.contains("day-pill") ? t : t.closest(".day-pill");
    trainingViewDay = Number(pill.dataset.day);
    renderTraining();
    return;
  }

  if (t.dataset && t.dataset.trainingId) {
    pushUndo();
    const id = t.dataset.trainingId;
    state.training.done[id] = !state.training.done[id];
    maybeAutoCheckGymBonus();
    renderAll();
    return;
  }

  if (t.classList.contains("ex-log-btn")) {
    const id = t.dataset.actualId;
    const row = t.closest(".ex-row");
    const input = row && row.querySelector(".ex-actual");
    const value = input ? input.value.trim() : "";
    pushUndo();
    if (value) state.trainingActuals[id] = value;
    // Logging your actual performance implies you did the exercise — one tap
    // instead of typing the actual *and* separately ticking the checkbox.
    state.training.done[id] = true;
    maybeAutoCheckGymBonus();
    renderAll();
    return;
  }

  if (t.dataset && t.dataset.historyId) {
    openExerciseHistory(t.dataset.historyId);
    return;
  }

  if (t.id === "history-modal-close" || t.id === "history-modal") {
    $("#history-modal").hidden = true;
    return;
  }

  if (t.id === "training-add-exercise") {
    const session = sessionForDay(trainingViewDay);
    if (!session) return;
    const val = (id) => ($("#" + id) ? $("#" + id).value.trim() : "");
    const name = val("cx-name");
    if (!name) {
      alert("Give the exercise a name.");
      return;
    }
    pushUndo();
    if (!state.customExercises[session.key]) state.customExercises[session.key] = [];
    state.customExercises[session.key].push({
      id: `${session.key}-c${uid()}`,
      section: val("cx-section"),
      name,
      sets: val("cx-sets") || "—",
      reps: val("cx-reps") || "—",
      weight: val("cx-weight") || "—",
      notes: val("cx-notes")
    });
    ["cx-section", "cx-name", "cx-sets", "cx-reps", "cx-weight", "cx-notes"].forEach((id) => {
      if ($("#" + id)) $("#" + id).value = "";
    });
    renderAll();
    return;
  }

  if (t.dataset && t.dataset.removeCustomId) {
    const removeId = t.dataset.removeCustomId;
    const sessionKey = removeId.slice(0, removeId.lastIndexOf("-"));
    const list = state.customExercises[sessionKey];
    if (!list) return;
    if (!confirm("Remove this added exercise?")) return;
    pushUndo();
    state.customExercises[sessionKey] = list.filter((e) => e.id !== removeId);
    delete state.training.done[removeId];
    delete state.trainingActuals[removeId];
    renderAll();
    return;
  }

  if (t.id === "training-report-btn") {
    const report = buildTrainingReport();
    const out = $("#training-report-output");
    const status = $("#training-report-status");
    out.value = report;
    out.hidden = false;
    out.style.height = "auto";
    out.style.height = Math.min(out.scrollHeight + 4, 360) + "px";
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(report).then(
        () => { if (status) { status.textContent = "Copied to clipboard — paste it to Claude."; status.hidden = false; } },
        () => { if (status) { status.textContent = "Select the text below and copy it."; status.hidden = false; } }
      );
    } else if (status) {
      status.textContent = "Select the text below and copy it.";
      status.hidden = false;
    }
    return;
  }

  if (t.dataset && t.dataset.rehabIdx) {
    pushUndo();
    const idx = Number(t.dataset.rehabIdx);
    state.training.rehabDone[idx] = !state.training.rehabDone[idx];
    renderAll();
    return;
  }

  if (t.dataset && t.dataset.pancakeIdx) {
    pushUndo();
    const idx = Number(t.dataset.pancakeIdx);
    state.training.pancakeDone[idx] = !state.training.pancakeDone[idx];
    renderAll();
    return;
  }

  if (t.id === "training-complete-all") {
    pushUndo();
    const session = sessionForDay(trainingViewDay);
    if (session) {
      sessionExerciseEntries(session).forEach((e) => { state.training.done[e.id] = true; });
      const todayIdx = todayDayIdx();
      if (trainingViewDay === todayIdx) state.gymSchedule[todayIdx].done = true;
    }
    renderAll();
    return;
  }

  const calCell = t.closest && t.closest(".cal-cell");
  if (calCell && calCell.dataset.date) {
    const iso = calCell.dataset.date;
    const dateObj = new Date(iso + "T00:00:00");
    const label = dateObj.toLocaleDateString(undefined, { weekday: "long", day: "2-digit", month: "short" });
    const existing = state.dayPlans[iso] || "";
    const plan = prompt(`High-level plan for ${label}:`, existing);
    if (plan === null) return;
    pushUndo();
    if (plan.trim()) state.dayPlans[iso] = plan.trim();
    else delete state.dayPlans[iso];
    renderAll();
    return;
  }

  if (t.id === "save-score-btn") {
    const date = $("#save-score-date").value || todayISO();
    const earned = Number($("#save-score-earned").value);
    const max = Number($("#save-score-max").value);
    if (Number.isNaN(earned) || Number.isNaN(max) || max < 0) {
      alert("Enter valid earned and max point values.");
      return;
    }
    pushUndo();
    state.dailyLog[date] = { earned, max };
    renderAll();
    return;
  }

  if (t.dataset && t.dataset.rightNowId) {
    pushUndo();
    const item = state.rightNow.find((x) => x.id === t.dataset.rightNowId);
    if (item) item.done = !item.done;
    renderAll();
    return;
  }

  if (t.dataset && t.dataset.rightNowDel) {
    if (!confirm("Delete this item?")) return;
    pushUndo();
    const idx = state.rightNow.findIndex((x) => x.id === t.dataset.rightNowDel);
    if (idx > -1) state.rightNow.splice(idx, 1);
    renderAll();
    return;
  }

  if (t.id === "add-right-now-btn") {
    const input = $("#add-right-now-label");
    const label = input.value.trim();
    if (!label) return;
    pushUndo();
    state.rightNow.push({ id: uid(), label, done: false });
    input.value = "";
    renderAll();
    return;
  }

  if (t.classList.contains("time-bump-btn")) {
    const row = t.closest(".task-row");
    const minutes = Number(t.dataset.bumpMinutes);
    if (row && row.dataset.id === timeShiftAnchorId && row.dataset.kind === "routine") {
      cascadeBumpTime(row.dataset.id, minutes);
    } else {
      const input = row && row.querySelector(".task-time, .plan-time");
      if (input) bumpTimeInput(input, minutes);
    }
    return;
  }

  if (t.classList.contains("task-label") && t.dataset.kind === "routine") {
    timeShiftAnchorId = (timeShiftAnchorId === t.dataset.id) ? null : t.dataset.id;
    applyTimeShiftHighlight();
    return;
  }

  if (t.id === "account-signup-btn" || t.id === "account-login-btn") {
    const statusEl = $("#account-status");
    if (!supabaseClient) {
      statusEl.textContent = "Sync isn't available right now (offline, or the Supabase library failed to load).";
      return;
    }
    const email = $("#account-email").value.trim();
    const password = $("#account-password").value;
    if (!email || !password) {
      statusEl.textContent = "Enter an email and password.";
      return;
    }
    const isSignUp = t.id === "account-signup-btn";
    statusEl.textContent = isSignUp ? "Signing up…" : "Logging in…";
    const action = isSignUp ? signUpAccount(email, password) : logInAccount(email, password);
    action.then(({ data, error }) => {
      if (error) {
        statusEl.textContent = error.message;
        return;
      }
      if (isSignUp && !data.session) {
        statusEl.textContent = "Check your email to confirm your account, then log in.";
        return;
      }
      statusEl.textContent = "";
    });
    return;
  }

  if (t.id === "account-logout-btn") {
    logOutAccount();
    return;
  }

  if (t.id === "add-project-btn") {
    const title = prompt("Project name:", "New Project");
    if (title === null) return;
    pushUndo();
    state.projects.push({ id: uid(), title: title.trim() || "Untitled Project", content: "" });
    renderAll();
    return;
  }

  if (t.dataset && t.dataset.projectRename) {
    const proj = state.projects.find((p) => p.id === t.dataset.projectRename);
    if (!proj) return;
    const title = prompt("Project name:", proj.title);
    if (title === null) return;
    pushUndo();
    proj.title = title.trim() || proj.title;
    renderAll();
    return;
  }

  if (t.dataset && t.dataset.projectDel) {
    if (!confirm("Delete this project?")) return;
    pushUndo();
    const idx = state.projects.findIndex((p) => p.id === t.dataset.projectDel);
    if (idx > -1) state.projects.splice(idx, 1);
    renderAll();
    return;
  }

  if (t.id === "news-refresh-btn") {
    fetchNews();
    return;
  }
});

// Project notes and logged workout "actuals" save via their own debounced
// listeners rather than the usual pushUndo()+renderAll() pattern —
// renderAll() rebuilds these lists from scratch, which would tear the
// focused input (and cursor position) out from under the user mid-keystroke.
let projectSaveTimer = null;
let exActualSaveTimer = null;
let rhrSaveTimer = null;
document.addEventListener("input", (e) => {
  const t = e.target;

  if (t.classList.contains("project-note")) {
    const proj = state.projects.find((p) => p.id === t.dataset.projectId);
    if (!proj) return;
    proj.content = t.value;
    if (projectSaveTimer) clearTimeout(projectSaveTimer);
    projectSaveTimer = setTimeout(saveState, 400);
    return;
  }

  if (t.classList.contains("ex-actual")) {
    state.trainingActuals[t.dataset.actualId] = t.value;
    if (exActualSaveTimer) clearTimeout(exActualSaveTimer);
    exActualSaveTimer = setTimeout(saveState, 400);
    return;
  }

  if (t.classList.contains("rhr-input")) {
    const day = t.dataset.day;
    if (!state.monthly.days[day]) state.monthly.days[day] = { gym: false, sauna: false, redlight: false, hbot: false, rhr: "" };
    state.monthly.days[day].rhr = t.value;
    if (rhrSaveTimer) clearTimeout(rhrSaveTimer);
    // Skip renderAll() like the other two above -- it would rebuild the
    // whole monthly table and yank focus out from under the number pad
    // mid-entry. The Avg RHR hint just goes stale until the next full
    // render (tab switch, midnight rollover, etc), same tradeoff as
    // project notes and logged actuals.
    rhrSaveTimer = setTimeout(saveState, 400);
    return;
  }
});

document.addEventListener("change", (e) => {
  const t = e.target;

  if (t.id === "save-score-date") {
    const earnedInput = $("#save-score-earned");
    const maxInput = $("#save-score-max");
    const isToday = applySaveScoreToday(t, earnedInput, maxInput);
    if (!isToday) {
      const existing = state.dailyLog[t.value];
      earnedInput.value = existing ? existing.earned : "";
      maxInput.value = existing ? existing.max : "";
    }
    return;
  }

  if (t.classList.contains("gym-type-select")) {
    pushUndo();
    const idx = Number(t.dataset.dayIdx);
    state.weeklySessionKeys[idx] = t.value || null;
    renderAll();
    return;
  }

  // Deliberately skip renderAll() for both time inputs below: it wipes and
  // rebuilds the whole list's DOM (el.innerHTML = ""), which yanks the
  // native time picker closed mid-interaction — on some browsers "change"
  // fires after each segment (hour, then minute), so a full re-render after
  // the first one used to slam the picker shut before you could finish.
  // Time doesn't feed into any score calculation, so nothing else needs to
  // visually update here anyway — just persist it.
  if (t.classList.contains("task-time")) {
    pushUndo();
    const list = getList(t.dataset.kind);
    const item = list.find((x) => x.id === t.dataset.id);
    if (item) item.time = t.value;
    saveState();
    updateUndoButton();
    return;
  }

  if (t.classList.contains("plan-time")) {
    pushUndo();
    state.tomorrowPlan.times[t.dataset.id] = t.value;
    saveState();
    updateUndoButton();
    return;
  }

  const isMutating = ["sleep-time", "last-meal-time", "gym-finish-time", "last-coffee-time", "no-scroll-check"].includes(t.id);
  if (isMutating) pushUndo();

  if (t.id === "sleep-time") state.timing.sleep = t.value;
  if (t.id === "last-meal-time") state.timing.lastMeal = t.value;
  if (t.id === "gym-finish-time") state.timing.gymFinish = t.value;
  if (t.id === "last-coffee-time") state.timing.lastCoffee = t.value;
  if (t.id === "no-scroll-check") state.timing.noScroll = t.checked;
  if (isMutating) renderAll();
});

// Enter submits the "add" row a field belongs to, same as clicking its
// button — inputs opt in via data-enter-target="<button id>". Left off
// textareas (e.g. the bulk-add box) so Enter there still just inserts a
// newline.
document.addEventListener("keydown", (e) => {
  if (e.key === "Escape") {
    const modal = $("#history-modal");
    if (modal && !modal.hidden) modal.hidden = true;
    return;
  }
  if (e.key !== "Enter") return;
  const t = e.target;
  if (t.tagName !== "INPUT" || !t.dataset.enterTarget) return;
  e.preventDefault();
  const btn = document.getElementById(t.dataset.enterTarget);
  if (btn) btn.click();
});

function switchTab(tab) {
  $$(".tab-btn").forEach((b) => b.classList.toggle("active", b.dataset.tab === tab));
  $$(".tab-panel").forEach((p) => p.classList.toggle("active", p.id === "tab-" + tab));
}

/* -------------------------------- Boot ----------------------------------- */

injectIcons();
renderAll();
initAuth();
fetchNews();

// Re-check for a date rollover periodically, in case the app is left open
// past midnight.
setInterval(() => {
  const before = state.currentDate;
  checkRollover();
  if (before !== state.currentDate) renderAll();
}, 60 * 1000);

// Keep the News Board reasonably fresh without refetching on every render.
setInterval(fetchNews, 15 * 60 * 1000);

if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => {
    navigator.serviceWorker.register("./sw.js").catch((err) => console.warn("SW registration failed", err));
  });
}
