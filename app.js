/* =========================================================================
   Daily Tracker — vanilla JS PWA
   Ported from "DailyTracker_v6" Excel workbook.
   All data lives in localStorage on this device (phone-only, per your choice).
   ========================================================================= */

const STORAGE_KEY = "dailyTracker.state.v1";

const DEFAULT_NOTES = "📚  REFERENCE & NOTES\n  Sections: 🧠 Mind & Mental Health Notes   |   🦴 QL Muscle Recovery Plan   |   🦶 Ankle Rehabilitation   |   💪 Strength Training Block   |   🎯 Overcoming Social Anxiety   |   🤜 Wrist Injury Recovery\n\n\n🧠  Mind & Mental Health Notes\n  🧠 2. Fixing the “Constant Brain Chatter” \n🧠 2. Fixing the “Constant Brain Chatter” \nThat nonstop mental noise you mentioned? That’s very real—and strongly linked to overstimulation. \nYour brain is used to: \nconstant input \nfast switching \nhigh dopamine spikes \nSo silence feels uncomfortable. \n🎯 Goal: \nNot “empty mind”—but quieter, more controlled thinking \n🔧 Simple System That Actually Works \n1. Daily “Mental Dump” (5–10 min) \nGrab your phone or paper and write: \neverything on your mind \nno structure, no filter \nThis clears mental backlog. \n2. Controlled Silence (Start small) \nDaily: \n5–10 minutes \nno phone, no music, no stimulation \nJust sit or walk. \n👉 At first your brain will race—that’s normal. Over time, it settles. \n3. Reduce Input Overload \nRight now: \nYouTube binge = mental chaos fuel \nBy cutting that down, you’ll notice: \nless noise \nbetter focus \nmore stable mood \n4. Focus Training \nPick ONE task daily: \nwork, reading, learning \nDo it: \nno switching \nno background noise \nEven 30–60 minutes helps rewire your attention span. \n\n\n🦴  QL Muscle Recovery Plan\n  Chronic tightness (9 months) — Comprehensive rehab protocol\nRIGHT QL RECOVERY PLAN\nChronic tightness (9 months) — Comprehensive rehab protocol\n\nWHY IT'S NOT RESOLVING\nWeak / inhibited glutes\nTight hip flexors (psoas)\nHip drop / pelvic imbalance\nSleep position\n\n⚠  RED FLAGS — SEE A PHYSIO IF:\n  •  Pain radiates down into the glute or leg (could involve lumbar spine, not just QL)\n  •  Pain is sharp or stabbing rather than tight / achy\n  •  Symptoms worsen despite 3–4 weeks of consistent rehab work\n  •  Numbness or tingling in the leg (nerve involvement)\n🥊  MUAY THAI NOTE\nCheck whether you consistently kick/knee with the right leg. Repeated hip extension under load is a direct aggravator. Balance volume between sides and prioritise post-training QL + hip flexor stretching on training days.\n\nRECOVERY TIMELINE\nPeriod\nWeek 1–2\nWeek 3–4\nWeek 5–6\nWeek 7–8\nWeek 9–12\n\n\n\n\n\n\n\n\n\n🦶  Ankle Rehabilitation\n  \nANKLE REHABILITATION PROGRAMME\nLateral ankle sprain (inversion) — Functional recovery to full loaded range-of-motion\nPhase\nPhase 1\nMobility\nPhase 2\nProprioception\nPhase 3\nStrength\nIMPORTANT NOTES\n• Phases overlap — start Phase 2 proprioception work from Week 2 once mobility is improving.\n• Pain guide: 0 = no pain, 10 = worst imaginable. Stop and rest if pain spikes above 4/10.\n• For lotus / seiza pain: use a yoga block under hips and reduce height by ~1 cm per week.\n• See a physio if lateral instability persists during cutting/side-step movements.\n• Muay Thai: shadow work and bag from Week 6 if Phase 2 goals are met. Sparring: Week 10+.\n\n\n💪  Strength Training Block\n  Front Lever · Handstand Push-Up · Leg Strength\n6-WEEK STRENGTH BLOCK — FULL PLAN\nFront Lever · Handstand Push-Up · Leg Strength\nYOUR 1RM ESTIMATES\nLift\nBarbell Overhead Press\nWeighted Pull-Up\nATG Barbell Squat\nBarbell Deadlift\nHip Thrust\nCalf Raise (endurance)\nWEEKLY SCHEDULE\nMonday\nDay 1 – Upper Push\nPROGRESSION PHASES\nPhase\nAccumulation\nIntensification\nPeak\nDeload\n\n\n🎯  Overcoming Social Anxiety\n  \nOVERCOMING SOCIAL ANXIETY\nA 6-Week Speed-Run Plan  •  From avoidance to confidence\nSocial anxiety shrinks through deliberate, repeated exposure — not willpower, not waiting to feel ready. Each week builds on the last. The goal isn't zero anxiety; it's proving to your nervous system that social situations are survivable, then normal.\nExpose Daily\nStack 2-3 micro-interactions daily — volume beats intensity\n6-WEEK PLAN AT A GLANCE\nWeek\nWeek 1\nWeek 2\nWeek 3\nWeek 4\nWeek 5\nWeek 6\nDAILY NON-NEGOTIABLES\nPhysical\nMental\nReadiness comes after action, not before.  Start today.\n\n\n🤜  Wrist Injury Recovery\n  Mechanism: Blocking kick → forceful wrist extension under load   |   Primary concern: Dorsal radiocarpal / intercarpal l\n🥊  Wrist Injury Recovery Plan  —  Muay Thai Dorsal Sprain\nMechanism: Blocking kick → forceful wrist extension under load   |   Primary concern: Dorsal radiocarpal / intercarpal ligament sprain\nPhase\nPhase 1\nAcute\nPhase 2\nSubacute\nPhase 3\nRehabilitation\nPhase 4\nReturn to\nTraining\n⚠️  SCAPHOID FRACTURE SCREENING — Check before proceeding with rehab\nTest\nAnatomical Snuffbox Test\nScaphoid Compression Test\nSwelling Pattern\nAction if suspicious\n🗓️  Wrist Recovery — Daily Tracker\nLog each day. Tick ✓ completed modalities. Rate pain 0 (none) → 10 (severe). Note ROM: Full / Partial / Limited / None in painful direction (extension).\nPHASE 1 — ACUTE (Days 1-5)\nDay\n1\n2\n3\nPHASE 2 — SUBACUTE (Days 6-14)\n4\n5\n6\n7\n8\n9\n10\n11\nPHASE 3 — REHABILITATION (Days 15-42)\n12\n13\n14\n15\n16\n17\n18\n19\n20\n21\n22\n23\n24\n25\n26\n27\n28\n29\n30\n31\n32\n33\n34\n35\n36\n37\n38\nPHASE 4 — RETURN TO TRAINING\n39\n40\n41\n42";

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
  cloud: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><path d="M7 18a4.5 4.5 0 0 1-.5-8.97A5.5 5.5 0 0 1 17 8.5a4 4 0 0 1-.7 7.94"/><path d="M16.3 16.44H7"/></svg>'
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
    title: "Pull A — Weighted PU & OAP",
    focus: "Primary: lats · biceps · scapular depressors · unilateral pulling strength",
    exercises: [
      { section: "MAIN LIFT", name: "Weighted Pull-Up (pronated, sh-width)", sets: 4, reps: "5–8", weight: "+20kg", notes: "+2.5kg when 4×6 achieved. Target: +30kg · Current: 6,5,4,4 at +20kg — sets 3–4 are the focus" },
      { section: "OAP PROGRESSION", name: "Archer Pull-Up", sets: 4, reps: "5–6 ea. side", weight: "Bodyweight", notes: "Add 2.5kg when 4×6 clean each side · Extend non-working arm progressively straighter each week" },
      { section: null, name: "One-Arm Negative (OAN)", sets: 3, reps: "3–5 ea. side", weight: "Bodyweight", notes: "Slow the eccentric — target 5s down · Jump to top, lower with one arm. Most direct OAP builder" },
      { section: null, name: "One-Arm Dead Hang", sets: 3, reps: "10–20s ea.", weight: "Bodyweight", notes: "Add load when 20s comfortable · Grip and shoulder stability. Alternate arms" },
      { section: "SUPPLEMENTARY PULLING", name: "Straight-Arm Lat Pulldown", sets: 4, reps: "10–12", weight: "50kg", notes: "+2.5kg when 4×12 · Arms locked. Front lever and OAP both need this" },
      { section: null, name: "Barbell Pendlay Row", sets: 4, reps: "6–8", weight: "60kg", notes: "+2.5kg when 4×8 · Horizontal torso. Reset each rep — don't skip this one" },
      { section: null, name: "Seated Cable Row (wide grip)", sets: 3, reps: "10", weight: "62.5kg", notes: "+2.5kg when 3×10 · Mid-back thickness and retraction endurance" },
      { section: "ACCESSORY", name: "Bicep Curl (drop set)", sets: 2, reps: "10 → failure", weight: "14kg → 8kg", notes: "Add 1kg top set when 12 reps before drop · OAP requires strong bicep under full load" },
      { section: null, name: "Cable Face Pull", sets: 3, reps: "15", weight: "17.5kg", notes: "+2.5kg when 3×15 · Shoulder health — don't skip" },
      { section: null, name: "Scapular Pull-Ups", sets: 3, reps: "10–12", weight: "Bodyweight", notes: "Add 5kg when 3×12 · Slow and deliberate. Depression strength for OAP entry" }
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
    title: "Push B — Handstand Skill",
    focus: "Primary: handstand balance · scapular control · shoulder endurance",
    exercises: [
      { section: "HANDSTAND SKILL — quality over quantity", name: "Freestanding Handstand Attempts", sets: "5–8", reps: "Max hold", weight: "Bodyweight", notes: "Target: consistent 20s → 30s → 60s · Kick up, find balance. Rest 90s between. Log best hold each session" },
      { section: null, name: "Wall Handstand (chest to wall)", sets: 3, reps: "45–60s", weight: "Bodyweight", notes: "Extend to 90s before moving on · Chest-to-wall = better body line than back-to-wall" },
      { section: null, name: "Handstand Shoulder Taps", sets: 3, reps: "8–12 taps", weight: "Bodyweight", notes: "Increase taps per set weekly · Against wall — builds unilateral shoulder stability for balance" },
      { section: null, name: "Pike Compression Hold", sets: 3, reps: "20–30s", weight: "Bodyweight / 5kg on legs", notes: "Extend hold, then add weight · Hip flexor and compression strength — feeds press-to-HS long term" },
      { section: "SHOULDER ACCESSORY", name: "Incline DB Press", sets: 3, reps: "10–12", weight: "20kg DBs", notes: "+2kg when 3×12 · Shoulder stability. Supporting work" },
      { section: null, name: "Cable Pull-Over (straight arm)", sets: 3, reps: "12", weight: "32.5kg", notes: "+2.5kg when 3×12 · Lat in lengthened position — straight-arm strength" },
      { section: null, name: "Rear Delt Fly (pause at top)", sets: 3, reps: "12–15", weight: "10kg DBs", notes: "+2kg when 3×15 · Posterior shoulder balance — protects under heavy pressing volume" },
      { section: null, name: "Band Pull-Aparts", sets: 3, reps: "20", weight: "Medium band", notes: "Progress band resistance · Shoulder health. Keep smooth" },
      { section: null, name: "Rotator Cuff Ext. Rotation", sets: 2, reps: "15 each", weight: "5kg", notes: "Maintain — prehab only · Never load this heavily. Shoulder longevity" }
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
  const key = WEEKLY_SESSION_KEYS[dayIdx];
  return key ? TRAINING_SESSIONS[key] : null;
}

function exerciseId(sessionKey, idx) {
  return `${sessionKey}-${idx}`;
}

let uidCounter = 1;
function uid() { return "id" + (uidCounter++) + "_" + Math.random().toString(36).slice(2, 7); }

function defaultState() {
  const today = todayISO();
  const d = new Date();
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
    sideMissions: [
      ["Shave", 5, ""],
      ["Make bed", 5, ""]
    ].map(([label, points, icon]) => ({ id: uid(), label, points, done: false, time: "", icon })),
    training: { done: {}, rehabDone: {}, pancakeDone: {} },
    dailyLog: {},
    dayPlans: {},
    tomorrowPlan: [],
    rightNow: [],
    monthly: {
      year: d.getFullYear(),
      month: d.getMonth() + 1, // 1-12
      days: {}
    },
    archive: [],
    notes: DEFAULT_NOTES
  };
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

function loadState() {
  const raw = localStorage.getItem(STORAGE_KEY);
  if (!raw) return defaultState();
  try {
    const parsed = JSON.parse(raw);
    // fill in any missing fields from a default (forward-compat safety)
    const base = defaultState();
    const merged = Object.assign({}, base, parsed, {
      timing: Object.assign({}, base.timing, parsed.timing),
      monthly: Object.assign({}, base.monthly, parsed.monthly),
      training: Object.assign({}, base.training, parsed.training),
      dayPlans: Object.assign({}, base.dayPlans, parsed.dayPlans)
    });
    // Gym schedule day *labels* always come fresh from the current program —
    // only the per-day "done" flags are worth keeping from a save.
    if (Array.isArray(parsed.gymSchedule)) {
      merged.gymSchedule = GYM_DAYS.map((day, i) => ({
        day,
        done: !!(parsed.gymSchedule[i] && parsed.gymSchedule[i].done)
      }));
    }
    migrateTaskIcons(merged.routineTasks, ROUTINE_ICON_MIGRATIONS);
    migrateTaskIcons(merged.meals, MEALS_ICON_MIGRATIONS);
    migrateTaskIcons(merged.sideMissions, {});
    return merged;
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

// Reapplies the same forward-compat merge/migration loadState() does for a
// freshly-parsed blob, so a state pulled from the cloud gets the same
// treatment as one loaded from localStorage.
function normalizeIncomingState(raw) {
  const base = defaultState();
  const merged = Object.assign({}, base, raw, {
    timing: Object.assign({}, base.timing, raw.timing),
    monthly: Object.assign({}, base.monthly, raw.monthly),
    training: Object.assign({}, base.training, raw.training),
    dayPlans: Object.assign({}, base.dayPlans, raw.dayPlans)
  });
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
      state = normalizeIncomingState(data.state);
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
// pass immediate=true for the one-off "seed the cloud" push right after login.
function pushStateToCloud(immediate) {
  if (!supabaseClient || !currentUser) return;
  if (cloudPushTimer) clearTimeout(cloudPushTimer);
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
  cloudPushTimer = setTimeout(doPush, 1500);
}

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

  // Reset the daily checklist for the new day.
  state.routineTasks.forEach((t) => (t.done = false));
  state.meals.forEach((t) => (t.done = false));
  state.sideMissions.forEach((t) => (t.done = false));
  state.gymSchedule.forEach((g) => (g.done = false));
  state.timing = { sleep: "", lastMeal: "", gymFinish: "", lastCoffee: "", noScroll: false };
  state.training = { done: {}, rehabDone: {}, pancakeDone: {} };

  // Anything queued in "Plan for Tomorrow" becomes part of the Daily Routine
  // Tasks now that tomorrow has arrived, then the queue clears.
  if (state.tomorrowPlan.length > 0) {
    state.tomorrowPlan.forEach((p) => {
      state.routineTasks.push({ id: uid(), label: p.label, points: p.points, done: false, time: "", icon: "" });
    });
    state.tomorrowPlan = [];
  }

  state.currentDate = today;
  trainingViewDay = todayDayIdx();
  saveState();
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
  state.archive.unshift({ label, gym, sauna, redlight, hbot, totalPts, bestDay });
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
  renderHeader();
  renderTaskSection("routine-list", state.routineTasks, "routine", true);
  renderTaskSection("meals-list", state.meals, "meals", false);
  renderTaskSection("side-list", state.sideMissions, "side", true);
  renderTiming();
  renderGymSchedule();
  renderCalendar();
  renderMonthly();
  renderArchive();
  renderTraining();
  renderSaveScore();
  renderRightNow();
  renderTomorrowPlan();
  $("#notes-area").value = state.notes;
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
      <span class="task-points ${t.points < 0 ? "neg" : ""}" data-kind="${kind}" data-id="${t.id}">${t.points > 0 ? "+" : ""}${t.points}</span>
      <button class="icon-btn edit-btn" data-kind="${kind}" data-id="${t.id}" aria-label="edit">✎</button>
      <button class="icon-btn del-btn" data-kind="${kind}" data-id="${t.id}" aria-label="delete">${iconTag("trash")}</button>
    `;
    el.appendChild(row);
  });
  if (reorderable) attachDragHandlers(el, kind);
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

function attachDragHandlers(containerEl, kind) {
  const rows = Array.from(containerEl.querySelectorAll(".task-row"));
  rows.forEach((row) => {
    const handle = row.querySelector(".drag-handle");
    if (!handle) return;

    handle.addEventListener("pointerdown", (e) => {
      e.preventDefault();
      pushUndo();
      const dragEl = row;
      const list = getList(kind);
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
              swapListItems(list, dragEl.dataset.id, next.dataset.id);
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
              swapListItems(list, dragEl.dataset.id, prev.dataset.id);
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

function renderGymSchedule() {
  const todayIdx = todayDayIdx();
  const el = $("#gym-schedule-list");
  el.innerHTML = "";
  state.gymSchedule.forEach((g, i) => {
    const session = sessionForDay(i);
    const row = document.createElement("div");
    row.className = "gym-row" + (i === todayIdx ? " today" : "");
    row.innerHTML = `
      <span class="gym-day">${g.day}${i === todayIdx ? " ← today" : ""}</span>
      <span class="gym-type">${esc(session ? session.shortTitle : "Rest")}</span>
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
  header.innerHTML = `<span>DATE</span><span>${iconTag("dumbbell")} GYM</span><span>${iconTag("sauna")} SAUNA</span><span>${iconTag("sun")} RED LIGHT</span><span>${iconTag("wind")} HBOT</span><span>PTS</span>`;
  el.appendChild(header);
  for (let day = 1; day <= daysInMonth; day++) {
    const key = String(day);
    const d = m.days[key] || { gym: false, sauna: false, redlight: false, hbot: false };
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
      <span>${pts}</span>
    `;
    el.appendChild(row);
  }
  $("#monthly-total").textContent = `Total: ${totalPts} pts`;
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
      <span>${iconTag("dumbbell")} ${a.gym} &nbsp; ${iconTag("sauna")} ${a.sauna} &nbsp; ${iconTag("sun")} ${a.redlight} &nbsp; ${iconTag("wind")} ${a.hbot}</span>
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
// Jot down tasks for tomorrow's Daily Routine Tasks today — they auto-feed
// into the real list the moment the date actually rolls over (checkRollover).

function renderTomorrowPlan() {
  const el = $("#tomorrow-plan-list");
  if (!el) return;
  el.innerHTML = "";
  state.tomorrowPlan.forEach((item) => {
    const row = document.createElement("div");
    row.className = "task-row";
    row.innerHTML = `
      <span class="task-label">${esc(item.label)}</span>
      <span class="task-points ${item.points < 0 ? "neg" : ""}">${item.points > 0 ? "+" : ""}${item.points}</span>
      <button class="icon-btn del-btn" data-tomorrow-del="${item.id}" aria-label="delete">${iconTag("trash")}</button>
    `;
    el.appendChild(row);
  });
}

/* ------------------------------- Training -------------------------------- */

function exerciseRowHtml(ex, checked, dataAttrs) {
  return `
    <div class="ex-row">
      ${ex.section ? `<div class="ex-section">${esc(ex.section)}</div>` : ""}
      <div class="ex-line">
        <button class="check-btn small" ${dataAttrs}>${checked ? "✓" : "✗"}</button>
        <div class="ex-info">
          <div class="ex-name">${esc(ex.name)}</div>
          <div class="ex-meta">${esc(String(ex.sets))} × ${esc(String(ex.reps))} &middot; ${esc(ex.weight)}</div>
          <div class="ex-notes">${esc(ex.notes)}</div>
        </div>
      </div>
    </div>
  `;
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
  if (!session) {
    sessionEl.innerHTML = `
      <h2>${iconTag("moon")} Rest Day</h2>
      <p class="hint">No lifting session scheduled${trainingViewDay === todayIdx ? " today" : ""}. Good day for the Pancake Program below, or just recover.</p>
    `;
  } else {
    const isToday = trainingViewDay === todayIdx;
    const doneCount = session.exercises.filter((_, i) => state.training.done[exerciseId(session.key, i)]).length;
    sessionEl.innerHTML = `
      <h2>${isToday ? iconTag("pin") + " Today: " : ""}${esc(session.title)}</h2>
      <p class="hint">${esc(session.focus)}</p>
      <p class="hint">${doneCount} / ${session.exercises.length} logged</p>
      <div id="training-exercise-list"></div>
      <button class="add-task-btn" id="training-complete-all" style="margin-top:10px;">${iconTag("checkCircle")} Mark all complete</button>
    `;
    const listEl = $("#training-exercise-list", sessionEl);
    session.exercises.forEach((ex, i) => {
      const id = exerciseId(session.key, i);
      const checked = !!state.training.done[id];
      listEl.insertAdjacentHTML("beforeend", exerciseRowHtml(ex, checked, `data-training-id="${id}"`));
    });
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

/* ------------------------------ Event wiring ----------------------------- */

function getList(kind) {
  if (kind === "routine") return state.routineTasks;
  if (kind === "meals") return state.meals;
  if (kind === "side") return state.sideMissions;
  return null;
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
    // Nice-to-have: once every exercise in *today's* session is logged,
    // auto-check the Gym Schedule box for today so the daily score reflects
    // it without an extra manual tap. One-directional — unticking a single
    // exercise afterward won't claw the bonus back.
    const todayIdx = todayDayIdx();
    const session = sessionForDay(trainingViewDay);
    if (session && trainingViewDay === todayIdx) {
      const allDone = session.exercises.every((_, i) => state.training.done[exerciseId(session.key, i)]);
      if (allDone) state.gymSchedule[todayIdx].done = true;
    }
    renderAll();
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
      session.exercises.forEach((_, i) => { state.training.done[exerciseId(session.key, i)] = true; });
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

  if (t.dataset && t.dataset.tomorrowDel) {
    pushUndo();
    const idx = state.tomorrowPlan.findIndex((x) => x.id === t.dataset.tomorrowDel);
    if (idx > -1) state.tomorrowPlan.splice(idx, 1);
    renderAll();
    return;
  }

  if (t.id === "add-tomorrow-btn") {
    const labelInput = $("#add-tomorrow-label");
    const pointsInput = $("#add-tomorrow-points");
    const label = labelInput.value.trim();
    const points = Number(pointsInput.value);
    if (!label || Number.isNaN(points)) return;
    pushUndo();
    state.tomorrowPlan.push({ id: uid(), label, points });
    labelInput.value = "";
    pointsInput.value = "5";
    renderAll();
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

  if (t.classList.contains("task-time")) {
    pushUndo();
    const list = getList(t.dataset.kind);
    const item = list.find((x) => x.id === t.dataset.id);
    if (item) item.time = t.value;
    renderAll();
    return;
  }

  const isMutating = ["sleep-time", "last-meal-time", "gym-finish-time", "last-coffee-time", "no-scroll-check", "notes-area"].includes(t.id);
  if (isMutating) pushUndo();

  if (t.id === "sleep-time") state.timing.sleep = t.value;
  if (t.id === "last-meal-time") state.timing.lastMeal = t.value;
  if (t.id === "gym-finish-time") state.timing.gymFinish = t.value;
  if (t.id === "last-coffee-time") state.timing.lastCoffee = t.value;
  if (t.id === "no-scroll-check") state.timing.noScroll = t.checked;
  if (t.id === "notes-area") state.notes = t.value;
  if (isMutating) renderAll();
});

function switchTab(tab) {
  $$(".tab-btn").forEach((b) => b.classList.toggle("active", b.dataset.tab === tab));
  $$(".tab-panel").forEach((p) => p.classList.toggle("active", p.id === "tab-" + tab));
}

/* -------------------------------- Boot ----------------------------------- */

injectIcons();
renderAll();
initAuth();

// Re-check for a date rollover periodically, in case the app is left open
// past midnight.
setInterval(() => {
  const before = state.currentDate;
  checkRollover();
  if (before !== state.currentDate) renderAll();
}, 60 * 1000);

if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => {
    navigator.serviceWorker.register("./sw.js").catch((err) => console.warn("SW registration failed", err));
  });
}
