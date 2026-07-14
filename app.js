/* =========================================================================
   Daily Tracker — vanilla JS PWA
   Ported from "DailyTracker_v6" Excel workbook.
   All data lives in localStorage on this device (phone-only, per your choice).
   ========================================================================= */

const STORAGE_KEY = "dailyTracker.state.v1";

const DEFAULT_NOTES = "📚  REFERENCE & NOTES\n  Sections: 🧠 Mind & Mental Health Notes   |   🦴 QL Muscle Recovery Plan   |   🦶 Ankle Rehabilitation   |   💪 Strength Training Block   |   🎯 Overcoming Social Anxiety   |   🤜 Wrist Injury Recovery\n\n\n🧠  Mind & Mental Health Notes\n  🧠 2. Fixing the “Constant Brain Chatter” \n🧠 2. Fixing the “Constant Brain Chatter” \nThat nonstop mental noise you mentioned? That’s very real—and strongly linked to overstimulation. \nYour brain is used to: \nconstant input \nfast switching \nhigh dopamine spikes \nSo silence feels uncomfortable. \n🎯 Goal: \nNot “empty mind”—but quieter, more controlled thinking \n🔧 Simple System That Actually Works \n1. Daily “Mental Dump” (5–10 min) \nGrab your phone or paper and write: \neverything on your mind \nno structure, no filter \nThis clears mental backlog. \n2. Controlled Silence (Start small) \nDaily: \n5–10 minutes \nno phone, no music, no stimulation \nJust sit or walk. \n👉 At first your brain will race—that’s normal. Over time, it settles. \n3. Reduce Input Overload \nRight now: \nYouTube binge = mental chaos fuel \nBy cutting that down, you’ll notice: \nless noise \nbetter focus \nmore stable mood \n4. Focus Training \nPick ONE task daily: \nwork, reading, learning \nDo it: \nno switching \nno background noise \nEven 30–60 minutes helps rewire your attention span. \n\n\n🦴  QL Muscle Recovery Plan\n  Chronic tightness (9 months) — Comprehensive rehab protocol\nRIGHT QL RECOVERY PLAN\nChronic tightness (9 months) — Comprehensive rehab protocol\n\nWHY IT'S NOT RESOLVING\nWeak / inhibited glutes\nTight hip flexors (psoas)\nHip drop / pelvic imbalance\nSleep position\n\n⚠  RED FLAGS — SEE A PHYSIO IF:\n  •  Pain radiates down into the glute or leg (could involve lumbar spine, not just QL)\n  •  Pain is sharp or stabbing rather than tight / achy\n  •  Symptoms worsen despite 3–4 weeks of consistent rehab work\n  •  Numbness or tingling in the leg (nerve involvement)\n🥊  MUAY THAI NOTE\nCheck whether you consistently kick/knee with the right leg. Repeated hip extension under load is a direct aggravator. Balance volume between sides and prioritise post-training QL + hip flexor stretching on training days.\n\nRECOVERY TIMELINE\nPeriod\nWeek 1–2\nWeek 3–4\nWeek 5–6\nWeek 7–8\nWeek 9–12\n\n\n\n\n\n\n\n\n\n🦶  Ankle Rehabilitation\n  \nANKLE REHABILITATION PROGRAMME\nLateral ankle sprain (inversion) — Functional recovery to full loaded range-of-motion\nPhase\nPhase 1\nMobility\nPhase 2\nProprioception\nPhase 3\nStrength\nIMPORTANT NOTES\n• Phases overlap — start Phase 2 proprioception work from Week 2 once mobility is improving.\n• Pain guide: 0 = no pain, 10 = worst imaginable. Stop and rest if pain spikes above 4/10.\n• For lotus / seiza pain: use a yoga block under hips and reduce height by ~1 cm per week.\n• See a physio if lateral instability persists during cutting/side-step movements.\n• Muay Thai: shadow work and bag from Week 6 if Phase 2 goals are met. Sparring: Week 10+.\n\n\n💪  Strength Training Block\n  Front Lever · Handstand Push-Up · Leg Strength\n6-WEEK STRENGTH BLOCK — FULL PLAN\nFront Lever · Handstand Push-Up · Leg Strength\nYOUR 1RM ESTIMATES\nLift\nBarbell Overhead Press\nWeighted Pull-Up\nATG Barbell Squat\nBarbell Deadlift\nHip Thrust\nCalf Raise (endurance)\nWEEKLY SCHEDULE\nMonday\nDay 1 – Upper Push\nPROGRESSION PHASES\nPhase\nAccumulation\nIntensification\nPeak\nDeload\n\n\n🎯  Overcoming Social Anxiety\n  \nOVERCOMING SOCIAL ANXIETY\nA 6-Week Speed-Run Plan  •  From avoidance to confidence\nSocial anxiety shrinks through deliberate, repeated exposure — not willpower, not waiting to feel ready. Each week builds on the last. The goal isn't zero anxiety; it's proving to your nervous system that social situations are survivable, then normal.\nExpose Daily\nStack 2-3 micro-interactions daily — volume beats intensity\n6-WEEK PLAN AT A GLANCE\nWeek\nWeek 1\nWeek 2\nWeek 3\nWeek 4\nWeek 5\nWeek 6\nDAILY NON-NEGOTIABLES\nPhysical\nMental\nReadiness comes after action, not before.  Start today.\n\n\n🤜  Wrist Injury Recovery\n  Mechanism: Blocking kick → forceful wrist extension under load   |   Primary concern: Dorsal radiocarpal / intercarpal l\n🥊  Wrist Injury Recovery Plan  —  Muay Thai Dorsal Sprain\nMechanism: Blocking kick → forceful wrist extension under load   |   Primary concern: Dorsal radiocarpal / intercarpal ligament sprain\nPhase\nPhase 1\nAcute\nPhase 2\nSubacute\nPhase 3\nRehabilitation\nPhase 4\nReturn to\nTraining\n⚠️  SCAPHOID FRACTURE SCREENING — Check before proceeding with rehab\nTest\nAnatomical Snuffbox Test\nScaphoid Compression Test\nSwelling Pattern\nAction if suspicious\n🗓️  Wrist Recovery — Daily Tracker\nLog each day. Tick ✓ completed modalities. Rate pain 0 (none) → 10 (severe). Note ROM: Full / Partial / Limited / None in painful direction (extension).\nPHASE 1 — ACUTE (Days 1-5)\nDay\n1\n2\n3\nPHASE 2 — SUBACUTE (Days 6-14)\n4\n5\n6\n7\n8\n9\n10\n11\nPHASE 3 — REHABILITATION (Days 15-42)\n12\n13\n14\n15\n16\n17\n18\n19\n20\n21\n22\n23\n24\n25\n26\n27\n28\n29\n30\n31\n32\n33\n34\n35\n36\n37\n38\nPHASE 4 — RETURN TO TRAINING\n39\n40\n41\n42";

const GYM_DAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];
const GYM_TYPES = ["Pull", "Push", "Legs", "Pull", "Push", "Legs", "Active Rest"];
const MONTH_NAMES = ["January","February","March","April","May","June","July","August","September","October","November","December"];

let uidCounter = 1;
function uid() { return "id" + (uidCounter++) + "_" + Math.random().toString(36).slice(2, 7); }

function defaultState() {
  const today = todayISO();
  const d = new Date();
  return {
    version: 1,
    currentDate: today,
    routineTasks: [
      ["🧴 Face Routine (AM)", 5],
      ["🧘 Stretch", 8],
      ["🦶 Ankle Training", 5],
      ["🚶 Walk / Steps", 10],
      ["📚 Audiobook", 4],
      ["🧘 Meditate", 12],
      ["🚿 Shower", 5],
      ["💊 Supplements", 5],
      ["🧹 Clean", 5],
      ["🚶 QL Recovery", 5],
      ["📵 Relapse", -80],
      ["🧠 Mental dump + Controlled silence", 10],
      ["💪 Gym Session", 15],
      ["🧴 Face Routine (PM)", 5],
      ["📚 Language learning", 10],
      ["🧠 Respondent surveys", 5],
      ["😤 Face your fears", 10],
      ["🌙 Evening Wind Down", 5],
      ["👨‍🍳 Meal prep", 8],
      ["🧴 Asthma Inhaler", 6],
      ["Water Floss", 2]
    ].map(([label, points]) => ({ id: uid(), label, points, done: false })),
    meals: [
      ["🥩 Healthy Meal 1", 8],
      ["🥦 Healthy Meal 2", 8],
      ["🫠 Healthy Meal 3", 8],
      ["🥦 No Junk Food?", 8]
    ].map(([label, points]) => ({ id: uid(), label, points, done: false })),
    timing: { sleep: "", lastMeal: "", gymFinish: "", lastCoffee: "", noScroll: false },
    gymSchedule: GYM_DAYS.map((day, i) => ({ day, type: GYM_TYPES[i], done: false })),
    sideMissions: [
      ["Shave", 5],
      ["Make bed", 5]
    ].map(([label, points]) => ({ id: uid(), label, points, done: false })),
    dailyLog: {},
    monthly: {
      year: d.getFullYear(),
      month: d.getMonth() + 1, // 1-12
      days: {}
    },
    archive: [],
    notes: DEFAULT_NOTES
  };
}

function todayISO(date) {
  const d = date || new Date();
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

function loadState() {
  const raw = localStorage.getItem(STORAGE_KEY);
  if (!raw) return defaultState();
  try {
    const parsed = JSON.parse(raw);
    // fill in any missing fields from a default (forward-compat safety)
    const base = defaultState();
    return Object.assign({}, base, parsed, {
      timing: Object.assign({}, base.timing, parsed.timing),
      monthly: Object.assign({}, base.monthly, parsed.monthly)
    });
  } catch (e) {
    console.error("Failed to parse saved state, starting fresh.", e);
    return defaultState();
  }
}

function saveState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

let state = loadState();

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
  const dow = new Date().getDay(); // 0=Sun..6=Sat
  const idx = dow === 0 ? 6 : dow - 1; // convert to 0=Mon..6=Sun
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

  state.currentDate = today;
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
  renderTaskSection("routine-list", state.routineTasks, "routine");
  renderTaskSection("meals-list", state.meals, "meals");
  renderTaskSection("side-list", state.sideMissions, "side");
  renderTiming();
  renderGymSchedule();
  renderCalendar();
  renderMonthly();
  renderArchive();
  $("#notes-area").value = state.notes;
  saveState();
}

function renderHeader() {
  const totals = computeTotals(state);
  const d = new Date();
  const dayName = d.toLocaleDateString(undefined, { weekday: "long" });
  const dateStr = d.toLocaleDateString(undefined, { day: "2-digit", month: "short", year: "numeric" });
  $("#today-label").textContent = `${dayName} ${dateStr}`;
  $("#score-label").textContent = `🏆 ${totals.earned} / ${totals.max} pts`;
  const pct = totals.max > 0 ? Math.max(0, Math.min(100, (totals.earned / totals.max) * 100)) : 0;
  $("#score-bar-fill").style.width = pct + "%";
}

function renderTaskSection(listId, tasks, kind) {
  const el = $("#" + listId);
  el.innerHTML = "";
  tasks.forEach((t) => {
    const row = document.createElement("div");
    row.className = "task-row" + (t.done ? " done" : "");
    row.innerHTML = `
      <button class="check-btn" data-kind="${kind}" data-id="${t.id}" aria-label="toggle">${t.done ? "✓" : "✗"}</button>
      <span class="task-label" data-kind="${kind}" data-id="${t.id}">${esc(t.label)}</span>
      <span class="task-points ${t.points < 0 ? "neg" : ""}" data-kind="${kind}" data-id="${t.id}">${t.points > 0 ? "+" : ""}${t.points}</span>
      <button class="icon-btn edit-btn" data-kind="${kind}" data-id="${t.id}" aria-label="edit">✎</button>
      <button class="icon-btn del-btn" data-kind="${kind}" data-id="${t.id}" aria-label="delete">🗑</button>
    `;
    el.appendChild(row);
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
  const dow = new Date().getDay();
  const todayIdx = dow === 0 ? 6 : dow - 1;
  const el = $("#gym-schedule-list");
  el.innerHTML = "";
  state.gymSchedule.forEach((g, i) => {
    const row = document.createElement("div");
    row.className = "gym-row" + (i === todayIdx ? " today" : "");
    row.innerHTML = `
      <span class="gym-day">${g.day}${i === todayIdx ? " ← today" : ""}</span>
      <span class="gym-type">${esc(g.type)}</span>
      <button class="check-btn small" data-idx="${i}" id="gym-done-${i}">${g.done ? "✓" : "✗"}</button>
    `;
    el.appendChild(row);
  });
  const bonus = gymBonus(state);
  $("#gym-bonus-label").textContent = `Schedule bonus: ${bonus.earned} / 10 pts`;
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
    let entry = state.dailyLog[iso];
    if (iso === todayISO()) entry = { earned: totals.earned, max: totals.max };
    if (entry && entry.max > 0) {
      const pct = Math.max(0, Math.min(1, entry.earned / entry.max));
      cell.style.background = heatColor(pct);
      cell.title = `${entry.earned} / ${entry.max} pts`;
    }
    if (iso === todayISO()) cell.classList.add("today");
    cell.innerHTML = `<span>${day}</span>`;
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
  header.innerHTML = `<span>DATE</span><span>💪 GYM</span><span>🧖 SAUNA</span><span>🔴 RED LIGHT</span><span>💨 HBOT</span><span>PTS</span>`;
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
      <span>💪 ${a.gym} &nbsp; 🧖 ${a.sauna} &nbsp; 🔴 ${a.redlight} &nbsp; 💨 ${a.hbot}</span>
      <span>Total: ${a.totalPts} pts &middot; Best day: ${a.bestDay} pts</span>
    `;
    el.appendChild(row);
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
  const t = e.target;

  if (t.classList.contains("check-btn") && t.dataset.kind) {
    const list = getList(t.dataset.kind);
    const item = list.find((x) => x.id === t.dataset.id);
    if (item) item.done = !item.done;
    renderAll();
    return;
  }

  if (t.classList.contains("del-btn")) {
    const list = getList(t.dataset.kind);
    const idx = list.findIndex((x) => x.id === t.dataset.id);
    if (idx > -1 && confirm("Delete this task?")) list.splice(idx, 1);
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
    item.label = newLabel.trim() || item.label;
    if (!Number.isNaN(n)) item.points = n;
    renderAll();
    return;
  }

  if (t.id && t.id.startsWith("gym-done-")) {
    const idx = Number(t.dataset.idx);
    state.gymSchedule[idx].done = !state.gymSchedule[idx].done;
    renderAll();
    return;
  }

  if (t.dataset && t.dataset.field && t.dataset.day) {
    const day = t.dataset.day;
    const field = t.dataset.field;
    if (!state.monthly.days[day]) state.monthly.days[day] = { gym: false, sauna: false, redlight: false, hbot: false };
    state.monthly.days[day][field] = !state.monthly.days[day][field];
    renderAll();
    return;
  }

  if (t.classList.contains("add-task-btn")) {
    const kind = t.dataset.kind;
    const list = getList(kind);
    const labelInput = $(`#add-${kind}-label`);
    const pointsInput = $(`#add-${kind}-points`);
    const label = labelInput.value.trim();
    const points = Number(pointsInput.value);
    if (!label || Number.isNaN(points)) return;
    list.push({ id: uid(), label, points, done: false });
    labelInput.value = "";
    pointsInput.value = "5";
    renderAll();
    return;
  }

  if (t.classList.contains("tab-btn")) {
    switchTab(t.dataset.tab);
    return;
  }
});

document.addEventListener("change", (e) => {
  const t = e.target;
  if (t.id === "sleep-time") state.timing.sleep = t.value;
  if (t.id === "last-meal-time") state.timing.lastMeal = t.value;
  if (t.id === "gym-finish-time") state.timing.gymFinish = t.value;
  if (t.id === "last-coffee-time") state.timing.lastCoffee = t.value;
  if (t.id === "no-scroll-check") state.timing.noScroll = t.checked;
  if (t.id === "notes-area") state.notes = t.value;
  renderAll();
});

function switchTab(tab) {
  $$(".tab-btn").forEach((b) => b.classList.toggle("active", b.dataset.tab === tab));
  $$(".tab-panel").forEach((p) => p.classList.toggle("active", p.id === "tab-" + tab));
}

/* -------------------------------- Boot ----------------------------------- */

renderAll();

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
