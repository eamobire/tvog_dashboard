// TVOG Ops Dashboard — Apps Script data endpoint
//
// Serves the TVOG Ops Google Sheets as one JSON object, so the standalone
// dashboard webapp can fetch live data from them. Three of the KPI rows
// (Net cash position, Monthly net cash flow, Weekly prayer meeting
// attendance, Bible study attendance) don't read their "Actual" from the
// KPIs sheet at all — they're computed live from Enoch's existing Finance,
// Prayer Meeting, and Bible Study trackers, via the LiveSource column on
// the KPIs sheet. Everything else still comes straight from that sheet.
//
// SETUP (one-time):
// 1. Open script.google.com, create a new project, paste this whole file in.
// 2. Deploy -> New deployment -> type "Web app".
//      Execute as:  Me
//      Who has access:  Anyone  (see the note at the bottom of this file
//        before picking "Anyone within [org]" instead — it changes how
//        the webapp has to call this endpoint)
// 3. Click Deploy, authorize when Google asks, then copy the Web app URL
//    it gives you (ends in /exec).
// 4. Paste that URL into index.html's API_URL constant (near the bottom of
//    that file's <script> block — open it in a text editor, not a browser,
//    to edit it; see the note further down).

const SHEET_IDS = {
  departments: '1BtQ8EUnZEZt2xqn_bZMTjdZO64Xkdu3jtLZRHuFDPyg', // Name | Lead | Owns
  kpis:        '16ZXOdpNat4AabWtyLE96kal6oC6BtZPeO4sPFR3Ad2U', // Department | KPI | Target | Actual | Status | Frequency | LiveSource
  nextSteps:   '110z4nmYUqWTLsmAXVnLlWgpLEqzmdjGo2tCNWAoPLDc', // Title | Detail | Owner

  // Enoch's existing, already-in-use trackers — read directly, not copied in.
  financeTracker:    '1EO9l3X9S-z-HTFLK0I6ZcFmeQr60qQqp7-KV8txdEjM', // TVOG Finance Tracker
  prayerTracker:     '1Uyy_fLt_NyzQLbBLFOgmMHcS1H9XNlnkX7iPheGlaGU', // TVOG Prayer Meeting Tracker
  bibleStudyTracker: '17VoFKW8IAr-UlSrgjJVvZxX1dutdXZa9m6RjBNGz4do'  // TVO Bible Study Tracker
};

function doGet(e) {
  const payload = {
    meta: {
      lastSynced: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
      dataStatus: 'live',
      source: 'TVOG Ops Google Sheets (live) + Finance/Prayer/Bible Study trackers (live)'
    },
    departments: readDepartments(),
    kpis: readKpis(),
    nextSteps: readNextSteps()
  };
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function readDepartments() {
  const rows = getRows(SHEET_IDS.departments);
  return rows.map(r => ({
    name: String(r[0] || '').trim(),
    lead: String(r[1] || '').trim(),
    owns: String(r[2] || '').trim()
  })).filter(d => d.name);
}

function readKpis() {
  const rows = getRows(SHEET_IDS.kpis);
  return rows.map(r => {
    const k = {
      department: String(r[0] || '').trim(),
      kpi: String(r[1] || '').trim(),
      target: String(r[2] || '').trim(),
      actual: r[3] ? String(r[3]).trim() : null,
      status: r[4] ? String(r[4]).trim() : null,
      frequency: String(r[5] || '').trim()
    };
    const source = r[6] ? String(r[6]).trim() : '';
    if (source) {
      const live = getLiveActual(source);
      if (live && live.actual != null) {
        k.actual = live.actual;
        k.status = live.status;
      }
    }
    return k;
  }).filter(k => k.department);
}

function readNextSteps() {
  const rows = getRows(SHEET_IDS.nextSteps);
  return rows.map(r => ({
    title: String(r[0] || '').trim(),
    detail: String(r[1] || '').trim(),
    owner: String(r[2] || '').trim()
  })).filter(s => s.title);
}

function getRows(sheetId) {
  const sheet = SpreadsheetApp.openById(sheetId).getSheets()[0];
  const values = sheet.getDataRange().getValues();
  values.shift(); // drop header row
  return values;
}

// -----------------------------------------------------------------------
// LIVE KPI SOURCES
//
// These read Enoch's existing trackers directly, so the dashboard shows
// real numbers without anyone copying a figure from one sheet to another.
// If you add another live-sourced KPI later: add a new "source" key here,
// a matching case below, and put that key in the KPIs sheet's LiveSource
// column on that row.
// -----------------------------------------------------------------------

function getLiveActual(source) {
  try {
    if (source === 'finance_cash') return financeCashPosition();
    if (source === 'finance_flow') return financeNetFlow();
    if (source === 'prayer_attendance') return trackerAttendance(SHEET_IDS.prayerTracker);
    if (source === 'bible_attendance') return trackerAttendance(SHEET_IDS.bibleStudyTracker);
  } catch (err) {
    return { actual: null, status: null };
  }
  return null;
}

// Finance Tracker, "Monthly Summary" tab: reads the "Year Total" row's
// Closing Balance (column I) — this already carries the running balance
// forward through every month, so it's always the current position.
function financeCashPosition() {
  const sheet = SpreadsheetApp.openById(SHEET_IDS.financeTracker).getSheetByName('Monthly Summary');
  const values = sheet.getDataRange().getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === 'Year Total') {
      const balance = Number(values[i][8]) || 0;
      return { actual: formatMoney(balance), status: balance > 0 ? 'On Track' : 'Behind' };
    }
  }
  return { actual: null, status: null };
}

// Finance Tracker, "Monthly Summary" tab: reads the most recent month that
// actually had activity (Total In or Total Out > 0) and returns its
// Net (In - Out), column H.
function financeNetFlow() {
  const sheet = SpreadsheetApp.openById(SHEET_IDS.financeTracker).getSheetByName('Monthly Summary');
  const values = sheet.getDataRange().getValues();
  let lastActive = null;
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === 'Year Total') break;
    const totalIn = Number(values[i][5]) || 0;
    const totalOut = Number(values[i][6]) || 0;
    if (totalIn > 0 || totalOut > 0) lastActive = values[i];
  }
  if (!lastActive) return { actual: null, status: null };
  const net = Number(lastActive[7]) || 0;
  const monthCell = lastActive[0];
  const monthLabel = (monthCell instanceof Date)
    ? Utilities.formatDate(monthCell, Session.getScriptTimeZone(), 'MMM yyyy')
    : String(monthCell);
  return { actual: formatMoney(net) + ' (' + monthLabel + ')', status: net >= 0 ? 'On Track' : 'At Risk' };
}

// Prayer Meeting / Bible Study trackers, "Attendance Log" tab: both share
// the same column layout (Date | Department/Location | Notes | Attendance
// | Prior Meeting | Comment | Month). Reads the last logged session and,
// where a prior count is on record, shows growth vs. it.
function trackerAttendance(sheetId) {
  const sheet = SpreadsheetApp.openById(sheetId).getSheetByName('Attendance Log');
  const values = sheet.getDataRange().getValues();
  let last = null;
  for (let i = 1; i < values.length; i++) { // skip header row
    if (values[i][0]) last = values[i];
  }
  if (!last) return { actual: null, status: null };
  const attendance = Number(last[3]) || 0;
  const prior = Number(last[4]) || 0;
  if (prior > 0) {
    const growth = ((attendance - prior) / prior) * 100;
    const sign = growth >= 0 ? '+' : '';
    return {
      actual: attendance + ' (' + sign + growth.toFixed(0) + '% vs prior)',
      status: growth >= 0 ? 'On Track' : 'At Risk'
    };
  }
  return { actual: String(attendance), status: 'Pending' };
}

function formatMoney(n) {
  n = Number(n) || 0;
  return 'GHS' + n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// -----------------------------------------------------------------------
// ACCESS NOTE
//
// "Who has access: Anyone" means anyone who has this /exec URL can read
// the JSON — nobody can discover it by guessing, but it isn't restricted
// to your organization either. That's the only setting a plain webapp on
// GitHub Pages can call directly with fetch(), because it can't complete
// a Google sign-in on its own.
//
// If you need real per-person access control instead (Google sign-in,
// restricted to specific people), the webapp needs a "Sign in with
// Google" step added to it, and this deployment changes to "Anyone
// within [your organization]". That's a heavier change — say the word
// and Claude can build that version instead.
// -----------------------------------------------------------------------
