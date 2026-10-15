/**
 * ============================================================================
 * BETTER BUDGET - WEEKLY BUDGET REPORT
 * ============================================================================
 * Emails a weekly summary of expense spending vs. monthly budgets.
 *
 * Uses the shared pieces of the tool:
 *   - readChartOfAccounts(), getCurrencySymbol(), getSetting()  (Settings and Setup.gs)
 *   - validateLedgerSheet() and findHeaderRow()                 (Monthly Reports.gs)
 *
 * Setup for the user: type a Monthly Budget next to the expense accounts they
 * want tracked (Chart of Accounts, column D), then use
 * Better Budget > Weekly Email > Turn on. Optional: set "Report email" in
 * the Settings table to send the report somewhere other than their own account.
 *
 * All names here start with "weekly"/"Weekly"/"WEEKLY" so they can't clash
 * with names in other files.
 */

const WEEKLY_CONFIG = {
  PACE_TOLERANCE: 10,                 // % points ahead of month progress before "Ahead of pace"
  PROJECTION_MIN_DAY: 7,              // don't judge pace before this day of the month
  TRIGGER_DAY: ScriptApp.WeekDay.MONDAY,
  TRIGGER_HOUR: 8,
};

const WEEKLY_MONTHS = ['January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December'];

// ---------------------------------------------------------------------------
// MAIN
// ---------------------------------------------------------------------------

function sendWeeklyReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();

  // "Today" as a plain local date, built from the spreadsheet's time zone
  const p = Utilities.formatDate(new Date(), tz, 'yyyy-M-d').split('-').map(Number);
  const today = new Date(p[0], p[1] - 1, p[2]);
  const endOfToday = new Date(p[0], p[1] - 1, p[2], 23, 59, 59);
  const weekStart = new Date(p[0], p[1] - 1, p[2] - 6);   // last 7 days incl. today
  const monthStart = new Date(p[0], p[1] - 1, 1);
  const daysInMonth = new Date(p[0], p[1], 0).getDate();
  const dayOfMonth = p[2];
  const monthPct = (dayOfMonth / daysInMonth) * 100;

  const accounts = weeklyReadAccounts_();

  // This month's ledger, plus last month's if the 7-day window crosses into it
  const monthRows = weeklyReadLedger_(weeklyFindLedgerSheet_(ss, today));
  let weekRows = monthRows;
  if (weekStart < monthStart) {
    weekRows = weekRows.concat(weeklyReadLedger_(weeklyFindLedgerSheet_(ss, weekStart)));
  }

  // Aggregate expense spend per account (expenses are net debits)
  const spend = {};
  const getSpend = function (code) {
    return spend[code] || (spend[code] = { month: 0, week: 0 });
  };

  monthRows.forEach(function (r) {
    if (!weeklyIsExpense_(accounts[r.code])) return;
    getSpend(r.code).month += r.debit - r.credit;
  });

  const bigTx = [];
  weekRows.forEach(function (r) {
    if (!weeklyIsExpense_(accounts[r.code])) return;
    if (!r.date || r.date < weekStart || r.date > endOfToday) return;
    const net = r.debit - r.credit;
    getSpend(r.code).week += net;
    if (net > 0) bigTx.push({ date: r.date, desc: r.desc, name: accounts[r.code].name, amount: net });
  });

  // Split into budgeted / unbudgeted
  const budgeted = [];
  const unbudgeted = [];
  Object.keys(spend).forEach(function (code) {
    const a = accounts[code];
    const s = spend[code];
    if (a.budget === null) {
      if (Math.abs(s.month) > 0.005 || Math.abs(s.week) > 0.005) {
        unbudgeted.push({ name: a.name, month: s.month, week: s.week });
      }
      return;
    }
    budgeted.push(weeklyAssess_(code, a, s, dayOfMonth, daysInMonth, monthPct));
  });

  // Budgeted accounts with no spending yet still belong in the report
  Object.keys(accounts).forEach(function (code) {
    const a = accounts[code];
    if (weeklyIsExpense_(a) && a.budget !== null && !spend[code]) {
      budgeted.push(weeklyAssess_(code, a, { month: 0, week: 0 }, dayOfMonth, daysInMonth, monthPct));
    }
  });

  const rank = { 'OVER BUDGET': 0, 'On pace to exceed': 1, 'Ahead of pace': 2, 'On track': 3 };
  budgeted.sort(function (a, b) { return rank[a.status] - rank[b.status] || b.pctUsed - a.pctUsed; });
  unbudgeted.sort(function (a, b) { return b.month - a.month; });
  bigTx.sort(function (a, b) { return b.amount - a.amount; });

  const html = weeklyBuildHtml_(budgeted, unbudgeted, bigTx.slice(0, 5), {
    ssName: ss.getName(), noBudgets: budgeted.length === 0, ssUrl: ss.getUrl(), weekStart: weekStart, today: today,
    dayOfMonth: dayOfMonth, daysInMonth: daysInMonth, monthPct: monthPct,
  });

  const flagged = budgeted.filter(function (r) { return r.status !== 'On track'; }).length;
  const subject = 'Weekly budget report - ' + WEEKLY_MONTHS[today.getMonth()].slice(0, 3) + ' ' +
    today.getDate() + ', ' + today.getFullYear() +
    (flagged ? ' (' + flagged + ' need attention)' : ' (all on track)');

  MailApp.sendEmail({
    to: weeklyRecipient_(),
    subject: subject,
    htmlBody: html,
  });
}

/** Menu: sends the report immediately. */
function sendWeeklyReportNow() {
  const ui = SpreadsheetApp.getUi();
  try {
    sendWeeklyReport();
  } catch (e) {
    ui.alert('Could not send the weekly report', String(e.message || e), ui.ButtonSet.OK);
    return;
  }
  SpreadsheetApp.getActiveSpreadsheet().toast('Weekly report emailed to ' + weeklyRecipient_() + '.', 'Done', 4);
}

/** Menu: Weekly Email > Turn on. Safe to repeat: replaces any earlier schedule. */
function turnOnWeeklyEmail() {
  const ui = SpreadsheetApp.getUi();
  createWeeklyReportTrigger();
  ui.alert('Weekly email is on',
    'A report will be emailed to ' + weeklyRecipient_() + ' every ' + weeklyDayName_() +
    ' around ' + weeklyHourLabel_() + '.\n\n' +
    'Tip: use Better Budget > Weekly Email > Send now to see what it looks like.\n' +
    'To send it to a different address, type it next to "Report email" in the Settings table on the Chart of Accounts.',
    ui.ButtonSet.OK);
}

/** Menu: Weekly Email > Turn off. */
function turnOffWeeklyEmail() {
  const ui = SpreadsheetApp.getUi();
  const removed = weeklyDeleteTriggers_();
  ui.alert(removed ? 'Weekly email is off' : 'Weekly email was already off',
    removed ? 'No more weekly reports will be sent. You can turn it back on at any time.'
            : 'There was no weekly schedule to remove.', ui.ButtonSet.OK);
}

/** Menu: Weekly Email > Status. */
function weeklyEmailStatus() {
  const ui = SpreadsheetApp.getUi();
  const on = weeklyHasTrigger_();
  ui.alert('Weekly email: ' + (on ? 'ON' : 'OFF'),
    on ? 'Sent to ' + weeklyRecipient_() + ' every ' + weeklyDayName_() + ' around ' + weeklyHourLabel_() + '.'
       : 'Use Better Budget > Weekly Email > Turn on to start receiving it.',
    ui.ButtonSet.OK);
}

/** Creates the weekly schedule (no dialogs, so it also works from the editor). */
function createWeeklyReportTrigger() {
  weeklyDeleteTriggers_();
  ScriptApp.newTrigger('sendWeeklyReport')
    .timeBased()
    .onWeekDay(WEEKLY_CONFIG.TRIGGER_DAY)
    .atHour(WEEKLY_CONFIG.TRIGGER_HOUR)
    .create();
}

function weeklyDeleteTriggers_() {
  let n = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sendWeeklyReport') { ScriptApp.deleteTrigger(t); n++; }
  });
  return n;
}

function weeklyHasTrigger_() {
  return ScriptApp.getProjectTriggers().some(function (t) {
    return t.getHandlerFunction() === 'sendWeeklyReport';
  });
}

function weeklyDayName_() {
  const d = String(WEEKLY_CONFIG.TRIGGER_DAY);   // "MONDAY"
  return d.charAt(0) + d.slice(1).toLowerCase();
}

function weeklyHourLabel_() {
  const h = WEEKLY_CONFIG.TRIGGER_HOUR;
  return (h % 12 === 0 ? 12 : h % 12) + (h < 12 ? ' am' : ' pm');
}

/** Report email setting if it looks like an address, otherwise the user's own account. */
function weeklyRecipient_() {
  let v = '';
  try { v = String(getSetting('Report email') || '').trim(); } catch (e) { v = ''; }
  if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v)) return v;
  return Session.getEffectiveUser().getEmail();
}

// ---------------------------------------------------------------------------
// ASSESSMENT
// ---------------------------------------------------------------------------

function weeklyAssess_(code, a, s, dayOfMonth, daysInMonth, monthPct) {
  const pctUsed = a.budget > 0 ? (s.month / a.budget) * 100 : 0;
  const projected = (s.month / dayOfMonth) * daysInMonth;
  let status = 'On track';
  if (s.month > a.budget) {
    status = 'OVER BUDGET';
  } else if (dayOfMonth >= WEEKLY_CONFIG.PROJECTION_MIN_DAY) {
    if (projected > a.budget) status = 'On pace to exceed';
    else if (pctUsed > monthPct + WEEKLY_CONFIG.PACE_TOLERANCE) status = 'Ahead of pace';
  }
  return {
    code: code, name: a.name, budget: a.budget, spent: s.month, week: s.week,
    pctUsed: pctUsed, projected: projected, remaining: a.budget - s.month, status: status,
  };
}

// ---------------------------------------------------------------------------
// DATA ACCESS
// ---------------------------------------------------------------------------

function weeklyIsExpense_(acc) {
  return !!acc && (acc.type === 'Expense' || acc.type === 'Expenses');
}

/** Chart of Accounts via the shared reader: { code: {name, type, budget|null} } */
function weeklyReadAccounts_() {
  const chart = readChartOfAccounts();
  const out = {};
  Object.keys(chart).forEach(function (code) {
    out[code] = { name: chart[code].name, type: chart[code].type, budget: chart[code].budget };
  });
  if (!Object.keys(out).length) {
    throw new Error('The Chart of Accounts is empty. Run Better Budget > Set Up / Upgrade Tool first.');
  }
  return out;
}

/**
 * Finds the General Ledger tab for the month containing `date`.
 * 1) tries a tab named "October 2026" style, 2) falls back to scanning every
 * tab's title block (B3 = month, B4 = year), the same cells your report engine uses.
 */
function weeklyFindLedgerSheet_(ss, date) {
  const month = WEEKLY_MONTHS[date.getMonth()];
  const year = String(date.getFullYear());

  const byName = ss.getSheetByName(month + ' ' + year);
  if (byName && validateLedgerSheet(byName).valid) return byName;

  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    const sh = sheets[i];
    if (!validateLedgerSheet(sh).valid) continue;
    const m = String(sh.getRange('B3').getValue()).trim().toLowerCase();
    const y = String(sh.getRange('B4').getValue()).trim();
    if (m === month.toLowerCase() && y === year) return sh;
  }
  return null;
}

/** Reads ledger rows: Date | Account | Description | Debit | Credit */
function weeklyReadLedger_(sheet) {
  if (!sheet) return [];
  const raw = sheet.getDataRange().getValues();
  const hIdx = findHeaderRow(raw);
  const start = (hIdx !== -1 ? hIdx : 7) + 1;   // your engine defaults to header on row 8

  const out = [];
  for (let i = start; i < raw.length; i++) {
    const r = raw[i];
    if (r[1] === '' || r[1] === null) continue;
    out.push({
      date: weeklyParseDate_(r[0]),
      code: String(r[1]).trim(),
      desc: String(r[2] || ''),
      debit: Number(r[3]) || 0,
      credit: Number(r[4]) || 0,
    });
  }
  return out;
}

/** Handles real dates and "2026-10-11" text (as sent by your sidebar). */
function weeklyParseDate_(v) {
  if (v instanceof Date) return v;
  if (typeof v === 'string') {
    const m = v.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  }
  const d = new Date(v);
  return isNaN(d.getTime()) ? null : d;
}

// ---------------------------------------------------------------------------
// EMAIL
// ---------------------------------------------------------------------------

let WEEKLY_CUR_CACHE_ = null;
function weeklyCurrency_() {
  if (WEEKLY_CUR_CACHE_ === null) {
    try { WEEKLY_CUR_CACHE_ = getCurrencySymbol(); } catch (e) { WEEKLY_CUR_CACHE_ = '$'; }
  }
  return WEEKLY_CUR_CACHE_;
}

function weeklyMoney_(n) {
  const neg = n < 0;
  const s = Math.abs(n).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return (neg ? '-' : '') + weeklyCurrency_() + s;
}

function weeklyEsc_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function weeklyShortDate_(d) {
  return WEEKLY_MONTHS[d.getMonth()].slice(0, 3) + ' ' + d.getDate();
}

function weeklyBuildHtml_(budgeted, unbudgeted, bigTx, meta) {
  const colors = {
    'OVER BUDGET': '#c62828', 'On pace to exceed': '#e65100',
    'Ahead of pace': '#b8860b', 'On track': '#2e7d32',
  };
  const sum = function (arr, k) { return arr.reduce(function (a, r) { return a + r[k]; }, 0); };

  const totalBudget = sum(budgeted, 'budget');
  const totalSpent = sum(budgeted, 'spent');
  const weekAll = sum(budgeted, 'week') + sum(unbudgeted, 'week');

  let h = '<div style="font-family:Arial,sans-serif;max-width:760px">';
  h += '<h2 style="margin-bottom:4px">Weekly budget report</h2>';
  h += '<p style="color:#555;margin-top:0">' + weeklyShortDate_(meta.weekStart) + ' - ' +
       weeklyShortDate_(meta.today) + ' &nbsp;|&nbsp; Day ' + meta.dayOfMonth + ' of ' +
       meta.daysInMonth + ' (' + meta.monthPct.toFixed(0) + '% of the month)</p>';

  h += '<p><b>Spent this week (all expenses):</b> ' + weeklyMoney_(weekAll) +
       '<br><b>Budgeted categories, month to date:</b> ' + weeklyMoney_(totalSpent) +
       ' of ' + weeklyMoney_(totalBudget) + ' (' +
       (totalBudget ? (totalSpent / totalBudget * 100).toFixed(0) : 0) + '%)</p>';

  if (meta.noBudgets) {
    h += '<p style="background:#fff8e1;padding:8px;border-radius:4px">No budgets are set yet. Type a monthly amount next to ' +
         'your expense accounts (Monthly Budget column on the Chart of Accounts) to get over-budget and pace alerts.</p>';
  }

  h += '<table style="border-collapse:collapse;width:100%;font-size:13px">';
  h += '<tr style="background:#eee;text-align:left"><th style="padding:6px">Category</th><th>Budget</th>' +
       '<th>Spent</th><th>% used</th><th>Left</th><th>This week</th><th>Projected</th><th>Status</th></tr>';
  budgeted.forEach(function (r) {
    h += '<tr style="border-bottom:1px solid #ddd">' +
      '<td style="padding:6px">' + weeklyEsc_(r.name) + '</td>' +
      '<td>' + weeklyMoney_(r.budget) + '</td>' +
      '<td>' + weeklyMoney_(r.spent) + '</td>' +
      '<td>' + r.pctUsed.toFixed(0) + '%</td>' +
      '<td>' + weeklyMoney_(r.remaining) + '</td>' +
      '<td>' + weeklyMoney_(r.week) + '</td>' +
      '<td>' + weeklyMoney_(r.projected) + '</td>' +
      '<td style="color:' + colors[r.status] + ';font-weight:bold">' + r.status + '</td></tr>';
  });
  h += '</table>';

  if (unbudgeted.length) {
    h += '<h3>Spending with no budget set</h3><ul>';
    unbudgeted.forEach(function (u) {
      h += '<li>' + weeklyEsc_(u.name) + ': ' + weeklyMoney_(u.month) + ' this month (' +
           weeklyMoney_(u.week) + ' this week)</li>';
    });
    h += '</ul>';
  }

  if (bigTx.length) {
    h += '<h3>Largest expenses this week</h3><ul>';
    bigTx.forEach(function (t) {
      h += '<li>' + weeklyShortDate_(t.date) + ' - ' + weeklyEsc_(t.desc || t.name) +
           ' (' + weeklyEsc_(t.name) + '): ' + weeklyMoney_(t.amount) + '</li>';
    });
    h += '</ul>';
  }

  h += '<p style="color:#777;font-size:11px">Projected = spend so far / days elapsed x days in month. ' +
       'Pace flags start on day ' + WEEKLY_CONFIG.PROJECTION_MIN_DAY + '. ' +
       '<a href="' + meta.ssUrl + '">Open ' + weeklyEsc_(meta.ssName || 'your sheet') + '</a></p></div>';
  return h;
}
