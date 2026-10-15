/**
 * ============================================================================
 * BETTER BUDGET - VERSION AND ABOUT
 * ============================================================================
 * NEW file. Bump BETTER_BUDGET_VERSION whenever you publish a new copy of the
 * template, so users (and you, when they send feedback) know what they have.
 */

const BETTER_BUDGET_VERSION = '1.0.0';

const BETTER_BUDGET_DISCLAIMER =
  'Better Budget is a personal budgeting aid. It is not professional financial, tax or legal advice, ' +
  'and it comes with no warranty. Check important figures yourself and keep your own backups ' +
  '(File > Make a copy).';

/** Menu: Better Budget > About. */
function showAbout() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('Better Budget v' + BETTER_BUDGET_VERSION,
    BETTER_BUDGET_DISCLAIMER + '\n\n' +
    'Your data stays in this spreadsheet. The only thing the tool ever sends out is the optional weekly ' +
    'email, and it goes to the address in your Settings (or your own Google account).',
    ui.ButtonSet.OK);
}
