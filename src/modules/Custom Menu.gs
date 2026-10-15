/**
 * Adds the "Better Budget" menu whenever the spreadsheet is opened.
 * (After pasting this in, reload the sheet once so the menu rebuilds.)
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Better Budget')
    .addItem('Make Transaction Entry', 'showSidebar')
    .addItem('Start New Month', 'startNextMonth')
    .addItem('Generate Monthly Reports', 'generateConsolidatedReports')
    .addItem('Generate Executive Dashboard', 'generateDashboard')
    .addSeparator()
    .addSubMenu(ui.createMenu('Weekly Email')
      .addItem('Turn on', 'turnOnWeeklyEmail')
      .addItem('Turn off', 'turnOffWeeklyEmail')
      .addItem('Status', 'weeklyEmailStatus')
      .addItem('Send now', 'sendWeeklyReportNow'))
    .addSeparator()
    .addItem('Set Up / Upgrade Tool', 'setUpOrUpgradeTool')
    .addItem('About', 'showAbout')
    .addToUi();
}
