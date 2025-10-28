// -----------------------------------------------------------------------------
// UIFunctions.gs
// -----------------------------------------------------------------------------

/**
 * Update your existing onOpen function to include the new menu item
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('COC/CTO System')
    .addItem('Dashboard', 'showDashboard')
    .addSeparator()
    .addItem('Record COC Earned', 'showMonthlyCOCEntry')
    .addItem('Record CTO Application', 'showCTORecordForm')
    .addItem('CTO Applications Manager', 'showCTOApplicationsManager')
    .addSeparator()
    .addItem('Employee Ledger', 'showEmployeeLedger')
    .addItem('Employee Manager', 'showEmployeeManager')
    .addItem('Holiday Manager', 'showHolidayManager')
    .addItem('Import', 'showHistoricalimport')
    .addSeparator()
    .addItem('Settings', 'showSettings')
    .addItem('Reports', 'showReports')
    .addSeparator()
    .addSubMenu(ui.createMenu('Admin Tools')
      .addItem('Run Data Migration (MONTH_YEAR)', 'runCOCRecordsMigration')
      .addItem('Debug COC Records', 'debugMariaOctober2025'))
    .addToUi();
}


/**
 * Opens the CTO Applications Manager
 */
function showCTOApplicationsManager() {
  const html = HtmlService.createHtmlOutputFromFile('CTOApplicationsManager')
    .setWidth(1400)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'CTO Applications Manager');
}


function showSettings() {
  const html = HtmlService.createHtmlOutputFromFile('Settings')
    .setWidth(800)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'System Settings');
}


/**
 * The following functions open their corresponding HTML files as modal
 * dialogs. The widths/heights can be adjusted as needed to fit the contents.
 */
function showDashboard() {
  const html = HtmlService.createHtmlOutputFromFile('Dashboard.html')
    .setWidth(1200)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Dashboard');
}


function showMonthlyCOCEntry() {
  const html = HtmlService.createHtmlOutputFromFile('MonthlyCOCEntry')
    .setWidth(1200)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Monthly COC Entry');
}


function showCTORecordForm() {
  const html = HtmlService.createHtmlOutputFromFile('CTORecordForm')
    .setWidth(1000)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'CTO Application');
}


function showEmployeeLedger() {
  const html = HtmlService.createHtmlOutputFromFile('EmployeeLedger')
    .setWidth(1200)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Employee Ledger');
}


function showEmployeeManager() {
  const html = HtmlService.createHtmlOutputFromFile('EmployeeManager')
    .setWidth(1200)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Employee Manager');
}


function showHistoricalimport() {
  const html = HtmlService.createHtmlOutputFromFile('Historicalimport')
    .setWidth(1200)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Historical Import');
}


function showReports() {
  const html = HtmlService.createHtmlOutputFromFile('Reports')
    .setWidth(1200)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Reports');
}


/**
 * Opens the Holiday Manager modal dialog. This interface allows HR staff
 * to add, edit and delete holidays and no‑work days. Holidays are stored
 * in a dedicated sheet and used to classify overtime dates correctly.
 */
function showHolidayManager() {
  const html = HtmlService.createHtmlOutputFromFile('HolidayManager')
    .setWidth(1000)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Holiday Manager');
}


function navigateToPage(pageName) {
  // We use the 'show' functions that already exist in your Code.gs
  try {
    if (pageName === 'MonthlyCOCEntry') {
      showMonthlyCOCEntry();
    } else if (pageName === 'CTORecordForm') {
      showCTORecordForm();
    } else if (pageName === 'EmployeeLedger') {
      showEmployeeLedger();
    } else if (pageName === 'EmployeeManager') {
      showEmployeeManager();
    } else if (pageName === 'Reports') {
      showReports();
    } else if (pageName === 'HolidayManager') {
      showHolidayManager();
    } else {
      // Fallback in case of an unknown pageName
      showDashboard();
    }
  } catch (e) {
    Logger.log('Error navigating to page: ' + pageName + '. Error: ' + e.message);
    // If it fails, just show the dashboard again
    showDashboard();
  }
}


