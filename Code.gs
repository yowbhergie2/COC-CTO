// -----------------------------------------------------------------------------
// Configuration
// -----------------------------------------------------------------------------

const DATABASE_ID = '1LIQRnQb7lL-6hdSpsOwq-XVwRM6Go6u4q6RkSB2ZCXo';

// This is the CORRECT column mapping for COC_Balance_Detail
const DETAIL_COLS = {
  ENTRY_ID: 0,            // Entry ID
  EMPLOYEE_ID: 1,         // Employee ID
  EMPLOYEE_NAME: 2,       // Employee Name
  CERTIFICATE_ID: 3,      // Certificate ID (NEW - links to monthly certificate)
  RECORD_ID: 4,           // Record ID (from COC_Records)
  MONTH_YEAR: 5,          // Month-Year (e.g., "2025-10")
  DATE_EARNED: 6,         // Date Rendered/Earned
  DAY_TYPE: 7,            // Day Type (Weekday, Weekend, Regular Holiday, etc.)
  HOURS_EARNED: 8,        // Hours Earned
  HOURS_USED: 9,          // Hours Used (for FIFO tracking)
  HOURS_REMAINING: 10,    // Hours Remaining
  CERTIFICATE_ISSUE_DATE: 11, // Certificate Issue Date (NEW - when certificate was issued)
  EXPIRATION_DATE: 12,    // Expiration Date (Certificate Issue Date + 1 year - 1 day)
  STATUS: 13,             // Status (Active, Used, Expired, Cancelled)
  DATE_CREATED: 14,       // Date Created
  CREATED_BY: 15,         // Created By
  LAST_UPDATED: 16,       // Last Updated
  NOTES: 17               // Notes/Remarks
};

// COC_Records sheet column mapping (22 columns)
const RECORD_COLS = {
  RECORD_ID: 0,           // Record ID (unique identifier)
  EMPLOYEE_ID: 1,         // Employee ID
  EMPLOYEE_NAME: 2,       // Employee Name
  MONTH_YEAR: 3,          // Month-Year (e.g., "2025-10")
  DATE_RENDERED: 4,       // Date Rendered (date COC was earned)
  DAY_TYPE: 5,            // Day Type (Weekday, Weekend, Holiday, etc.)
  AM_IN: 6,               // AM In time
  AM_OUT: 7,              // AM Out time
  PM_IN: 8,               // PM In time
  PM_OUT: 9,              // PM Out time
  HOURS_WORKED: 10,       // Hours Worked
  MULTIPLIER: 11,         // Multiplier (based on day type)
  COC_EARNED: 12,         // COC Earned (hours)
  CERTIFICATE_ID: 13,     // Certificate ID (linked to COC_Certificates)
  DATE_RECORDED: 14,      // Date Recorded
  EXPIRATION_DATE: 15,    // Expiration Date
  STATUS: 16,             // Status (Pending, Active, Used, Expired, Cancelled)
  APPROVED_BY: 17,        // Approved By
  APPROVED_DATE: 18,      // Approved Date
  CREATED_BY: 19,         // Created By
  LAST_MODIFIED: 20,      // Last Modified
  MODIFIED_BY: 21         // Modified By
};

// COC_Certificates sheet column mapping (13 columns)
const CERT_COLS = {
  CERTIFICATE_ID: 0,      // Certificate ID (unique identifier)
  EMPLOYEE_ID: 1,         // Employee ID
  EMPLOYEE_NAME: 2,       // Employee Name
  MONTH_YEAR: 3,          // Month-Year (e.g., "2025-10")
  TOTAL_COC_EARNED: 4,    // Total COC Earned (hours)
  NUMBER_OF_RECORDS: 5,   // Number of Records (count)
  ISSUE_DATE: 6,          // Issue Date
  EXPIRATION_DATE: 7,     // Expiration Date
  CERTIFICATE_URL: 8,     // Certificate URL (Google Doc)
  PDF_URL: 9,             // PDF URL (exported PDF)
  STATUS: 10,             // Status (Active, Cancelled, etc.)
  CREATED_DATE: 11,       // Created Date
  CREATED_BY: 12          // Created By
};

// Employees sheet column mapping
const EMP_COLS = {
  EMPLOYEE_ID: 0,         // Employee ID
  FIRST_NAME: 1,          // First Name
  MIDDLE_INITIAL: 2,      // Middle Initial
  LAST_NAME: 3,           // Last Name
  SUFFIX: 4,              // Suffix (Jr., Sr., etc.)
  POSITION: 5,            // Position
  OFFICE: 6,              // Office
  STATUS: 7               // Status (Active, Inactive, etc.)
};

// COC_Ledger sheet column mapping
const LEDGER_COLS = {
  LEDGER_ID: 0,           // Ledger ID (unique identifier)
  EMPLOYEE_ID: 1,         // Employee ID
  EMPLOYEE_NAME: 2,       // Employee Name
  TRANSACTION_DATE: 3,    // Transaction Date
  TRANSACTION_TYPE: 4,    // Transaction Type (Earned, Used, Expired, Adjusted)
  REFERENCE_ID: 5,        // Reference ID (Certificate ID, CTO ID, etc.)
  BALANCE_BEFORE: 6,      // Balance Before
  COC_EARNED: 7,          // COC Earned (hours)
  CTO_USED: 8,            // CTO Used (hours)
  COC_EXPIRED: 9,         // COC Expired (hours)
  BALANCE_ADJUSTMENT: 10, // Balance Adjustment (hours)
  BALANCE_AFTER: 11,      // Balance After
  MONTH_YEAR_EARNED: 12,  // Month-Year Earned (e.g., "2025-10")
  EXPIRATION_DATE: 13,    // Expiration Date
  PROCESSED_BY: 14,       // Processed By
  PROCESSED_DATE: 15,     // Processed Date
  REMARKS: 16             // Remarks/Notes
};

// Status constants
const STATUS_PENDING = 'Pending';
const STATUS_ACTIVE = 'Active';
const STATUS_USED = 'Used';
const STATUS_EXPIRED = 'Expired';
const STATUS_CANCELLED = 'Cancelled';

// Transaction type constants
const TR_TYPE_EARNED = 'Earned';
const TR_TYPE_USED = 'Used';
const TR_TYPE_EXPIRED = 'Expired';
const TR_TYPE_ADJUSTED = 'Adjusted';



// -----------------------------------------------------------------------------
// Main Entry Point
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


