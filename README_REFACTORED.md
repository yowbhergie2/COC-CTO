# COC Tracking System - Refactored Structure

## Overview

The original `Code.gs` (7,363 lines, 262KB) has been refactored into **9 organized files** totaling 142 functions. Each file groups related functionality for better maintainability and clarity.

---

## File Structure

### 1. **Code.gs** (6 KB)
**Purpose**: Main entry point with configuration and initialization

**Contents**:
- Configuration constants (DATABASE_ID, column mappings, status constants)
- Column mapping objects: DETAIL_COLS, RECORD_COLS, CERT_COLS, EMP_COLS, LEDGER_COLS
- Status constants: STATUS_PENDING, STATUS_ACTIVE, STATUS_USED, etc.
- Transaction type constants: TR_TYPE_EARNED, TR_TYPE_USED, etc.
- `onOpen()` - Main menu initialization function

**Functions**: 1
- `onOpen()`

---

### 2. **UIFunctions.gs** (4.5 KB)
**Purpose**: User interface and menu functions

**Contents**: All functions that display UI elements or handle navigation

**Functions**: 12
- `showCTOApplicationsManager()`
- `showSettings()`
- `showDashboard()`
- `showMonthlyCOCEntry()`
- `showCTORecordForm()`
- `showEmployeeLedger()`
- `showEmployeeManager()`
- `showHistoricalimport()`
- `showReports()`
- `showHolidayManager()`
- `navigateToPage()`

**Dependencies**: None (these functions are called by onOpen and menu items)

---

### 3. **DataFunctions.gs** (35 KB)
**Purpose**: Core data operations and spreadsheet interactions

**Contents**: Functions for CRUD operations on sheets, employee management, and data retrieval

**Functions**: 22
- `getDatabase()`
- `getSheetDataNoHeader()`
- `ensureCOCBalanceDetailSheet()`
- `ensureLedgerSheet()`
- `getEmployeeById()`
- `getEmployeeDetails()`
- `listEmployees()`
- `addEmployee()`
- `addEmployeeWithFIFO()`
- `updateEmployee()`
- `getLedgerForEmployee()`
- `getReportData()`
- `listHolidays()`
- `addHoliday()`
- `updateHoliday()`
- `deleteHoliday()`
- `getDashboardStats()`
- `getRecentActivities()`
- `addCOCToBalanceDetail()`
- `getCurrentCOCBalance()`
- `getCurrentCOCBalanceFromDetail()`
- `getCOCBalanceBreakdown()`

**Dependencies**:
- Uses: HelperFunctions (formatDate, generateIds, getScriptTimeZone)
- Used by: API, BusinessLogic, Certificates

---

### 4. **HelperFunctions.gs** (13 KB)
**Purpose**: Utility functions and helpers

**Contents**: Formatters, generators, validators, and utility functions

**Functions**: 20
- `getScriptTimeZone()`
- `formatDate()`
- `formatLongDate()`
- `formatInclusiveDates()`
- `getCurrentUserEmail()`
- `generateUniqueId()`
- `generateEmployeeId()`
- `generateRecordId()`
- `generateCTOId()`
- `generateLedgerId()`
- `generateCOCDetailEntryId()`
- `generateLedgerEntryId()`
- `padNumber()`
- `getDayType()`
- `getDayTypeEnhanced()`
- `getSettings()`
- `getDropdownOptions()`
- `parseMonthYear()`
- `timeToMinutes()`
- `calculateBalanceFallback()`

**Dependencies**:
- Uses: DataFunctions (getDatabase)
- Used by: All other files

---

### 5. **BusinessLogic.gs** (50 KB)
**Purpose**: Business rules and processing logic

**Contents**: Core business logic for COC/CTO processing, FIFO calculations, validations

**Functions**: 16
- `recordCOCEntries()`
- `recordCTOApplication()`
- `recordCTOApplicationWithFIFO()`
- `cancelCTOApplication()`
- `cancelCOCRecord()`
- `consumeCOCWithFIFO()`
- `deductCOCHoursFIFO()`
- `restoreCOCHoursFIFO()`
- `checkAndExpireCOC()`
- `checkExpiredCOC()`
- `checkMonthlyLimitForMonth()`
- `checkTotalBalanceLimitForEmployee()`
- `validateCTOUpdate()`
- `checkEmployeeFIFO()`
- `calculateOvertimeForDate()`
- `calculateCOCForDate()`

**Dependencies**:
- Uses: DataFunctions, HelperFunctions
- Used by: API, Certificates

---

### 6. **Certificates.gs** (18 KB)
**Purpose**: Certificate generation and management

**Contents**: Functions for generating and managing COC certificates

**Functions**: 4
- `calculateCertificateExpiration()`
- `generateMonthlyCOCCertificate()`
- `generateCOCCertificate()`
- `generateCertificateDocument()`

**Dependencies**:
- Uses: DataFunctions, HelperFunctions, BusinessLogic
- Used by: API

---

### 7. **API.gs** (96 KB)
**Purpose**: External API layer

**Contents**: All `api*` functions that serve as the public interface

**Functions**: 50
- Employee APIs: `apiAddEmployee()`, `apiGetEmployee()`, `apiUpdateEmployee()`, `apiListEmployees()`, etc.
- Balance APIs: `apiGetBalance()`, `apiGetBalanceFromDetail()`, `apiGetCOCBalanceBreakdown()`
- COC APIs: `apiRecordCOC()`, `apiListCOCRecordsForMonth()`, `apiGenerateMonthlyCOCCertificate()`
- CTO APIs: `apiRecordCTO()`, `apiGetEmployeeCTOApplications()`, `apiUpdateCTOApplication()`, `apiCancelCTOApplication()`
- Holiday APIs: `apiListHolidays()`, `apiAddHoliday()`, `apiUpdateHoliday()`, `apiDeleteHoliday()`
- Report APIs: `apiGetReport()`, `apiGetDashboardStats()`, `apiGetRecentActivities()`
- FIFO APIs: `apiFIFOIntegrityCheck()`, `apiFIFOFix()`, `apiFIFOEmployeeReport()`
- Utility APIs: `apiGetDayType()`, `apiCalculateOvertime()`, `apiGetDropdownOptions()`
- And 30+ more...

**Dependencies**:
- Uses: All other function files
- Used by: External callers (web apps, triggers, etc.)

---

### 8. **MigrationFunctions.gs** (8.5 KB)
**Purpose**: Data migration utilities

**Contents**: Functions for migrating data between schema versions

**Functions**: 3
- `migrateCOCRecordsMonthYear()`
- `migrateExistingInitialBalances()`
- `runCOCRecordsMigration()`

**Dependencies**:
- Uses: DataFunctions, HelperFunctions
- Used by: Manual execution for data migrations

---

### 9. **DebugAndTest.gs** (14 KB)
**Purpose**: Testing and diagnostic functions

**Contents**: Test functions and diagnostic utilities

**Functions**: 15
- `debugCOCRecords()`
- `debugMariaOctober2025()`
- `debugAllThreeFunctions()`
- `testGetBalance()`
- `testApiGetLedgerDirect()`
- `testGetEmployeeCTOApplications()`
- `testGetAllCTOApplications()`
- `testCancelCTOApplication()`
- `testUpdateCTOApplication()`
- `testAllSerializations()`
- `testFIFOIntegrityCheck()`
- `diagnosticCheckLedgerForEMP002()`
- `diagnosticCheckBalanceForEMP002()`
- `diagnosticCheckApiGetLedger()`
- `runAllDiagnostics()`

**Dependencies**:
- Uses: All other files
- Used by: Manual testing and debugging

---

### 10. **Migration.gs** (existing file, 683 lines)
**Purpose**: Specific COC balance migration scripts

**Contents**: Step-by-step migration functions for COC balance detail
- `runMigrationStep1_Initialize()`
- `runMigrationStep2_MigrateInitialBalances()`
- `runMigrationStep3_MigrateCOCRecords()`
- `verifyMigrationForEmployee()`
- `runCompleteMigration()`
- And more...

**Note**: This file already exists and contains specific migration workflows. Keep as-is.

---

## Dependency Graph

```
Code.gs (Config + onOpen)
    ↓
UIFunctions.gs → (displays UI)
    ↓
API.gs (External Interface)
    ↓
    ├─→ DataFunctions.gs (CRUD Operations)
    ├─→ BusinessLogic.gs (Processing)
    ├─→ Certificates.gs (Certificate Generation)
    ├─→ MigrationFunctions.gs (Migrations)
    └─→ DebugAndTest.gs (Testing)
         ↓
    HelperFunctions.gs (Used by all)
```

---

## Global Variables/Constants

The following are defined in **Code.gs** and available to all files:

### Constants
- `DATABASE_ID` - Spreadsheet ID
- `DETAIL_COLS` - COC_Balance_Detail column mapping
- `RECORD_COLS` - COC_Records column mapping
- `CERT_COLS` - COC_Certificates column mapping
- `EMP_COLS` - Employees column mapping
- `LEDGER_COLS` - COC_Ledger column mapping
- `STATUS_*` - Status constants (PENDING, ACTIVE, USED, EXPIRED, CANCELLED)
- `TR_TYPE_*` - Transaction type constants (EARNED, USED, EXPIRED, ADJUSTED)

---

## How to Use

### Option 1: Import All Files into Google Apps Script
1. Create a new Google Apps Script project
2. Delete the default `Code.gs`
3. Add all 9 .gs files to the project:
   - File → New → Script file (for each)
   - Copy and paste the contents

### Option 2: Keep Existing Migration.gs
1. If you already have `Migration.gs`, keep it
2. Add the other 9 files
3. Total: 10 files

### Option 3: Namespace Approach (Optional)
If you want to prevent naming conflicts, you can wrap functions in namespace objects:

```javascript
// In HelperFunctions.gs
const Helpers = {
  formatDate: function(date) { ... },
  generateUniqueId: function(prefix) { ... },
  // ...
};

// Usage in other files
const formattedDate = Helpers.formatDate(new Date());
```

---

## Notes

### 1. No Logic Changes
- All functions maintain their original logic
- All comments and variable names are preserved
- Only organizational changes made

### 2. Function Discovery
- Google Apps Script automatically discovers all functions across files
- Functions can call each other regardless of which file they're in
- The `onOpen()` function in Code.gs will still work

### 3. Execution Order
- Google Apps Script doesn't have a specific file loading order
- All files are concatenated at runtime
- Ensure global variables (from Code.gs) are defined before use

### 4. Testing Recommendation
- Test the `onOpen()` function first to ensure menu loads
- Test a few API functions to verify cross-file function calls work
- Run diagnostic functions from DebugAndTest.gs to verify system integrity

### 5. File Size Reduction
- Original: 1 file × 262 KB = 262 KB total
- Refactored: 9 files averaging 28 KB each
- Easier to navigate and maintain

---

## Summary Statistics

| File | Size | Functions | Purpose |
|------|------|-----------|---------|
| Code.gs | 6 KB | 1 | Configuration + Entry Point |
| UIFunctions.gs | 4.5 KB | 12 | User Interface |
| DataFunctions.gs | 35 KB | 22 | Data Operations |
| HelperFunctions.gs | 13 KB | 20 | Utilities |
| BusinessLogic.gs | 50 KB | 16 | Business Rules |
| Certificates.gs | 18 KB | 4 | Certificate Management |
| API.gs | 96 KB | 50 | External API Layer |
| MigrationFunctions.gs | 8.5 KB | 3 | Data Migrations |
| DebugAndTest.gs | 14 KB | 15 | Testing & Diagnostics |
| **TOTAL** | **245 KB** | **143** | - |

---

## Questions?

If you encounter any issues:
1. Check that all files are uploaded to Google Apps Script
2. Verify that function names haven't been accidentally changed
3. Check the Apps Script execution log for errors
4. Use the diagnostic functions in DebugAndTest.gs to verify system state

---

**Last Updated**: 2025-10-28
**Refactored From**: Code.gs (7,363 lines, 142 functions)
