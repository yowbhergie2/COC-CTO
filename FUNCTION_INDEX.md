# Function Index - Quick Reference

This index shows which file each function is located in.

## Code.gs (1 function)
- `onOpen()` - Main menu initialization

## UIFunctions.gs (12 functions)
- `navigateToPage(pageName)`
- `showCTOApplicationsManager()`
- `showCTORecordForm()`
- `showDashboard()`
- `showEmployeeLedger()`
- `showEmployeeManager()`
- `showHistoricalimport()`
- `showHolidayManager()`
- `showMonthlyCOCEntry()`
- `showReports()`
- `showSettings()`

## DataFunctions.gs (22 functions)
- `addCOCToBalanceDetail(employeeId, employeeName, recordId, dateEarned, hoursEarned)`
- `addEmployee(data)`
- `addEmployeeWithFIFO(data)`
- `addHoliday(date, type, description)`
- `deleteHoliday(rowNumber)`
- `ensureCOCBalanceDetailSheet()`
- `ensureLedgerSheet()`
- `getCOCBalanceBreakdown(employeeId)`
- `getCurrentCOCBalance(employeeId)`
- `getCurrentCOCBalanceFromDetail(employeeId)`
- `getDashboardStats()`
- `getDatabase()`
- `getEmployeeById(employeeId)`
- `getEmployeeDetails(employeeId)`
- `getLedgerForEmployee(employeeId)`
- `getRecentActivities(limit)`
- `getReportData(type, startDate, endDate)`
- `getSheetDataNoHeader(sheetName)`
- `listEmployees(includeInactive)`
- `listHolidays()`
- `updateEmployee(employeeId, data)`
- `updateHoliday(rowNumber, date, type, description)`

## HelperFunctions.gs (20 functions)
- `calculateBalanceFallback(employeeId)`
- `formatDate(date)`
- `formatInclusiveDates(dateArray)`
- `formatLongDate(date)`
- `generateCOCDetailEntryId()`
- `generateCTOId()`
- `generateEmployeeId()`
- `generateLedgerEntryId()`
- `generateRecordId()`
- `generateUniqueId(prefix)`
- `getCurrentUserEmail()`
- `getDayType(date)`
- `getDayTypeEnhanced(date)`
- `getDropdownOptions()`
- `getScriptTimeZone()`
- `getSettings()`
- `padNumber(num, length)`
- `parseMonthYear(monthYear)`
- `timeToMinutes(timeStr)`

## BusinessLogic.gs (16 functions)
- `calculateCOCForDate(date, hoursWorked, timeIn, timeOut)`
- `calculateOvertimeForDate(date, amIn, amOut, pmIn, pmOut)`
- `cancelCOCRecord(recordId, remarks)`
- `cancelCTOApplication(ctoId, remarks)`
- `checkAndExpireCOC()`
- `checkEmployeeFIFO(employeeId, detailData, ctoData)`
- `checkExpiredCOC()`
- `checkMonthlyLimitForMonth(employeeId, monthYear, hoursToAdd)`
- `checkTotalBalanceLimitForEmployee(employeeId, hoursToAdd)`
- `consumeCOCWithFIFO(employeeId, hoursToConsume, reference)`
- `deductCOCHoursFIFO(employeeId, hoursToDeduct, referenceId)` (appears twice in original)
- `recordCOCEntries(employeeId, month, year, entries)`
- `recordCTOApplication(employeeId, hours, startDate, endDate, remarks)`
- `recordCTOApplicationWithFIFO(employeeId, hours, startDate, endDate, remarks)`
- `restoreCOCHoursFIFO(employeeId, hoursToRestore, referenceId)`
- `validateCTOUpdate(application, newHours, newStartDateStr, newEndDateStr)`

## Certificates.gs (4 functions)
- `calculateCertificateExpiration(issueDate)`
- `generateCertificateDocument(certificateId, empDetails, records, issueDate, expirationDate)`
- `generateCOCCertificate(recordId)`
- `generateMonthlyCOCCertificate(employeeId, monthYear)`

## API.gs (50 functions)
- `apiAddEmployee(data)`
- `apiAddHoliday(date, type, description, halfdayTime, suspensionTime, remarks)`
- `apiCalculateOvertime(year, month, day, amIn, amOut, pmIn, pmOut)`
- `apiCalculateOvertimeForDate(year, month, day, amIn, amOut, pmIn, pmOut)`
- `apiCancelCOC(recordId, remarks)`
- `apiCancelCTO(ctoId, remarks)`
- `apiCancelCTOApplication(applicationId)` (appears twice in original)
- `apiCheckAndExpireCOC()`
- `apiDeleteCOCRecord(recordId, reason)`
- `apiDeleteHoliday(rowNumber)`
- `apiFIFOEmployeeReport(employeeId)`
- `apiFIFOFix(employeeId, dryRun)`
- `apiFIFOIntegrityCheck(employeeId)`
- `apiGenerateCOCCertificate(recordId)`
- `apiGenerateMonthlyCOCCertificate(employeeId, month, year, issueDateString)`
- `apiGetAllCTOApplications()`
- `apiGetBalance(empId)`
- `apiGetBalanceFromDetail(empId)`
- `apiGetCOCBalanceBreakdown(employeeId)`
- `apiGetDashboardStats()`
- `apiGetDayType(year, month, day)`
- `apiGetDropdownOptions()`
- `apiGetEmployee(employeeId)`
- `apiGetEmployeeCOCStats(employeeId, month, year)`
- `apiGetEmployeeCTOApplications(employeeId)`
- `apiGetEmployeesWithExpiringCOC()`
- `apiGetHistoricalImports()`
- `apiGetLedger(employeeId)`
- `apiGetMonthlyCertificate(employeeId, month, year)`
- `apiGetRecentActivities(limit)`
- `apiGetReport(type, startDate, endDate)`
- `apiGetSignatories()`
- `apiImportHistoricalCOC(data)`
- `apiImportHistoricalCOCBatch(csvData)`
- `apiInitializeCOCBalanceDetail()`
- `apiListCOCRecordsForMonth(employeeId, month, year)`
- `apiListCOCRecordsForMonth_OLD(employeeId, month, year)`
- `apiListEmployees(includeInactive)`
- `apiListEmployeesForDropdown()`
- `apiListHolidays()`
- `apiMigrateExistingInitialBalances()`
- `apiRecordCOC(employeeId, month, year, entries)`
- `apiRecordCOCWithValidation_OLD(employeeId, month, year, entries)`
- `apiRecordCOC_OLD_WRAPPER(employeeId, month, year, entries)`
- `apiRecordCTO(employeeId, hours, startDate, endDate, remarks)`
- `apiRecordCTOWithFIFO(employeeId, hours, startDate, endDate, remarks)`
- `apiSaveSignatories(signatories)`
- `apiUpdateCTOApplication(applicationId, newHours, newStartDate, newEndDate, newRemarks)`
- `apiUpdateEmployee(employeeId, data)`
- `apiUpdateHoliday(rowNumber, date, type, description, halfdayTime, suspensionTime, remarks)`

## MigrationFunctions.gs (3 functions)
- `migrateCOCRecordsMonthYear()`
- `migrateExistingInitialBalances()`
- `runCOCRecordsMigration()`

## DebugAndTest.gs (15 functions)
- `debugAllThreeFunctions()`
- `debugCOCRecords(employeeId, month, year)`
- `debugMariaOctober2025()`
- `diagnosticCheckApiGetLedger()`
- `diagnosticCheckBalanceForEMP002()`
- `diagnosticCheckLedgerForEMP002()`
- `runAllDiagnostics()`
- `testAllSerializations()`
- `testApiGetLedgerDirect()`
- `testCancelCTOApplication()`
- `testFIFOIntegrityCheck()`
- `testGetAllCTOApplications()`
- `testGetBalance()`
- `testGetEmployeeCTOApplications()` (appears twice in original)
- `testUpdateCTOApplication()`

---

**Total**: 143 functions across 9 files
