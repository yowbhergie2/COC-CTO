// -----------------------------------------------------------------------------
// API.gs
// -----------------------------------------------------------------------------

/**
 * API: Add employee with FIFO tracking
 * This replaces apiAddEmployee
 */
function apiAddEmployee(data) {
  return addEmployeeWithFIFO(data);
}


/**
 * API: Add a new holiday
 * @param {Date} date The date of the holiday
 * @param {string} type The type of holiday
 * @param {string} description Optional description
 * @param {string} halfdayTime Optional time for half-day holidays
 * @param {string} suspensionTime Optional time for work suspension
 * @param {string} remarks Optional additional remarks
 * @return {Object} Success result
 */
function apiAddHoliday(date, type, description, halfdayTime, suspensionTime, remarks) {
  try {
    const db = getDatabase();
    let sheet = db.getSheetByName('Holidays');

    if (!sheet) {
      sheet = db.insertSheet('Holidays');
      sheet.getRange(1, 1, 1, 6).setValues([[
        'Date', 'Type', 'Description', 'Half-day Start Time', 'Suspension Time', 'Remarks'
      ]]);
      sheet.setFrozenRows(1);
    }

    // Validate inputs
    if (!date || !(date instanceof Date)) {
      throw new Error('Invalid date provided');
    }

    if (!type) {
      throw new Error('Holiday type is required');
    }

    // Check for duplicate date
    const existingData = sheet.getDataRange().getValues();
    const TIME_ZONE = getScriptTimeZone();
    const dateStr = Utilities.formatDate(new Date(date), TIME_ZONE, 'yyyy-MM-dd');
    
    for (let i = 1; i < existingData.length; i++) {
      const existingDate = existingData[i][0];
      if (existingDate && Utilities.formatDate(new Date(existingDate), TIME_ZONE, 'yyyy-MM-dd') === dateStr) {
        throw new Error('A holiday already exists for this date. Please edit the existing entry instead.');
      }
    }

    // Add new row
    const newRow = [
      new Date(date),
      type,
      description || '',
      halfdayTime || '',
      suspensionTime || '',
      remarks || ''
    ];

    sheet.appendRow(newRow);

    Logger.log('Holiday added successfully: ' + dateStr + ' - ' + type);
    return { success: true, message: 'Holiday added successfully' };

  } catch (error) {
    Logger.log('ERROR in apiAddHoliday: ' + error.message + '\nStack: ' + error.stack);
    throw new Error('Failed to add holiday: ' + error.message);
  }
}


function apiCalculateOvertime(year, month, day, amIn, amOut, pmIn, pmOut) {
  const date = new Date(year, month - 1, day);
  return calculateOvertimeForDate(date, amIn, amOut, pmIn, pmOut);
}


/**
 * API wrapper for calculateOvertimeForDate
 * This function is called from the frontend to automatically detect day type
 * and calculate COC when times are entered
 *
 * @param {number} year - The year (e.g., 2025)
 * @param {number} month - The month (1-12)
 * @param {number} day - The day of the month (1-31)
 * @param {string} amIn - AM in time (HH:mm format)
 * @param {string} amOut - AM out time (HH:mm format)
 * @param {string} pmIn - PM in time (HH:mm format)
 * @param {string} pmOut - PM out time (HH:mm format)
 * @return {Object} Result with dayType, hoursWorked, multiplier, cocEarned
 */
function apiCalculateOvertimeForDate(year, month, day, amIn, amOut, pmIn, pmOut) {
  const date = new Date(year, month - 1, day);
  return calculateOvertimeForDate(date, amIn, amOut, pmIn, pmOut);
}


function apiCancelCTO(ctoId, remarks) {
  return cancelCTOApplication(ctoId, remarks);
}


function apiCancelCOC(recordId, remarks) {
  return cancelCOCRecord(recordId, remarks);
}


/**
 * Cancel a CTO application and restore the COC hours using FIFO
 */
function apiCancelCTOApplication(applicationId) {
  try {
    Logger.log('=== Cancelling CTO Application: ' + applicationId + ' ===');
    
    const db = getDatabase();
    const ctoSheet = db.getSheetByName('CTO_Applications');
    const ledgerSheet = db.getSheetByName('COC_Ledger');
    const detailSheet = db.getSheetByName('COC_Balance_Detail');
    
    if (!ctoSheet) throw new Error('CTO_Applications sheet not found');
    if (!ledgerSheet) throw new Error('COC_Ledger sheet not found');
    if (!detailSheet) throw new Error('COC_Balance_Detail sheet not found');
    
    // Find the application
    const ctoData = ctoSheet.getDataRange().getValues();
    let appRow = -1;
    let application = null;
    
    for (let i = 1; i < ctoData.length; i++) {
      if (ctoData[i][0] === applicationId) {
        appRow = i + 1; // Sheet row (1-indexed)
        application = {
          id: ctoData[i][0],
          employeeId: ctoData[i][1],
          employeeName: ctoData[i][2],
          hours: parseFloat(ctoData[i][4]) || 0,
          status: ctoData[i][10]
        };
        break;
      }
    }
    
    if (!application) {
      throw new Error('CTO application not found: ' + applicationId);
    }
    
    // Check if already cancelled
    if (application.status === 'Cancelled') {
      return {
        success: false,
        message: 'This application is already cancelled'
      };
    }
    
    // --- START MODIFICATION ---
    // Allow cancelling 'Pending' or 'Approved' applications
    if (application.status !== 'Pending' && application.status !== 'Approved') {
      return {
        success: false,
        // Updated error message
        message: 'Only Pending or Approved applications can be cancelled. This application is ' + application.status
      };
    }
    // --- END MODIFICATION ---
    
    Logger.log('Application found: ' + application.employeeName + ', Status: ' + application.status + ', Hours: ' + application.hours);
    
    // Update CTO application status
    ctoSheet.getRange(appRow, 11).setValue('Cancelled'); // Column K (Status)
    ctoSheet.getRange(appRow, 12).setValue('Cancelled by user on ' + new Date()); // Column L (Remarks)
    
    Logger.log('✓ CTO application status updated to Cancelled');
    
    // --- START MODIFICATION: Restore hours ONLY if it was Approved ---
    // If it was Pending, no hours were deducted yet, so no need to restore.
    let restoredHours = 0; // Initialize restoredHours
    if (application.status === 'Approved') {
        Logger.log('Restoring hours because status was Approved.');
        // Find and reverse the ledger entry
        const ledgerData = ledgerSheet.getDataRange().getValues();
        let ledgerRow = -1;

        for (let i = ledgerData.length - 1; i >= 1; i--) {
            if (ledgerData[i][1] === application.employeeId &&
                ledgerData[i][5] === applicationId &&
                ledgerData[i][4] === 'CTO Used') {
                ledgerRow = i + 1;
                break;
            }
        }

        if (ledgerRow > 0) {
            // Mark ledger entry as cancelled
            ledgerSheet.getRange(ledgerRow, 5).setValue('CTO Cancelled (Approved)'); // Clarify transaction type
            ledgerSheet.getRange(ledgerRow, 13).setValue('Cancelled on ' + new Date()); // Remarks
            Logger.log('✓ Original CTO Used ledger entry marked as cancelled');
        } else {
             Logger.log('⚠ Original CTO Used ledger entry not found, proceeding with balance restoration.');
        }

        // Restore the COC hours in detail sheet (reverse FIFO deduction)
        try {
          // --- Call the separate restore function ---
          restoredHours = restoreCOCHoursFIFO(application.employeeId, application.hours, applicationId);
          if (restoredHours > 0) {
             Logger.log('✓ Restored ' + restoredHours + ' hours to COC balance via FIFO.');
          } else {
             Logger.log('⚠ No hours were restored via FIFO. This might be okay if the deduction logic changed or was manual.');
          }
        } catch (restoreError) {
           Logger.log('ERROR during FIFO restore: ' + restoreError.message);
           // Even if restore fails, continue with ledger update to reflect cancellation attempt
        }

        // Add a reversal entry to the ledger ONLY if hours were deducted (status was Approved)
        const newBalance = getCurrentCOCBalance(application.employeeId); // Recalculate balance AFTER potential restore
        const ledgerEntry = [
          generateLedgerEntryId(),
          application.employeeId,
          application.employeeName,
          new Date(),
          'CTO Cancelled (Restore)', // Clearer transaction type
          applicationId,
          application.hours, // COC Earned (restored)
          0, // CTO Used
          newBalance,
          '', // Month-Year Earned
          '', // Expiration Date
          Session.getActiveUser().getEmail(),
          'CTO application (Approved) cancelled, ' + application.hours + ' hrs restored'
        ];
        ledgerSheet.appendRow(ledgerEntry);
        Logger.log('✓ Added restoration entry to ledger');

    } else {
       // If status was 'Pending', just log that no restoration needed
       Logger.log('Status was Pending, no hours to restore.');
       // Optionally add a simpler ledger entry just marking cancellation
        const newBalance = getCurrentCOCBalance(application.employeeId); // Get current balance
        const ledgerEntry = [
          generateLedgerEntryId(),
          application.employeeId,
          application.employeeName,
          new Date(),
          'CTO Cancelled (Pending)', // Clearer transaction type
          applicationId,
          0, // COC Earned
          0, // CTO Used
          newBalance, // Balance remains unchanged
          '', // Month-Year Earned
          '', // Expiration Date
          Session.getActiveUser().getEmail(),
          'CTO application (Pending) cancelled before approval.'
        ];
        ledgerSheet.appendRow(ledgerEntry);
        Logger.log('✓ Added ledger entry for Pending cancellation.');
    }
    // --- END MODIFICATION ---
    
    Logger.log('=== CTO Cancellation Complete ===');
    
    return {
      success: true,
      // --- MODIFICATION: Updated success message based on status ---
      message: application.status === 'Approved' 
         ? 'CTO application cancelled successfully. ' + application.hours + ' hours have been restored to the employee\'s COC balance.'
         : 'Pending CTO application cancelled successfully.'
      // --- END MODIFICATION ---
    };
    
  } catch (error) {
    Logger.log('ERROR in apiCancelCTOApplication: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    
    return {
      success: false,
      message: 'Failed to cancel CTO application: ' + error.message
    };
  }
}


/**
 * API: Check expired COC
 */
function apiCheckAndExpireCOC() {
  return checkAndExpireCOC();
}


/**
 * Deletes (marks as Cancelled) a COC record.
 * This is used for pending records that haven't been certificated yet.
 *
 * @param {string} recordId - The record ID to delete
 * @param {string} reason - Reason for deletion
 * @returns {Object} Result object
 */
function apiDeleteCOCRecord(recordId, reason) {
  if (!recordId) throw new Error("Record ID is required.");
  if (!reason) throw new Error("Reason for deletion is required.");

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const db = SpreadsheetApp.openById(DATABASE_ID);
    const recordsSheet = db.getSheetByName('COC_Records');
    const data = recordsSheet.getDataRange().getValues();

    const currentUser = getCurrentUserEmail();
    const now = new Date();

    // Find the record
    let rowIndex = -1;
    let recordData = null;

    for (let i = 1; i < data.length; i++) {
      if (data[i][RECORD_COLS.RECORD_ID] === recordId) {
        rowIndex = i;
        recordData = data[i];
        break;
      }
    }

    if (rowIndex === -1) {
      throw new Error("Record not found.");
    }

    // Check if record is already certificated
    const certificateId = recordData[RECORD_COLS.CERTIFICATE_ID];
    if (certificateId) {
      throw new Error("Cannot delete a record that has already been certificated. Please contact administrator if you need to cancel a certificated record.");
    }

    // Check current status
    const currentStatus = recordData[RECORD_COLS.STATUS];
    if (currentStatus === STATUS_CANCELLED) {
      throw new Error("This record has already been cancelled.");
    }

    // Update the record to mark it as Cancelled
    // We don't actually delete it to maintain audit trail
    const row = rowIndex + 1; // Convert to 1-based index
    recordsSheet.getRange(row, RECORD_COLS.STATUS + 1).setValue(STATUS_CANCELLED);
    recordsSheet.getRange(row, RECORD_COLS.LAST_MODIFIED + 1).setValue(now);
    recordsSheet.getRange(row, RECORD_COLS.MODIFIED_BY + 1).setValue(currentUser);

    // Add a note/remark in a comments column if available, or we can add it to the APPROVED_BY column as a workaround
    // Since we don't have a dedicated "Remarks" column in COC_Records, we'll store it in the ledger

    // Create a ledger entry for this cancellation
    const ledgerSheet = db.getSheetByName('COC_Ledger');
    if (ledgerSheet) {
      const employeeId = recordData[RECORD_COLS.EMPLOYEE_ID];
      const employeeName = recordData[RECORD_COLS.EMPLOYEE_NAME];
      const cocEarned = parseFloat(recordData[RECORD_COLS.COC_EARNED] || 0);
      const monthYear = recordData[RECORD_COLS.MONTH_YEAR];

      const ledgerId = generateUniqueId("LDG-");
      const ledgerRow = new Array(15).fill('');
      ledgerRow[LEDGER_COLS.LEDGER_ID] = ledgerId;
      ledgerRow[LEDGER_COLS.EMPLOYEE_ID] = employeeId;
      ledgerRow[LEDGER_COLS.EMPLOYEE_NAME] = employeeName;
      ledgerRow[LEDGER_COLS.TRANSACTION_DATE] = now;
      ledgerRow[LEDGER_COLS.TRANSACTION_TYPE] = 'COC Cancelled';
      ledgerRow[LEDGER_COLS.REFERENCE_ID] = recordId;
      ledgerRow[LEDGER_COLS.BALANCE_BEFORE] = 0; // Since it was pending, it never affected balance
      ledgerRow[LEDGER_COLS.COC_EARNED] = 0;
      ledgerRow[LEDGER_COLS.CTO_USED] = 0;
      ledgerRow[LEDGER_COLS.COC_EXPIRED] = 0;
      ledgerRow[LEDGER_COLS.BALANCE_ADJUSTMENT] = 0;
      ledgerRow[LEDGER_COLS.BALANCE_AFTER] = 0;
      ledgerRow[LEDGER_COLS.MONTH_YEAR_EARNED] = monthYear;
      ledgerRow[LEDGER_COLS.PROCESSED_BY] = currentUser;
      ledgerRow[14] = `Cancelled pending COC record (${cocEarned.toFixed(2)} hrs). Reason: ${reason}`; // Remarks column

      ledgerSheet.getRange(ledgerSheet.getLastRow() + 1, 1, 1, ledgerRow.length).setValues([ledgerRow]);
    }

    Logger.log(`COC Record ${recordId} cancelled by ${currentUser}. Reason: ${reason}`);

    return {
      success: true,
      message: "Record cancelled successfully."
    };

  } catch (e) {
    Logger.log(`Error in apiDeleteCOCRecord: ${e}`);
    throw new Error(`Failed to delete record: ${e.message}`);
  } finally {
    lock.releaseLock();
  }
}


/**
 * API: Delete a holiday
 * @param {number} rowNumber The row number to delete
 * @return {Object} Success result
 */
function apiDeleteHoliday(rowNumber) {
  try {
    const db = getDatabase();
    const sheet = db.getSheetByName('Holidays');
    
    if (!sheet) {
      throw new Error('Holidays sheet not found');
    }

    // Validate row number
    if (!rowNumber || rowNumber < 2) {
      throw new Error('Invalid row number');
    }

    if (rowNumber > sheet.getLastRow()) {
      throw new Error('Row number does not exist');
    }

    // Delete the row
    sheet.deleteRow(rowNumber);

    Logger.log('Holiday deleted successfully from row: ' + rowNumber);
    return { success: true, message: 'Holiday deleted successfully' };

  } catch (error) {
    Logger.log('ERROR in apiDeleteHoliday: ' + error.message + '\nStack: ' + error.stack);
    throw new Error('Failed to delete holiday: ' + error.message);
  }
}


function apiGenerateCOCCertificate(recordId) {
  return generateCOCCertificate(recordId);
}


/**
 * Generates a single monthly certificate for all uncertificated records.
 * This is the new function called by the "Generate" button.
 */
function apiGenerateMonthlyCOCCertificate(employeeId, month, year, issueDateString) {
  if (!employeeId || !month || !year || !issueDateString) {
    throw new Error("Employee ID, month, year, and issue date are required.");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const db = SpreadsheetApp.openById(DATABASE_ID);
    const recordsSheet = db.getSheetByName('COC_Records');
    const recordsData = recordsSheet.getDataRange().getValues();
    const certSheet = db.getSheetByName('COC_Certificates');
    const detailSheet = db.getSheetByName('COC_Balance_Detail');
    const ledgerSheet = db.getSheetByName('COC_Ledger');
    
    const settings = getSheetDataNoHeader('Settings');
    const validityMonths = parseInt(settings.find(r => r[0] === 'COC_VALIDITY_MONTHS')[1] || 12);

    const currentUser = getCurrentUserEmail();
    const now = new Date();
    const monthYear = `${year}-${String(month).padStart(2, '0')}`;
    
    // 1. Get Employee Details
    const empDetails = getEmployeeDetails(employeeId);
    if (!empDetails) throw new Error("Employee not found.");

    // 2. Find records to certify
    const recordsToCertifyIndices = []; // Store row indices (1-based)
    const recordsToCertifyData = [];
    
    // Start from row 2 (index 1) to skip header
    for (let i = 1; i < recordsData.length; i++) {
      const row = recordsData[i];
      if (row[RECORD_COLS.EMPLOYEE_ID] === employeeId &&
          row[RECORD_COLS.MONTH_YEAR] === monthYear &&
          (row[RECORD_COLS.STATUS] === STATUS_PENDING || row[RECORD_COLS.CERTIFICATE_ID] === '')) {
        recordsToCertifyIndices.push(i + 1); // Store 1-based index
        recordsToCertifyData.push(row);
      }
    }

    if (recordsToCertifyData.length === 0) {
      throw new Error("No pending COC records found for this employee and month to certify.");
    }

    // 3. Calculate totals and dates
    const totalNewCOC = recordsToCertifyData.reduce((sum, r) => sum + parseFloat(r[RECORD_COLS.COC_EARNED] || 0), 0);
    const numRecords = recordsToCertifyData.length;
    
    const issueDate = new Date(issueDateString);
    const expirationDate = new Date(issueDate);
    expirationDate.setFullYear(expirationDate.getFullYear() + validityMonths);
    expirationDate.setDate(expirationDate.getDate() - 1);

    const certificateId = generateUniqueId("CERT-");

    // 4. Generate the Google Doc
    const doc = generateCertificateDocument(certificateId, empDetails, recordsToCertifyData, issueDate, expirationDate);

    // 5. Create new rows for other sheets
    const newCertRow = new Array(13).fill('');
    newCertRow[CERT_COLS.CERTIFICATE_ID] = certificateId;
    newCertRow[CERT_COLS.EMPLOYEE_ID] = employeeId;
    newCertRow[CERT_COLS.EMPLOYEE_NAME] = empDetails.fullName;
    newCertRow[CERT_COLS.MONTH_YEAR] = monthYear;
    newCertRow[CERT_COLS.TOTAL_COC_EARNED] = totalNewCOC;
    newCertRow[CERT_COLS.NUMBER_OF_RECORDS] = numRecords;
    newCertRow[CERT_COLS.ISSUE_DATE] = issueDate;
    newCertRow[CERT_COLS.EXPIRATION_DATE] = expirationDate;
    newCertRow[CERT_COLS.CERTIFICATE_URL] = doc.url;
    newCertRow[CERT_COLS.PDF_URL] = doc.pdfUrl;
    newCertRow[CERT_COLS.STATUS] = STATUS_ACTIVE;
    newCertRow[CERT_COLS.CREATED_DATE] = now;
    newCertRow[CERT_COLS.CREATED_BY] = currentUser;
    certSheet.appendRow(newCertRow);

    const newDetailRows = [];
    recordsToCertifyData.forEach(record => {
      const newDetailRow = new Array(19).fill('');
      newDetailRow[DETAIL_COLS.ENTRY_ID] = generateUniqueId("COCD-");
      newDetailRow[DETAIL_COLS.EMPLOYEE_ID] = employeeId;
      newDetailRow[DETAIL_COLS.EMPLOYEE_NAME] = empDetails.fullName;
      newDetailRow[DETAIL_COLS.CERTIFICATE_ID] = certificateId;
      newDetailRow[DETAIL_COLS.RECORD_ID] = record[RECORD_COLS.RECORD_ID];
      newDetailRow[DETAIL_COLS.MONTH_YEAR] = monthYear;
      newDetailRow[DETAIL_COLS.DATE_EARNED] = record[RECORD_COLS.DATE_RENDERED];
      newDetailRow[DETAIL_COLS.DAY_TYPE] = record[RECORD_COLS.DAY_TYPE];
      newDetailRow[DETAIL_COLS.HOURS_EARNED] = record[RECORD_COLS.COC_EARNED];
      newDetailRow[DETAIL_COLS.HOURS_USED] = 0;
      newDetailRow[DETAIL_COLS.HOURS_REMAINING] = record[RECORD_COLS.COC_EARNED];
      newDetailRow[DETAIL_COLS.CERTIFICATE_ISSUE_DATE] = issueDate;
      newDetailRow[DETAIL_COLS.EXPIRATION_DATE] = expirationDate;
      newDetailRow[DETAIL_COLS.STATUS] = STATUS_ACTIVE;
      newDetailRow[DETAIL_COLS.DATE_CREATED] = record[RECORD_COLS.DATE_RECORDED];
      newDetailRow[DETAIL_COLS.CREATED_BY] = record[RECORD_COLS.CREATED_BY];
      newDetailRow[DETAIL_COLS.LAST_UPDATED] = now;
      newDetailRow[DETAIL_COLS.NOTES] = `Part of ${monthYear} certificate. Expires ${Utilities.formatDate(expirationDate, "GMT+8", "yyyy-MM-dd")}.`;
      
      newDetailRows.push(newDetailRow);
    });
    if (newDetailRows.length > 0) {
      detailSheet.getRange(detailSheet.getLastRow() + 1, 1, newDetailRows.length, newDetailRows[0].length).setValues(newDetailRows);
    }

    // 6. Create Ledger Entry
    const currentBalance = apiGetBalance(employeeId);
    const newBalance = currentBalance + totalNewCOC;
    const newLedgerRow = new Array(19).fill('');
    newLedgerRow[LEDGER_COLS.LEDGER_ID] = generateUniqueId("LED-");
    newLedgerRow[LEDGER_COLS.EMPLOYEE_ID] = employeeId;
    newLedgerRow[LEDGER_COLS.EMPLOYEE_NAME] = empDetails.fullName;
    newLedgerRow[LEDGER_COLS.TRANSACTION_DATE] = issueDate;
    newLedgerRow[LEDGER_COLS.TRANSACTION_TYPE] = TR_TYPE_EARNED;
    newLedgerRow[LEDGER_COLS.REFERENCE_ID] = certificateId;
    newLedgerRow[LEDGER_COLS.BALANCE_BEFORE] = currentBalance;
    newLedgerRow[LEDGER_COLS.COC_EARNED] = totalNewCOC;
    newLedgerRow[LEDGER_COLS.CTO_USED] = 0;
    newLedgerRow[LEDGER_COLS.COC_EXPIRED] = 0;
    newLedgerRow[LEDGER_COLS.BALANCE_ADJUSTMENT] = 0;
    newLedgerRow[LEDGER_COLS.BALANCE_AFTER] = newBalance;
    newLedgerRow[LEDGER_COLS.MONTH_YEAR_EARNED] = monthYear;
    newLedgerRow[LEDGER_COLS.EXPIRATION_DATE] = expirationDate;
    newLedgerRow[LEDGER_COLS.PROCESSED_BY] = currentUser;
    newLedgerRow[LEDGER_COLS.PROCESSED_DATE] = now;
    newLedgerRow[LEDGER_COLS.REMARKS] = `Certificate ${certificateId} issued for ${totalNewCOC} hrs.`;
    ledgerSheet.appendRow(newLedgerRow);

    // 7. Update original COC_Records
    recordsToCertifyIndices.forEach(rowIndex => {
      recordsSheet.getRange(rowIndex, RECORD_COLS.CERTIFICATE_ID + 1).setValue(certificateId);
      recordsSheet.getRange(rowIndex, RECORD_COLS.EXPIRATION_DATE + 1).setValue(expirationDate);
      recordsSheet.getRange(rowIndex, RECORD_COLS.STATUS + 1).setValue(STATUS_ACTIVE);
      recordsSheet.getRange(rowIndex, RECORD_COLS.LAST_MODIFIED + 1).setValue(now);
      recordsSheet.getRange(rowIndex, RECORD_COLS.MODIFIED_BY + 1).setValue(currentUser);
    });

    // 8. Return success
    return {
      certificateId: certificateId,
      numRecords: numRecords,
      totalHours: totalNewCOC
    };

  } catch (e) {
    Logger.log(`Error in apiGenerateMonthlyCOCCertificate: ${e}\nStack: ${e.stack}`);
    throw new Error(`Failed to generate certificate: ${e.message}`);
  } finally {
    lock.releaseLock();
  }
}


/**
 * Get all CTO applications from the database with proper serialization
 */
function apiGetAllCTOApplications() {
  try {
    const db = getDatabase();
    const ctoSheet = db.getSheetByName('CTO_Applications');
    
    if (!ctoSheet || ctoSheet.getLastRow() < 2) {
      Logger.log('No CTO applications found');
      return [];
    }
    
    const data = ctoSheet.getDataRange().getValues();
    const applications = [];
    const TIME_ZONE = getScriptTimeZone();
    
    // Helper to format dates safely
    function safeFormatDate(value) {
      if (!value) return '';
      try {
        if (value instanceof Date) {
          return Utilities.formatDate(value, TIME_ZONE, 'yyyy-MM-dd');
        }
        return String(value);
      } catch (e) {
        return String(value);
      }
    }
    
    // Process each row (skip header)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // Skip empty rows
      if (!row[0]) continue;
      
      try {
        // Ensure all data is serializable
        const app = {
          appId: String(row[0] || ''),
          employeeId: String(row[1] || ''),
          employeeName: String(row[2] || ''),
          office: String(row[3] || ''),
          hours: parseFloat(row[4]) || 0,
          startDate: safeFormatDate(row[5]),
          endDate: safeFormatDate(row[6]),
          appliedDate: safeFormatDate(row[7]),
          approvedBy: String(row[8] || ''),
          approvedDate: safeFormatDate(row[9]),
          status: String(row[10] || 'Pending'),
          remarks: String(row[11] || '')
        };
        
        applications.push(app);
        
      } catch (rowError) {
        Logger.log('Error processing CTO application row ' + (i + 1) + ': ' + rowError.message);
        continue;
      }
    }
    
    Logger.log('Loaded ' + applications.length + ' CTO applications');
    return applications;
    
  } catch (error) {
    Logger.log('ERROR in apiGetAllCTOApplications: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    return [];
  }
}


/**
 * Get employee's current COC balance
 * This reads from the COC_Ledger to get the most recent balance
 * 
 * @param {string} empId - Employee ID (e.g., "EMP008")
 * @returns {number} Current COC balance in hours
 */
function apiGetBalance(empId) {
  try {
    // Use the correct spreadsheet ID
    let ss;
    if (typeof CONFIG !== 'undefined' && CONFIG.SPREADSHEET_ID) {
      ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    } else if (typeof SPREADSHEET_ID !== 'undefined') {
      ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    } else {
      const SPREADSHEET_ID = '1LIQRnQb7lL-6hdSpsOwq-XVwRM6Go6u4q6RkSB2ZCXo';
      ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    }
    
    // Read from COC_Ledger sheet (NOT COC_Balance_Detail)
    const ledgerSheet = ss.getSheetByName('COC_Ledger');
    
    if (!ledgerSheet) {
      Logger.log('COC_Ledger sheet not found');
      return 0;
    }
    
    const data = ledgerSheet.getDataRange().getValues();
    const headers = data[0];
    
    // Find column indices
    const empIdCol = headers.indexOf('Employee ID');
    const transDateCol = headers.indexOf('Transaction Date');
    const balanceCol = headers.indexOf('Balance After'); // Use the correct header name
    
    if (empIdCol === -1 || balanceCol === -1 || transDateCol === -1) {
      Logger.log('Required columns not found in COC_Ledger');
      return 0;
    }
    
    // Find the most recent transaction for this employee
    let latestBalance = 0;
    let latestDate = null;
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      if (row[empIdCol] === empId) {
        const transDate = new Date(row[transDateCol]);
        const balance = parseFloat(row[balanceCol]) || 0;
        
        if (!latestDate || transDate > latestDate) {
          latestDate = transDate;
          latestBalance = balance;
        }
      }
    }
    
    Logger.log('Balance for ' + empId + ': ' + latestBalance);
    return latestBalance;
    
  } catch (error) {
    Logger.log('Error in apiGetBalance: ' + error.toString());
    return 0;
  }
}


/**
 * Alternative: Get balance from COC_Balance_Detail (FIFO method)
 * Use this if you want to calculate from remaining hours in detail sheet
 */
function apiGetBalanceFromDetail(empId) {
  try {
    let ss;
    if (typeof CONFIG !== 'undefined' && CONFIG.SPREADSHEET_ID) {
      ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    } else if (typeof SPREADSHEET_ID !== 'undefined') {
      ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    } else {
      const SPREADSHEET_ID = '1LIQRnQb7lL-6hdSpsOwq-XVwRM6Go6u4q6RkSB2ZCXo';
      ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    }
    
    const detailSheet = ss.getSheetByName('COC_Balance_Detail');
    
    if (!detailSheet) {
      Logger.log('COC_Balance_Detail sheet not found');
      return 0;
    }
    
    const data = detailSheet.getDataRange().getValues();
    const headers = data[0];
    
    const empIdCol = headers.indexOf('Employee ID');
    const hoursRemainingCol = headers.indexOf('Hours Remaining');
    const statusCol = headers.indexOf('Status');
    
    if (empIdCol === -1 || hoursRemainingCol === -1) {
      Logger.log('Required columns not found in COC_Balance_Detail');
      return 0;
    }
    
    let totalBalance = 0;
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      if (row[empIdCol] === empId) {
        const status = row[statusCol];
        // Only count Active entries
        if (status === 'Active') {
          const hoursRemaining = parseFloat(row[hoursRemainingCol]) || 0;
          totalBalance += hoursRemaining;
        }
      }
    }
    
    Logger.log('Balance from detail for ' + empId + ': ' + totalBalance);
    return totalBalance;
    
  } catch (error) {
    Logger.log('Error in apiGetBalanceFromDetail: ' + error.toString());
    return 0;
  }
}


/**
 * API: Get COC balance breakdown
 */
function apiGetCOCBalanceBreakdown(employeeId) {
  try {
    const breakdown = getCOCBalanceBreakdown(employeeId);
    
    // Convert to serializable format
    const safeBreakdown = [];
    
    if (breakdown && Array.isArray(breakdown)) {
      breakdown.forEach(function(item) {
        safeBreakdown.push({
          entryId: String(item.entryId || ''),
          recordId: String(item.recordId || ''),
          dateEarned: String(item.dateEarned || ''),
          hoursEarned: parseFloat(item.hoursEarned) || 0,
          hoursRemaining: parseFloat(item.hoursRemaining) || 0,
          expirationDate: String(item.expirationDate || ''),
          daysUntilExpiration: parseInt(item.daysUntilExpiration) || 0
        });
      });
    }
    
    return safeBreakdown;
    
  } catch (error) {
    Logger.log('ERROR in apiGetCOCBalanceBreakdown: ' + error.message);
    return [];
  }
}


/**
 * Server API wrapper for dashboard stats.
 *
 * @return {Object} Stats for dashboard display.
*/
function apiGetDashboardStats() {
  return getDashboardStats();
}


/**
 * API: Get day type only (Weekday, Weekend, Holiday)
 * Used for immediate day type detection without needing time entries
 *
 * @param {number} year - The year (e.g., 2025)
 * @param {number} month - The month (1-12)
 * @param {number} day - The day of the month (1-31)
 * @return {string} The day type (Weekday, Weekend, Regular Holiday, Special Non-Working Holiday)
 */
function apiGetDayType(year, month, day) {
  const date = new Date(year, month - 1, day);
  return getDayType(date); // This will now work
}


/**
 * API wrapper for getting dropdown options for Positions and Offices.
 * @return {Object} An object { positions: [...], offices: [...] }.
 */
function apiGetDropdownOptions() {
  return getDropdownOptions();
}


function apiGetEmployee(employeeId) {
  return getEmployeeById(employeeId);
}


/**
 * Gets the total COC earned for the specified month.
 * @param {string} employeeId The ID of the employee.
 * @param {number} month The month (1-12) from the dropdown.
 * @param {number} year The year (e.g., 2025) from the dropdown.
 * @returns {object} An object containing { currentMonthTotal, monthlyRemaining, etc. }.
 */
function apiGetEmployeeCOCStats(employeeId, month, year) {
  const db = getDatabase(); // Assuming you have this function
  const detailSheet = db.getSheetByName('COC_Balance_Detail');
  const TIME_ZONE = getScriptTimeZone(); // Assuming you have this function

  // Get current balance (This is the TOTAL balance, which is fine)
  const currentBalance = getCurrentCOCBalanceFromDetail(employeeId); // Assuming you have this function

  // Get current month's COC total
  const now = new Date(); // We use this *only* as a fallback

  // --- START OF FIX ---
  
  // 1. Use the passed-in month and year. If they are missing, use the current month/year.
  const targetMonth = month || (now.getMonth() + 1);
  const targetYear = year || now.getFullYear();

  // 2. Create the target Month-Year string in YYYY-MM format
  //    This format MUST match your 'COC_Balance_Detail' sheet's 'Month Year' column (Column F)
  //    Based on your CSV, it looks like it's "YYYY-MM" (e.g., "2025-10").
  const targetMonthYear = `${targetYear}-${String(targetMonth).padStart(2, '0')}`;
  
  // --- END OF FIX ---

  let currentMonthTotal = 0;

  if (detailSheet && detailSheet.getLastRow() > 1) {
    const data = detailSheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const empId = row[1]; // Employee ID (Column B)
      
      // --- START OF FIX 2 ---
      // We read the 'Month Year' column directly. It's much faster and more reliable.
      // This assumes 'Month Year' is in Column F (index 5)
      const rowMonthYear = row[5]; 
      const hoursEarned = parseFloat(row[8]) || 0; // 'Hours Earned' (Column I)

      // 3. Compare against the targetMonthYear string
      if (empId === employeeId && rowMonthYear === targetMonthYear) {
        currentMonthTotal += hoursEarned;
      }
      // --- END OF FIX 2 ---
    }
  }

  // This return object matches your original function
  return {
    employeeId: employeeId,
    currentBalance: currentBalance,
    balanceLimit: 120,
    balanceRemaining: 120 - currentBalance,
    currentMonthTotal: currentMonthTotal, // This will now be for the SELECTED month
    monthlyLimit: 40,
    monthlyRemaining: 40 - currentMonthTotal,
    canAddThisMonth: 40 - currentMonthTotal,
    canAddTotal: Math.min(40 - currentMonthTotal, 120 - currentBalance)
  };
}


/**
 * Get all CTO applications for a specific employee
 * This should be added to your Code.gs file
 * 
 * @param {string} empId - Employee ID (e.g., "EMP008")
 * @returns {Array} Array of CTO application objects
 */
/**
 * FIXED: Get all CTO applications for a specific employee
 * Enhanced with robust date handling for inconsistent formats
 * 
 * @param {string} empId - Employee ID (e.g., "EMP008")
 * @returns {Array} Array of CTO application objects
 */
function apiGetEmployeeCTOApplications(employeeId) {
  try {
    // Get the original result (you already have this function somewhere)
    const db = getDatabase();
    const ctoSheet = db.getSheetByName('CTO_Applications');
    
    if (!ctoSheet || ctoSheet.getLastRow() < 2) {
      return [];
    }
    
    const data = ctoSheet.getDataRange().getValues();
    const applications = [];
    
    // Normalize search ID
    const searchId = String(employeeId).trim();
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowEmpId = String(row[1] || '').trim();
      
      if (rowEmpId === searchId) {
        applications.push({
          appId: String(row[0] || ''),
          employeeId: String(row[1] || ''),
          employeeName: String(row[2] || ''),
          office: String(row[3] || ''),
          hours: parseFloat(row[4]) || 0,
          startDate: String(row[5] || ''),
          endDate: String(row[6] || ''),
          appliedDate: String(row[7] || ''),
          status: String(row[10] || ''),
          remarks: String(row[11] || '')
        });
      }
    }
    
    return applications;
    
  } catch (error) {
    Logger.log('ERROR in apiGetEmployeeCTOApplications: ' + error.message);
    return [];
  }
}


/**
 * ============================================================================
 * API: Get Employees with Expiring COC (within 30 days)
 * ============================================================================
 * Retrieves a list of employees who have active COC entries expiring within
 * the next 30 days, along with the total hours expiring for each employee
 * and the date of the earliest expiration.
 *
 * @return {Array<Object>} An array of objects: { employeeId, employeeName, totalHoursExpiring, earliestExpiryDate }
 */
function apiGetEmployeesWithExpiringCOC() {
  try {
    const db = getDatabase();
    const detailSheet = ensureCOCBalanceDetailSheet(); // Ensure the sheet exists
    const TIME_ZONE = getScriptTimeZone();

    if (!detailSheet || detailSheet.getLastRow() < 2) {
      Logger.log('COC_Balance_Detail sheet is empty or not found.');
      return [];
    }

    const data = detailSheet.getDataRange().getValues();
    const expiringMap = {}; // Use a map to group by employeeId

    const today = new Date();
    today.setHours(0, 0, 0, 0); // Start of today
    const thirtyDaysFromNow = new Date(today);
    thirtyDaysFromNow.setDate(today.getDate() + 30); // End of the 30-day window

    Logger.log(`Checking for expirations between ${today.toISOString()} and ${thirtyDaysFromNow.toISOString()}`);

    // Iterate through detail entries (skip header row)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const employeeId = String(row[DETAIL_COLS.EMPLOYEE_ID] || '').trim();
      const employeeName = String(row[DETAIL_COLS.EMPLOYEE_NAME] || '').trim();
      const hoursRemaining = parseFloat(row[DETAIL_COLS.HOURS_REMAINING]) || 0;
      const expirationDateVal = row[DETAIL_COLS.EXPIRATION_DATE];
      const status = String(row[DETAIL_COLS.STATUS] || '').trim();

      // Skip if no employee ID, status is not Active, or no hours remaining
      if (!employeeId || status !== 'Active' || hoursRemaining <= 0) {
        continue;
      }

      // --- Robust Date Parsing for Expiration Date ---
      let expirationDate = null;
      if (expirationDateVal instanceof Date && !isNaN(expirationDateVal.getTime())) {
          expirationDate = new Date(expirationDateVal); // Clone to avoid modifying original
      } else if (expirationDateVal) {
          try {
              const parsed = new Date(expirationDateVal);
              if (!isNaN(parsed.getTime())) {
                  expirationDate = parsed;
              } else {
                 Logger.log(`Skipping row ${i+1} due to invalid expiration date format: ${expirationDateVal}`);
                 continue; // Skip if date is invalid
              }
          } catch (e) {
               Logger.log(`Error parsing expiration date at row ${i+1}: ${expirationDateVal}`);
               continue; // Skip on parsing error
          }
      } else {
           Logger.log(`Skipping row ${i+1} due to missing expiration date.`);
           continue; // Skip if no expiration date
      }
      expirationDate.setHours(0,0,0,0); // Use start of the expiration day for comparison
      // --- End Date Parsing ---


      // Check if the expiration date is within the next 30 days (inclusive of today)
      if (expirationDate >= today && expirationDate <= thirtyDaysFromNow) {
        // If employee not yet in map, initialize
        if (!expiringMap[employeeId]) {
          expiringMap[employeeId] = {
            employeeId: employeeId,
            employeeName: employeeName,
            totalHoursExpiring: 0,
            earliestExpiryDate: expirationDate // Initialize with the first one found
          };
        }

        // Add hours to the total for this employee
        expiringMap[employeeId].totalHoursExpiring += hoursRemaining;

        // Update the earliest expiration date if this one is earlier
        if (expirationDate < expiringMap[employeeId].earliestExpiryDate) {
          expiringMap[employeeId].earliestExpiryDate = expirationDate;
        }
         Logger.log(`Found expiring entry for ${employeeId}: ${hoursRemaining} hrs on ${Utilities.formatDate(expirationDate, TIME_ZONE, 'yyyy-MM-dd')}`);
      }
    }

    // Convert map to array
    const expiringList = Object.values(expiringMap);

    // Sort the list (e.g., by earliest expiry date, then by name)
    expiringList.sort((a, b) => {
      if (a.earliestExpiryDate.getTime() !== b.earliestExpiryDate.getTime()) {
        return a.earliestExpiryDate - b.earliestExpiryDate; // Sort by date first
      }
      return a.employeeName.localeCompare(b.employeeName); // Then by name
    });

     // Format the date string for the final output
     const formattedList = expiringList.map(item => ({
       ...item,
       earliestExpiryDate: Utilities.formatDate(item.earliestExpiryDate, TIME_ZONE, 'yyyy-MM-dd') // Format date as YYYY-MM-DD string
     }));


    Logger.log(`Found ${formattedList.length} employees with COC expiring within 30 days.`);
    return formattedList; // Return the array

  } catch (error) {
    Logger.log('ERROR in apiGetEmployeesWithExpiringCOC: ' + error.message + '\nStack: ' + error.stack);
    return []; // Return empty array on error
  }
}


/**
 * API: Get historical import records (last 50)
 * @return {Array<Object>} Array of import history objects
 */
function apiGetHistoricalImports() {
  try {
    const db = getDatabase();
    const detailSheet = db.getSheetByName('COC_Balance_Detail');
    
    if (!detailSheet || detailSheet.getLastRow() < 2) {
      return [];
    }

    const data = detailSheet.getDataRange().getValues();
    const results = [];
    const TIME_ZONE = getScriptTimeZone();

    // Look for entries with "Historical Import" in remarks
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const remarks = String(row[13] || '').toLowerCase();
      
      if (remarks.includes('historical') || remarks.includes('import') || remarks.includes('migrat')) {
        results.push({
          employeeId: row[1],
          employeeName: row[2],
          monthYear: row[3],
          certificateDate: Utilities.formatDate(new Date(row[4]), TIME_ZONE, 'MMM dd, yyyy'),
          hoursEarned: Number(row[5]).toFixed(2),
          hoursUsed: Number(row[6]).toFixed(2),
          hoursRemaining: Number(row[7]).toFixed(2),
          status: row[8],
          importedDate: Utilities.formatDate(new Date(row[11]), TIME_ZONE, 'MMM dd, yyyy HH:mm')
        });
      }
    }

    // Sort by import date (most recent first)
    results.sort((a, b) => {
      const dateA = new Date(a.importedDate);
      const dateB = new Date(b.importedDate);
      return dateB - dateA;
    });

    // Return last 50 entries
    return results.slice(0, 50);

  } catch (error) {
    Logger.log('ERROR in apiGetHistoricalImports: ' + error.message + '\nStack: ' + error.stack);
    throw new Error('Failed to load import history: ' + error.message);
  }
}


function apiGetLedger(employeeId) {
  try {
    const result = getLedgerForEmployee(employeeId);
    
    // Convert to serializable format
    const safeResult = {
      balance: parseFloat(result.balance) || 0,
      entries: []
    };
    
    if (result.entries && Array.isArray(result.entries)) {
      result.entries.forEach(function(entry) {
        safeResult.entries.push({
          ledgerId: String(entry.ledgerId || ''),
          employeeId: String(entry.employeeId || ''),
          employeeName: String(entry.employeeName || ''),
          transactionDate: String(entry.transactionDate || ''),
          transactionType: String(entry.transactionType || ''),
          referenceId: String(entry.referenceId || ''),
          cocEarned: parseFloat(entry.cocEarned) || 0,
          ctoUsed: parseFloat(entry.ctoUsed) || 0,
          cocBalance: parseFloat(entry.cocBalance) || 0,
          monthYearEarned: String(entry.monthYearEarned || ''),
          expirationDate: String(entry.expirationDate || ''),
          processedBy: String(entry.processedBy || ''),
          remarks: String(entry.remarks || '')
        });
      });
    }
    
    return safeResult;
    
  } catch (error) {
    Logger.log('ERROR in apiGetLedger: ' + error.message);
    return { balance: 0, entries: [] };
  }
}


/**
 * Gets the monthly certificate information for display.
 * Returns null if no certificate exists for the month.
 *
 * @param {string} employeeId - The employee ID
 * @param {number} month - Month (1-12)
 * @param {number} year - Year (e.g., 2025)
 * @returns {Object|null} Certificate object or null
 */
function apiGetMonthlyCertificate(employeeId, month, year) {
  if (!employeeId || !month || !year) return null;

  try {
    const monthYear = `${year}-${String(month).padStart(2, '0')}`;
    const certsData = getSheetDataNoHeader('COC_Certificates');

    // Find certificate for this employee and month
    const cert = certsData.find(c =>
      c[CERT_COLS.EMPLOYEE_ID] === employeeId &&
      c[CERT_COLS.MONTH_YEAR] === monthYear
    );

    if (!cert) return null;

    return {
      certificateId: cert[CERT_COLS.CERTIFICATE_ID],
      totalCOC: parseFloat(cert[CERT_COLS.TOTAL_COC_EARNED] || 0),
      numRecords: cert[CERT_COLS.NUMBER_OF_RECORDS],
      issueDate: cert[CERT_COLS.ISSUE_DATE] ?
        Utilities.formatDate(new Date(cert[CERT_COLS.ISSUE_DATE]), "GMT+8", "MMM dd, yyyy") : null,
      expirationDate: cert[CERT_COLS.EXPIRATION_DATE] ?
        Utilities.formatDate(new Date(cert[CERT_COLS.EXPIRATION_DATE]), "GMT+8", "MMM dd, yyyy") : null,
      certificateUrl: cert[CERT_COLS.CERTIFICATE_URL],
      pdfUrl: cert[CERT_COLS.PDF_URL],
      status: cert[CERT_COLS.STATUS]
    };

  } catch (e) {
    Logger.log(`Error in apiGetMonthlyCertificate: ${e}`);
    return null;
  }
}


/**
 * API wrapper for recent ledger activity. Accepts an optional limit.
 *
 * @param {number} limit Maximum number of records to return.
 * @return {Array<Object>} Recent activities.
 */
function apiGetRecentActivities(limit) {
  return getRecentActivities(limit);
}


function apiGetReport(type, startDate, endDate) {
  return getReportData(type, startDate, endDate);
}


/**
 * Gets the current signatories from the Settings sheet.
 */
function apiGetSignatories() {
  try {
    const settings = getSheetDataNoHeader('Signatories'); // Changed from 'Settings'
    const getSetting = (key, defaultValue = "") => {
      const row = settings.find(r => r[0] === key);
      return row ? row[1] : defaultValue;
    };

    return {
      issuedByName: getSetting("SIGNATORY_ISSUED_BY_NAME", "NIDA O. TRINIDAD"),
      issuedByPosition: getSetting("SIGNATORY_ISSUED_BY_POSITION", "Administrative Officer V")
      // REMOVED: notedByName and notedByPosition
    };
  } catch (e) {
    Logger.log(`Error in apiGetSignatories: ${e}`);
    throw new Error(`Failed to get signatories: ${e.message}`);
  }
}


/**
 * API: Import a single historical COC entry
 * @param {Object} data Import data object
 * @return {Object} Success result
 */
function apiImportHistoricalCOC(data) {
  try {
    const db = getDatabase();
    const detailSheet = ensureCOCBalanceDetailSheet();
    const ledgerSheet = ensureLedgerSheet();
    
    if (!detailSheet) {
      throw new Error('COC_Balance_Detail sheet not found');
    }

    // Validate employee
    const employee = getEmployeeById(data.employeeId);
    if (!employee) {
      throw new Error('Employee not found: ' + data.employeeId);
    }

    // Validate data
    if (!data.monthYear || !data.certificateDate || !data.hoursEarned) {
      throw new Error('Missing required fields');
    }

    if (data.hoursUsed > data.hoursEarned) {
      throw new Error('Hours Used cannot exceed Hours Earned');
    }

    const hoursRemaining = data.hoursEarned - data.hoursUsed;

    // Generate IDs
    const recordId = generateRecordId();
    const certificateId = 'CERT-' + Utilities.formatDate(new Date(), getScriptTimeZone(), 'yyyyMMddHHmmssSSS');
    
    // Parse dates
    const certDate = new Date(data.certificateDate);
    const expDate = new Date(data.expirationDate);
    const importDate = new Date();

    // Create the detail entry
    const detailRow = [
      recordId,                    // A: Record ID
      data.employeeId,             // B: Employee ID
      employee.fullName,           // C: Employee Name
      data.monthYear,              // D: Month-Year Earned
      certDate,                    // E: Certificate Date
      data.hoursEarned,            // F: Hours Earned
      data.hoursUsed,              // G: Hours Used
      hoursRemaining,              // H: Hours Remaining
      data.status,                 // I: Status
      expDate,                     // J: Expiration Date
      certificateId,               // K: Certificate ID
      importDate,                  // L: Date Created
      Session.getActiveUser().getEmail(), // M: Created By
      data.remarks || 'Historical Import' // N: Remarks
    ];

    detailSheet.appendRow(detailRow);
    Logger.log('Historical COC entry imported: ' + recordId);

    // Add ledger entry if there are hours
    if (data.hoursEarned > 0) {
      const ledgerRow = [
        generateLedgerEntryId(),
        data.employeeId,
        employee.fullName,
        importDate,
        'Historical Import',
        recordId,
        data.hoursEarned,
        0, // No CTO used in import
        getCurrentCOCBalance(data.employeeId),
        data.monthYear,
        expDate,
        Session.getActiveUser().getEmail(),
        data.remarks || 'Historical data import'
      ];
      
      ledgerSheet.appendRow(ledgerRow);
      Logger.log('Ledger entry created for historical import');
    }

    return { 
      success: true, 
      message: 'Historical COC entry imported successfully',
      recordId: recordId
    };

  } catch (error) {
    Logger.log('ERROR in apiImportHistoricalCOC: ' + error.message + '\nStack: ' + error.stack);
    throw new Error('Failed to import historical COC: ' + error.message);
  }
}


/**
 * API: Import multiple historical COC entries from CSV
 * @param {Array} csvData Array of CSV rows (first row is headers)
 * @return {Object} Result with success and error counts
 */
function apiImportHistoricalCOCBatch(csvData) {
  try {
    if (!csvData || csvData.length < 2) {
      throw new Error('CSV data is empty or invalid');
    }

    let successCount = 0;
    let errorCount = 0;
    const errors = [];

    // Expected headers: Employee ID, Month-Year, Certificate Date, Hours Earned, Hours Used, Status, Remarks
    const headers = csvData[0];
    
    // Skip header row, process data rows
    for (let i = 1; i < csvData.length; i++) {
      const row = csvData[i];
      
      // Skip empty rows
      if (!row[0] || row[0].trim() === '') {
        continue;
      }

      try {
        // Parse row data
        const employeeId = row[0].trim();
        const monthYear = row[1].trim();
        const certificateDate = row[2].trim();
        const hoursEarned = parseFloat(row[3]) || 0;
        const hoursUsed = parseFloat(row[4]) || 0;
        const status = row[5].trim() || 'Active';
        const remarks = row[6] || 'CSV Import';

        // Calculate expiration date (Certificate Date + 1 year - 1 day)
        const certDate = new Date(certificateDate);
        const expDate = new Date(certDate);
        expDate.setFullYear(expDate.getFullYear() + 1);
        expDate.setDate(expDate.getDate() - 1);

        // Import the entry
        const data = {
          employeeId: employeeId,
          monthYear: monthYear,
          certificateDate: certificateDate,
          expirationDate: expDate.toISOString().split('T')[0],
          hoursEarned: hoursEarned,
          hoursUsed: hoursUsed,
          status: status,
          remarks: remarks
        };

        apiImportHistoricalCOC(data);
        successCount++;

      } catch (rowError) {
        errorCount++;
        errors.push(`Row ${i + 1}: ${rowError.message}`);
        Logger.log(`Error importing row ${i + 1}: ${rowError.message}`);
      }
    }

    const result = {
      success: true,
      successCount: successCount,
      errorCount: errorCount,
      errors: errors
    };

    if (errorCount > 0) {
      Logger.log(`Batch import completed with ${errorCount} errors: ${errors.join('; ')}`);
    }

    return result;

  } catch (error) {
    Logger.log('ERROR in apiImportHistoricalCOCBatch: ' + error.message + '\nStack: ' + error.stack);
    throw new Error('Failed to import CSV: ' + error.message);
  }
}


/**
 * API: Initialize COC_Balance_Detail sheet
 */
function apiInitializeCOCBalanceDetail() {
  ensureCOCBalanceDetailSheet();
  return { success: true, message: 'COC_Balance_Detail sheet initialized' };
}


/**
 * Lists all recorded COC entries for a given employee and month.
 * This is called by loadExistingRecords() in the HTML.
 */
function apiListCOCRecordsForMonth(employeeId, month, year) {
  if (!employeeId || !month || !year) throw new Error("Employee ID, month, and year are required.");

  try {
    const monthYear = `${year}-${String(month).padStart(2, '0')}`;
    const recordsData = getSheetDataNoHeader('COC_Records');
    const certsData = getSheetDataNoHeader('COC_Certificates');

    Logger.log(`apiListCOCRecordsForMonth: Searching for empId="${employeeId}", month=${month}, year=${year}, monthYear="${monthYear}"`);
    Logger.log(`Total records in sheet: ${recordsData.length}`);

    // Create a map of Certificate IDs to their URLs for easy lookup
    const certMap = new Map();
    certsData.forEach(cert => {
      certMap.set(cert[CERT_COLS.CERTIFICATE_ID], {
        url: cert[CERT_COLS.CERTIFICATE_URL],
        pdf: cert[CERT_COLS.PDF_URL]
      });
    });

    // Filter records - EXCLUDE CANCELLED status
    const employeeMonthRecords = recordsData.filter(r => {
      const rowEmpId = String(r[RECORD_COLS.EMPLOYEE_ID] || '').trim();
      const rowMonthYear = String(r[RECORD_COLS.MONTH_YEAR] || '').trim();
      const rowStatus = String(r[RECORD_COLS.STATUS] || '').trim();

      const empMatch = rowEmpId === employeeId;
      const monthMatch = rowMonthYear === monthYear;
      const notCancelled = rowStatus !== STATUS_CANCELLED;

      if (empMatch) {
        Logger.log(`  Row: empId="${rowEmpId}", monthYear="${rowMonthYear}", status="${rowStatus}", matches=${empMatch && monthMatch && notCancelled}`);
      }

      return empMatch && monthMatch && notCancelled;
    });

    Logger.log(`Found ${employeeMonthRecords.length} matching records`);

    // Format the records for the client
    const formattedRecords = employeeMonthRecords.map(r => {
      const recordId = r[RECORD_COLS.RECORD_ID];
      const certificateId = r[RECORD_COLS.CERTIFICATE_ID];
      const certInfo = certMap.get(certificateId);

      return {
        recordId: recordId,
        displayDate: Utilities.formatDate(new Date(r[RECORD_COLS.DATE_RENDERED]), "GMT+8", "MMM dd, yyyy"),
        dayType: r[RECORD_COLS.DAY_TYPE],
        hoursWorked: parseFloat(r[RECORD_COLS.HOURS_WORKED] || 0),
        cocEarned: parseFloat(r[RECORD_COLS.COC_EARNED] || 0),
        certificateId: certificateId,
        certificateUrl: certInfo ? certInfo.url : null,
        pdfUrl: certInfo ? (certInfo.pdf || certInfo.url) : null // Fallback pdf to url
      };
    });

    return formattedRecords;

  } catch (e) {
    Logger.log(`Error in apiListCOCRecordsForMonth: ${e}`);
    throw new Error(`Failed to list COC records: ${e.message}`);
  }
}


 * OLD VERSION - DEPRECATED
 * This function has been replaced by the newer version at line 6383+.
 * Kept here for reference but should not be used.
 * The newer version uses MONTH_YEAR column filtering instead of parsing dates.
 */
/*
function apiListCOCRecordsForMonth_OLD(employeeId, month, year) {
  const db = getDatabase();
  const recSheet = db.getSheetByName('COC_Records');
  if (!recSheet) throw new Error('COC_Records sheet not found');

  const certSheet = db.getSheetByName('COC_Certificates');
  const TIME_ZONE = getScriptTimeZone(); // Get timezone

  // Get all records
  const recData = recSheet.getDataRange().getValues();

  // Get all certificates (if sheet exists)
  const certData = certSheet ? certSheet.getDataRange().getValues() : [];

  // Build certificate lookup by recordId
  const certMap = {};
  for (let i = 1; i < certData.length; i++) {
    const recordId = certData[i][0];
    certMap[recordId] = {
      certificateUrl: certData[i][6],
      pdfUrl: certData[i][7],
      issuedDate: certData[i][8]
    };
  }

  const targetMonth = month - 1; // JS months are 0-indexed
  const targetYear = year;
  const results = [];

  Logger.log('apiListCOCRecordsForMonth: Looking for employeeId=' + employeeId + ', month=' + month + ', year=' + year);

  for (let i = 1; i < recData.length; i++) {
    const row = recData[i];
    const recEmployeeId = String(row[1] || '').trim(); // Convert to string and trim
    const status = String(row[15] || '').trim();

    // Skip if not matching employee or if cancelled
    if (recEmployeeId !== employeeId) continue;
    if (status.toLowerCase() === 'cancelled') continue;

    const recordId = row[0];

    // Parse date carefully
    let dateRendered;
    if (row[4] instanceof Date) {
      dateRendered = new Date(row[4]);
    } else if (row[4]) {
      dateRendered = new Date(row[4]);
      if (isNaN(dateRendered.getTime())) {
        Logger.log(`WARNING: Invalid date at row ${i}: ${row[4]}`);
        continue;
      }
    } else {
      Logger.log(`WARNING: Missing date at row ${i}`);
      continue;
    }

    // Check if date matches target month/year
    if (dateRendered.getMonth() !== targetMonth || dateRendered.getFullYear() !== targetYear) {
      continue;
    }

    const dayType = row[5] || '';
    const hoursWorked = parseFloat(row[10]) || 0;
    const cocEarned = parseFloat(row[12]) || 0;

    // Get certificate info if exists
    const cert = certMap[recordId];

    // FIX: Use Utilities.formatDate() instead of formatDate()
    let displayDate;
    try {
      displayDate = Utilities.formatDate(dateRendered, TIME_ZONE, 'MMM dd, yyyy');
    } catch (e) {
      // Fallback if formatting fails
      displayDate = dateRendered.toLocaleDateString();
      Logger.log(`WARNING: Date formatting failed for ${dateRendered}, using fallback`);
    }

    // Format certificate date if exists
    let certificateDisplayDate = null;
    if (cert && cert.issuedDate) {
      try {
        const certDate = new Date(cert.issuedDate);
        if (!isNaN(certDate.getTime())) {
          certificateDisplayDate = Utilities.formatDate(certDate, TIME_ZONE, 'MMM dd, yyyy');
        }
      } catch (e) {
        Logger.log(`WARNING: Certificate date formatting failed for ${cert.issuedDate}`);
      }
    }

    results.push({
      recordId: recordId,
      displayDate: displayDate, // NOW PROPERLY FORMATTED!
      dayType: dayType,
      hoursWorked: hoursWorked,
      cocEarned: cocEarned,
      certificateUrl: cert ? cert.certificateUrl : null,
      pdfUrl: cert ? cert.pdfUrl : null,
      certificateDisplayDate: certificateDisplayDate
    });
  }

  // Sort by date
  results.sort((a, b) => {
    // Extract dates from displayDate strings for sorting
    const dateA = new Date(a.displayDate);
    const dateB = new Date(b.displayDate);
    return dateA - dateB;
  });

  Logger.log(`Found ${results.length} records for employee ${employeeId} in ${month}/${year}`);
  return results;
}


/**
* API wrapper for listing employees. Accepts an optional boolean flag to
 * include inactive employees. Without parameters the list defaults to
 * active employees only. This change allows HTML clients to request
 * either view while remaining backwards compatible with existing calls.
 *
 * @param {boolean} includeInactive Whether to include inactive employees.
 * @return {Array<Object>} List of employee objects.
 */
function apiListEmployees(includeInactive) {
  // Explicitly coerce undefined/null to false. When a parameter is passed
  // through google.script.run it will arrive as the first argument.
  return listEmployees(Boolean(includeInactive));
}


/**
 * API: List employees for dropdown (simplified)
 * @return {Array<Object>} Array of employee objects with id and fullName
 */
function apiListEmployeesForDropdown() {
  try {
    const db = getDatabase();
    const empSheet = db.getSheetByName('Employees');
    
    if (!empSheet || empSheet.getLastRow() < 2) {
      return [];
    }

    const data = empSheet.getDataRange().getValues();
    const results = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const status = row[6]; // Assuming status is in column G
      
      // Only include active employees
      if (status !== 'Active') continue;
      
      results.push({
        id: row[0],        // Employee ID
        fullName: row[1]   // Full Name
      });
    }

    // Sort by name
    results.sort((a, b) => a.fullName.localeCompare(b.fullName));

    return results;

  } catch (error) {
    Logger.log('ERROR in apiListEmployeesForDropdown: ' + error.message + '\nStack: ' + error.stack);
    throw new Error('Failed to load employees: ' + error.message);
  }
}


/**
 * API: List all holidays
 * @return {Array<Object>} Array of holiday objects
 */
function apiListHolidays() {
  try {
    const db = getDatabase();
    const sheet = db.getSheetByName('Holidays');
    
    if (!sheet || sheet.getLastRow() < 2) {
      return [];
    }

    const data = sheet.getDataRange().getValues();
    const results = [];
    const TIME_ZONE = getScriptTimeZone();

    // Skip header row
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const date = row[0]; // Column A: Date
      
      if (!date) continue; // Skip empty rows

      results.push({
        rowNumber: i + 1,
        date: date,
        dateISO: Utilities.formatDate(new Date(date), TIME_ZONE, 'yyyy-MM-dd'),
        type: row[1] || '',           // Column B: Type
        description: row[2] || '',     // Column C: Description
        halfdayTime: row[3] || '',     // Column D: Half-day Start Time
        suspensionTime: row[4] || '',  // Column E: Work Suspension Time
        remarks: row[5] || ''          // Column F: Remarks
      });
    }

    // Sort by date (most recent first for better UX)
    results.sort((a, b) => new Date(b.date) - new Date(a.date));

    return results;

  } catch (error) {
    Logger.log('ERROR in apiListHolidays: ' + error.message + '\nStack: ' + error.stack);
    throw new Error('Failed to load holidays: ' + error.message);
  }
}


/**
 * API: Run migration
 */
function apiMigrateExistingInitialBalances() {
  return migrateExistingInitialBalances();
}


/**
 * Records new COC entries from the form submission.
 * This just saves the records as "Pending".
 * The "Generate Certificate" step will finalize them.
 */
function apiRecordCOC(employeeId, month, year, entries) {
  if (!employeeId || !month || !year) throw new Error("Employee ID, month, and year are required.");
  if (!entries || entries.length === 0) throw new Error("At least one entry is required.");

  const lock = LockService.getScriptLock();
  lock.waitLock(30000); // Wait 30 seconds

  try {
    const monthYear = `${year}-${String(month).padStart(2, '0')}`;
    const currentUser = getCurrentUserEmail();
    const now = new Date();
    
    // 1. Get Employee Details
    const empDetails = getEmployeeDetails(employeeId);
    if (!empDetails) throw new Error("Employee not found.");

    let totalNewCOC = 0;
    let addedCount = 0;
    const newRecordRows = [];

    // 2. Process each entry
    for (const entry of entries) {
      // Server-side validation and calculation
      const result = apiCalculateOvertimeForDate(year, month, entry.day, entry.amIn, entry.amOut, entry.pmIn, entry.pmOut);
      
      if (result.cocEarned > 0) {
        const recordId = generateUniqueId("COC-");
        const dateRendered = new Date(year, month - 1, entry.day);

        const newRow = new Array(22).fill(''); // Initialize empty row
        newRow[RECORD_COLS.RECORD_ID] = recordId;
        newRow[RECORD_COLS.EMPLOYEE_ID] = employeeId;
        newRow[RECORD_COLS.EMPLOYEE_NAME] = empDetails.fullName;
        newRow[RECORD_COLS.MONTH_YEAR] = monthYear;
        newRow[RECORD_COLS.DATE_RENDERED] = dateRendered;
        newRow[RECORD_COLS.DAY_TYPE] = result.dayType;
        newRow[RECORD_COLS.AM_IN] = entry.amIn;
        newRow[RECORD_COLS.AM_OUT] = entry.amOut;
        newRow[RECORD_COLS.PM_IN] = entry.pmIn;
        newRow[RECORD_COLS.PM_OUT] = entry.pmOut;
        newRow[RECORD_COLS.HOURS_WORKED] = result.hoursWorked;
        newRow[RECORD_COLS.MULTIPLIER] = result.multiplier;
        newRow[RECORD_COLS.COC_EARNED] = result.cocEarned;
        newRow[RECORD_COLS.CERTIFICATE_ID] = ''; // Will be set upon generation
        newRow[RECORD_COLS.DATE_RECORDED] = now;
        newRow[RECORD_COLS.EXPIRATION_DATE] = ''; // Will be set upon generation
        newRow[RECORD_COLS.STATUS] = STATUS_PENDING; // Set to Pending
        newRow[RECORD_COLS.APPROVED_BY] = ''; // Assumes an approval step, or can be auto-approved
        newRow[RECORD_COLS.APPROVED_DATE] = '';
        newRow[RECORD_COLS.CREATED_BY] = currentUser;
        newRow[RECORD_COLS.LAST_MODIFIED] = now;
        newRow[RECORD_COLS.MODIFIED_BY] = currentUser;

        newRecordRows.push(newRow);
        totalNewCOC += result.cocEarned;
        addedCount++;
      }
    }

    if (newRecordRows.length > 0) {
      const recordsSheet = SpreadsheetApp.openById(DATABASE_ID).getSheetByName('COC_Records');
      recordsSheet.getRange(recordsSheet.getLastRow() + 1, 1, newRecordRows.length, newRecordRows[0].length).setValues(newRecordRows);
    } else {
      throw new Error("No valid COC entries to record.");
    }

    return {
      added: addedCount,
      totalNewCOC: totalNewCOC
    };

  } catch (e) {
    Logger.log(`Error in apiRecordCOC: ${e}`);
    throw new Error(`Failed to record COC entries: ${e.message}`);
  } finally {
    lock.releaseLock();
  }
}


 * OLD VERSION - DEPRECATED
 * This function has been replaced by the newer version at line 6436+.
 * The newer version uses the updated schema with MONTH_YEAR column and improved logic.
 */
/*
function apiRecordCOCWithValidation_OLD(employeeId, month, year, entries) {
  const db = getDatabase();
  const TIME_ZONE = getScriptTimeZone();

  // Get employee
  const employee = getEmployeeById(employeeId);
  if (!employee) {
    throw new Error('Employee not found');
  }

  // Get settings
  const settings = getSettings();

  // Validate inputs
  if (!month || !year || !entries || entries.length === 0) {
    throw new Error('Invalid parameters: month, year, and entries are required');
  }

  // Calculate total COC that will be added
  let totalNewCOC = 0;
  const calculatedEntries = [];

  for (const entry of entries) {
    // Calculate overtime for each entry
    const date = new Date(year, month - 1, entry.day);
    const result = calculateOvertimeForDate(date, entry.amIn, entry.amOut, entry.pmIn, entry.pmOut);
    // --- REMOVED STRAY 'G' ---
    totalNewCOC += result.cocEarned;
    calculatedEntries.push({
      day: entry.day,
      ...result
    });
  }

  // === START OF NEWLY INSERTED CODE ===
  // === CHECK FOR DUPLICATE DATES ===
  const cocRecordsSheet = db.getSheetByName('COC_Records');
  if (!cocRecordsSheet) throw new Error('COC_Records sheet not found');

  const recordsData = cocRecordsSheet.getDataRange().getValues();
  const existingDates = {};
  const duplicateDetails = [];

  // Build map of existing dates for this employee (exclude cancelled)
  for (let i = 1; i < recordsData.length; i++) {
    const row = recordsData[i];
    const empId = row[1];
    const status = row[15];

    if (empId === employeeId && status !== 'Cancelled') {
      const dateRendered = new Date(row[4]);
      const dateKey = Utilities.formatDate(dateRendered, TIME_ZONE, 'yyyy-MM-dd');

      existingDates[dateKey] = {
        date: dateRendered,
        dayType: row[5],
        hoursWorked: parseFloat(row[10]) || 0,
        cocEarned: parseFloat(row[12]) || 0
      };
    }
  }

  // Check each entry for duplicates
  calculatedEntries.forEach(calc => {
    const dateToCheck = new Date(year, month - 1, calc.day);
    const dateKey = Utilities.formatDate(dateToCheck, TIME_ZONE, 'yyyy-MM-dd');

    if (existingDates[dateKey]) {
      const existing = existingDates[dateKey];
      duplicateDetails.push({
        date: Utilities.formatDate(existing.date, TIME_ZONE, 'MMMM dd, yyyy'),
        dayType: existing.dayType,
        hoursWorked: existing.hoursWorked,
        cocEarned: existing.cocEarned
      });
    }
  });

  // If duplicates found, throw enhanced error message
  if (duplicateDetails.length > 0) {
    let errorMsg = 'The following date(s) already have COC records:\n\n';

    duplicateDetails.forEach(dup => {
      errorMsg += `• <strong>${dup.date}</strong> (${dup.dayType}, `;
      errorMsg += `<span style="color: #dc2626; font-weight: bold;">${dup.hoursWorked.toFixed(1)} hours</span>, `;
      errorMsg += `<span style="color: #16a34a; font-weight: bold;">${dup.cocEarned.toFixed(1)} COC</span>)\n`;
    });

    errorMsg += '\n<strong>To update these records:</strong>\n';
    errorMsg += '<ol style="margin: 8px 0; padding-left: 20px;">';
    errorMsg += '<li>Remove these entries from your current form, OR</li>';
    errorMsg += '<li>Delete the existing records from the database first, then resubmit</li>';
    errorMsg += '</ol>\n';
    errorMsg += '<div style="background-color: #fef2f2; border-left: 4px solid #dc2626; padding: 12px; margin-top: 12px;">';
    errorMsg += '<strong>Note:</strong> You cannot have multiple entries for the same date.';
    errorMsg += '</div>';

    throw new Error(errorMsg);
  }
  // === END OF NEWLY INSERTED CODE ===


  // === VALIDATION: Check COC limits ===
  // Use MM-yyyy format consistently
  const monthYear = String(month).padStart(2, '0') + '-' + String(year);

  // Check monthly limit
  const monthlyCheck = checkMonthlyLimitForMonth(employeeId, monthYear, totalNewCOC);
  if (!monthlyCheck.valid) {
    throw new Error(monthlyCheck.message);
  }

  // Check total balance limit
  const balanceCheck = checkTotalBalanceLimitForEmployee(employeeId, totalNewCOC);
  if (!balanceCheck.valid) {
    throw new Error(balanceCheck.message);
  }

  // === Validations passed, proceed with recording ===

  // const cocRecordsSheet = db.getSheetByName('COC_Records'); // <-- This is now defined above in the new code
  const ledgerSheet = db.getSheetByName('COC_Ledger');
  const detailSheet = ensureCOCBalanceDetailSheet();

  // if (!cocRecordsSheet) throw new Error('COC_Records sheet not found'); // <-- This is now defined above
  if (!ledgerSheet) throw new Error('COC_Ledger sheet not found');

  let runningBalance = getCurrentCOCBalanceFromDetail(employeeId);
  const recordRows = [];
  const ledgerRows = [];
  const detailRows = [];

  // Process each entry
  calculatedEntries.forEach(calc => {
    // Skip entries that didn't earn any COC
    if (calc.cocEarned <= 0) return;

    const recordId = generateRecordId();
    const dateRendered = new Date(year, month - 1, calc.day);
    const monthYearEarned = Utilities.formatDate(dateRendered, TIME_ZONE, 'MM-yyyy');

    runningBalance += calc.cocEarned;

    // COC_Records row
    recordRows.push([
      recordId,
      employeeId,
      employee.fullName,
      monthYearEarned,
      dateRendered,
      calc.dayType,
      calc.amIn || '',
      calc.amOut || '',
      calc.pmIn || '',
      calc.pmOut || '',
      calc.hoursWorked,
      calc.multiplier,
      calc.cocEarned,
      new Date(),
      '',
      'Active'
    ]);

    // COC_Ledger row
    const ledgerId = generateLedgerId();
    ledgerRows.push([
      ledgerId,
      employeeId,
      employee.fullName,
      new Date(),
      'COC Earned',
      recordId,
      calc.cocEarned,
      0,
      runningBalance,
      monthYearEarned,
      '',
      Session.getActiveUser().getEmail(),
      `COC earned on ${Utilities.formatDate(dateRendered, TIME_ZONE, 'MMMM dd, yyyy')} (${calc.dayType}, ${calc.hoursWorked.toFixed(2)} hrs × ${calc.multiplier})`
    ]);

    // COC_Balance_Detail row (for FIFO tracking)
    detailRows.push([
      recordId,
      employeeId,
      employee.fullName,
      monthYearEarned,
      '',
      calc.cocEarned,
      0,
      calc.cocEarned,
      'Active',
      '',
      '',
      new Date(),
      Session.getActiveUser().getEmail(),
      `COC earned from ${calc.dayType} overtime. ${monthlyCheck.message}`
    ]);
  });

  // Batch write to sheets
  if (recordRows.length > 0) {
    cocRecordsSheet.getRange(cocRecordsSheet.getLastRow() + 1, 1, recordRows.length, recordRows[0].length)
      .setValues(recordRows);
  }

  if (ledgerRows.length > 0) {
    ledgerSheet.getRange(ledgerSheet.getLastRow() + 1, 1, ledgerRows.length, ledgerRows[0].length)
      .setValues(ledgerRows);
  }

  if (detailRows.length > 0) {
    detailSheet.getRange(detailSheet.getLastRow() + 1, 1, detailRows.length, detailRows[0].length)
      .setValues(detailRows);
  }

  return {
    success: true,
    added: calculatedEntries.length,
    totalNewCOC: totalNewCOC,
    balanceAfter: runningBalance
  };
}


function apiRecordCOC_OLD_WRAPPER(employeeId, month, year, entries) {
  return apiRecordCOCWithValidation_OLD(employeeId, month, year, entries);
}


function apiRecordCTO(employeeId, hours, startDate, endDate, remarks) {
  return recordCTOApplication(employeeId, hours, startDate, endDate, remarks);
}


/**
 * API: Record CTO with FIFO
 */
function apiRecordCTOWithFIFO(employeeId, hours, startDate, endDate, remarks) {
  return recordCTOApplicationWithFIFO(employeeId, hours, startDate, endDate, remarks);
}


/**
 * Saves the signatories to the Settings sheet.
 */
function apiSaveSignatories(signatories) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const db = SpreadsheetApp.openById(DATABASE_ID);
    const sheet = db.getSheetByName('Signatories'); // Changed from 'Settings'
    const data = sheet.getDataRange().getValues();

    const settingsToUpdate = {
      "SIGNATORY_ISSUED_BY_NAME": signatories.issuedByName,
      "SIGNATORY_ISSUED_BY_POSITION": signatories.issuedByPosition
      // REMOVED: notedByName and notedByPosition
    };

    const keysToUpdate = Object.keys(settingsToUpdate);
    let updatedCount = 0;

    // Update existing keys
    for (let i = 1; i < data.length; i++) { // Start from row 2 (index 1)
      const key = data[i][0];
      if (keysToUpdate.includes(key)) {
        sheet.getRange(i + 1, 2).setValue(settingsToUpdate[key]); // Column 2 (B) is the value
        sheet.getRange(i + 1, 4).setValue(new Date()); // Column 4 (D) is Last Updated
        sheet.getRange(i + 1, 5).setValue(getCurrentUserEmail()); // Column 5 (E) is Updated By
        
        // Remove key from list once updated
        keysToUpdate.splice(keysToUpdate.indexOf(key), 1);
        updatedCount++;
      }
    }

    // Add new keys if they didn't exist
    const newRows = [];
    for (const key of keysToUpdate) {
      newRows.push([
        key,
        settingsToUpdate[key],
        `Signatory for COC Certificates`, // Description
        new Date(), // Last Updated
        getCurrentUserEmail() // Updated By
      ]);
    }

    if (newRows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
    }
    
    return { success: true, updated: updatedCount, added: newRows.length };

  } catch (e) {
    Logger.log(`Error in apiSaveSignatories: ${e}`);
    throw new Error(`Failed to save signatories: ${e.message}`);
  } finally {
    lock.releaseLock();
  }
}


function apiUpdateCTOApplication(applicationId, newHours, newStartDate, newEndDate, newRemarks) {
  try {
    Logger.log('=== Updating CTO Application: ' + applicationId + ' ===');
    Logger.log('New hours: ' + newHours);
    Logger.log('New dates: ' + newStartDate + ' to ' + newEndDate);

    const db = getDatabase();
    const ctoSheet = db.getSheetByName('CTO_Applications');
    const ledgerSheet = db.getSheetByName('COC_Ledger'); // Ensure ledger sheet is available

    if (!ctoSheet) throw new Error('CTO_Applications sheet not found');
    if (!ledgerSheet) throw new Error('COC_Ledger sheet not found'); // Added check

    // Find the application
    const ctoData = ctoSheet.getDataRange().getValues();
    let appRow = -1;
    let application = null;

    for (let i = 1; i < ctoData.length; i++) {
      if (ctoData[i][0] === applicationId) {
        appRow = i + 1; // Sheet row (1-indexed)
        application = {
          id: ctoData[i][0],
          employeeId: ctoData[i][1],
          employeeName: ctoData[i][2],
          oldHours: parseFloat(ctoData[i][4]) || 0,
          oldStartDate: ctoData[i][5], // Keep original format for comparison/logging
          oldEndDate: ctoData[i][6],   // Keep original format for comparison/logging
          oldRemarks: ctoData[i][11],
          status: ctoData[i][10]
        };
        break;
      }
    }

    if (!application) {
      throw new Error('CTO application not found: ' + applicationId);
    }

    // Check if already cancelled
    if (application.status === 'Cancelled') {
      return {
        success: false,
        message: 'Cannot update a cancelled application'
      };
    }

    // Allow updating Approved or Pending
     if (application.status !== 'Pending' && application.status !== 'Approved') {
        return {
            success: false,
            message: 'Only Pending or Approved applications can be updated. This application is ' + application.status
        };
    }


    Logger.log('Found application: ' + application.employeeName);
    Logger.log('Old hours: ' + application.oldHours + ', New hours: ' + newHours);

    // --- START VALIDATION ---
    const validationError = validateCTOUpdate(application, newHours, newStartDate, newEndDate);
     if (validationError) {
        return { success: false, message: validationError };
    }
    // --- END VALIDATION ---


    // --- Apply Updates ---
    const TIME_ZONE = getScriptTimeZone();
    const updateTimestamp = new Date(); // Use consistent timestamp for updates

    // Update CTO_Applications sheet
    ctoSheet.getRange(appRow, 5).setValue(newHours); // Column E (Hours)
    ctoSheet.getRange(appRow, 6).setValue(new Date(newStartDate)); // Column F (Start Date) - Ensure it's a Date object
    ctoSheet.getRange(appRow, 7).setValue(new Date(newEndDate));   // Column G (End Date) - Ensure it's a Date object
    // Append update info to remarks
    const updateRemark = `\n[Updated on ${Utilities.formatDate(updateTimestamp, TIME_ZONE, 'yyyy-MM-dd HH:mm')} by ${Session.getActiveUser().getEmail()}] ${newRemarks || '(No additional remarks)'}`;
    ctoSheet.getRange(appRow, 12).setValue((application.oldRemarks || '') + updateRemark); // Column L (Remarks)

    Logger.log('✓ CTO application updated in database');

    // --- Ledger and FIFO Adjustments (Only if status is Approved and hours changed) ---
     let hoursDifference = 0;
     if (application.status === 'Approved') {
        hoursDifference = newHours - application.oldHours;
        Logger.log(`Status is Approved. Hours difference: ${hoursDifference}`);

        if (hoursDifference !== 0) {
             // If hours increased, deduct using FIFO
            if (hoursDifference > 0) {
                try {
                    deductCOCHoursFIFO(application.employeeId, hoursDifference, applicationId + '-UPDATE'); // Use a distinct reference for update adjustment
                    Logger.log('✓ Deducted additional ' + hoursDifference + ' hours using FIFO for update.');
                } catch (deductError) {
                     Logger.log('ERROR during FIFO deduction for update: ' + deductError.message);
                     // Rollback changes? Or just return error? For now, return error.
                     // It might be better to revert the CTO sheet change here.
                    return { success: false, message: 'Failed to apply update: ' + deductError.message };
                }
            }
            // If hours decreased, restore using FIFO
            else if (hoursDifference < 0) {
                try {
                    restoreCOCHoursFIFO(application.employeeId, Math.abs(hoursDifference), applicationId); // Restore based on original ID
                    Logger.log('✓ Restored ' + Math.abs(hoursDifference) + ' hours to FIFO due to update.');
                } catch (restoreError) {
                     Logger.log('ERROR during FIFO restoration for update: ' + restoreError.message);
                     // Rollback changes? Or just return error? For now, return error.
                    return { success: false, message: 'Failed to apply update: ' + restoreError.message };
                }
            }
        }
    } else {
         Logger.log('Status is Pending. No ledger/FIFO adjustments needed yet.');
    }


    // --- Add a specific Ledger Entry for the UPDATE action itself ---
    const currentBalance = getCurrentCOCBalance(application.employeeId); // Get balance AFTER potential adjustments
    const updateLedgerRemark = `CTO Updated: ${applicationId}. ` +
                             (hoursDifference !== 0 ? `Hours changed ${application.oldHours} -> ${newHours} (${hoursDifference > 0 ? '+' : ''}${hoursDifference.toFixed(2)}). ` : '') +
                             `Dates changed. ${newRemarks || ''}`;

    const updateLedgerEntry = [
      generateLedgerEntryId(),
      application.employeeId,
      application.employeeName,
      updateTimestamp, // Use the consistent timestamp
      'CTO Updated', // New Transaction Type
      applicationId, // Reference the original CTO ID
      (hoursDifference < 0) ? Math.abs(hoursDifference) : 0, // Restore is like earning COC back
      (hoursDifference > 0) ? hoursDifference : 0, // Increase is like using more CTO
      currentBalance,
      '', // Month-Year Earned (N/A for update)
      '', // Expiration Date (N/A for update)
      Session.getActiveUser().getEmail(),
      updateLedgerRemark.trim() // Trim potential extra space at the end
    ];
    ledgerSheet.appendRow(updateLedgerEntry);
    Logger.log('✓ Added "CTO Updated" entry to ledger');
    // --- End Ledger Update ---


    Logger.log('=== CTO Update Complete ===');

    return {
      success: true,
      message: 'CTO application updated successfully'
    };

  } catch (error) {
    Logger.log('ERROR in apiUpdateCTOApplication: ' + error.message);
    Logger.log('Stack: ' + error.stack);

    return {
      success: false,
      message: 'Failed to update CTO application: ' + error.message
    };
  }
}


function apiUpdateEmployee(employeeId, data) {
  return updateEmployee(employeeId, data);
}


/**
 * API: Update an existing holiday
 * @param {number} rowNumber The row number to update
 * @param {Date} date The date of the holiday
 * @param {string} type The type of holiday
 * @param {string} description Optional description
 * @param {string} halfdayTime Optional time for half-day holidays
 * @param {string} suspensionTime Optional time for work suspension
 * @param {string} remarks Optional additional remarks
 * @return {Object} Success result
 */
function apiUpdateHoliday(rowNumber, date, type, description, halfdayTime, suspensionTime, remarks) {
  try {
    const db = getDatabase();
    const sheet = db.getSheetByName('Holidays');
    
    if (!sheet) {
      throw new Error('Holidays sheet not found');
    }

    // Validate inputs
    if (!rowNumber || rowNumber < 2) {
      throw new Error('Invalid row number');
    }

    if (!date || !(date instanceof Date)) {
      throw new Error('Invalid date provided');
    }

    if (!type) {
      throw new Error('Holiday type is required');
    }

    // Check for duplicate date (excluding current row)
    const existingData = sheet.getDataRange().getValues();
    const TIME_ZONE = getScriptTimeZone();
    const dateStr = Utilities.formatDate(new Date(date), TIME_ZONE, 'yyyy-MM-dd');
    
    for (let i = 1; i < existingData.length; i++) {
      if (i + 1 === rowNumber) continue; // Skip current row
      const existingDate = existingData[i][0];
      if (existingDate && Utilities.formatDate(new Date(existingDate), TIME_ZONE, 'yyyy-MM-dd') === dateStr) {
        throw new Error('A holiday already exists for this date. Please choose a different date.');
      }
    }

    // Update the row
    const range = sheet.getRange(rowNumber, 1, 1, 6);
    range.setValues([[
      new Date(date),
      type,
      description || '',
      halfdayTime || '',
      suspensionTime || '',
      remarks || ''
    ]]);

    Logger.log('Holiday updated successfully: ' + dateStr + ' - ' + type);
    return { success: true, message: 'Holiday updated successfully' };

  } catch (error) {
    Logger.log('ERROR in apiUpdateHoliday: ' + error.message + '\nStack: ' + error.stack);
    throw new Error('Failed to update holiday: ' + error.message);
  }
}


/**
 * API: Run FIFO integrity check for all employees or specific employee
 * @param {string} employeeId Optional - specific employee to check, or null for all
 * @return {Object} Report with findings and discrepancies
 */
function apiFIFOIntegrityCheck(employeeId) {
  try {
    const db = getDatabase();
    const detailSheet = db.getSheetByName('COC_Balance_Detail');
    const ctoSheet = db.getSheetByName('CTO_Applications');
    
    if (!detailSheet) {
      throw new Error('COC_Balance_Detail sheet not found');
    }
    
    if (!ctoSheet) {
      throw new Error('CTO_Applications sheet not found');
    }

    Logger.log('=== Starting FIFO Integrity Check ===');
    Logger.log('Employee filter: ' + (employeeId || 'ALL'));

    const detailData = detailSheet.getDataRange().getValues();
    const ctoData = ctoSheet.getDataRange().getValues();
    
    const report = {
      checkDate: new Date(),
      employeeId: employeeId || 'ALL',
      totalEmployeesChecked: 0,
      totalDiscrepancies: 0,
      discrepancyDetails: [],
      integrityIssues: [],
      summary: ''
    };

    // Get unique employees to check
    const employeesToCheck = new Set();
    for (let i = 1; i < detailData.length; i++) {
      const empId = detailData[i][DETAIL_COLS.EMPLOYEE_ID];
      if (!employeeId || empId === employeeId) {
        employeesToCheck.add(empId);
      }
    }

    // Check each employee
    for (const empId of employeesToCheck) {
      Logger.log(`Checking employee: ${empId}`);
      const employeeIssues = checkEmployeeFIFO(empId, detailData, ctoData);
      
      if (employeeIssues.length > 0) {
        report.totalDiscrepancies++;
        report.integrityIssues.push({
          employeeId: empId,
          issues: employeeIssues
        });
      }
      
      report.totalEmployeesChecked++;
    }

    // Generate summary
    if (report.totalDiscrepancies === 0) {
      report.summary = `✓ All ${report.totalEmployeesChecked} employee(s) passed FIFO integrity check. No discrepancies found.`;
    } else {
      report.summary = `⚠ Found FIFO discrepancies for ${report.totalDiscrepancies} out of ${report.totalEmployeesChecked} employee(s).`;
    }

    Logger.log('=== FIFO Integrity Check Complete ===');
    Logger.log(report.summary);

    return report;

  } catch (error) {
    Logger.log('ERROR in apiFIFOIntegrityCheck: ' + error.message + '\nStack: ' + error.stack);
    throw new Error('FIFO integrity check failed: ' + error.message);
  }
}


/**
 * API: Fix FIFO discrepancies for an employee
 * This will recalculate all COC balances using proper FIFO logic
 * @param {string} employeeId Employee ID to fix
 * @param {boolean} dryRun If true, only simulate the fix without applying changes
 * @return {Object} Result of the fix operation
 */
function apiFIFOFix(employeeId, dryRun) {
  try {
    if (!employeeId) {
      throw new Error('Employee ID is required');
    }

    const db = getDatabase();
    const detailSheet = db.getSheetByName('COC_Balance_Detail');
    const ctoSheet = db.getSheetByName('CTO_Applications');
    
    Logger.log(`=== ${dryRun ? 'Simulating' : 'Executing'} FIFO Fix for ${employeeId} ===`);

    const detailData = detailSheet.getDataRange().getValues();
    const ctoData = ctoSheet.getDataRange().getValues();

    // Get all COC entries for this employee
    const cocEntries = [];
    for (let i = 1; i < detailData.length; i++) {
      const row = detailData[i];
      if (row[DETAIL_COLS.EMPLOYEE_ID] === employeeId) {
        let certificateDate = row[DETAIL_COLS.CERTIFICATE_DATE];
        if (!(certificateDate instanceof Date) || isNaN(certificateDate.getTime())) {
          certificateDate = parseMonthYear(row[DETAIL_COLS.MONTH_YEAR]) || new Date(row[DETAIL_COLS.DATE_CREATED]);
        }

        cocEntries.push({
          rowNumber: i + 1,
          recordId: row[DETAIL_COLS.RECORD_ID],
          certificateDate: certificateDate,
          hoursEarned: parseFloat(row[DETAIL_COLS.HOURS_EARNED]) || 0,
          currentUsed: parseFloat(row[DETAIL_COLS.HOURS_USED]) || 0,
          currentRemaining: parseFloat(row[DETAIL_COLS.HOURS_REMAINING]) || 0,
          status: row[DETAIL_COLS.STATUS]
        });
      }
    }

    // Sort by certificate date (FIFO)
    cocEntries.sort((a, b) => a.certificateDate - b.certificateDate);

    // Reset all to unused state (for recalculation)
    cocEntries.forEach(entry => {
      entry.newUsed = 0;
      entry.newRemaining = entry.hoursEarned;
      entry.newStatus = 'Active';
    });

    // Get all approved CTO applications
    const ctoApplications = [];
    for (let i = 1; i < ctoData.length; i++) {
      const row = ctoData[i];
      if (row[1] === employeeId && row[8] === 'Approved') {
        ctoApplications.push({
          applicationId: row[0],
          applicationDate: new Date(row[2]),
          hoursRequested: parseFloat(row[3]) || 0
        });
      }
    }

    // Sort by application date
    ctoApplications.sort((a, b) => a.applicationDate - b.applicationDate);

    // Apply FIFO deduction
    for (const cto of ctoApplications) {
      let remainingToDeduct = cto.hoursRequested;
      
      for (let i = 0; i < cocEntries.length && remainingToDeduct > 0.01; i++) {
        if (cocEntries[i].newRemaining > 0) {
          const hoursToDeduct = Math.min(cocEntries[i].newRemaining, remainingToDeduct);
          cocEntries[i].newUsed += hoursToDeduct;
          cocEntries[i].newRemaining -= hoursToDeduct;
          
          // Update status
          if (cocEntries[i].newRemaining < 0.01) {
            cocEntries[i].newStatus = 'Fully Used';
          } else if (cocEntries[i].newUsed > 0.01) {
            cocEntries[i].newStatus = 'Partially Used';
          }
          
          remainingToDeduct -= hoursToDeduct;
        }
      }
    }

    // Prepare result
    const result = {
      employeeId: employeeId,
      dryRun: dryRun,
      changes: [],
      summary: ''
    };

    // Compare and apply changes
    for (const entry of cocEntries) {
      const usedChanged = Math.abs(entry.currentUsed - entry.newUsed) > 0.01;
      const remainingChanged = Math.abs(entry.currentRemaining - entry.newRemaining) > 0.01;
      const statusChanged = entry.status !== entry.newStatus;
      
      if (usedChanged || remainingChanged || statusChanged) {
        result.changes.push({
          recordId: entry.recordId,
          certificateDate: entry.certificateDate,
          before: {
            used: entry.currentUsed.toFixed(2),
            remaining: entry.currentRemaining.toFixed(2),
            status: entry.status
          },
          after: {
            used: entry.newUsed.toFixed(2),
            remaining: entry.newRemaining.toFixed(2),
            status: entry.newStatus
          }
        });

        // Apply changes if not dry run
        if (!dryRun) {
          detailSheet.getRange(entry.rowNumber, DETAIL_COLS.HOURS_USED + 1).setValue(entry.newUsed);
          detailSheet.getRange(entry.rowNumber, DETAIL_COLS.HOURS_REMAINING + 1).setValue(entry.newRemaining);
          detailSheet.getRange(entry.rowNumber, DETAIL_COLS.STATUS + 1).setValue(entry.newStatus);
        }
      }
    }

    if (result.changes.length === 0) {
      result.summary = '✓ No FIFO discrepancies found. No changes needed.';
    } else {
      result.summary = `${dryRun ? 'Would update' : 'Updated'} ${result.changes.length} COC record(s) to match FIFO order.`;
    }

    Logger.log(result.summary);
    return result;

  } catch (error) {
    Logger.log('ERROR in apiFIFOFix: ' + error.message + '\nStack: ' + error.stack);
    throw new Error('FIFO fix failed: ' + error.message);
  }
}


/**
 * API: Get detailed FIFO report for a specific employee
 * @param {string} employeeId Employee ID
 * @return {Object} Detailed report
 */
function apiFIFOEmployeeReport(employeeId) {
  try {
    if (!employeeId) {
      throw new Error('Employee ID is required');
    }

    const db = getDatabase();
    const detailSheet = db.getSheetByName('COC_Balance_Detail');
    const ctoSheet = db.getSheetByName('CTO_Applications');
    const TIME_ZONE = getScriptTimeZone();

    const detailData = detailSheet.getDataRange().getValues();
    const ctoData = ctoSheet.getDataRange().getValues();

    // Get employee info
    const employee = getEmployeeById(employeeId);
    if (!employee) {
      throw new Error('Employee not found');
    }

    // Get COC entries
    const cocEntries = [];
    for (let i = 1; i < detailData.length; i++) {
      const row = detailData[i];
      if (row[DETAIL_COLS.EMPLOYEE_ID] === employeeId) {
        const certificateDateVal = row[DETAIL_COLS.CERTIFICATE_DATE];
        const expirationVal = row[DETAIL_COLS.EXPIRATION_DATE];
        const certificateDate = certificateDateVal instanceof Date && !isNaN(certificateDateVal.getTime())
          ? certificateDateVal
          : parseMonthYear(row[DETAIL_COLS.MONTH_YEAR]) || new Date(row[DETAIL_COLS.DATE_CREATED]);
        const expirationDate = expirationVal instanceof Date && !isNaN(expirationVal.getTime())
          ? expirationVal
          : null;

        cocEntries.push({
          recordId: row[DETAIL_COLS.RECORD_ID],
          monthYear: row[DETAIL_COLS.MONTH_YEAR],
          certificateDate: Utilities.formatDate(certificateDate, TIME_ZONE, 'MMM dd, yyyy'),
          hoursEarned: Number(row[DETAIL_COLS.HOURS_EARNED]).toFixed(2),
          hoursUsed: Number(row[DETAIL_COLS.HOURS_USED]).toFixed(2),
          hoursRemaining: Number(row[DETAIL_COLS.HOURS_REMAINING]).toFixed(2),
          status: row[DETAIL_COLS.STATUS],
          expirationDate: expirationDate ? Utilities.formatDate(expirationDate, TIME_ZONE, 'MMM dd, yyyy') : '—'
        });
      }
    }

    // Sort by certificate date
    cocEntries.sort((a, b) => new Date(a.certificateDate) - new Date(b.certificateDate));

    // Get CTO applications
    const ctoApps = [];
    for (let i = 1; i < ctoData.length; i++) {
      const row = ctoData[i];
      if (row[1] === employeeId && row[8] === 'Approved') {
        ctoApps.push({
          applicationId: row[0],
          applicationDate: Utilities.formatDate(new Date(row[2]), TIME_ZONE, 'MMM dd, yyyy'),
          hoursUsed: Number(row[3]).toFixed(2),
          status: row[8]
        });
      }
    }

    const report = {
      employeeId: employeeId,
      employeeName: employee.fullName,
      totalCOCEarned: cocEntries.reduce((sum, e) => sum + parseFloat(e.hoursEarned), 0).toFixed(2),
      totalCOCUsed: cocEntries.reduce((sum, e) => sum + parseFloat(e.hoursUsed), 0).toFixed(2),
      totalCOCRemaining: cocEntries.reduce((sum, e) => sum + parseFloat(e.hoursRemaining), 0).toFixed(2),
      cocEntries: cocEntries,
      ctoApplications: ctoApps,
      integrityCheck: checkEmployeeFIFO(employeeId, detailData, ctoData)
    };

    return report;

  } catch (error) {
    Logger.log('ERROR in apiFIFOEmployeeReport: ' + error.message + '\nStack: ' + error.stack);
    throw new Error('Failed to generate FIFO report: ' + error.message);
  }
}


