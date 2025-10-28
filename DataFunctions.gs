// -----------------------------------------------------------------------------
// DataFunctions.gs
// -----------------------------------------------------------------------------

/**
 * Returns the database spreadsheet object. All data reads/writes should go
 * through this helper so that future changes (like migrating to another file)
 * require editing only one location.
 *
 * @return {Spreadsheet} The database spreadsheet.
 */
function getDatabase() {
  try {
    if (DATABASE_ID && typeof DATABASE_ID === 'string' && DATABASE_ID !== 'REPLACE_WITH_DATABASE_SHEET_ID') {
      // --- MODIFICATION: Added error logging ---
      const db = SpreadsheetApp.openById(DATABASE_ID);
      db.getName(); // This simple call will trigger auth/permission errors
      return db;
      // --- END MODIFICATION ---
    }
  } catch (err) {
    // --- MODIFICATION: Added a visible log for debugging ---
    // This will write the *actual* error to the Apps Script console.
    Logger.log('WARNING: Could not open database by ID (' + DATABASE_ID + '). Error: ' + err.message + '. Falling back to active spreadsheet.');
    // --- END MODIFICATION ---
  }
  // Fall back to the active spreadsheet if openById fails
  return SpreadsheetApp.getActive();
}


/**
 * Gets all data from a sheet excluding the header row.
 * Returns an array of arrays representing the data rows.
 *
 * @param {string} sheetName The name of the sheet to retrieve data from
 * @return {Array<Array>} Array of row arrays (excluding header)
 */
function getSheetDataNoHeader(sheetName) {
  const db = getDatabase();
  const sheet = db.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" not found in database.`);
  }

  const data = sheet.getDataRange().getValues();

  // Return empty array if sheet only has header or is empty
  if (data.length <= 1) {
    return [];
  }

  // Return all rows except the first one (header)
  return data.slice(1);
}


/**
 * Creates or ensures the COC_Balance_Detail sheet exists
 * This sheet tracks individual COC entries for FIFO consumption
 */
function ensureCOCBalanceDetailSheet() {
  const db = getDatabase();
  let detailSheet = db.getSheetByName('COC_Balance_Detail');

  if (!detailSheet) {
    detailSheet = db.insertSheet('COC_Balance_Detail');

    const headers = [
      'Record ID', 'Employee ID', 'Employee Name', 'Month-Year Earned',
      'Certificate Date', 'Hours Earned', 'Hours Used', 'Hours Remaining',
      'Status', 'Expiration Date', 'Certificate ID', 'Date Created',
      'Created By', 'Remarks'
    ];

    detailSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    detailSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    detailSheet.setFrozenRows(1);
  }

  return detailSheet;
}


/**
 * Ensure Ledger sheet exists
 * @return {Sheet} The ledger sheet
 */
function ensureLedgerSheet() {
  const db = getDatabase();
  let sheet = db.getSheetByName('Ledger');
  
  if (!sheet) {
    sheet = db.insertSheet('Ledger');
    // Add headers
    const headers = [
      'Ledger ID', 'Employee ID', 'Employee Name', 'Transaction Date', 
      'Transaction Type', 'Reference ID', 'COC Earned', 'CTO Used', 
      'Running Balance', 'Month-Year Earned', 'Expiration Date', 
      'Processed By', 'Remarks'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    Logger.log('Ledger sheet created');
  }
  
  return sheet;
}


/**
 * Retrieves an employee record by ID. The returned object contains the row
 * number to facilitate updates along with all employee fields. Returns null
 * if the employee does not exist.
 *
 * @param {string} employeeId The employee ID (e.g. EMP001).
 * @return {Object|null} The employee record or null if not found.
 */
function getEmployeeById(employeeId) {
  Logger.log('getEmployeeById called for: ' + employeeId);
  const db = getDatabase();
  const sheet = db.getSheetByName('Employees');
  const TIME_ZONE = getScriptTimeZone(); // Added this for the formatter
  // Read up to column K
  const data = sheet.getRange(1, 1, sheet.getLastRow(), 11).getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === employeeId) {
      const row = data[i];
      const fullName = row[2] + (row[3] ? ' ' + row[3] + ' ' : ' ') + row[1] + (row[4] ? ' ' + row[4] : '');
      return {
        row: i + 1,
        id: row[0],
        lastName: row[1],
        firstName: row[2],
        middleInitial: row[3],
        suffix: row[4],
        position: row[5],
        office: row[6],
        email: row[7],
        status: row[8],
        initialBalance: parseFloat(row[9]) || 0,
        initialBalanceDate: row[10] ? Utilities.formatDate(new Date(row[10]), TIME_ZONE, 'yyyy-MM-dd') : '',
        fullName: fullName.trim()
      };
    }
  }
  return null;
}


/**
 * Gets full employee details from the 'Employees' sheet.
 * @param {string} employeeId The Employee ID.
 * @returns {object} An object with fullName, position, and office.
 */
function getEmployeeDetails(employeeId) {
  const data = getSheetDataNoHeader('Employees');
  const empRow = data.find(r => r[EMP_COLS.EMPLOYEE_ID] === employeeId);
  if (!empRow) return null;
  
  const middle = empRow[EMP_COLS.MIDDLE_INITIAL] ? ` ${empRow[EMP_COLS.MIDDLE_INITIAL]}.` : '';
  const suffix = empRow[EMP_COLS.SUFFIX] ? ` ${empRow[EMP_COLS.SUFFIX]}` : '';
  
  return {
    fullName: `${empRow[EMP_COLS.FIRST_NAME]}${middle} ${empRow[EMP_COLS.LAST_NAME]}${suffix}`,
    position: empRow[EMP_COLS.POSITION],
    office: empRow[EMP_COLS.OFFICE]
  };
}


/**
 * Returns a list of employees. Optionally exclude inactive employees. Each
 * employee object includes the employee ID, full name, position, office,
 * email, status, *total* initial COC balance, and the earliest "as-of" date
 * for that balance.
 *
 * @param {boolean} includeInactive When true inactive employees are also
 * included in the list. Default false.
 * @return {Array<Object>} An array of employee objects.
 */
function listEmployees(includeInactive) {
  // --- NEW DEBUG LOG ---
  Logger.log('apiListEmployees called. includeInactive = ' + includeInactive);
  // ---
  const db = getDatabase();
  // --- NEW DEBUG LOG ---
  // This will tell us if we are reading the correct spreadsheet.
  Logger.log('getDatabase() returned spreadsheet: ' + db.getName());
  // ---
  const sheet = db.getSheetByName('Employees');

  // --- NEW DEBUG LOG ---
  if (!sheet) {
    Logger.log('ERROR: "Employees" sheet was NOT FOUND in ' + db.getName());
    return []; // Return early if sheet not found
  }
  Logger.log('"Employees" sheet was found.');
  // ---

  // Read up to column K (InitialBalanceDate)
  // --- NEW DEBUG LOG ---
  const lastRow = sheet.getLastRow();
  Logger.log('Last row in "Employees" sheet is: ' + lastRow);
  if (lastRow < 2) {
      Logger.log('No data rows found (lastRow < 2). Returning empty list.');
      return [];
  }
  // ---
  const rows = sheet.getRange(1, 1, lastRow, 11).getValues();
  const employees = [];
  const TIME_ZONE = getScriptTimeZone(); // Added this for the formatter
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const id = row[0];
    const lastName = row[1];
    const firstName = row[2];
    const middleInitial = row[3];
    const suffix = row[4];
    const position = row[5];
    const office = row[6];
    const email = row[7];
    const status = row[8];
    const initialBalance = parseFloat(row[9]) || 0; // This is the TOTAL initial balance
    const initialBalanceDate = row[10] ? Utilities.formatDate(new Date(row[10]), TIME_ZONE, 'yyyy-MM-dd') : '';

    if (!includeInactive && status !== 'Active') continue;
    const fullName = firstName + (middleInitial ? ' ' + middleInitial + ' ' : ' ') + lastName + (suffix ? ' ' + suffix : '');

    employees.push({
      id: id,
      fullName: fullName.trim(),
      position: position,
      office: office,
      email: email,
      status: status,
      initialBalance: initialBalance,
      initialBalanceDate: initialBalanceDate, // This is the *earliest* date for reference
      lastName: lastName,
      firstName: firstName,
      middleInitial: middleInitial,
      suffix: suffix
    });
  }
  return employees;
}


/**
 * Adds a new employee to the database. (Original, non-FIFO version)
 *
 * --- MODIFIED (Option C) ---
 * - Now accepts `data.initialBalanceEntries` which is an array of objects:
 * `[{amount: 5.0, date: '2024-10-01'}, ...]`.
 * - It calculates the *total* initial balance and the *earliest* date to
 * store in the 'Employees' sheet (Cols J & K).
 * - It then loops through the `initialBalanceEntries` array and creates a
 * *separate ledger entry for each one*, correctly calculating expirations
 * and running balances.
 *
 * @param {Object} data An object with employee details and
 * `initialBalanceEntries` array.
 * @return {string} The generated employee ID.
 */
function addEmployee(data) {
  const db = getDatabase();
  const sheet = db.getSheetByName('Employees');
  if (!sheet) throw new Error('Employees sheet not found in database');

  const id = generateEmployeeId();
  const fullName = data.firstName + (data.middleInitial ? ' ' + data.middleInitial + ' ' : ' ') + data.lastName + (data.suffix ? ' ' + data.suffix : '');

  // --- NEW LOGIC for Option C ---
  const entries = data.initialBalanceEntries || [];

  // 1. Calculate Total Balance and Earliest Date
  let totalInitialBalance = 0;
  let earliestDate = null;
  const validDates = [];

  entries.forEach(entry => {
      const amount = parseFloat(entry.amount) || 0;
      if (amount > 0 && entry.date) {
          totalInitialBalance += amount;
          validDates.push(new Date(entry.date));
      }
  });

  if (validDates.length > 0) {
      // Find the minimum (earliest) date
      earliestDate = new Date(Math.min.apply(null, validDates));
  }
  // --- END NEW LOGIC ---

  // 2. Append to 'Employees' sheet
  sheet.appendRow([
    id,
    data.lastName,
    data.firstName,
    data.middleInitial || '',
    data.suffix || '',
    data.position,
    data.office,
    data.email,
    data.status || 'Active',
    totalInitialBalance, // Store the calculated total
    earliestDate // Store the earliest date for reference
  ]);

  // 3. Create multiple ledger entries if balance > 0
  if (totalInitialBalance > 0) {
    const ledgerSheet = db.getSheetByName('COC_Ledger');
    if (!ledgerSheet) throw new Error('COC_Ledger sheet not found');
    const settings = getSettings();
    const validityMonths = parseInt(settings['COC_VALIDITY_MONTHS']) || 12;
    const TIME_ZONE = getScriptTimeZone(); // Added this

    let runningBalance = 0; // New employee's balance starts at 0
    const ledgerRows = []; // To batch-write for efficiency

    // Sort entries by date to ensure correct running balance
    const sortedEntries = entries
      .filter(e => (parseFloat(e.amount) || 0) > 0 && e.date)
      .sort((a, b) => new Date(a.date) - new Date(b.date));

    sortedEntries.forEach(entry => {
      const amount = parseFloat(entry.amount);
      const earnedDate = new Date(entry.date);
      const ledgerId = generateLedgerId(); // Generates unique ID with milliseconds
      const monthYear = Utilities.formatDate(earnedDate, TIME_ZONE, 'MM-yyyy');

      const expiration = new Date(earnedDate);
      expiration.setMonth(expiration.getMonth() + validityMonths);

      runningBalance += amount; // Increment running balance

      const remarks = 'Initial COC balance (Earned: ' + formatDate(earnedDate) + ')';

      ledgerRows.push([
        ledgerId,
        id,
        fullName.trim(),
        new Date(), // transactionDate (today)
        'Initial Balance',
        'INIT-BALANCE-' + ledgerId.substring(4, 12), // Reference
        amount, // cocEarned
        0, // ctoUsed
        runningBalance, // cocBalance (cumulative)
        monthYear, // monthYearEarned
        expiration, // expirationDate
        Session.getActiveUser().getEmail(),
        remarks
      ]);
    });

    // 4. Batch-write all ledger rows
    if (ledgerRows.length > 0) {
       ledgerSheet.getRange(ledgerSheet.getLastRow() + 1, 1, ledgerRows.length, ledgerRows[0].length).setValues(ledgerRows);
    }
  }
  return id;
}


/**
 * MODIFIED: Enhanced addEmployee to create COC_Balance_Detail entries
 * This replaces the existing addEmployee function starting at line 345
 */
function addEmployeeWithFIFO(data) {
  const db = getDatabase();
  const sheet = db.getSheetByName('Employees');
  if (!sheet) throw new Error('Employees sheet not found in database');

  const id = generateEmployeeId();
  const fullName = data.firstName + (data.middleInitial ? ' ' + data.middleInitial + ' ' : ' ') +
                     data.lastName + (data.suffix ? ' ' + data.suffix : '');

  const entries = data.initialBalanceEntries || [];

  // Calculate Total Balance and Earliest Date
  let totalInitialBalance = 0;
  let earliestDate = null;
  const validDates = [];

  entries.forEach(entry => {
    const amount = parseFloat(entry.amount) || 0;
    if (amount > 0 && entry.date) {
      totalInitialBalance += amount;
      validDates.push(new Date(entry.date));
    }
  });

  if (validDates.length > 0) {
    earliestDate = new Date(Math.min.apply(null, validDates));
  }

  // Append to 'Employees' sheet
  sheet.appendRow([
    id,
    data.lastName,
    data.firstName,
    data.middleInitial || '',
    data.suffix || '',
    data.position,
    data.office,
    data.email,
    data.status || 'Active',
    totalInitialBalance,
    earliestDate
  ]);

  // Create ledger and detail entries
  if (totalInitialBalance > 0) {
    const ledgerSheet = db.getSheetByName('COC_Ledger');
    const detailSheet = ensureCOCBalanceDetailSheet();

    if (!ledgerSheet) throw new Error('COC_Ledger sheet not found');

    const settings = getSettings();
    const validityMonths = parseInt(settings['COC_VALIDITY_MONTHS']) || 12;
    const TIME_ZONE = getScriptTimeZone();

    let runningBalance = 0;
    const ledgerRows = [];
    const detailRows = [];

    // Sort entries by date
    const sortedEntries = entries
      .filter(e => (parseFloat(e.amount) || 0) > 0 && e.date)
      .sort((a, b) => new Date(a.date) - new Date(b.date));

    sortedEntries.forEach(entry => {
      const amount = parseFloat(entry.amount);
      const earnedDate = new Date(entry.date);
      const ledgerId = generateLedgerId();
      const monthYear = Utilities.formatDate(earnedDate, TIME_ZONE, 'MM-yyyy');

      const certificateDate = new Date(earnedDate);
      const expiration = calculateCertificateExpiration(certificateDate);

      runningBalance += amount;

      const remarks = 'Initial COC balance (Earned: ' + formatDate(earnedDate) + ')';
      const recordId = 'INITIAL-' + id + '-' + Utilities.formatDate(earnedDate, TIME_ZONE, 'yyyyMM');

      // Ledger entry
      ledgerRows.push([
        ledgerId,
        id,
        fullName.trim(),
        new Date(),
        'Initial Balance',
        recordId,
        amount,
        0,
        runningBalance,
        monthYear,
        expiration,
        Session.getActiveUser().getEmail(),
        remarks
      ]);

      // Detail entry (for FIFO tracking)
      detailRows.push([
        recordId,
        id,
        fullName.trim(),
        monthYear,
        certificateDate,
        amount,
        0,
        amount,
        'Active',
        expiration,
        'CERT-' + Utilities.formatDate(certificateDate, TIME_ZONE, 'yyyyMMddHHmmssSSS'),
        new Date(),
        Session.getActiveUser().getEmail(),
        remarks
      ]);
    });

    // Batch write
    if (ledgerRows.length > 0) {
      ledgerSheet.getRange(ledgerSheet.getLastRow() + 1, 1, ledgerRows.length, ledgerRows[0].length)
        .setValues(ledgerRows);
    }

    if (detailRows.length > 0) {
      detailSheet.getRange(detailSheet.getLastRow() + 1, 1, detailRows.length, detailRows[0].length)
        .setValues(detailRows);
    }
  }

  return id;
}


/**
 * Updates an existing employee’s details.
 *
 * --- IMPORTANT ---
 * This function intentionally DOES NOT update the initial balance (Cols J & K).
 * Initial balances are considered historical data tied to the ledger and
 * should not be changed after employee creation.
 *
 * @param {string} employeeId The ID of the employee to update.
 * @param {Object} data The fields to update (e.g., lastName, firstName, etc.).
 * @return {boolean} True on success.
 */
function updateEmployee(employeeId, data) {
  const db = getDatabase();
  const sheet = db.getSheetByName('Employees');
  if (!sheet) throw new Error('Employees sheet not found');
  const employee = getEmployeeById(employeeId);
  if (!employee) throw new Error('Employee not found');

  // Build updated row values (B to I)
  const updated = [];
  updated.push(data.lastName !== undefined ? data.lastName : employee.lastName);
  updated.push(data.firstName !== undefined ? data.firstName : employee.firstName);
  updated.push(data.middleInitial !== undefined ? data.middleInitial : employee.middleInitial);
  updated.push(data.suffix !== undefined ? data.suffix : employee.suffix);
  updated.push(data.position !== undefined ? data.position : employee.position);
  updated.push(data.office !== undefined ? data.office : employee.office);
  updated.push(data.email !== undefined ? data.email : employee.email);
  updated.push(data.status !== undefined ? data.status : employee.status);

  // This updates 8 columns, starting from column 2 (B to I)
  sheet.getRange(employee.row, 2, 1, updated.length).setValues([updated]);
  return true;
}


/**
 * Retrieves the ledger for a specific employee ordered newest first. Each
 * ledger entry is returned as an object with formatted dates. The current
 * balance is computed using getCurrentCOCBalance so that the UI always shows
 * up‑to‑date information.
 *
 * @param {string} employeeId The employee ID.
 * @return {Object} Contains the current balance and an array of ledger entries.
 */
/**
 * FIXED: Retrieves the ledger for a specific employee with robust date handling
 * Handles multiple date formats: ISO dates, US dates, and mixed formats
 *
 * @param {string} employeeId The employee ID.
 * @return {Object} Contains the current balance and an array of ledger entries.
 */
function getLedgerForEmployee(employeeId) {
  try {
    const db = getDatabase();
    const ledgerSheet = db.getSheetByName('COC_Ledger');
    
    if (!ledgerSheet) {
      return { balance: 0, entries: [] };
    }
    
    if (ledgerSheet.getLastRow() <= 1) {
      const balance = getCurrentCOCBalance(employeeId);
      return { balance: balance, entries: [] };
    }
    
    const data = ledgerSheet.getDataRange().getValues();
    const TIME_ZONE = getScriptTimeZone();
    const entries = [];
    
    // CRITICAL FIX: Normalize employee ID for comparison
    const searchId = String(employeeId).trim();
    
    /**
     * Helper to safely parse dates
     */
    function safeDateParse(value) {
      if (!value) return null;
      
      try {
        if (value instanceof Date) {
          return isNaN(value.getTime()) ? null : value;
        }
        
        if (typeof value === 'string') {
          value = value.trim();
          let date = new Date(value);
          if (!isNaN(date.getTime())) {
            return date;
          }
          
          const usDateMatch = value.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
          if (usDateMatch) {
            const [, month, day, year] = usDateMatch;
            date = new Date(year, parseInt(month) - 1, day);
            if (!isNaN(date.getTime())) {
              return date;
            }
          }
        }
        
        const date = new Date(value);
        return isNaN(date.getTime()) ? null : date;
        
      } catch (e) {
        return null;
      }
    }
    
    /**
     * Helper to format dates safely
     */
    function safeFormatDate(value, format) {
      const date = safeDateParse(value);
      if (!date) return 'N/A';
      
      try {
        return Utilities.formatDate(date, TIME_ZONE, format || 'MMMM dd, yyyy');
      } catch (e) {
        return date.toLocaleDateString();
      }
    }
    
    // Process ledger rows in reverse order (newest first)
    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      
      // Skip empty rows
      if (!row || row.length === 0 || !row[0]) {
        continue;
      }
      
      // CRITICAL FIX: Normalize row employee ID before comparison
      const rowEmpId = String(row[1] || '').trim();
      
      if (rowEmpId !== searchId) {
        continue;
      }
      
      try {
        const entry = {
          ledgerId: row[0] || '',
          employeeId: row[1] || '',
          employeeName: row[2] || '',
          transactionDate: safeFormatDate(row[3], 'MMMM dd, yyyy'),
          transactionType: row[4] || '',
          referenceId: row[5] || '',
          cocEarned: parseFloat(row[6]) || 0,
          ctoUsed: parseFloat(row[7]) || 0,
          cocBalance: parseFloat(row[8]) || 0,
          monthYearEarned: row[9] || '',
          expirationDate: row[10] ? safeFormatDate(row[10], 'MMMM dd, yyyy') : '',
          processedBy: row[11] || '',
          remarks: row[12] || ''
        };
        
        entries.push(entry);
        
      } catch (rowError) {
        Logger.log('Error processing row ' + i + ': ' + rowError.message);
        continue;
      }
    }
    
    const balance = getCurrentCOCBalance(employeeId);
    
    return { 
      balance: balance, 
      entries: entries 
    };
    
  } catch (error) {
    Logger.log('CRITICAL ERROR in getLedgerForEmployee: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    
    return { 
      balance: 0, 
      entries: [],
      error: error.message 
    };
  }
}


 * “monthly” which summarizes COC earned and CTO used by employee over a
 * specified date range. Additional report types can be added by extending
 * this function.
 *
 * @param {string} type The report type (e.g. "monthly").
 * @param {Date} startDate Inclusive start date of the report.
 * @param {Date} endDate Inclusive end date of the report.
 * @return {Array<Object>} An array of summary objects keyed by employee.
 */
function getReportData(type, startDate, endDate) {
  const db = getDatabase();
  const employees = listEmployees(false);
  if (type === 'monthly') {
    const recSheet = db.getSheetByName('COC_Records');
    const ctoSheet = db.getSheetByName('CTO_Applications');
    const recData = recSheet ? recSheet.getDataRange().getValues() : [];
    const ctoData = ctoSheet ? ctoSheet.getDataRange().getValues() : [];
    const summary = [];
    employees.forEach(emp => {
      let cocTotal = 0;
      let ctoTotal = 0;
      // Sum COC earned within range
      for (let i = 1; i < recData.length; i++) {
        const row = recData[i];
        if (row[1] === emp.id) {
          const d = new Date(row[4]);
          if (d >= startDate && d <= endDate) {
            cocTotal += parseFloat(row[12]) || 0;
          }
        }
      }
      // Sum CTO used within range
      for (let i = 1; i < ctoData.length; i++) {
        const row = ctoData[i];
        if (row[1] === emp.id && row[10] === 'Approved') {
          const d = new Date(row[9]);
          if (d >= startDate && d <= endDate) {
            ctoTotal += parseFloat(row[4]) || 0;
          }
        }
      }
      summary.push({
        employeeId: emp.id,
        employeeName: emp.fullName,
        cocEarned: cocTotal,
        ctoUsed: ctoTotal,
        net: cocTotal - ctoTotal
      });
    });
    return summary;
  }
  return [];
}


/**
* Retrieves a list of holidays from the Holidays sheet. Each entry contains
 * the row number (for editing/deletion), the date, type and description.
 * If the Holidays sheet does not exist an empty array is returned. Dates
 * are returned as JavaScript Date objects to allow the client to format
 * appropriately.
 *
 * @return {Array<Object>} A list of holiday entries.
 */
function listHolidays() {
  const db = getDatabase();
  const sheet = db.getSheetByName('Holidays');
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const dateValue = row[0];
    // Only process rows with a valid date
    if (!dateValue) continue;
    result.push({
      rowNumber: i + 1, // 1‑based index including header
      date: new Date(dateValue),
      type: row[1],
      description: row[2] || ''
    });
  }
  return result;
}


/**
 * Adds a new holiday to the Holidays sheet. If the sheet does not exist
 * it will be created with appropriate headers.
 *
 * @param {Date} date The holiday date.
 * @param {string} type The holiday type (e.g. "Regular", "Special Non‑Working", "Local", "No Work").
 * @param {string} description Optional description.
 * @return {boolean} True on success.
 */
function addHoliday(date, type, description) {
  const db = getDatabase();
  let sheet = db.getSheetByName('Holidays');
  if (!sheet) {
    sheet = db.insertSheet('Holidays');
    // Create header row
    sheet.getRange(1, 1, 1, 3).setValues([
      ['Date', 'Type', 'Description']
    ]);
  }
  sheet.appendRow([date, type, description || '']);
  return true;
}


/**
 * Updates an existing holiday entry. The caller must supply the row number
 * (including the header row) along with the new values. If the row does
 * not exist an error is thrown.
 *
 * @param {number} rowNumber The row number to update (1‑based, includes header).
 * @param {Date} date The new date.
 * @param {string} type The new type.
 * @param {string} description The new description.
 * @return {boolean} True on success.
*/
function updateHoliday(rowNumber, date, type, description) {
  const db = getDatabase();
  const sheet = db.getSheetByName('Holidays');
  if (!sheet) throw new Error('Holidays sheet not found');
  const lastRow = sheet.getLastRow();
  if (rowNumber < 2 || rowNumber > lastRow) throw new Error('Invalid row number');
  sheet.getRange(rowNumber, 1, 1, 3).setValues([[date, type, description || '']]);
  return true;
}


/**
 * Deletes a holiday entry by row number. Throws an error if the row is
 * invalid. Deleting the header row is not permitted.
 *
 * @param {number} rowNumber The row number to delete.
 * @return {boolean} True on success.
 */
function deleteHoliday(rowNumber) {
  const db = getDatabase();
  const sheet = db.getSheetByName('Holidays');
  if (!sheet) throw new Error('Holidays sheet not found');
  const lastRow = sheet.getLastRow();
  if (rowNumber < 2 || rowNumber > lastRow) throw new Error('Invalid row number');
  sheet.deleteRow(rowNumber);
  return true;
}


 * count of active COC records, count of CTO applications and the number of
 * active COC entries expiring within the next 30 days. The calculation is
 * performed server‑side to avoid transferring large datasets to the client.
 *
 * @return {Object} Dashboard statistics.
 */
function getDashboardStats() {
  const db = getDatabase();
  // Active employees
  const employees = listEmployees(false);
  const employeeCount = employees.length;
  // COC records
  let activeCOCCount = 0;
  let expiringSoonCount = 0;
  const recSheet = db.getSheetByName('COC_Records');
  if (recSheet) {
    const recData = recSheet.getDataRange().getValues();
    const today = new Date();
    const thirty = 30 * 24 * 60 * 60 * 1000;
    for (let i = 1; i < recData.length; i++) {
      const status = recData[i][15];
      if (status === 'Active') {
        activeCOCCount++;
        const exp = recData[i][14];
        if (exp) {
          const expDate = new Date(exp);
          const diff = expDate.getTime() - today.getTime();
          if (diff >= 0 && diff <= thirty) {
            expiringSoonCount++;
          }
        }
      }
    }
  }
  // CTO applications
  let ctoCount = 0;
  const ctoSheet = db.getSheetByName('CTO_Applications');
  if (ctoSheet) {
    const ctoData = ctoSheet.getDataRange().getValues();
    ctoCount = Math.max(0, ctoData.length - 1);
  }
  return {
    employees: employeeCount,
    activeCOC: activeCOCCount,
    ctoApplications: ctoCount,
    expiringSoon: expiringSoonCount
  };
}


/**
 * Retrieves the latest ledger transactions across all employees. This is used
 * by the dashboard to show a recap of recent activity. The results are
 * returned in reverse chronological order (newest first). Each entry includes
 * the transaction date, employee name, type, reference and remarks.
 *
 * @param {number} limit The maximum number of entries to return.
 * @return {Array<Object>} An array of recent ledger entries.
 */
function getRecentActivities(limit) {
  const db = getDatabase();
  const ledgerSheet = db.getSheetByName('COC_Ledger');
  if (!ledgerSheet) return [];
  const data = ledgerSheet.getDataRange().getValues();
  const entries = [];
  // Iterate backwards to gather newest first
  for (let i = data.length - 1; i >= 1 && entries.length < (limit || 10); i--) {
    const row = data[i];
    entries.push({
      transactionDate: formatDate(row[3]),
      employeeName: row[2],
      transactionType: row[4],
      referenceId: row[5],
      remarks: row[12]
    });
  }
  return entries;
}


/**
 * Adds a COC entry to COC_Balance_Detail when COC is earned
 * Call this from recordCOCEntries after creating COC_Records entry
 */
function addCOCToBalanceDetail(employeeId, employeeName, recordId, dateEarned, hoursEarned) {
  const detailSheet = ensureCOCBalanceDetailSheet();
  const TIME_ZONE = getScriptTimeZone();

  const earnedDate = new Date(dateEarned);
  const monthYearEarned = Utilities.formatDate(earnedDate, TIME_ZONE, 'MM-yyyy');
  const now = new Date();
  const createdBy = Session.getActiveUser().getEmail();

  detailSheet.appendRow([
    recordId,                            // Record ID (matches COC_Records ID)
    employeeId,
    employeeName,
    monthYearEarned,
    '',                                  // Certificate Date (set once certificate is issued)
    hoursEarned,
    0,                                   // Hours Used
    hoursEarned,                         // Hours Remaining
    'Active',
    '',                                  // Expiration Date (computed after certificate issuance)
    '',                                  // Certificate ID
    now,
    createdBy,
    'COC earned from ' + recordId
  ]);
}


/**
 * Returns the current COC balance for a given employee. First tries to read
 * the most recent balance from the ledger sheet (Column I). If no ledger
 * records exist for the employee, it falls back to the 'initialBalance'
 * value in the 'Employees' sheet (Column J).
 *
 * @param {string} employeeId The employee ID.
 * @return {number} The current COC balance (in hours).
 */
function getCurrentCOCBalance(employeeId) {
  try {
    const db = getDatabase();
    const ledgerSheet = db.getSheetByName('COC_Ledger');
    
    if (!ledgerSheet) {
      return 0;
    }

    if (ledgerSheet.getLastRow() <= 1) {
      return calculateBalanceFallback(employeeId);
    }
    
    const ledgerData = ledgerSheet.getDataRange().getValues();
    
    // CRITICAL FIX: Normalize employee ID for comparison
    const searchId = String(employeeId).trim();
    
    // Search from bottom for the latest transaction
    for (let i = ledgerData.length - 1; i >= 1; i--) {
      const row = ledgerData[i];
      
      // Skip empty rows
      if (!row || row.length === 0 || !row[0]) {
        continue;
      }
      
      // CRITICAL FIX: Normalize row employee ID before comparison
      const rowEmpId = String(row[1] || '').trim();
      
      if (rowEmpId === searchId) {
        const balance = parseFloat(row[8]) || 0;
        return balance < 0 ? 0 : balance;
      }
    }

    // No ledger entries found, use fallback
    return calculateBalanceFallback(employeeId);
    
  } catch (error) {
    Logger.log('ERROR in getCurrentCOCBalance: ' + error.message);
    return 0;
  }
}


/**
 * Get current COC balance from COC_Balance_Detail (active entries only)
 * More accurate than ledger-based calculation
 */
function getCurrentCOCBalanceFromDetail(employeeId) {
  const db = getDatabase();
  const detailSheet = db.getSheetByName('COC_Balance_Detail');

  if (!detailSheet || detailSheet.getLastRow() < 2) {
    // Fallback to original method if detail sheet doesn't exist
    return getCurrentCOCBalance(employeeId);
  }

  const data = detailSheet.getDataRange().getValues();
  let totalRemaining = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const empId = row[1]; // Employee ID
    const hoursRemaining = row[6]; // Hours Remaining
    const status = row[8]; // Status

    if (empId === employeeId && status === 'Active' && hoursRemaining > 0) {
      totalRemaining += hoursRemaining;
    }
  }

  return totalRemaining;
}


/**
 * FIXED: Get COC Balance breakdown with robust date handling
 */
function getCOCBalanceBreakdown(employeeId) {
  const detailSheet = ensureCOCBalanceDetailSheet();
  const data = detailSheet.getDataRange().getValues();
  const TIME_ZONE = getScriptTimeZone();
  
  /**
   * Helper to safely parse dates
   */
  function safeDateParse(value) {
    if (!value) return null;
    try {
      if (value instanceof Date) return isNaN(value.getTime()) ? null : value;
      if (typeof value === 'string') {
        const date = new Date(value);
        if (!isNaN(date.getTime())) return date;
        const usMatch = value.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
        if (usMatch) {
          const [, m, d, y] = usMatch;
          return new Date(y, parseInt(m) - 1, d);
        }
      }
      return new Date(value);
    } catch (e) {
      return null;
    }
  }
  
  function safeFormatDate(value) {
    const date = safeDateParse(value);
    if (!date) return 'N/A';
    try {
      return Utilities.formatDate(date, TIME_ZONE, 'yyyy-MM-dd');
    } catch (e) {
      return date.toISOString().split('T')[0];
    }
  }

  const breakdown = [];
  
  // CRITICAL FIX: Normalize employee ID for comparison
  const searchId = String(employeeId).trim();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    // CRITICAL FIX: Normalize row employee ID before comparison
    const rowEmpId = String(row[1] || '').trim();
    const status = row[8];
    const hoursRemaining = parseFloat(row[6]) || 0;
    
    if (rowEmpId === searchId && status === 'Active' && hoursRemaining > 0) {
      const expDate = safeDateParse(row[7]);
      const daysUntil = expDate ? Math.ceil((expDate - new Date()) / (1000 * 60 * 60 * 24)) : -999;
      
      breakdown.push({
        entryId: row[0],
        recordId: row[3],
        dateEarned: safeFormatDate(row[4]),
        hoursEarned: parseFloat(row[5]) || 0,
        hoursRemaining: hoursRemaining,
        expirationDate: safeFormatDate(row[7]),
        daysUntilExpiration: daysUntil
      });
    }
  }

  breakdown.sort((a, b) => new Date(a.dateEarned) - new Date(b.dateEarned));
  return breakdown;
}


