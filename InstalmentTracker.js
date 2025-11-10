// ===============================================
// INSTALMENTTRACKER.GS - Instalment Tracking System
// ===============================================

function createInstalmentTrackerSheet(ss) {
  const sheet = ss.insertSheet('Instalment Tracker');
  const headers = [
    'Student Name', 'Course', 'Full Price', 'Amount Paid', 'Instalments Paid', 
    'Last Payment Date', 'Next Payment Due', 'Payment Complete', 'Completion Date'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Format headers
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#f0f0f0');
  
  // Position tab after Awaiting Employer Invoice, before Failed orders
  positionInstalmentTrackerTab(ss, sheet);
  
  Logger.log('Created Instalment Tracker sheet');
  return sheet;
}

function processInstalmentPayment(studentName, course, fullPrice, actualPrice, paymentDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let trackerSheet = ss.getSheetByName('Instalment Tracker');
  
  if (!trackerSheet) {
    trackerSheet = createInstalmentTrackerSheet(ss);
  }
  
  // Find existing student record
  const existingRowIndex = findStudentRecord(trackerSheet, studentName, course, fullPrice);
  
  if (existingRowIndex > 0) {
    // Update existing record
    updateStudentRecord(trackerSheet, existingRowIndex, actualPrice, paymentDate);
  } else {
    // Create new record
    createStudentRecord(trackerSheet, studentName, course, fullPrice, actualPrice, paymentDate);
  }
}

function findStudentRecord(sheet, studentName, course, fullPrice) {
  const dataRange = sheet.getDataRange();
  if (dataRange.getNumRows() <= 1) return 0; // No data
  
  const allData = dataRange.getValues();
  const dataRows = allData.slice(1); // Skip header
  
  // Look for matching student, course, and full price
  for (let i = 0; i < dataRows.length; i++) {
    const row = dataRows[i];
    if (row[0] === studentName && row[1] === course && Number(row[2]) === Number(fullPrice)) {
      return i + 2; // +2 for header and 1-indexed
    }
  }
  
  return 0; // Not found
}

function createStudentRecord(sheet, studentName, course, fullPrice, actualPrice, paymentDate) {
  const instalmentCount = getInstalmentCount(actualPrice, fullPrice);
  const nextPaymentDue = calculateNextPaymentDate(paymentDate);
  
  const newRecord = [
    studentName,
    course,
    Number(fullPrice),
    Number(actualPrice),
    instalmentCount,
    paymentDate,
    nextPaymentDue,
    'In Progress', // Payment Complete
    '' // Completion Date
  ];
  
  // Add to end of sheet
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, 1, newRecord.length).setValues([newRecord]);
  
  Logger.log(`Created new instalment record: ${studentName} - ${course} - Payment ${instalmentCount}`);
}

function updateStudentRecord(sheet, rowIndex, actualPrice, paymentDate) {
  const currentData = sheet.getRange(rowIndex, 1, 1, 9).getValues()[0];
  
  // Update totals
  const currentAmountPaid = Number(currentData[3]);
  const newAmountPaid = currentAmountPaid + Number(actualPrice);
  const newInstalmentCount = Number(currentData[4]) + 1;
  const fullPrice = Number(currentData[2]);
  
  // Check if payment is complete
  const isComplete = newAmountPaid >= fullPrice;
  const completionStatus = isComplete ? 'Complete' : 'In Progress';
  const completionDate = isComplete ? new Date() : '';
  const nextPaymentDue = isComplete ? '' : calculateNextPaymentDate(paymentDate);
  
  // Update the record
  sheet.getRange(rowIndex, 4).setValue(newAmountPaid); // Amount Paid
  sheet.getRange(rowIndex, 5).setValue(newInstalmentCount); // Instalments Paid
  sheet.getRange(rowIndex, 6).setValue(paymentDate); // Last Payment Date
  sheet.getRange(rowIndex, 7).setValue(nextPaymentDue); // Next Payment Due
  sheet.getRange(rowIndex, 8).setValue(completionStatus); // Payment Complete
  sheet.getRange(rowIndex, 9).setValue(completionDate); // Completion Date
  
  if (isComplete) {
    Logger.log(`COMPLETED: ${currentData[0]} finished paying for ${currentData[1]} - Total: ${newAmountPaid}`);
  } else {
    Logger.log(`Updated: ${currentData[0]} - ${currentData[1]} - Payment ${newInstalmentCount} - Next due: ${nextPaymentDue}`);
  }
}

function getInstalmentCount(actualPrice, fullPrice) {
  const actual = Number(actualPrice);
  const full = Number(fullPrice);
  
  // Determine instalment number based on price patterns
  if (full === 997) { // Platinum
    if (actual === 397) return 1;
    if (actual === 300) return 2; // We'll update this to correct number when we update existing records
  } else if (full === 647) { // Revision
    if (actual === 347) return 1;
    if (actual === 300) return 2;
  } else if (full === 597) { // Tuition
    if (actual === 297) return 1;
    if (actual === 300) return 2;
  }
  
  return 1; // Default to first instalment
}

function calculateNextPaymentDate(lastPaymentDate) {
  const nextDate = new Date(lastPaymentDate);
  nextDate.setMonth(nextDate.getMonth() + 1); // Add 1 month
  return nextDate;
}

function cleanupCompletedPayments() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trackerSheet = ss.getSheetByName('Instalment Tracker');
  
  if (!trackerSheet) {
    Logger.log('Instalment Tracker sheet not found');
    return;
  }
  
  const dataRange = trackerSheet.getDataRange();
  if (dataRange.getNumRows() <= 1) {
    Logger.log('No data in Instalment Tracker to clean up');
    return;
  }
  
  const allData = dataRange.getValues();
  const dataRows = allData.slice(1); // Skip header
  const currentDate = new Date();
  const thirtyDaysAgo = new Date(currentDate.getTime() - (30 * 24 * 60 * 60 * 1000));
  
  Logger.log('Cleaning up completed payments older than 30 days...');
  
  // Process rows in reverse order to avoid index issues when deleting
  for (let i = dataRows.length - 1; i >= 0; i--) {
    const row = dataRows[i];
    const completionStatus = row[7]; // Payment Complete
    const completionDate = row[8]; // Completion Date
    
    if (completionStatus === 'Complete' && completionDate && new Date(completionDate) < thirtyDaysAgo) {
      const rowIndex = i + 2; // +2 for header and 1-indexed
      trackerSheet.deleteRow(rowIndex);
      Logger.log(`Deleted completed payment record: ${row[0]} - ${row[1]} (completed ${completionDate})`);
    }
  }
  
  Logger.log('Cleanup completed');
}

function setupInstalmentTracking() {
  // Set up trigger to clean up completed payments weekly
  ScriptApp.newTrigger('cleanupCompletedPayments')
    .timeBased()
    .everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(7)
    .create();
  
  Logger.log('Instalment tracking cleanup trigger set up');
}

function updateInstalmentTrackerFromMonthlySheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();
  
  Logger.log('Scanning monthly sheets for instalment payments...');
  
  allSheets.forEach(sheet => {
    const sheetName = sheet.getName();
    if (isMonthlySheetName(sheetName)) {
      processMonthlySheetForInstalments(sheet);
    }
  });
  
  Logger.log('Finished updating Instalment Tracker from monthly sheets');
}

function processMonthlySheetForInstalments(sheet) {
  const dataRange = sheet.getDataRange();
  if (dataRange.getNumRows() <= 1) return;
  
  const allData = dataRange.getValues();
  const headers = allData[0];
  const dataRows = allData.slice(1);
  
  // Find column indices
  const nameCol = headers.indexOf('Name');
  const courseCol = headers.indexOf('Course');
  const fullPriceCol = headers.indexOf('Full Price');
  const actualPriceCol = headers.indexOf('Actual Price');
  const paymentPlanCol = headers.indexOf('Payment Plan');
  const dateCol = headers.indexOf('Date');
  
  if (nameCol === -1 || courseCol === -1 || fullPriceCol === -1 || 
      actualPriceCol === -1 || paymentPlanCol === -1 || dateCol === -1) {
    Logger.log(`${sheet.getName()} missing required columns for instalment processing`);
    return;
  }
  
  // Process each row that has a payment plan
  dataRows.forEach(row => {
    const hasPaymentPlan = row[paymentPlanCol] === 'Y';
    
    if (hasPaymentPlan) {
      const studentName = row[nameCol];
      const course = row[courseCol];
      const fullPrice = row[fullPriceCol];
      const actualPrice = row[actualPriceCol];
      const paymentDate = new Date(row[dateCol]);
      
      if (studentName && course && fullPrice && actualPrice) {
        processInstalmentPayment(studentName, course, fullPrice, actualPrice, paymentDate);
      }
    }
  });
}

function debugInstalmentTracker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trackerSheet = ss.getSheetByName('Instalment Tracker');
  
  if (!trackerSheet) {
    Logger.log('Instalment Tracker sheet not found');
    return;
  }
  
  const dataRange = trackerSheet.getDataRange();
  const values = dataRange.getValues();
  
  Logger.log('=== Instalment Tracker Debug ===');
  values.forEach((row, index) => {
    if (index === 0) {
      Logger.log(`Headers: ${row.join(', ')}`);
    } else {
      Logger.log(`Row ${index + 1}: ${row.join(', ')}`);
    }
  });
}