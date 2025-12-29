// ===============================================
// INSTALMENTTRACKER.GS - Instalment Tracking System
// UPDATED: Platinum course now £1047 (was £997) with £397, £350, £300 installments
// UPDATED: Added Tuition/Revision Plus (822/522/300) support
// FIXED: Now searches for earlier payments when £300 or £350 is received
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
    // IMPORTANT FIX: If this is a £300 or £350 payment, they MUST have earlier payments
    // (otherwise they wouldn't be in the monthly sheets at all)
    const actual = Number(actualPrice);
    
    if (actual === 300 || actual === 350) {
      Logger.log(`⚠️ Received £${actual} payment for ${studentName} with no existing tracker record`);
      Logger.log(`   This student MUST have earlier payments - searching all monthly sheets...`);
      
      // Search for ALL payments for this student
      const allPayments = searchForAllStudentPayments(ss, studentName, course, fullPrice);
      
      if (allPayments.length > 0) {
        Logger.log(`   ✅ Found ${allPayments.length} total payment(s) for this student`);
        Logger.log(`   Creating complete record with full payment history...`);
        
        // Create record with all payments
        createStudentRecordWithAllPayments(trackerSheet, studentName, course, fullPrice, allPayments);
      } else {
        Logger.log(`   ❌ ERROR: No payments found - this should never happen!`);
        Logger.log(`   Creating record anyway but this needs investigation`);
        createStudentRecord(trackerSheet, studentName, course, fullPrice, actualPrice, paymentDate);
      }
    } else {
      // Normal case: first payment (397/347/297/522) or full payment
      createStudentRecord(trackerSheet, studentName, course, fullPrice, actualPrice, paymentDate);
    }
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

  // If actual equals full, it's a single full payment
  if (actual === full) {
    return 1;
  }

  // Determine instalment number based on price patterns

  // UPDATED: Platinum (1047)
  if (full === 1047) {
    if (actual === 397) return 1;
    if (actual === 350) return 1; // Will be corrected by searchForAllStudentPayments
    if (actual === 300) return 1; // Will be corrected by searchForAllStudentPayments
  } 
  
  // Tuition/Revision Plus (822)
  else if (full === 822) {
    if (actual === 522) return 1;
    if (actual === 300) return 1; // Will be corrected by searchForAllStudentPayments
  } 
  
  // Revision (647)
  else if (full === 647) {
    if (actual === 347) return 1;
    if (actual === 300) return 1; // Will be corrected by searchForAllStudentPayments
  } 
  
  // Tuition (597)
  else if (full === 597) {
    if (actual === 297) return 1;
    if (actual === 300) return 1; // Will be corrected by searchForAllStudentPayments
  }

  return 1; // Default to first instalment
}

function calculateNextPaymentDate(lastPaymentDate) {
  const nextDate = new Date(lastPaymentDate);
  nextDate.setMonth(nextDate.getMonth() + 1); // Add 1 month
  return nextDate;
}

// NEW FUNCTION: Search monthly sheets for ALL payments for this student
function searchForAllStudentPayments(ss, studentName, course, fullPrice) {
  const allSheets = ss.getSheets();
  const monthlySheets = allSheets.filter(sheet => isMonthlySheetName(sheet.getName()));
  
  const allPayments = [];
  
  monthlySheets.forEach(sheet => {
    const sheetData = sheet.getDataRange().getValues();
    const sheetHeaders = sheetData[0];
    const sheetRows = sheetData.slice(1);
    
    const nameCol = sheetHeaders.indexOf('Name');
    const fullPriceCol = sheetHeaders.indexOf('Full Price');
    const actualPriceCol = sheetHeaders.indexOf('Actual Price');
    const dateCol = sheetHeaders.indexOf('Date');
    
    if (nameCol === -1 || actualPriceCol === -1) return;
    
    sheetRows.forEach(row => {
      // Match by name and full price
      if (row[nameCol] === studentName && Number(row[fullPriceCol]) === Number(fullPrice)) {
        allPayments.push({
          amount: Number(row[actualPriceCol]),
          date: new Date(row[dateCol]),
          sheetName: sheet.getName()
        });
      }
    });
  });
  
  // Sort by date (oldest first)
  allPayments.sort((a, b) => a.date - b.date);
  
  return allPayments;
}

// NEW FUNCTION: Create student record with ALL payments found
function createStudentRecordWithAllPayments(sheet, studentName, course, fullPrice, allPayments) {
  // Calculate totals from all payments
  const totalPaid = allPayments.reduce((sum, p) => sum + p.amount, 0);
  const instalmentCount = allPayments.length;
  
  // Get the last payment date (most recent)
  const lastPaymentDate = allPayments[allPayments.length - 1].date;
  
  const isComplete = totalPaid >= Number(fullPrice);
  const nextPaymentDue = isComplete ? '' : calculateNextPaymentDate(lastPaymentDate);
  const completionStatus = isComplete ? 'Complete' : 'In Progress';
  const completionDate = isComplete ? new Date() : '';
  
  const newRecord = [
    studentName,
    course,
    Number(fullPrice),
    totalPaid,
    instalmentCount,
    lastPaymentDate,
    nextPaymentDue,
    completionStatus,
    completionDate
  ];
  
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, 1, newRecord.length).setValues([newRecord]);
  
  Logger.log(`✅ Created complete instalment record: ${studentName} - ${course}`);
  Logger.log(`   Found payment history: ${allPayments.map(p => `£${p.amount}`).join(', ')}`);
  Logger.log(`   Total: £${totalPaid} (${instalmentCount} instalments)`);
  Logger.log(`   Status: ${completionStatus}`);
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