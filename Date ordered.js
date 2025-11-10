// ===============================================
// DATE-ORDERED INSERTION FOR MONTHLY SHEETS
// ===============================================

function addRowToMonthlySheetInDateOrder(monthlySheet, newRowData) {
  const dataRange = monthlySheet.getDataRange();
  
  // If this is the first data row (only header exists), just add it
  if (dataRange.getNumRows() <= 1) {
    monthlySheet.getRange(2, 1, 1, newRowData.length).setValues([newRowData]);
    Logger.log(`Added first row to ${monthlySheet.getName()}`);
    return;
  }
  
  const allData = dataRange.getValues();
  const headers = allData[0];
  const dataRows = allData.slice(1);
  
  // Find date column index
  const dateCol = headers.indexOf('Date');
  if (dateCol === -1) {
    Logger.log(`Warning: No Date column found in ${monthlySheet.getName()}, adding to end`);
    const lastRow = monthlySheet.getLastRow();
    monthlySheet.getRange(lastRow + 1, 1, 1, newRowData.length).setValues([newRowData]);
    return;
  }
  
  const newDate = new Date(newRowData[dateCol]);
  
  // Find the correct insertion point (oldest to newest)
  let insertIndex = -1;
  
  for (let i = 0; i < dataRows.length; i++) {
    const existingDate = new Date(dataRows[i][dateCol]);
    
    // If new date is older than or equal to current row date, insert here
    if (newDate.getTime() <= existingDate.getTime()) {
      insertIndex = i;
      break;
    }
  }
  
  // If no insertion point found, add to the end
  if (insertIndex === -1) {
    insertIndex = dataRows.length;
  }
  
  // Insert the new row at the correct position
  const insertRowNumber = insertIndex + 2; // +2 for header and 1-indexed sheets
  
  // Insert a new row at the correct position
  monthlySheet.insertRowBefore(insertRowNumber);
  
  // Add the data to the newly inserted row
  monthlySheet.getRange(insertRowNumber, 1, 1, newRowData.length).setValues([newRowData]);
  
  Logger.log(`Added row for ${newRowData[1]} to ${monthlySheet.getName()} at position ${insertRowNumber} (date: ${newDate.toDateString()})`);
}

// ===============================================
// UPDATED MOVEMENT FUNCTIONS WITH DATE ORDERING
// ===============================================

function moveToMonthlySheetFromSorting(ss, originalData) {
  const date = new Date(originalData[0]);
  const monthName = getMonthName(date.getMonth());
  const year = date.getFullYear();
  const sheetName = `${monthName} ${year}`;

  let monthlySheet = ss.getSheetByName(sheetName);
  if (!monthlySheet) {
    monthlySheet = createMonthlySheet(ss, sheetName);
    positionMonthlySheetTab(ss, monthlySheet, date);
  }

  // Get payment plan information
  const paymentInfo = getPaymentPlanInfo(originalData[4], originalData[5]); // Full Price, Actual Price

  // Prepare row data with course filled in
  const rowWithCourse = [...originalData];
  rowWithCourse[2] = paymentInfo.course; // Use course from payment plan detection

  // Add calculated columns
  const fmeFee = calculateFMEFee(originalData[4], originalData[5]);
  const stripeFee = calculateStripeFee(originalData[5]);
  const expectedIncome = calculateExpectedIncome(originalData[5], fmeFee, stripeFee);

  const completeRow = [
    ...rowWithCourse,
    fmeFee,
    stripeFee,
    '', // Actual Stripe Fee (empty)
    expectedIncome,
    paymentInfo.isPaymentPlan ? 'Y' : '', // Payment Plan
    paymentInfo.instalment, // Instalment
    '' // Comment (empty)
  ];

  // FIXED: Add row in correct date order
  addRowToMonthlySheetInDateOrder(monthlySheet, completeRow);

  // Process for student engagement transfer
  const studentData = {
    name: originalData[1], // Name column
    sitting: originalData[3], // Sitting column
    actualPrice: originalData[5], // Actual Price column
    course: paymentInfo.course
  };

  processStudentForEngagement(studentData, sheetName);

  if (paymentInfo.isPaymentPlan) {
    Logger.log(`Moved data from sorting to ${sheetName} sheet - Payment Plan: ${paymentInfo.instalment} - Processed for Engagement - SORTED BY DATE`);
  } else {
    Logger.log(`Moved data from sorting to ${sheetName} sheet - Full Payment - Processed for Engagement - SORTED BY DATE`);
  }
}

function moveToMonthlySheet(ss, row) {
  const date = new Date(row[0]);
  const monthName = getMonthName(date.getMonth());
  const year = date.getFullYear();
  const sheetName = `${monthName} ${year}`;

  let monthlySheet = ss.getSheetByName(sheetName);
  if (!monthlySheet) {
    monthlySheet = createMonthlySheet(ss, sheetName);
    positionMonthlySheetTab(ss, monthlySheet, date);
  }

  // Get payment plan information
  const paymentInfo = getPaymentPlanInfo(row[4], row[5]); // Full Price, Actual Price

  // Prepare row data with course filled in
  const rowWithCourse = [...row];
  rowWithCourse[2] = paymentInfo.course; // Use course from payment plan detection

  // Add calculated columns
  const fmeFee = calculateFMEFee(row[4], row[5]);
  const stripeFee = calculateStripeFee(row[5]);
  const expectedIncome = calculateExpectedIncome(row[5], fmeFee, stripeFee);

  const completeRow = [
    ...rowWithCourse,
    fmeFee,
    stripeFee,
    '', // Actual Stripe Fee (empty)
    expectedIncome,
    paymentInfo.isPaymentPlan ? 'Y' : '', // Payment Plan
    paymentInfo.instalment, // Instalment
    '' // Comment (empty)
  ];

  // FIXED: Add row in correct date order
  addRowToMonthlySheetInDateOrder(monthlySheet, completeRow);

  // Process for student engagement transfer
  const studentData = {
    name: row[1], // Name column
    sitting: row[3], // Sitting column
    actualPrice: row[5], // Actual Price column
    course: paymentInfo.course
  };

  Logger.log(`🎓 Attempting engagement transfer for ${studentData.name} - Price: £${studentData.actualPrice}`);

  // Call the engagement transfer function
  processStudentForEngagement(studentData, sheetName);

  // If this is an instalment payment, add to Instalment Tracker
  if (paymentInfo.isPaymentPlan) {
    const studentName = row[1]; // Name column
    processInstalmentPayment(studentName, paymentInfo.course, row[4], row[5], date);
    Logger.log(`Added to ${sheetName} sheet - Payment Plan: ${paymentInfo.instalment} - Added to Instalment Tracker - Engagement Transfer Attempted - SORTED BY DATE`);
  } else {
    Logger.log(`Added to ${sheetName} sheet - Full Payment - Engagement Transfer Attempted - SORTED BY DATE`);
  }
}

function moveToMonthlySheetFromInvoice(ss, originalData) {
  const date = new Date(originalData[0]);
  const monthName = getMonthName(date.getMonth());
  const year = date.getFullYear();
  const sheetName = `${monthName} ${year}`;

  let monthlySheet = ss.getSheetByName(sheetName);
  if (!monthlySheet) {
    monthlySheet = createMonthlySheet(ss, sheetName);
    positionMonthlySheetTab(ss, monthlySheet, date);
  }

  // Get payment plan information (using updated actual price = full price)
  const paymentInfo = getPaymentPlanInfo(originalData[4], originalData[5]); // Full Price, Actual Price

  // Prepare row data with course filled in
  const rowWithCourse = [...originalData];
  rowWithCourse[2] = paymentInfo.course; // Use course from payment plan detection

  // Add calculated columns (using the updated actual price which now equals full price)
  const fmeFee = calculateFMEFee(originalData[4], originalData[5]);
  const stripeFee = calculateStripeFee(originalData[5]);
  const expectedIncome = calculateExpectedIncome(originalData[5], fmeFee, stripeFee);

  const completeRow = [
    ...rowWithCourse,
    fmeFee,
    stripeFee,
    '', // Actual Stripe Fee (empty)
    expectedIncome,
    paymentInfo.isPaymentPlan ? 'Y' : '', // Payment Plan (should be empty since actual=full)
    paymentInfo.instalment, // Instalment (should be empty since actual=full)
    '' // Comment (empty)
  ];

  // FIXED: Add row in correct date order
  addRowToMonthlySheetInDateOrder(monthlySheet, completeRow);

  // Process for student engagement transfer
  const studentData = {
    name: originalData[1], // Name column
    sitting: originalData[3], // Sitting column
    actualPrice: originalData[5], // Actual Price column (now equals full price)
    course: paymentInfo.course
  };

  processStudentForEngagement(studentData, sheetName);

  Logger.log(`Moved invoice data to ${sheetName} sheet with full payment (invoice paid) - Processed for Engagement - SORTED BY DATE`);
}

// ===============================================
// UTILITY FUNCTIONS FOR DATE ORDERING
// ===============================================

function sortExistingMonthlySheetByDate(monthlySheetName) {
  Logger.log(`🔄 Sorting existing ${monthlySheetName} by date order...`);
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(monthlySheetName);
  
  if (!sheet) {
    Logger.log(`❌ Sheet ${monthlySheetName} not found`);
    return;
  }
  
  const dataRange = sheet.getDataRange();
  if (dataRange.getNumRows() <= 1) {
    Logger.log(`No data to sort in ${monthlySheetName}`);
    return;
  }
  
  const allData = dataRange.getValues();
  const headers = allData[0];
  const dataRows = allData.slice(1);
  
  // Find date column
  const dateCol = headers.indexOf('Date');
  if (dateCol === -1) {
    Logger.log(`❌ No Date column found in ${monthlySheetName}`);
    return;
  }
  
  // Sort data rows by date (oldest first)
  dataRows.sort((a, b) => {
    const dateA = new Date(a[dateCol]);
    const dateB = new Date(b[dateCol]);
    return dateA.getTime() - dateB.getTime(); // Ascending order (oldest first)
  });
  
  // Clear existing data and write sorted data
  sheet.clear();
  
  // Write headers first
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Format headers
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#f0f0f0');
  
  // Write sorted data
  if (dataRows.length > 0) {
    sheet.getRange(2, 1, dataRows.length, headers.length).setValues(dataRows);
  }
  
  // Force Sitting column to text format
  const sittingCol = headers.indexOf('Sitting');
  if (sittingCol !== -1) {
    const sittingColumn = sheet.getRange(`${String.fromCharCode(65 + sittingCol)}:${String.fromCharCode(65 + sittingCol)}`);
    sittingColumn.setNumberFormat('@');
  }
  
  Logger.log(`✅ Sorted ${dataRows.length} rows by date in ${monthlySheetName} (oldest first)`);
}

function sortAllMonthlySheetsByDate() {
  Logger.log('🔄 SORTING ALL MONTHLY SHEETS BY DATE');
  Logger.log('=====================================');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();
  
  // Find all monthly sheets
  const monthlySheets = allSheets.filter(sheet => isMonthlySheetName(sheet.getName()));
  
  Logger.log(`Found ${monthlySheets.length} monthly sheets to sort`);
  
  monthlySheets.forEach(sheet => {
    sortExistingMonthlySheetByDate(sheet.getName());
  });
  
  Logger.log('🎉 All monthly sheets sorted by date!');
}

// ===============================================
// TESTING FUNCTIONS
// ===============================================

function testDateOrderedInsertion() {
  Logger.log('🧪 TESTING DATE-ORDERED INSERTION');
  Logger.log('==================================');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Test with sample data with different dates
  const testRows = [
    [new Date('2025-08-20'), 'Test Student 1', 'Platinum', 'December 2025', 997, 397, 'Standard'], // Newer
    [new Date('2025-08-15'), 'Test Student 2', 'Revision', 'December 2025', 647, 347, 'Standard'], // Older
    [new Date('2025-08-18'), 'Test Student 3', 'Tuition', 'December 2025', 597, 297, 'Standard']  // Middle
  ];
  
  Logger.log('Adding test students in random order...');
  
  testRows.forEach((row, index) => {
    Logger.log(`  ${index + 1}. ${row[1]} - Date: ${row[0].toDateString()}`);
    moveToMonthlySheet(ss, row);
  });
  
  Logger.log('\n✅ Test complete - check your August 2025 sheet');
  Logger.log('Students should appear in date order: Test Student 2 (15th), Test Student 3 (18th), Test Student 1 (20th)');
}