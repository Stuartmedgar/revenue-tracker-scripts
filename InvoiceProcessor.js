// ===============================================
// INVOICEPROCESSOR.GS - Awaiting Employer Invoice Logic (Complete with Student Engagement and Fixed Date Ordering)
// ===============================================

function processInvoicePaidRow(ss, awaitingSheet, row, rowIndex) {
  // Get original data (first 7 columns)
  let originalData = row.slice(0, 7);

  // Step 1: Change Actual Price to equal Full Price
  const fullPrice = originalData[4]; // Column E
  originalData[5] = fullPrice; // Set Actual Price (Column F) to Full Price value

  Logger.log(`Updated Actual Price from ${row[5]} to ${fullPrice} for invoice payment`);

  // Step 2: Move to monthly sheet using the updated data
  moveToMonthlySheetFromInvoice(ss, originalData);

  // Step 3: Delete from Awaiting Employer Invoice sheet
  deleteAwaitingInvoiceRow(awaitingSheet, rowIndex);
  cleanupCheckboxTracking(ss, row, 'AwaitingInvoice');

  Logger.log(`Row ${rowIndex}: Invoice paid - moved to monthly sheet with full price`);
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

  // FIXED: Add to END of sheet instead of row 2 (oldest first, newest last)
  const lastRow = monthlySheet.getLastRow();
  monthlySheet.getRange(lastRow + 1, 1, 1, completeRow.length).setValues([completeRow]);

  // Process for student engagement transfer
  const studentData = {
    name: originalData[1], // Name column
    sitting: originalData[3], // Sitting column
    actualPrice: originalData[5], // Actual Price column (now equals full price)
    course: paymentInfo.course
  };

  processStudentForEngagement(studentData, sheetName);

  Logger.log(`Moved invoice data to ${sheetName} sheet with full payment (invoice paid) - Processed for Engagement`);
}

function deleteAwaitingInvoiceRow(sheet, rowIndex) {
  try {
    sheet.deleteRow(rowIndex);
    Logger.log(`Deleted row ${rowIndex} from Awaiting Employer Invoice sheet`);
  } catch (error) {
    Logger.log(`Error deleting awaiting invoice row ${rowIndex}: ${error.toString()}`);
  }
}

function createAwaitingEmployerInvoiceSheet(ss) {
  const sheet = ss.insertSheet('Awaiting Employer Invoice');

  const headers = [
    'Date', 'Name', 'Course', 'Sitting', 'Full Price', 'Actual Price', 'Order Type',
    'Invoice Paid', 'Reminder'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Format headers
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#f0f0f0');

  // Force Sitting column to text format
  const sittingColumn = sheet.getRange('D:D');
  sittingColumn.setNumberFormat('@');

  // Position the tab correctly (after Sort, before Failed orders)
  positionAwaitingInvoiceTab(ss, sheet);

  Logger.log('Created Awaiting Employer Invoice sheet with text formatting');

  return sheet;
}