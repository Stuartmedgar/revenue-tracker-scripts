// ===============================================
// DATAPROCESSOR.GS - Data Processing Logic (Complete with Student Engagement and Fixed Date Ordering)
// UPDATED: Added Tuition/Revision Plus (822) support
// ===============================================

function processDataEntries() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dataEntrySheet = ss.getSheetByName('Data Entry');

    if (!dataEntrySheet) {
      Logger.log('Data Entry sheet not found');
      return;
    }

    // Get all data from Data Entry sheet (excluding header)
    const dataRange = dataEntrySheet.getDataRange();
    if (dataRange.getNumRows() <= 1) {
      Logger.log('No data to process');
      return;
    }

    const allData = dataRange.getValues();
    const headers = allData[0];
    const dataRows = allData.slice(1);

    // Track entry times for new rows
    trackDataEntryTimes(ss, dataRows);

    // Process each row (in reverse order to avoid index issues when deleting)
    for (let i = dataRows.length - 1; i >= 0; i--) {
      const row = dataRows[i];
      const rowIndex = i + 2; // +2 because we removed header and sheets are 1-indexed

      // Check if row has been in the sheet for at least 5 minutes
      if (shouldProcessRow(ss, row, rowIndex)) {
        processIndividualRow(ss, dataEntrySheet, row, rowIndex);
      }
    }

  } catch (error) {
    Logger.log('Error in processDataEntries: ' + error.toString());
  }
}

function processDataEntriesTestMode() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dataEntrySheet = ss.getSheetByName('Data Entry');

    if (!dataEntrySheet) {
      Logger.log('Data Entry sheet not found');
      return;
    }

    // Get all data from Data Entry sheet (excluding header)
    const dataRange = dataEntrySheet.getDataRange();
    if (dataRange.getNumRows() <= 1) {
      Logger.log('No data to process');
      return;
    }

    const allData = dataRange.getValues();
    const dataRows = allData.slice(1);

    Logger.log(`Found ${dataRows.length} rows to process in test mode`);

    // Process each row (in reverse order to avoid index issues when deleting)
    for (let i = dataRows.length - 1; i >= 0; i--) {
      const row = dataRows[i];
      const rowIndex = i + 2; // +2 because we removed header and sheets are 1-indexed

      // Skip empty rows
      if (!row.some(cell => cell !== '')) {
        Logger.log(`Skipping empty row ${rowIndex}`);
        continue;
      }

      Logger.log(`Processing row ${rowIndex} in test mode`);
      processIndividualRowTestMode(ss, dataEntrySheet, row, rowIndex);
    }

  } catch (error) {
    Logger.log('Error in processDataEntriesTestMode: ' + error.toString());
  }
}

function trackDataEntryTimes(ss, dataRows) {
  // Get or create a hidden sheet to track entry times
  let trackingSheet = ss.getSheetByName('_DataEntryTimes');
  if (!trackingSheet) {
    trackingSheet = ss.insertSheet('_DataEntryTimes');
    trackingSheet.hideSheet();
    trackingSheet.getRange(1, 1, 1, 3).setValues([['RowData', 'EntryTime', 'RowIndex']]);
  }

  const currentTime = new Date();
  const existingTimes = trackingSheet.getDataRange().getValues().slice(1); // Skip header

  // Check each data row to see if it's new
  dataRows.forEach((row, index) => {
    const rowIndex = index + 2; // Actual sheet row number
    const rowKey = JSON.stringify(row); // Use row data as unique identifier

    // Check if this row is already tracked
    const isTracked = existingTimes.some(trackedRow => trackedRow[0] === rowKey);

    if (!isTracked && row.some(cell => cell !== '')) { // Only track non-empty rows
      // Add new entry time
      const lastRow = trackingSheet.getLastRow();
      trackingSheet.getRange(lastRow + 1, 1, 1, 3).setValues([[rowKey, currentTime, rowIndex]]);
      Logger.log(`Tracking new entry at row ${rowIndex}`);
    }
  });
}

function shouldProcessRow(ss, row, rowIndex) {
  // Skip empty rows
  if (!row.some(cell => cell !== '')) return false;

  // Get tracking sheet
  const trackingSheet = ss.getSheetByName('_DataEntryTimes');
  if (!trackingSheet) return false;

  const rowKey = JSON.stringify(row);
  const trackingData = trackingSheet.getDataRange().getValues().slice(1); // Skip header

  // Find when this row was first entered
  const trackedRow = trackingData.find(trackedRow => trackedRow[0] === rowKey);

  if (!trackedRow) {
    // Row not tracked yet, don't process
    return false;
  }

  const entryTime = new Date(trackedRow[1]);
  const currentTime = new Date();
  const timeDifference = currentTime.getTime() - entryTime.getTime();
  const fiveMinutesInMs = 5 * 60 * 1000;

  const shouldProcess = timeDifference >= fiveMinutesInMs;

  if (shouldProcess) {
    Logger.log(`Row ${rowIndex} ready for processing (entered ${Math.round(timeDifference / 60000)} minutes ago)`);
  }

  return shouldProcess;
}

function processIndividualRow(ss, dataEntrySheet, row, rowIndex) {
  const [date, name, course, sitting, fullPrice, actualPrice, orderType] = row;

  // First check: Failed orders
  if (isFailedOrder(orderType)) {
    moveToFailedOrders(ss, row);
    deleteRowFromDataEntry(dataEntrySheet, rowIndex);
    cleanupTrackingEntry(ss, row);
    return;
  }

  // Second check: Valid conditions for monthly sheet
  if (meetsMonthlySheetConditions(date, name, fullPrice, actualPrice)) {
    moveToMonthlySheet(ss, row);
    deleteRowFromDataEntry(dataEntrySheet, rowIndex);
    cleanupTrackingEntry(ss, row);
    return;
  }

  // If neither condition met, move to Sorting sheet
  moveToSortingSheet(ss, row);
  deleteRowFromDataEntry(dataEntrySheet, rowIndex);
  cleanupTrackingEntry(ss, row);
}

function processIndividualRowTestMode(ss, dataEntrySheet, row, rowIndex) {
  const [date, name, course, sitting, fullPrice, actualPrice, orderType] = row;

  // First check: Failed orders
  if (isFailedOrder(orderType)) {
    moveToFailedOrders(ss, row);
    deleteRowFromDataEntry(dataEntrySheet, rowIndex);
    Logger.log(`Row ${rowIndex}: Moved to Failed Orders`);
    return;
  }

  // Second check: Valid conditions for monthly sheet
  if (meetsMonthlySheetConditions(date, name, fullPrice, actualPrice)) {
    moveToMonthlySheet(ss, row);
    deleteRowFromDataEntry(dataEntrySheet, rowIndex);
    Logger.log(`Row ${rowIndex}: Moved to Monthly Sheet`);
    return;
  }

  // If neither condition met, move to Sorting sheet
  moveToSortingSheet(ss, row);
  deleteRowFromDataEntry(dataEntrySheet, rowIndex);
  Logger.log(`Row ${rowIndex}: Moved to Sort Sheet`);
}

function cleanupTrackingEntry(ss, row) {
  try {
    const trackingSheet = ss.getSheetByName('_DataEntryTimes');
    if (!trackingSheet) return;

    const rowKey = JSON.stringify(row);
    const trackingData = trackingSheet.getDataRange();
    const values = trackingData.getValues();

    // Find and delete the tracking entry
    for (let i = 1; i < values.length; i++) { // Start from 1 to skip header
      if (values[i][0] === rowKey) {
        trackingSheet.deleteRow(i + 1); // +1 because sheet rows are 1-indexed
        Logger.log('Cleaned up tracking entry');
        break;
      }
    }
  } catch (error) {
    Logger.log('Error cleaning up tracking entry: ' + error.toString());
  }
}

function isFailedOrder(orderType) {
  if (!orderType) return false;
  return orderType.toString().toLowerCase().includes('failed');
}

function meetsMonthlySheetConditions(date, name, fullPrice, actualPrice) {
  // Check if required fields have data
  if (!date || !name || !fullPrice || !actualPrice) {
    return false;
  }

  // UPDATED: Check if fullPrice is 997, 822, 647, or 597
  const validFullPrices = [997, 822, 647, 597];
  if (!validFullPrices.includes(Number(fullPrice))) {
    return false;
  }

  // UPDATED: Check if actualPrice is 997, 822, 647, 597, 522, 397, 347, 300, or 297
  const validActualPrices = [997, 822, 647, 597, 522, 397, 347, 300, 297];
  if (!validActualPrices.includes(Number(actualPrice))) {
    return false;
  }

  return true;
}

function moveToFailedOrders(ss, row) {
  let failedSheet = ss.getSheetByName('Failed orders');
  if (!failedSheet) {
    failedSheet = createFailedOrdersSheet(ss);
  }

  // FIXED: Add to END of sheet (oldest first, newest last)
  const lastRow = failedSheet.getLastRow();
  failedSheet.getRange(lastRow + 1, 1, 1, row.length).setValues([row]);

  Logger.log('Moved failed order to Failed orders sheet');
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

  // FIXED: Add to END of sheet instead of row 2 (oldest first, newest last)
  const lastRow = monthlySheet.getLastRow();
  monthlySheet.getRange(lastRow + 1, 1, 1, completeRow.length).setValues([completeRow]);

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
    Logger.log(`Added to ${sheetName} sheet - Payment Plan: ${paymentInfo.instalment} - Added to Instalment Tracker - Engagement Transfer Attempted`);
  } else {
    Logger.log(`Added to ${sheetName} sheet - Full Payment - Engagement Transfer Attempted`);
  }
}

function moveToSortingSheet(ss, row) {
  let sortingSheet = ss.getSheetByName('Sort');
  if (!sortingSheet) {
    sortingSheet = createSortingSheet(ss);
  }

  // FIXED: Add to END of sheet (oldest first, newest last)
  const lastRow = sortingSheet.getLastRow();
  sortingSheet.getRange(lastRow + 1, 1, 1, row.length).setValues([row]);

  // Add checkboxes in columns H, I, J (8, 9, 10)
  addCheckboxesToSortingRow(sortingSheet, lastRow + 1);

  Logger.log('Moved data to Sorting sheet with checkboxes');
}

function addCheckboxesToSortingRow(sheet, rowNumber) {
  // Add checkboxes to columns H, I, J (Manually Entered, Employer Invoice, Other)
  try {
    const checkboxRange = sheet.getRange(rowNumber, 8, 1, 3); // Columns H, I, J
    checkboxRange.insertCheckboxes();
    Logger.log(`Added checkboxes to row ${rowNumber} in Sort sheet using insertCheckboxes()`);
  } catch (error1) {
    Logger.log('insertCheckboxes() failed for Sort sheet: ' + error1.toString());
    try {
      // Try data validation method
      for (let col = 8; col <= 10; col++) {
        const cell = sheet.getRange(rowNumber, col);
        const rule = SpreadsheetApp.newDataValidation()
          .requireCheckbox()
          .build();
        cell.setDataValidation(rule);
        cell.setValue(false);
      }
      Logger.log(`Added checkboxes to row ${rowNumber} using data validation`);
    } catch (error2) {
      Logger.log('Data validation failed for Sort sheet: ' + error2.toString());
      // Fallback to boolean values
      const checkboxRange = sheet.getRange(rowNumber, 8, 1, 3);
      checkboxRange.setValues([[false, false, false]]);
      Logger.log(`Set boolean values as fallback for row ${rowNumber}`);
    }
  }
}