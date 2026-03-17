// ===============================================
// CHECKBOXHANDLERS.GS - Sorting Sheet Checkbox Logic (Complete with Student Engagement + Fixed Date Ordering)
// ===============================================

function setupCheckboxTriggers() {
  // Clear existing checkbox triggers
  clearCheckboxTriggers();

  // Set up time-based trigger to check for checkbox changes every 5 minutes
  ScriptApp.newTrigger('processCheckboxChanges')
    .timeBased()
    .everyMinutes(5)
    .create();

  Logger.log('Checkbox processing triggers set up successfully');
}

function clearCheckboxTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'processCheckboxChanges') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function processCheckboxChanges() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Process Sorting sheet checkboxes
    processSortingSheetCheckboxes(ss);

    // Process Awaiting Employer Invoice checkboxes
    processAwaitingInvoiceCheckboxes(ss);

  } catch (error) {
    Logger.log('Error in processCheckboxChanges: ' + error.toString());
  }
}

function processSortingSheetCheckboxes(ss) {
  const sortingSheet = ss.getSheetByName('Sort');

  if (!sortingSheet) {
    Logger.log('Sort sheet not found');
    return;
  }

  const dataRange = sortingSheet.getDataRange();
  if (dataRange.getNumRows() <= 1) {
    Logger.log('No data in Sort sheet to process');
    return;
  }

  const allData = dataRange.getValues();
  const dataRows = allData.slice(1); // Skip header

  // Track checkbox states and process changes
  trackCheckboxStates(ss, dataRows, 'Sort');

  // Process rows with checked boxes (in reverse order for safe deletion)
  for (let i = dataRows.length - 1; i >= 0; i--) {
    const row = dataRows[i];
    const rowIndex = i + 2; // +2 because we removed header and sheets are 1-indexed

    // Check if any checkbox is checked and if it's been 5 minutes
    if (shouldProcessCheckboxRow(ss, row, rowIndex, 'Sort')) {
      processSortingCheckboxRow(ss, sortingSheet, row, rowIndex);
    }
  }
}

function processSortingCheckboxRow(ss, sortingSheet, row, rowIndex) {
  // Get checkbox values (columns 8, 9, 10 = H, I, J)
  const manuallyEntered = row[7]; // Column H
  const employerInvoice = row[8]; // Column I
  const other = row[9]; // Column J

  // Get original data (first 7 columns)
  const originalData = row.slice(0, 7);

  if (manuallyEntered === true) {
    // Move to appropriate monthly sheet
    moveToMonthlySheetFromSorting(ss, originalData);
    deleteSortingRow(sortingSheet, rowIndex);
    cleanupCheckboxTracking(ss, row, 'Sort');
    Logger.log(`Row ${rowIndex}: Manually Entered - moved to monthly sheet`);

  } else if (employerInvoice === true) {
    // Move to Awaiting Employer Invoice sheet
    moveToAwaitingEmployerInvoice(ss, originalData);
    deleteSortingRow(sortingSheet, rowIndex);
    cleanupCheckboxTracking(ss, row, 'Sort');
    Logger.log(`Row ${rowIndex}: Employer Invoice - moved to awaiting invoice sheet`);

  } else if (other === true) {
    // Delete the row
    deleteSortingRow(sortingSheet, rowIndex);
    cleanupCheckboxTracking(ss, row, 'Sort');
    Logger.log(`Row ${rowIndex}: Other - deleted`);
  }
}

function moveToMonthlySheetFromSorting(ss, originalData) {
  // Use the same logic as the original monthly sheet move
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

  // FIXED: Add to END of sheet instead of row 2 (oldest first, newest last)
  const lastRow = monthlySheet.getLastRow();
  monthlySheet.getRange(lastRow + 1, 1, 1, completeRow.length).setValues([completeRow]);

  // Process for student engagement transfer
  const studentData = {
    name: originalData[1], // Name column
    sitting: extractSittingDate(originalData[3]), // Extract date from full Item Name
    actualPrice: originalData[5], // Actual Price column
    course: paymentInfo.course
  };

  processStudentForEngagement(studentData, sheetName);

  if (paymentInfo.isPaymentPlan) {
    Logger.log(`Moved data from sorting to ${sheetName} sheet - Payment Plan: ${paymentInfo.instalment} - Processed for Engagement`);
  } else {
    Logger.log(`Moved data from sorting to ${sheetName} sheet - Full Payment - Processed for Engagement`);
  }
}

function moveToAwaitingEmployerInvoice(ss, originalData) {
  let awaitingSheet = ss.getSheetByName('Awaiting Employer Invoice');

  if (!awaitingSheet) {
    awaitingSheet = createAwaitingEmployerInvoiceSheet(ss);
  }

  // FIXED: Add to END of sheet (oldest first, newest last)
  const lastRow = awaitingSheet.getLastRow();
  awaitingSheet.getRange(lastRow + 1, 1, 1, originalData.length).setValues([originalData]);

  // Add reminder text in column I (column 9)
  awaitingSheet.getRange(lastRow + 1, 9).setValue('Change date to invoice paid date');

  // Force checkbox insertion with multiple attempts
  insertCheckboxInAwaitingSheet(awaitingSheet, lastRow + 1);

  Logger.log('Moved data to Awaiting Employer Invoice sheet');
}

function deleteSortingRow(sheet, rowIndex) {
  try {
    sheet.deleteRow(rowIndex);
    Logger.log(`Deleted row ${rowIndex} from Sort sheet`);
  } catch (error) {
    Logger.log(`Error deleting sorting row ${rowIndex}: ${error.toString()}`);
  }
}

// ===============================================
// CHECKBOX STATE TRACKING
// ===============================================

function trackCheckboxStates(ss, dataRows, sheetName) {
  // Get or create tracking sheet for checkbox states
  let trackingSheet = ss.getSheetByName(`_${sheetName}CheckboxTimes`);

  if (!trackingSheet) {
    trackingSheet = ss.insertSheet(`_${sheetName}CheckboxTimes`);
    trackingSheet.hideSheet();
    trackingSheet.getRange(1, 1, 1, 4).setValues([['RowData', 'CheckboxChangeTime', 'RowIndex', 'CheckboxStates']]);
  }

  const currentTime = new Date();
  const existingTracking = trackingSheet.getDataRange().getValues().slice(1); // Skip header

  dataRows.forEach((row, index) => {
    const rowIndex = index + 2;
    const rowKey = JSON.stringify(row.slice(0, 7)); // Only track original data, not checkboxes
    const checkboxStates = JSON.stringify(row.slice(7)); // Track checkbox states separately

    // Find existing tracking entry
    const existingEntry = existingTracking.find(tracked => tracked[0] === rowKey);

    if (existingEntry) {
      // Check if checkbox states have changed
      if (existingEntry[3] !== checkboxStates) {
        // Update the tracking entry with new checkbox states and time
        const trackingRowIndex = existingTracking.indexOf(existingEntry) + 2; // +2 for header and 1-indexed
        trackingSheet.getRange(trackingRowIndex, 2, 1, 3).setValues([[currentTime, rowIndex, checkboxStates]]);
        Logger.log(`Updated checkbox state tracking for row ${rowIndex}`);
      }
    } else if (row.some(cell => cell !== '')) {
      // New row - add to tracking
      const lastRow = trackingSheet.getLastRow();
      trackingSheet.getRange(lastRow + 1, 1, 1, 4).setValues([[rowKey, null, rowIndex, checkboxStates]]);
      Logger.log(`Started tracking checkboxes for row ${rowIndex}`);
    }
  });
}

function shouldProcessCheckboxRow(ss, row, rowIndex, sheetName) {
  // Skip empty rows
  if (!row.some(cell => cell !== '')) return false;

  // Check if any checkbox is checked
  const checkboxValues = row.slice(7); // Get checkbox columns
  const hasCheckedBox = checkboxValues.some(value => value === true);

  if (!hasCheckedBox) return false;

  // Get tracking sheet
  const trackingSheet = ss.getSheetByName(`_${sheetName}CheckboxTimes`);
  if (!trackingSheet) return false;

  const rowKey = JSON.stringify(row.slice(0, 7));
  const trackingData = trackingSheet.getDataRange().getValues().slice(1);

  // Find tracking entry
  const trackedRow = trackingData.find(tracked => tracked[0] === rowKey);

  if (!trackedRow || !trackedRow[1]) {
    // No checkbox change time recorded yet
    return false;
  }

  const changeTime = new Date(trackedRow[1]);
  const currentTime = new Date();
  const timeDifference = currentTime.getTime() - changeTime.getTime();
  const fiveMinutesInMs = 5 * 60 * 1000;

  const shouldProcess = timeDifference >= fiveMinutesInMs;

  if (shouldProcess) {
    Logger.log(`Checkbox row ${rowIndex} ready for processing (changed ${Math.round(timeDifference / 60000)} minutes ago)`);
  }

  return shouldProcess;
}

function insertCheckboxInAwaitingSheet(sheet, rowNumber) {
  const checkboxCell = sheet.getRange(rowNumber, 8); // Column H

  Logger.log(`Attempting to insert checkbox in row ${rowNumber}, column H`);

  // Method 1: Direct insertCheckboxes()
  try {
    checkboxCell.clearContent(); // Clear any existing content
    checkboxCell.insertCheckboxes();

    // Verify it worked
    const validation = checkboxCell.getDataValidation();
    if (validation && validation.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.CHECKBOX) {
      Logger.log('SUCCESS: Checkbox inserted using insertCheckboxes()');
      return;
    } else {
      Logger.log('insertCheckboxes() completed but no validation found, trying method 2');
    }

  } catch (error) {
    Logger.log(`insertCheckboxes() failed: ${error.toString()}, trying method 2`);
  }

  // Method 2: Data validation approach
  try {
    checkboxCell.clearContent();
    const rule = SpreadsheetApp.newDataValidation()
      .requireCheckbox()
      .build();
    checkboxCell.setDataValidation(rule);
    checkboxCell.setValue(false);

    // Verify it worked
    const validation = checkboxCell.getDataValidation();
    if (validation && validation.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.CHECKBOX) {
      Logger.log('SUCCESS: Checkbox created using data validation');
      return;
    } else {
      Logger.log('Data validation completed but no checkbox found, trying method 3');
    }

  } catch (error) {
    Logger.log(`Data validation failed: ${error.toString()}, trying method 3`);
  }

  // Method 3: Force with flush
  try {
    checkboxCell.clearContent();
    SpreadsheetApp.flush(); // Force all pending operations to complete
    checkboxCell.insertCheckboxes();
    SpreadsheetApp.flush(); // Force the checkbox insertion to complete

    // Final verification
    const validation = checkboxCell.getDataValidation();
    if (validation && validation.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.CHECKBOX) {
      Logger.log('SUCCESS: Checkbox inserted with flush');
      return;
    } else {
      Logger.log('All methods failed - checkbox not created');
    }

  } catch (error) {
    Logger.log(`Final method failed: ${error.toString()}`);
  }

  Logger.log('WARNING: Could not create checkbox, cell will show boolean value');
}

function cleanupCheckboxTracking(ss, row, sheetName) {
  try {
    const trackingSheet = ss.getSheetByName(`_${sheetName}CheckboxTimes`);
    if (!trackingSheet) return;

    const rowKey = JSON.stringify(row.slice(0, 7));
    const trackingData = trackingSheet.getDataRange();
    const values = trackingData.getValues();

    // Find and delete the tracking entry
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === rowKey) {
        trackingSheet.deleteRow(i + 1);
        Logger.log('Cleaned up checkbox tracking entry');
        break;
      }
    }

  } catch (error) {
    Logger.log('Error cleaning up checkbox tracking: ' + error.toString());
  }
}

// ===============================================
// AWAITING INVOICE PROCESSING
// ===============================================

function processAwaitingInvoiceCheckboxes(ss) {
  const awaitingSheet = ss.getSheetByName('Awaiting Employer Invoice');

  if (!awaitingSheet) {
    Logger.log('Awaiting Employer Invoice sheet not found');
    return;
  }

  const dataRange = awaitingSheet.getDataRange();
  if (dataRange.getNumRows() <= 1) {
    Logger.log('No data in Awaiting Employer Invoice sheet to process');
    return;
  }

  const allData = dataRange.getValues();
  const dataRows = allData.slice(1); // Skip header

  // Track checkbox states
  trackCheckboxStates(ss, dataRows, 'AwaitingInvoice');

  // Process rows with checked "Invoice Paid" boxes (in reverse order for safe deletion)
  for (let i = dataRows.length - 1; i >= 0; i--) {
    const row = dataRows[i];
    const rowIndex = i + 2; // +2 because we removed header and sheets are 1-indexed

    // Check if invoice paid checkbox is checked and if it's been 5 minutes
    if (shouldProcessInvoiceRow(ss, row, rowIndex)) {
      processInvoicePaidRow(ss, awaitingSheet, row, rowIndex);
    }
  }
}

function shouldProcessInvoiceRow(ss, row, rowIndex) {
  // Skip empty rows
  if (!row.some(cell => cell !== '')) return false;

  // Check if invoice paid checkbox is checked (column 8 = H)
  const invoicePaid = row[7];
  if (invoicePaid !== true) return false;

  // Get tracking sheet
  const trackingSheet = ss.getSheetByName('_AwaitingInvoiceCheckboxTimes');
  if (!trackingSheet) return false;

  const rowKey = JSON.stringify(row.slice(0, 7)); // Original data only
  const trackingData = trackingSheet.getDataRange().getValues().slice(1);

  // Find tracking entry
  const trackedRow = trackingData.find(tracked => tracked[0] === rowKey);

  if (!trackedRow || !trackedRow[1]) {
    // No checkbox change time recorded yet
    return false;
  }

  const changeTime = new Date(trackedRow[1]);
  const currentTime = new Date();
  const timeDifference = currentTime.getTime() - changeTime.getTime();
  const fiveMinutesInMs = 5 * 60 * 1000;

  const shouldProcess = timeDifference >= fiveMinutesInMs;

  if (shouldProcess) {
    Logger.log(`Invoice row ${rowIndex} ready for processing (paid ${Math.round(timeDifference / 60000)} minutes ago)`);
  }

  return shouldProcess;
}

// ===============================================
// DEBUGGING FUNCTIONS
// ===============================================

function debugSortSheetCheckboxes() {
  Logger.log('=== DEBUGGING SORT SHEET CHECKBOXES ===');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sortingSheet = ss.getSheetByName('Sort');

  if (!sortingSheet) {
    Logger.log('❌ Sort sheet not found');
    return;
  }

  const dataRange = sortingSheet.getDataRange();
  if (dataRange.getNumRows() <= 1) {
    Logger.log('❌ No data in Sort sheet');
    return;
  }

  const allData = dataRange.getValues();
  const headers = allData[0];
  const dataRows = allData.slice(1);

  Logger.log(`📊 Found ${dataRows.length} data rows`);
  Logger.log(`📋 Headers: ${headers.join(', ')}`);

  // Check each row for checkbox values
  dataRows.forEach((row, index) => {
    const rowIndex = index + 2;
    const manuallyEntered = row[7]; // Column H
    const employerInvoice = row[8]; // Column I
    const other = row[9]; // Column J

    Logger.log(`Row ${rowIndex}:`);
    Logger.log(` Name: ${row[1]}`);
    Logger.log(` Manually Entered (H): ${manuallyEntered} (type: ${typeof manuallyEntered})`);
    Logger.log(` Employer Invoice (I): ${employerInvoice} (type: ${typeof employerInvoice})`);
    Logger.log(` Other (J): ${other} (type: ${typeof other})`);

    // Check if any checkbox is true
    if (manuallyEntered === true || employerInvoice === true || other === true) {
      Logger.log(` ✅ This row has checked boxes and should be processed`);
    } else {
      Logger.log(` ⏸️ This row has no checked boxes`);
    }
    Logger.log('');
  });
}

function forceProcessSortCheckboxes() {
  Logger.log('🚀 FORCING IMMEDIATE CHECKBOX PROCESSING (NO DELAY)');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sortingSheet = ss.getSheetByName('Sort');

  if (!sortingSheet) {
    Logger.log('❌ Sort sheet not found');
    return;
  }

  const dataRange = sortingSheet.getDataRange();
  if (dataRange.getNumRows() <= 1) {
    Logger.log('❌ No data in Sort sheet');
    return;
  }

  const allData = dataRange.getValues();
  const dataRows = allData.slice(1);

  // Process rows with checked boxes (in reverse order for safe deletion)
  let processedCount = 0;

  for (let i = dataRows.length - 1; i >= 0; i--) {
    const row = dataRows[i];
    const rowIndex = i + 2;

    const manuallyEntered = row[7]; // Column H
    const employerInvoice = row[8]; // Column I
    const other = row[9]; // Column J

    Logger.log(`🔍 Checking row ${rowIndex}: ${row[1]}`);
    Logger.log(` Checkboxes: H=${manuallyEntered}, I=${employerInvoice}, J=${other}`);

    if (manuallyEntered === true) {
      Logger.log(`✅ Processing "Manually Entered" for row ${rowIndex}`);
      const originalData = row.slice(0, 7);
      moveToMonthlySheetFromSorting(ss, originalData);
      deleteSortingRow(sortingSheet, rowIndex);
      processedCount++;
      Logger.log(`✅ Row ${rowIndex}: Moved to monthly sheet`);

    } else if (employerInvoice === true) {
      Logger.log(`✅ Processing "Employer Invoice" for row ${rowIndex}`);
      const originalData = row.slice(0, 7);
      moveToAwaitingEmployerInvoice(ss, originalData);
      deleteSortingRow(sortingSheet, rowIndex);
      processedCount++;
      Logger.log(`✅ Row ${rowIndex}: Moved to Awaiting Employer Invoice`);

    } else if (other === true) {
      Logger.log(`✅ Processing "Other" for row ${rowIndex}`);
      deleteSortingRow(sortingSheet, rowIndex);
      processedCount++;
      Logger.log(`✅ Row ${rowIndex}: Deleted`);

    } else {
      Logger.log(`⏸️ Row ${rowIndex}: No checkboxes checked, skipping`);
    }
  }

  Logger.log(`🎉 Processing complete! Processed ${processedCount} rows.`);
}

function processCheckboxesImmediatelySimple() {
  try {
    Logger.log('🚀 STARTING SIMPLE IMMEDIATE CHECKBOX PROCESSING');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sortingSheet = ss.getSheetByName('Sort');

    if (!sortingSheet) {
      Logger.log('❌ Sort sheet not found');
      return;
    }

    const dataRange = sortingSheet.getDataRange();
    if (dataRange.getNumRows() <= 1) {
      Logger.log('❌ No data in Sort sheet');
      return;
    }

    const allData = dataRange.getValues();
    const dataRows = allData.slice(1); // Skip header

    Logger.log(`📊 Found ${dataRows.length} rows to check`);

    // Process in reverse order to avoid index issues when deleting
    for (let i = dataRows.length - 1; i >= 0; i--) {
      const row = dataRows[i];
      const rowIndex = i + 2; // +2 for header and 1-indexed sheets

      // Get checkbox values
      const manuallyEntered = row[7]; // Column H
      const employerInvoice = row[8]; // Column I
      const other = row[9]; // Column J

      // Skip empty rows
      if (!row[1]) { // If no name, skip
        Logger.log(`⏸️ Row ${rowIndex}: Empty row, skipping`);
        continue;
      }

      Logger.log(`🔍 Row ${rowIndex} (${row[1]}): H=${manuallyEntered}, I=${employerInvoice}, J=${other}`);

      // Process based on checkbox state
      if (manuallyEntered === true) {
        Logger.log(`✅ Processing "Manually Entered" for ${row[1]}`);
        const originalData = row.slice(0, 7);
        moveToMonthlySheetFromSorting(ss, originalData);
        sortingSheet.deleteRow(rowIndex);
        Logger.log(`✅ ${row[1]} moved to monthly sheet and deleted from Sort`);

      } else if (employerInvoice === true) {
        Logger.log(`✅ Processing "Employer Invoice" for ${row[1]}`);
        const originalData = row.slice(0, 7);
        moveToAwaitingEmployerInvoice(ss, originalData);
        sortingSheet.deleteRow(rowIndex);
        Logger.log(`✅ ${row[1]} moved to Awaiting Employer Invoice and deleted from Sort`);

      } else if (other === true) {
        Logger.log(`✅ Processing "Other" for ${row[1]} - will delete`);
        sortingSheet.deleteRow(rowIndex);
        Logger.log(`✅ ${row[1]} deleted from Sort sheet`);

      } else {
        Logger.log(`⏸️ ${row[1]}: No checkboxes checked, leaving in Sort sheet`);
      }
    }

    Logger.log('🎉 Simple checkbox processing complete!');

  } catch (error) {
    Logger.log(`❌ Error in simple checkbox processing: ${error.toString()}`);
  }
}

// ===============================================
// TEST FUNCTIONS
// ===============================================

function testCheckboxProcessingNow() {
  Logger.log('Running checkbox test (bypassing 5-minute delay)...');
  processCheckboxChangesTestMode();
  Logger.log('Checkbox test complete - check the logs!');
}

function processCheckboxChangesTestMode() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Process Sorting sheet checkboxes immediately
    processSortingSheetCheckboxesTestMode(ss);

    // Process Awaiting Employer Invoice checkboxes immediately
    processAwaitingInvoiceCheckboxesTestMode(ss);

  } catch (error) {
    Logger.log('Error in processCheckboxChangesTestMode: ' + error.toString());
  }
}

function processAwaitingInvoiceCheckboxesTestMode(ss) {
  const awaitingSheet = ss.getSheetByName('Awaiting Employer Invoice');

  if (!awaitingSheet) {
    Logger.log('Awaiting Employer Invoice sheet not found');
    return;
  }

  const dataRange = awaitingSheet.getDataRange();
  if (dataRange.getNumRows() <= 1) {
    Logger.log('No data in Awaiting Invoice sheet to process');
    return;
  }

  const allData = dataRange.getValues();
  const dataRows = allData.slice(1);

  Logger.log(`Found ${dataRows.length} rows in Awaiting Invoice sheet for checkbox processing`);

  // Process rows with checked "Invoice Paid" boxes (in reverse order for safe deletion)
  for (let i = dataRows.length - 1; i >= 0; i--) {
    const row = dataRows[i];
    const rowIndex = i + 2;

    // Check if invoice paid checkbox is checked (column 8 = H)
    const invoicePaid = row[7]; // Column H (Invoice Paid)

    if (invoicePaid === true) {
      Logger.log(`Processing invoice paid row ${rowIndex} in test mode`);
      processInvoicePaidRow(ss, awaitingSheet, row, rowIndex);
    }
  }
}

function processSortingSheetCheckboxesTestMode(ss) {
  const sortingSheet = ss.getSheetByName('Sort');

  if (!sortingSheet) {
    Logger.log('Sort sheet not found');
    return;
  }

  const dataRange = sortingSheet.getDataRange();
  if (dataRange.getNumRows() <= 1) {
    Logger.log('No data in Sort sheet to process');
    return;
  }

  const allData = dataRange.getValues();
  const dataRows = allData.slice(1);

  Logger.log(`Found ${dataRows.length} rows in Sort sheet for checkbox processing`);

  // Process rows with checked boxes (in reverse order for safe deletion)
  for (let i = dataRows.length - 1; i >= 0; i--) {
    const row = dataRows[i];
    const rowIndex = i + 2;

    // Check if any checkbox is checked
    const checkboxValues = row.slice(7);
    const hasCheckedBox = checkboxValues.some(value => value === true);

    if (hasCheckedBox) {
      Logger.log(`Processing checkbox row ${rowIndex} in test mode`);
      processSortingCheckboxRow(ss, sortingSheet, row, rowIndex);
    }
  }
}