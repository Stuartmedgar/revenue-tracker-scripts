// ===============================================
// CODE.GS - Main Controller (Updated with Date Ordering Fixes)
// ===============================================

function setupRevenueTriggers() {
  // Clear existing triggers to avoid duplicates
  clearExistingTriggers();

  // Set up time-based trigger to run every 5 minutes
  ScriptApp.newTrigger('processDataEntries')
    .timeBased()
    .everyMinutes(5)
    .create();

  Logger.log('Revenue tracker triggers set up successfully');
}

function clearExistingTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'processDataEntries') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function initializeRevenueTracker() {
  Logger.log('Initializing Revenue Tracker Phase 3...');

  setupRevenueTriggers();
  setupCheckboxTriggers();
  setupInstalmentTracking();
  setupWeeklyEmailMonitoring();

  Logger.log('Revenue Tracker Phase 3 initialized successfully!');
  Logger.log('Email monitoring set for Mondays at 7 AM UK time');
}

function testWeeklyEmail() {
  Logger.log('Testing weekly email generation...');
  sendWeeklyMonitoringEmail();
  Logger.log('Test email sent - check your inbox!');
}

function debugEmailSystem() {
  Logger.log('=== EMAIL SYSTEM DEBUG ===');

  try {
    // Check what data would be found
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const currentDate = new Date();

    Logger.log(`Current date: ${currentDate}`);
    Logger.log(`30 days ago: ${new Date(currentDate.getTime() - (30 * 24 * 60 * 60 * 1000))}`);
    Logger.log(`7 days ago: ${new Date(currentDate.getTime() - (7 * 24 * 60 * 60 * 1000))}`);

    // Check each data source
    const unpaidInvoices = findUnpaidInvoices(ss, currentDate);
    const missedInstalments = findMissedInstalments(ss, currentDate);
    const incompletePayments = findIncompletePayments(ss, currentDate);

    Logger.log(`Found ${unpaidInvoices.length} unpaid invoices`);
    Logger.log(`Found ${missedInstalments.length} missed instalments`);
    Logger.log(`Found ${incompletePayments.length} incomplete payments`);

    const totalIssues = unpaidInvoices.length + missedInstalments.length + incompletePayments.length;
    Logger.log(`Total issues: ${totalIssues}`);

    if (totalIssues === 0) {
      Logger.log('❌ NO ISSUES FOUND - This is why no email was sent!');
      Logger.log('The system only sends emails when there are issues to report.');
    } else {
      Logger.log('✅ Issues found - email should be sent');

      // Try sending a test email with simplified approach
      Logger.log('Attempting to send test email...');

      const emailBody = generateEmailBody(unpaidInvoices, missedInstalments, incompletePayments, currentDate);
      const subject = `Revenue Tracker Alert TEST - ${formatDate(currentDate)}`;

      MailApp.sendEmail({
        subject: subject,
        htmlBody: emailBody
      });

      Logger.log('Test email sent to spreadsheet owner!');
    }

  } catch (error) {
    Logger.log(`❌ ERROR: ${error.toString()}`);
  }
}

function forceTestEmail() {
  Logger.log('Forcing test email with sample data...');

  try {
    const currentDate = new Date();

    // Create fake sample data for testing
    const sampleUnpaidInvoices = [
      {
        name: 'Test Student 1',
        course: 'Platinum',
        amount: 997,
        daysOverdue: 35,
        date: new Date('2025-06-30')
      }
    ];

    const sampleMissedInstalments = [
      {
        name: 'Test Student 2',
        course: 'Revision',
        amountPaid: 347,
        nextDue: new Date('2025-07-20'),
        daysOverdue: 15
      }
    ];

    const sampleIncompletePayments = [
      {
        name: 'Test Student 3',
        course: 'Tuition',
        fullPrice: 597,
        amountPaid: 550,
        shortfall: 47
      }
    ];

    const emailBody = generateEmailBody(sampleUnpaidInvoices, sampleMissedInstalments, sampleIncompletePayments, currentDate);
    const subject = `Revenue Tracker - FORCED TEST EMAIL - ${formatDate(currentDate)}`;

    MailApp.sendEmail({
      subject: subject,
      htmlBody: emailBody
    });

    Logger.log(`Forced test email sent with sample data`);

  } catch (error) {
    Logger.log(`Error sending forced test email: ${error.toString()}`);
  }
}

function previewWeeklyEmail() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentDate = new Date();

  const unpaidInvoices = findUnpaidInvoices(ss, currentDate);
  const missedInstalments = findMissedInstalments(ss, currentDate);
  const incompletePayments = findIncompletePayments(ss, currentDate);

  const emailBody = generateEmailBody(unpaidInvoices, missedInstalments, incompletePayments, currentDate);

  Logger.log('=== EMAIL PREVIEW ===');
  Logger.log(`Unpaid Invoices: ${unpaidInvoices.length}`);
  Logger.log(`Missed Instalments: ${missedInstalments.length}`);
  Logger.log(`Incomplete Payments: ${incompletePayments.length}`);
  Logger.log('Email body generated - ready to send');

  return emailBody;
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

function testInstalmentTracking() {
  Logger.log('=== Testing Instalment Tracking ===');

  // Test creating new instalment record
  processInstalmentPayment('Test Student', 'Platinum', 997, 397, new Date('2025-07-15'));

  // Test updating existing record
  processInstalmentPayment('Test Student', 'Platinum', 997, 300, new Date('2025-08-15'));

  // Test completing payment
  processInstalmentPayment('Test Student', 'Platinum', 997, 300, new Date('2025-09-15'));

  Logger.log('Test complete - check Instalment Tracker sheet');
}

function testProcessingNow() {
  Logger.log('Running manual test (bypassing 5-minute delay)...');
  processDataEntriesTestMode();
  Logger.log('Manual test complete - check the logs!');
}

function testCheckboxProcessingNow() {
  Logger.log('Running checkbox test (bypassing 5-minute delay)...');
  processCheckboxChangesTestMode();
  Logger.log('Checkbox test complete - check the logs!');
}

function testSimpleCheckbox() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let testSheet = ss.getSheetByName('CheckboxTest');

  if (!testSheet) {
    testSheet = ss.insertSheet('CheckboxTest');
  }

  // Clear any existing content
  testSheet.clear();

  // Test different methods
  testSheet.getRange(1, 1).setValue('Method 1 - insertCheckboxes():');
  try {
    const cell1 = testSheet.getRange(1, 2);
    cell1.insertCheckboxes();
    Logger.log('Method 1 worked');
  } catch (error) {
    Logger.log('Method 1 failed: ' + error.toString());
  }

  testSheet.getRange(2, 1).setValue('Method 2 - Data Validation:');
  try {
    const cell2 = testSheet.getRange(2, 2);
    const rule = SpreadsheetApp.newDataValidation()
      .requireCheckbox()
      .build();
    cell2.setDataValidation(rule);
    cell2.setValue(false);
    Logger.log('Method 2 worked');
  } catch (error) {
    Logger.log('Method 2 failed: ' + error.toString());
  }

  testSheet.getRange(3, 1).setValue('Method 3 - Direct checkbox:');
  try {
    const cell3 = testSheet.getRange(3, 2);
    cell3.insertCheckboxes(true, false); // checked=true, unchecked=false
    Logger.log('Method 3 worked');
  } catch (error) {
    Logger.log('Method 3 failed: ' + error.toString());
  }

  Logger.log('Checkbox test complete - check the CheckboxTest sheet');
}

function fixExistingCheckboxes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const awaitingSheet = ss.getSheetByName('Awaiting Employer Invoice');

  if (!awaitingSheet) {
    Logger.log('Awaiting Employer Invoice sheet not found');
    return;
  }

  const dataRange = awaitingSheet.getDataRange();
  const numRows = dataRange.getNumRows();

  Logger.log(`Fixing checkboxes for ${numRows - 1} data rows`);

  // Fix all rows except header
  for (let row = 2; row <= numRows; row++) {
    const cell = awaitingSheet.getRange(row, 8); // Column H
    const currentValue = cell.getValue();

    if (typeof currentValue === 'boolean') {
      Logger.log(`Fixing row ${row} - converting boolean to checkbox`);
      insertCheckboxInAwaitingSheet(awaitingSheet, row);
    }
  }

  Logger.log('Finished fixing existing checkboxes');
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

// ===============================================
// DATE ORDERING UTILITY FUNCTIONS
// ===============================================

function sortAllMonthlySheetsByDate() {
  Logger.log('=== SORTING ALL MONTHLY SHEETS BY DATE ===');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();
  
  // Find all monthly sheets
  const monthlySheets = allSheets.filter(sheet => isMonthlySheetName(sheet.getName()));
  
  Logger.log(`Found ${monthlySheets.length} monthly sheets to sort`);
  
  monthlySheets.forEach(sheet => {
    const sheetName = sheet.getName();
    Logger.log(`\n📋 Sorting ${sheetName}...`);
    
    const dataRange = sheet.getDataRange();
    if (dataRange.getNumRows() <= 1) {
      Logger.log('  No data to sort');
      return;
    }
    
    const allData = dataRange.getValues();
    const headers = allData[0];
    const dataRows = allData.slice(1);
    
    if (dataRows.length === 0) {
      Logger.log('  No data rows to sort');
      return;
    }
    
    // Find date column
    const dateCol = headers.indexOf('Date');
    if (dateCol === -1) {
      Logger.log('  ❌ No Date column found');
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
    
    Logger.log(`  ✅ Sorted ${dataRows.length} rows by date (oldest first)`);
  });
  
  Logger.log('\n🎉 All monthly sheets sorted by date!');
}

function testDateOrdering() {
  Logger.log('=== TESTING DATE ORDERING ===');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Test with sample data with different dates
  const testRows = [
    [new Date('2025-08-20'), 'Test Student 1', 'Platinum', 'December 2025', 997, 397, 'Standard'], // Newer
    [new Date('2025-08-15'), 'Test Student 2', 'Revision', 'December 2025', 647, 347, 'Standard'], // Older
    [new Date('2025-08-18'), 'Test Student 3', 'Tuition', 'December 2025', 597, 297, 'Standard']  // Middle
  ];
  
  Logger.log('Adding test students in random order...');
  testRows.forEach((row, index) => {
    Logger.log(`  ${index + 1}. ${row[1]} - Date: ${row[0]}`);
    moveToMonthlySheet(ss, row);
  });
  
  Logger.log('\n✅ Test complete - check your August 2025 sheet');
  Logger.log('Students should appear in date order: Test Student 2 (15th), Test Student 3 (18th), Test Student 1 (20th)');
}

// ===============================================
// SITTING COLUMN FIX - EXISTING FUNCTIONS
// ===============================================

function fixAllSittingColumns() {
  Logger.log('Fixing all Sitting columns to prevent date conversion...');
  fixSittingColumnFormat();
  Logger.log('All Sitting columns fixed!');
}

function testPaymentPlanDetection() {
  Logger.log('=== Testing Payment Plan Detection ===');

  // Test cases
  const testCases = [
    { full: 997, actual: 997, expected: 'Full payment' },
    { full: 997, actual: 397, expected: 'Instalment 1 of 3' },
    { full: 997, actual: 300, expected: 'Instalment 2 or 3 of 3' },
    { full: 647, actual: 647, expected: 'Full payment' },
    { full: 647, actual: 347, expected: 'Instalment 1 of 2' },
    { full: 647, actual: 300, expected: 'Instalment 2 of 2' },
    { full: 597, actual: 597, expected: 'Full payment' },
    { full: 597, actual: 297, expected: 'Instalment 1 of 2' },
    { full: 597, actual: 300, expected: 'Instalment 2 of 2' }
  ];

  testCases.forEach(test => {
    const result = detectPaymentPlan(test.full, test.actual);
    Logger.log(`Full: ${test.full}, Actual: ${test.actual}`);
    Logger.log(` Result: ${result.course} - Payment Plan: ${result.isPaymentPlan ? 'Y' : 'N'} - ${result.instalment}`);
    Logger.log(` Expected: ${test.expected}`);
    Logger.log('');
  });

  Logger.log('Payment plan detection test complete');
}

function updateExistingMonthlySheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();

  Logger.log('Starting to update existing monthly sheets with payment plan detection...');

  allSheets.forEach(sheet => {
    const sheetName = sheet.getName();
    if (isMonthlySheetName(sheetName)) {
      updateMonthlySheetPaymentPlans(sheet);
    }
  });

  Logger.log('Finished updating all existing monthly sheets');
}

function updateMonthlySheetPaymentPlans(sheet) {
  const dataRange = sheet.getDataRange();

  if (dataRange.getNumRows() <= 1) {
    Logger.log(`${sheet.getName()} has no data to update`);
    return;
  }

  const allData = dataRange.getValues();
  const headers = allData[0];
  const dataRows = allData.slice(1);

  // Find column indices
  const fullPriceCol = headers.indexOf('Full Price');
  const actualPriceCol = headers.indexOf('Actual Price');
  const paymentPlanCol = headers.indexOf('Payment Plan');
  const instalmentCol = headers.indexOf('Instalment');

  if (fullPriceCol === -1 || actualPriceCol === -1 || paymentPlanCol === -1 || instalmentCol === -1) {
    Logger.log(`${sheet.getName()} missing required columns`);
    return;
  }

  Logger.log(`Updating ${dataRows.length} rows in ${sheet.getName()}`);

  // Process each data row
  dataRows.forEach((row, index) => {
    const rowNumber = index + 2; // +2 for header and 1-indexed
    const fullPrice = row[fullPriceCol];
    const actualPrice = row[actualPriceCol];

    if (fullPrice && actualPrice) {
      const paymentInfo = detectPaymentPlan(fullPrice, actualPrice);

      // Update Payment Plan column
      const paymentPlanValue = paymentInfo.isPaymentPlan ? 'Y' : '';
      sheet.getRange(rowNumber, paymentPlanCol + 1).setValue(paymentPlanValue);

      // Update Instalment column
      sheet.getRange(rowNumber, instalmentCol + 1).setValue(paymentInfo.instalment);

      if (paymentInfo.isPaymentPlan) {
        Logger.log(`Row ${rowNumber}: Detected ${paymentInfo.course} payment plan - ${paymentInfo.instalment}`);
      }
    }
  });

  Logger.log(`Completed updating ${sheet.getName()}`);
}

function authorizeEmailPermissions() {
  try {
    MailApp.sendEmail({
      to: 'stuartmedgarwork@gmail.com',
      subject: 'Test - Revenue Tracker Authorization',
      body: 'This email confirms your Revenue Tracker has email permissions.'
    });

    Logger.log('Authorization successful - test email sent!');

  } catch (error) {
    Logger.log('Authorization needed: ' + error.toString());
  }
}