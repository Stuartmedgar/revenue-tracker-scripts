// ===============================================
// FIXED SORT SHEET CHECKBOX SYSTEM
// ===============================================

function fixSortSheetCompletely() {
  Logger.log('🔧 COMPLETELY FIXING SORT SHEET CHECKBOX SYSTEM');
  Logger.log('================================================');
  
  // Step 1: Clear all existing triggers
  Logger.log('\n1️⃣ Clearing existing triggers...');
  clearAllCheckboxTriggers();
  
  // Step 2: Reset all tracking
  Logger.log('\n2️⃣ Resetting tracking systems...');
  resetAllTrackingSystems();
  
  // Step 3: Set up fresh triggers
  Logger.log('\n3️⃣ Setting up fresh triggers...');
  setupImprovedCheckboxTriggers();
  
  // Step 4: Fix existing checkboxes
  Logger.log('\n4️⃣ Fixing existing checkboxes...');
  fixExistingSortCheckboxes();
  
  Logger.log('\n✅ SORT SHEET COMPLETELY FIXED!');
  Logger.log('💡 Test by checking a checkbox and waiting 5 minutes');
  Logger.log('💡 Or run testSortCheckboxesImmediately() to test without delay');
}

function clearAllCheckboxTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  let deletedCount = 0;
  
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'processCheckboxChanges' ||
        trigger.getHandlerFunction() === 'processCheckboxChangesImproved') {
      ScriptApp.deleteTrigger(trigger);
      deletedCount++;
    }
  });
  
  Logger.log(`Deleted ${deletedCount} old checkbox triggers`);
}

function resetAllTrackingSystems() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Delete old tracking sheets
  const trackingSheets = ['_SortCheckboxTimes', '_AwaitingInvoiceCheckboxTimes', '_DataEntryTimes'];
  
  trackingSheets.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      ss.deleteSheet(sheet);
      Logger.log(`Deleted tracking sheet: ${sheetName}`);
    }
  });
  
  Logger.log('All tracking systems reset');
}

function setupImprovedCheckboxTriggers() {
  // Set up improved trigger that runs every 5 minutes (Google's minimum allowed)
  ScriptApp.newTrigger('processCheckboxChangesImproved')
    .timeBased()
    .everyMinutes(5)
    .create();
    
  Logger.log('Improved checkbox trigger set up (every 5 minutes)');
}

function processCheckboxChangesImproved() {
  try {
    Logger.log('🔄 Running improved checkbox processing...');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Process Sort sheet checkboxes
    processSortSheetImproved(ss);
    
    // Process Awaiting Invoice checkboxes
    processAwaitingInvoiceImproved(ss);
    
  } catch (error) {
    Logger.log(`❌ Error in improved checkbox processing: ${error.toString()}`);
  }
}

function processSortSheetImproved(ss) {
  const sortSheet = ss.getSheetByName('Sort');
  if (!sortSheet) {
    Logger.log('Sort sheet not found');
    return;
  }
  
  const dataRange = sortSheet.getDataRange();
  if (dataRange.getNumRows() <= 1) {
    Logger.log('No data in Sort sheet');
    return;
  }
  
  const allData = dataRange.getValues();
  const dataRows = allData.slice(1); // Skip header
  
  Logger.log(`📊 Checking ${dataRows.length} rows in Sort sheet`);
  
  // Track checkbox states with improved tracking
  trackCheckboxStatesImproved(ss, dataRows, 'Sort');
  
  // Process rows with checked boxes (reverse order for safe deletion)
  let processedCount = 0;
  for (let i = dataRows.length - 1; i >= 0; i--) {
    const row = dataRows[i];
    const rowIndex = i + 2; // +2 for header and 1-indexed
    
    // Skip empty rows
    if (!row[1]) continue; // No name
    
    const manuallyEntered = row[7]; // Column H
    const employerInvoice = row[8];  // Column I  
    const other = row[9];           // Column J
    
    // Check if any checkbox is checked
    const hasCheckedBox = manuallyEntered === true || employerInvoice === true || other === true;
    
    if (hasCheckedBox) {
      Logger.log(`🔍 Row ${rowIndex} (${row[1]}): Checkboxes - H=${manuallyEntered}, I=${employerInvoice}, J=${other}`);
      
      // Check if enough time has passed (improved method)
      if (shouldProcessCheckboxRowImproved(ss, row, rowIndex)) {
        Logger.log(`✅ Processing row ${rowIndex} - time requirement met`);
        
        // Process the row
        processCheckboxRowImproved(ss, sortSheet, row, rowIndex);
        processedCount++;
      } else {
        Logger.log(`⏳ Row ${rowIndex} - waiting for time delay`);
      }
    }
  }
  
  if (processedCount > 0) {
    Logger.log(`✅ Processed ${processedCount} rows from Sort sheet`);
  } else {
    Logger.log('No rows were ready for processing');
  }
}

function trackCheckboxStatesImproved(ss, dataRows, sheetName) {
  // Get or create improved tracking sheet
  let trackingSheet = ss.getSheetByName(`_${sheetName}CheckboxTimes`);
  
  if (!trackingSheet) {
    trackingSheet = ss.insertSheet(`_${sheetName}CheckboxTimes`);
    trackingSheet.hideSheet();
    trackingSheet.getRange(1, 1, 1, 5).setValues([['RowKey', 'FirstCheckTime', 'RowIndex', 'StudentName', 'CheckboxStates']]);
  }
  
  const currentTime = new Date();
  const existingData = trackingSheet.getDataRange().getValues().slice(1); // Skip header
  
  dataRows.forEach((row, index) => {
    const rowIndex = index + 2;
    const studentName = row[1];
    
    // Skip empty rows
    if (!studentName) return;
    
    const rowKey = JSON.stringify(row.slice(0, 7)); // Original data only
    const checkboxStates = JSON.stringify([row[7], row[8], row[9]]); // Just checkbox values
    
    // Check if any checkbox is checked
    const hasCheckedBox = row[7] === true || row[8] === true || row[9] === true;
    
    // Find existing tracking entry
    const existingEntryIndex = existingData.findIndex(tracked => tracked[0] === rowKey);
    
    if (hasCheckedBox) {
      if (existingEntryIndex === -1) {
        // New checkbox change - start tracking
        const lastRow = trackingSheet.getLastRow();
        trackingSheet.getRange(lastRow + 1, 1, 1, 5).setValues([[
          rowKey, 
          currentTime, 
          rowIndex, 
          studentName, 
          checkboxStates
        ]]);
        Logger.log(`📝 Started tracking checkbox change for ${studentName} (Row ${rowIndex})`);
      } else {
        // Update existing entry if checkbox states changed
        const existingStates = existingData[existingEntryIndex][4];
        if (existingStates !== checkboxStates) {
          const trackingRowIndex = existingEntryIndex + 2; // +2 for header and 1-indexed
          trackingSheet.getRange(trackingRowIndex, 5).setValue(checkboxStates);
          Logger.log(`📝 Updated checkbox states for ${studentName} (Row ${rowIndex})`);
        }
      }
    } else {
      // No checkboxes checked - remove tracking if exists
      if (existingEntryIndex !== -1) {
        const trackingRowIndex = existingEntryIndex + 2;
        trackingSheet.deleteRow(trackingRowIndex);
        Logger.log(`🗑️ Removed tracking for ${studentName} - no checkboxes checked`);
      }
    }
  });
}

function shouldProcessCheckboxRowImproved(ss, row, rowIndex) {
  const trackingSheet = ss.getSheetByName('_SortCheckboxTimes');
  if (!trackingSheet) return false;
  
  const rowKey = JSON.stringify(row.slice(0, 7));
  const trackingData = trackingSheet.getDataRange().getValues().slice(1);
  
  // Find the tracking entry
  const trackedEntry = trackingData.find(tracked => tracked[0] === rowKey);
  
  if (!trackedEntry || !trackedEntry[1]) {
    Logger.log(`No tracking found for row ${rowIndex}`);
    return false;
  }
  
  const firstCheckTime = new Date(trackedEntry[1]);
  const currentTime = new Date();
  const timeDifference = currentTime.getTime() - firstCheckTime.getTime();
  const fiveMinutesInMs = 5 * 60 * 1000;
  
  const minutesWaited = Math.round(timeDifference / 60000);
  const shouldProcess = timeDifference >= fiveMinutesInMs;
  
  Logger.log(`Row ${rowIndex}: Waited ${minutesWaited} minutes, Should process: ${shouldProcess}`);
  
  return shouldProcess;
}

function processCheckboxRowImproved(ss, sortSheet, row, rowIndex) {
  const originalData = row.slice(0, 7); // First 7 columns
  const manuallyEntered = row[7]; // Column H
  const employerInvoice = row[8];  // Column I
  const other = row[9];           // Column J
  
  try {
    if (manuallyEntered === true) {
      Logger.log(`📤 Moving ${row[1]} to monthly sheet (Manually Entered)`);
      moveToMonthlySheetFromSorting(ss, originalData);
      
    } else if (employerInvoice === true) {
      Logger.log(`📤 Moving ${row[1]} to Awaiting Employer Invoice`);
      moveToAwaitingEmployerInvoiceImproved(ss, originalData);
      
    } else if (other === true) {
      Logger.log(`🗑️ Deleting ${row[1]} (Other)`);
      // Just delete, no move needed
    }
    
    // Delete the row from Sort sheet
    sortSheet.deleteRow(rowIndex);
    
    // Clean up tracking
    cleanupTrackingEntryImproved(ss, row, 'Sort');
    
    Logger.log(`✅ Successfully processed ${row[1]} from Sort sheet`);
    
  } catch (error) {
    Logger.log(`❌ Error processing row ${rowIndex}: ${error.toString()}`);
  }
}

function moveToAwaitingEmployerInvoiceImproved(ss, originalData) {
  let awaitingSheet = ss.getSheetByName('Awaiting Employer Invoice');
  
  if (!awaitingSheet) {
    awaitingSheet = createAwaitingEmployerInvoiceSheet(ss);
  }
  
  // Add to end of sheet
  const lastRow = awaitingSheet.getLastRow();
  const newRowIndex = lastRow + 1;
  
  // Add the data
  awaitingSheet.getRange(newRowIndex, 1, 1, originalData.length).setValues([originalData]);
  
  // Add reminder text in column I
  awaitingSheet.getRange(newRowIndex, 9).setValue('Change date to invoice paid date');
  
  // Add checkbox in column H with improved method
  addCheckboxImproved(awaitingSheet, newRowIndex, 8);
  
  Logger.log(`✅ Added ${originalData[1]} to Awaiting Employer Invoice sheet`);
}

function addCheckboxImproved(sheet, row, column) {
  const cell = sheet.getRange(row, column);
  
  try {
    // Method 1: Direct insertion
    cell.insertCheckboxes();
    Logger.log(`✅ Checkbox added to row ${row}, column ${column}`);
    
  } catch (error) {
    Logger.log(`Method 1 failed: ${error.toString()}`);
    
    try {
      // Method 2: Data validation
      const rule = SpreadsheetApp.newDataValidation()
        .requireCheckbox()
        .build();
      cell.setDataValidation(rule);
      cell.setValue(false);
      Logger.log(`✅ Checkbox added via validation to row ${row}, column ${column}`);
      
    } catch (error2) {
      Logger.log(`Method 2 failed: ${error2.toString()}`);
      // Set as boolean fallback
      cell.setValue(false);
      Logger.log(`⚠️ Set boolean value as fallback for row ${row}, column ${column}`);
    }
  }
}

function cleanupTrackingEntryImproved(ss, row, sheetName) {
  try {
    const trackingSheet = ss.getSheetByName(`_${sheetName}CheckboxTimes`);
    if (!trackingSheet) return;
    
    const rowKey = JSON.stringify(row.slice(0, 7));
    const trackingData = trackingSheet.getDataRange().getValues();
    
    // Find and delete the tracking entry
    for (let i = 1; i < trackingData.length; i++) {
      if (trackingData[i][0] === rowKey) {
        trackingSheet.deleteRow(i + 1);
        Logger.log(`🧹 Cleaned up tracking for ${row[1]}`);
        break;
      }
    }
  } catch (error) {
    Logger.log(`Error cleaning up tracking: ${error.toString()}`);
  }
}

function processAwaitingInvoiceImproved(ss) {
  const awaitingSheet = ss.getSheetByName('Awaiting Employer Invoice');
  if (!awaitingSheet) return;
  
  const dataRange = awaitingSheet.getDataRange();
  if (dataRange.getNumRows() <= 1) return;
  
  const allData = dataRange.getValues();
  const dataRows = allData.slice(1);
  
  // Track and process invoice paid checkboxes
  trackCheckboxStatesImproved(ss, dataRows, 'AwaitingInvoice');
  
  for (let i = dataRows.length - 1; i >= 0; i--) {
    const row = dataRows[i];
    const rowIndex = i + 2;
    
    const invoicePaid = row[7]; // Column H
    
    if (invoicePaid === true && shouldProcessCheckboxRowImproved(ss, row, rowIndex)) {
      processInvoicePaidRowImproved(ss, awaitingSheet, row, rowIndex);
    }
  }
}

function processInvoicePaidRowImproved(ss, awaitingSheet, row, rowIndex) {
  // Update actual price to equal full price for invoice payments
  let originalData = row.slice(0, 7);
  const fullPrice = originalData[4];
  originalData[5] = fullPrice; // Set actual price = full price
  
  Logger.log(`📤 Moving ${originalData[1]} from invoice to monthly sheet with full payment`);
  
  // Move to monthly sheet
  moveToMonthlySheetFromInvoice(ss, originalData);
  
  // Delete from awaiting sheet
  awaitingSheet.deleteRow(rowIndex);
  
  // Clean up tracking
  cleanupTrackingEntryImproved(ss, row, 'AwaitingInvoice');
}

function fixExistingSortCheckboxes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sortSheet = ss.getSheetByName('Sort');
  
  if (!sortSheet) {
    Logger.log('Sort sheet not found');
    return;
  }
  
  const dataRange = sortSheet.getDataRange();
  if (dataRange.getNumRows() <= 1) {
    Logger.log('No data in Sort sheet to fix');
    return;
  }
  
  Logger.log('🔧 Fixing existing checkboxes in Sort sheet...');
  
  // Fix checkboxes in columns H, I, J for all data rows
  for (let row = 2; row <= dataRange.getNumRows(); row++) {
    for (let col = 8; col <= 10; col++) { // Columns H, I, J
      const cell = sortSheet.getRange(row, col);
      const currentValue = cell.getValue();
      
      // If it's not already a proper checkbox, fix it
      if (typeof currentValue === 'string' || currentValue === '') {
        addCheckboxImproved(sortSheet, row, col);
      }
    }
  }
  
  Logger.log('✅ Fixed all checkboxes in Sort sheet');
}

// ===============================================
// TESTING AND DEBUGGING FUNCTIONS
// ===============================================

function testSortCheckboxesImmediately() {
  Logger.log('🧪 TESTING SORT CHECKBOXES IMMEDIATELY (NO DELAY)');
  Logger.log('================================================');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sortSheet = ss.getSheetByName('Sort');
  
  if (!sortSheet) {
    Logger.log('❌ Sort sheet not found');
    return;
  }
  
  const dataRange = sortSheet.getDataRange();
  if (dataRange.getNumRows() <= 1) {
    Logger.log('❌ No data in Sort sheet');
    return;
  }
  
  const allData = dataRange.getValues();
  const dataRows = allData.slice(1);
  
  Logger.log(`📊 Found ${dataRows.length} rows to check`);
  
  let processedCount = 0;
  
  // Process in reverse order to avoid index issues
  for (let i = dataRows.length - 1; i >= 0; i--) {
    const row = dataRows[i];
    const rowIndex = i + 2;
    
    // Skip empty rows
    if (!row[1]) continue;
    
    const manuallyEntered = row[7];
    const employerInvoice = row[8];
    const other = row[9];
    
    Logger.log(`🔍 Row ${rowIndex} (${row[1]}): H=${manuallyEntered}, I=${employerInvoice}, J=${other}`);
    
    if (manuallyEntered === true || employerInvoice === true || other === true) {
      Logger.log(`✅ Processing ${row[1]} immediately (test mode)`);
      processCheckboxRowImproved(ss, sortSheet, row, rowIndex);
      processedCount++;
    }
  }
  
  Logger.log(`🎉 Test complete! Processed ${processedCount} rows immediately`);
}

function debugSortCheckboxTracking() {
  Logger.log('🔍 DEBUGGING SORT CHECKBOX TRACKING');
  Logger.log('==================================');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Check Sort sheet data
  const sortSheet = ss.getSheetByName('Sort');
  if (sortSheet) {
    const dataRange = sortSheet.getDataRange();
    const allData = dataRange.getValues();
    const dataRows = allData.slice(1);
    
    Logger.log(`📊 Sort sheet has ${dataRows.length} data rows`);
    
    dataRows.forEach((row, index) => {
      const rowIndex = index + 2;
      if (row[1]) { // Has name
        const hasChecked = row[7] === true || row[8] === true || row[9] === true;
        Logger.log(`Row ${rowIndex} (${row[1]}): Checked=${hasChecked}, H=${row[7]}, I=${row[8]}, J=${row[9]}`);
      }
    });
  }
  
  // Check tracking sheet
  const trackingSheet = ss.getSheetByName('_SortCheckboxTimes');
  if (trackingSheet) {
    const trackingData = trackingSheet.getDataRange().getValues();
    Logger.log(`\n📋 Tracking sheet has ${trackingData.length - 1} entries`);
    
    trackingData.forEach((row, index) => {
      if (index === 0) {
        Logger.log(`Headers: ${row.join(', ')}`);
      } else {
        const studentName = row[3];
        const checkTime = row[1];
        const minutesAgo = checkTime ? Math.round((new Date() - new Date(checkTime)) / 60000) : 'Unknown';
        Logger.log(`Entry ${index}: ${studentName}, ${minutesAgo} minutes ago`);
      }
    });
  } else {
    Logger.log('\n❌ No tracking sheet found');
  }
}

function resetSortSheetCompletely() {
  Logger.log('🔄 COMPLETELY RESETTING SORT SHEET SYSTEM');
  Logger.log('=========================================');
  
  fixSortSheetCompletely();
  
  Logger.log('\n🎯 WHAT TO DO NEXT:');
  Logger.log('1. Check a checkbox in the Sort sheet');
  Logger.log('2. Wait 5 minutes OR run testSortCheckboxesImmediately()');
  Logger.log('3. The row should move to the appropriate sheet');
  Logger.log('4. Run debugSortCheckboxTracking() to see what\'s happening');
}