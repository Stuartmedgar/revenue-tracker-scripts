// ===============================================
// ENHANCED DEBUGGING FOR SORT SHEET ISSUES
// ===============================================

function diagnoseSortSheetIssues() {
  Logger.log('🔍 COMPREHENSIVE SORT SHEET DIAGNOSIS');
  Logger.log('=====================================');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sortSheet = ss.getSheetByName('Sort');
  
  if (!sortSheet) {
    Logger.log('❌ Sort sheet not found');
    return;
  }
  
  // Step 1: Check raw data and checkbox values
  Logger.log('\n1️⃣ CHECKING RAW DATA AND CHECKBOXES');
  Logger.log('===================================');
  
  const dataRange = sortSheet.getDataRange();
  const allData = dataRange.getValues();
  const headers = allData[0];
  const dataRows = allData.slice(1);
  
  Logger.log(`📊 Headers: ${headers.join(', ')}`);
  Logger.log(`📊 Total rows: ${dataRows.length}`);
  
  dataRows.forEach((row, index) => {
    const rowIndex = index + 2;
    
    if (row[1]) { // Has name
      Logger.log(`\nRow ${rowIndex}: ${row[1]}`);
      Logger.log(`  Date: ${row[0]}`);
      Logger.log(`  Course: ${row[2]}`);
      Logger.log(`  Sitting: ${row[3]}`);
      Logger.log(`  Full Price: ${row[4]}`);
      Logger.log(`  Actual Price: ${row[5]}`);
      Logger.log(`  Order Type: ${row[6]}`);
      Logger.log(`  Checkbox H (Manually Entered): ${row[7]} (type: ${typeof row[7]})`);
      Logger.log(`  Checkbox I (Employer Invoice): ${row[8]} (type: ${typeof row[8]})`);
      Logger.log(`  Checkbox J (Other): ${row[9]} (type: ${typeof row[9]})`);
      
      // Check if checkboxes are actually checkboxes
      const cellH = sortSheet.getRange(rowIndex, 8);
      const cellI = sortSheet.getRange(rowIndex, 9);
      const cellJ = sortSheet.getRange(rowIndex, 10);
      
      const validationH = cellH.getDataValidation();
      const validationI = cellI.getDataValidation();
      const validationJ = cellJ.getDataValidation();
      
      Logger.log(`  Cell H validation: ${validationH ? 'Has checkbox validation' : 'NO validation'}`);
      Logger.log(`  Cell I validation: ${validationI ? 'Has checkbox validation' : 'NO validation'}`);
      Logger.log(`  Cell J validation: ${validationJ ? 'Has checkbox validation' : 'NO validation'}`);
      
      // Determine if row should be processed
      const hasCheckedBox = row[7] === true || row[8] === true || row[9] === true;
      Logger.log(`  Should be processed: ${hasCheckedBox ? 'YES' : 'NO'}`);
    }
  });
  
  // Step 2: Check triggers
  Logger.log('\n2️⃣ CHECKING TRIGGERS');
  Logger.log('===================');
  
  const triggers = ScriptApp.getProjectTriggers();
  const checkboxTriggers = triggers.filter(t => 
    t.getHandlerFunction() === 'processCheckboxChanges' || 
    t.getHandlerFunction() === 'processCheckboxChangesImproved'
  );
  
  Logger.log(`Found ${checkboxTriggers.length} checkbox triggers`);
  checkboxTriggers.forEach((trigger, index) => {
    Logger.log(`Trigger ${index + 1}: ${trigger.getHandlerFunction()}`);
  });
  
  // Step 3: Check tracking sheet
  Logger.log('\n3️⃣ CHECKING TRACKING SHEET');
  Logger.log('==========================');
  
  const trackingSheet = ss.getSheetByName('_SortCheckboxTimes');
  if (trackingSheet) {
    const trackingData = trackingSheet.getDataRange().getValues();
    Logger.log(`Tracking entries: ${trackingData.length - 1}`);
    
    trackingData.forEach((row, index) => {
      if (index === 0) {
        Logger.log(`Headers: ${row.join(', ')}`);
      } else {
        Logger.log(`Entry ${index}: Student=${row[3]}, Time=${row[1]}, States=${row[4]}`);
      }
    });
  } else {
    Logger.log('❌ No tracking sheet found');
  }
}

function forceProcessAllCheckedBoxes() {
  Logger.log('🚀 FORCE PROCESSING ALL CHECKED BOXES (NO DELAYS)');
  Logger.log('================================================');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sortSheet = ss.getSheetByName('Sort');
  
  if (!sortSheet) {
    Logger.log('❌ Sort sheet not found');
    return;
  }
  
  const dataRange = sortSheet.getDataRange();
  const allData = dataRange.getValues();
  const dataRows = allData.slice(1);
  
  Logger.log(`📊 Checking ${dataRows.length} rows for checked boxes`);
  
  let processedCount = 0;
  let skippedCount = 0;
  
  // Process in reverse order to avoid index issues when deleting
  for (let i = dataRows.length - 1; i >= 0; i--) {
    const row = dataRows[i];
    const rowIndex = i + 2;
    
    // Skip empty rows
    if (!row[1]) {
      Logger.log(`⏸️ Row ${rowIndex}: Empty, skipping`);
      skippedCount++;
      continue;
    }
    
    const manuallyEntered = row[7];
    const employerInvoice = row[8];
    const other = row[9];
    
    Logger.log(`\n🔍 Row ${rowIndex} (${row[1]}):`);
    Logger.log(`  H=${manuallyEntered}, I=${employerInvoice}, J=${other}`);
    
    if (manuallyEntered === true) {
      Logger.log(`  ✅ Processing: Manually Entered`);
      try {
        const originalData = row.slice(0, 7);
        moveToMonthlySheetFromSorting(ss, originalData);
        sortSheet.deleteRow(rowIndex);
        processedCount++;
        Logger.log(`  ✅ Success: Moved to monthly sheet`);
      } catch (error) {
        Logger.log(`  ❌ Error: ${error.toString()}`);
      }
      
    } else if (employerInvoice === true) {
      Logger.log(`  ✅ Processing: Employer Invoice`);
      try {
        const originalData = row.slice(0, 7);
        moveToAwaitingEmployerInvoiceSimple(ss, originalData);
        sortSheet.deleteRow(rowIndex);
        processedCount++;
        Logger.log(`  ✅ Success: Moved to Awaiting Employer Invoice`);
      } catch (error) {
        Logger.log(`  ❌ Error: ${error.toString()}`);
      }
      
    } else if (other === true) {
      Logger.log(`  ✅ Processing: Other (Delete)`);
      try {
        sortSheet.deleteRow(rowIndex);
        processedCount++;
        Logger.log(`  ✅ Success: Deleted row`);
      } catch (error) {
        Logger.log(`  ❌ Error: ${error.toString()}`);
      }
      
    } else {
      Logger.log(`  ⏸️ No checkboxes checked, skipping`);
      skippedCount++;
    }
  }
  
  Logger.log(`\n🎯 SUMMARY:`);
  Logger.log(`✅ Processed: ${processedCount} rows`);
  Logger.log(`⏸️ Skipped: ${skippedCount} rows`);
}

function moveToAwaitingEmployerInvoiceSimple(ss, originalData) {
  let awaitingSheet = ss.getSheetByName('Awaiting Employer Invoice');
  
  if (!awaitingSheet) {
    Logger.log('Creating Awaiting Employer Invoice sheet...');
    awaitingSheet = createAwaitingEmployerInvoiceSheet(ss);
  }
  
  // Add to end of sheet
  const lastRow = awaitingSheet.getLastRow();
  const newRowIndex = lastRow + 1;
  
  // Add the data
  awaitingSheet.getRange(newRowIndex, 1, 1, originalData.length).setValues([originalData]);
  
  // Add reminder text in column I
  awaitingSheet.getRange(newRowIndex, 9).setValue('Change date to invoice paid date');
  
  // Try to add checkbox in column H
  const checkboxCell = awaitingSheet.getRange(newRowIndex, 8);
  try {
    checkboxCell.insertCheckboxes();
  } catch (error) {
    checkboxCell.setValue(false); // Fallback to boolean
  }
  
  Logger.log(`Added ${originalData[1]} to Awaiting Employer Invoice`);
}

function fixAllCheckboxesInSort() {
  Logger.log('🔧 FIXING ALL CHECKBOXES IN SORT SHEET');
  Logger.log('=====================================');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sortSheet = ss.getSheetByName('Sort');
  
  if (!sortSheet) {
    Logger.log('❌ Sort sheet not found');
    return;
  }
  
  const dataRange = sortSheet.getDataRange();
  const numRows = dataRange.getNumRows();
  
  if (numRows <= 1) {
    Logger.log('No data rows to fix');
    return;
  }
  
  Logger.log(`Fixing checkboxes for ${numRows - 1} data rows`);
  
  // Fix checkboxes in columns H, I, J for all data rows
  for (let row = 2; row <= numRows; row++) {
    const nameCell = sortSheet.getRange(row, 2);
    if (!nameCell.getValue()) continue; // Skip empty rows
    
    Logger.log(`Fixing checkboxes for row ${row}`);
    
    for (let col = 8; col <= 10; col++) { // Columns H, I, J
      const cell = sortSheet.getRange(row, col);
      const currentValue = cell.getValue();
      
      // Clear the cell and add proper checkbox
      cell.clearContent();
      cell.clearDataValidations();
      
      try {
        cell.insertCheckboxes();
        Logger.log(`  ✅ Fixed checkbox in column ${String.fromCharCode(64 + col)}`);
      } catch (error) {
        // Fallback to data validation method
        const rule = SpreadsheetApp.newDataValidation()
          .requireCheckbox()
          .build();
        cell.setDataValidation(rule);
        cell.setValue(false);
        Logger.log(`  ⚠️ Used validation fallback for column ${String.fromCharCode(64 + col)}`);
      }
    }
  }
  
  Logger.log('✅ All checkboxes fixed');
}

function testSingleRowMovement() {
  Logger.log('🧪 TESTING SINGLE ROW MOVEMENT');
  Logger.log('==============================');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Test data
  const testData = [
    new Date('2025-01-08'),
    'Test Student Movement',
    'Platinum',
    'December 2025',
    997,
    397,
    'Standard'
  ];
  
  Logger.log('Testing movement to monthly sheet...');
  try {
    moveToMonthlySheetFromSorting(ss, testData);
    Logger.log('✅ Monthly sheet movement successful');
  } catch (error) {
    Logger.log(`❌ Monthly sheet movement failed: ${error.toString()}`);
  }
  
  Logger.log('Testing movement to awaiting invoice sheet...');
  try {
    moveToAwaitingEmployerInvoiceSimple(ss, testData);
    Logger.log('✅ Awaiting invoice movement successful');
  } catch (error) {
    Logger.log(`❌ Awaiting invoice movement failed: ${error.toString()}`);
  }
}

function checkSortSheetPermissions() {
  Logger.log('🔐 CHECKING SORT SHEET PERMISSIONS');
  Logger.log('==================================');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sortSheet = ss.getSheetByName('Sort');
  
  if (!sortSheet) {
    Logger.log('❌ Sort sheet not found');
    return;
  }
  
  try {
    // Test reading
    const testRead = sortSheet.getRange('A1').getValue();
    Logger.log('✅ Can read from Sort sheet');
    
    // Test writing
    const testCell = sortSheet.getRange('Z1');
    testCell.setValue('Permission Test');
    const readBack = testCell.getValue();
    testCell.clearContent();
    
    if (readBack === 'Permission Test') {
      Logger.log('✅ Can write to Sort sheet');
    } else {
      Logger.log('❌ Cannot write to Sort sheet');
    }
    
    // Test deleting rows
    const lastRow = sortSheet.getLastRow();
    if (lastRow > 1) {
      Logger.log('✅ Can potentially delete rows (has data)');
    } else {
      Logger.log('⚠️ No data rows to test deletion');
    }
    
  } catch (error) {
    Logger.log(`❌ Permission error: ${error.toString()}`);
  }
}