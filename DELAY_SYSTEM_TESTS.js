// ===============================================
// TEST THE 5-MINUTE DELAY SYSTEM
// ===============================================

function testFullDelaySystem() {
  Logger.log('🧪 TESTING COMPLETE 5-MINUTE DELAY SYSTEM');
  Logger.log('==========================================');
  
  // Step 1: Check current setup
  Logger.log('\n1️⃣ CHECKING CURRENT SETUP...');
  diagnoseFiveMinuteDelay();
  
  // Step 2: Check triggers
  Logger.log('\n2️⃣ CHECKING TRIGGERS...');
  checkTriggerStatus();
  
  // Step 3: Simulate checkbox tracking
  Logger.log('\n3️⃣ SIMULATING CHECKBOX TRACKING...');
  manuallyTrackCheckboxChange();
  
  // Step 4: Test immediate processing (what should happen after 5 minutes)
  Logger.log('\n4️⃣ TESTING WHAT HAPPENS AFTER 5 MINUTES...');
  testProcessingAfterDelay();
  
  Logger.log('\n🎯 TEST COMPLETE - Review the logs above for issues');
}

function testProcessingAfterDelay() {
  Logger.log('⏰ Simulating what should happen after 5 minutes...');
  
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
  
  Logger.log(`📊 Checking ${dataRows.length} rows for processing eligibility...`);
  
  let eligibleRows = 0;
  
  // Check each row to see if it would be processed
  dataRows.forEach((row, index) => {
    const rowIndex = index + 2;
    
    if (!row[1]) return; // Skip empty rows
    
    const hasCheckedBox = row[7] === true || row[8] === true || row[9] === true;
    
    if (hasCheckedBox) {
      // Test the shouldProcessCheckboxRow function
      const shouldProcess = shouldProcessCheckboxRow(ss, row, rowIndex, 'Sort');
      
      if (shouldProcess) {
        eligibleRows++;
        Logger.log(`✅ Row ${rowIndex} (${row[1]}) WOULD BE PROCESSED`);
        
        // Show what would happen
        if (row[7] === true) {
          Logger.log(`  → Would move to monthly sheet`);
        } else if (row[8] === true) {
          Logger.log(`  → Would move to Awaiting Employer Invoice`);
        } else if (row[9] === true) {
          Logger.log(`  → Would be deleted`);
        }
      } else {
        Logger.log(`⏳ Row ${rowIndex} (${row[1]}) NOT READY YET`);
      }
    }
  });
  
  if (eligibleRows === 0) {
    Logger.log('🚨 NO ROWS ARE ELIGIBLE FOR PROCESSING');
    Logger.log('💡 This means the tracking system isn\'t working properly');
  } else {
    Logger.log(`🎯 ${eligibleRows} rows would be processed if the trigger was working`);
  }
}

function fixDelaySystemCompletely() {
  Logger.log('🔧 COMPLETELY FIXING THE 5-MINUTE DELAY SYSTEM');
  Logger.log('================================================');
  
  // Step 1: Clear old triggers
  Logger.log('\n1️⃣ Clearing old triggers...');
  clearCheckboxTriggers();
  
  // Step 2: Reset tracking
  Logger.log('\n2️⃣ Resetting tracking system...');
  resetAllCheckboxTracking();
  
  // Step 3: Set up fresh triggers
  Logger.log('\n3️⃣ Setting up fresh triggers...');
  setupCheckboxTriggers();
  
  // Step 4: Verify setup
  Logger.log('\n4️⃣ Verifying setup...');
  checkTriggerStatus();
  
  Logger.log('\n✅ DELAY SYSTEM COMPLETELY RESET AND FIXED!');
  Logger.log('💡 Now try checking a checkbox and wait 5 minutes');
  Logger.log('💡 You can also run testFullDelaySystem() to verify it\'s working');
}

function watchCheckboxChangesRealTime() {
  Logger.log('👀 WATCHING FOR CHECKBOX CHANGES IN REAL TIME');
  Logger.log('==============================================');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sortingSheet = ss.getSheetByName('Sort');
  
  if (!sortingSheet) {
    Logger.log('❌ Sort sheet not found');
    return;
  }

  // Take a snapshot of current checkbox states
  const dataRange = sortingSheet.getDataRange();
  if (dataRange.getNumRows() <= 1) {
    Logger.log('❌ No data in Sort sheet');
    return;
  }

  const allData = dataRange.getValues();
  const dataRows = allData.slice(1);
  
  Logger.log('📸 Taking snapshot of current checkbox states...');
  
  const currentStates = [];
  
  dataRows.forEach((row, index) => {
    const rowIndex = index + 2;
    
    if (!row[1]) return; // Skip empty rows
    
    const checkboxStates = {
      rowIndex: rowIndex,
      name: row[1],
      manuallyEntered: row[7],
      employerInvoice: row[8],
      other: row[9]
    };
    
    currentStates.push(checkboxStates);
    
    const hasChecked = checkboxStates.manuallyEntered === true || 
                      checkboxStates.employerInvoice === true || 
                      checkboxStates.other === true;
    
    if (hasChecked) {
      Logger.log(`Row ${rowIndex} (${checkboxStates.name}): CHECKED - H=${checkboxStates.manuallyEntered}, I=${checkboxStates.employerInvoice}, J=${checkboxStates.other}`);
    } else {
      Logger.log(`Row ${rowIndex} (${checkboxStates.name}): unchecked`);
    }
  });
  
  Logger.log('\n💡 NOW GO CHECK SOME BOXES AND RUN watchCheckboxChangesRealTime() AGAIN TO SEE CHANGES');
  
  // Store the current states for comparison
  const trackingSheet = ss.getSheetByName('_SortCheckboxTimes');
  if (trackingSheet) {
    Logger.log('\n📊 Current tracking sheet contents:');
    const trackingData = trackingSheet.getDataRange().getValues();
    trackingData.forEach((row, index) => {
      if (index === 0) {
        Logger.log(`Headers: ${row.join(', ')}`);
      } else {
        Logger.log(`Entry ${index}: ${row.join(' | ')}`);
      }
    });
  } else {
    Logger.log('\n⚠️ No tracking sheet exists yet');
  }
}