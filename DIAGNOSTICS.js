// ===============================================
// CHECKBOX 5-MINUTE DELAY DIAGNOSTICS
// ===============================================

function diagnoseFiveMinuteDelay() {
  Logger.log('🔍 DIAGNOSING 5-MINUTE DELAY SYSTEM');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sortingSheet = ss.getSheetByName('Sort');
  
  if (!sortingSheet) {
    Logger.log('❌ Sort sheet not found');
    return;
  }

  // Check if tracking sheet exists
  const trackingSheet = ss.getSheetByName('_SortCheckboxTimes');
  
  if (!trackingSheet) {
    Logger.log('❌ PROBLEM FOUND: Tracking sheet "_SortCheckboxTimes" does not exist!');
    Logger.log('💡 This is likely why the 5-minute delay isn\'t working');
    Logger.log('💡 The system needs this hidden sheet to track when checkboxes were changed');
    return;
  }
  
  Logger.log('✅ Tracking sheet exists');
  
  // Check tracking sheet contents
  const trackingData = trackingSheet.getDataRange().getValues();
  Logger.log(`📊 Tracking sheet has ${trackingData.length - 1} entries (excluding header)`);
  
  if (trackingData.length <= 1) {
    Logger.log('⚠️ POTENTIAL PROBLEM: No tracking entries found');
    Logger.log('💡 This suggests checkbox changes aren\'t being recorded');
  }
  
  // Show current tracking entries
  Logger.log('\n📋 CURRENT TRACKING ENTRIES:');
  trackingData.forEach((row, index) => {
    if (index === 0) {
      Logger.log(`Headers: ${row.join(', ')}`);
    } else {
      const rowData = row[0];
      const changeTime = row[1];
      const rowIndex = row[2];
      const checkboxStates = row[3];
      
      if (changeTime) {
        const minutesAgo = Math.round((new Date() - new Date(changeTime)) / 60000);
        Logger.log(`Entry ${index}: Row ${rowIndex}, Changed ${minutesAgo} minutes ago`);
        Logger.log(`  Checkbox states: ${checkboxStates}`);
      } else {
        Logger.log(`Entry ${index}: Row ${rowIndex}, No change time recorded`);
      }
    }
  });
  
  // Check current Sort sheet data
  const dataRange = sortingSheet.getDataRange();
  if (dataRange.getNumRows() <= 1) {
    Logger.log('\n❌ No data in Sort sheet to check');
    return;
  }

  const allData = dataRange.getValues();
  const dataRows = allData.slice(1);
  
  Logger.log('\n📊 CURRENT SORT SHEET STATUS:');
  dataRows.forEach((row, index) => {
    const rowIndex = index + 2;
    const studentName = row[1];
    
    if (!studentName) return; // Skip empty rows
    
    const manuallyEntered = row[7];  // Column H
    const employerInvoice = row[8];  // Column I  
    const other = row[9];            // Column J
    
    const hasCheckedBox = manuallyEntered === true || employerInvoice === true || other === true;
    
    if (hasCheckedBox) {
      Logger.log(`Row ${rowIndex} (${studentName}): HAS CHECKED BOXES`);
      Logger.log(`  H=${manuallyEntered}, I=${employerInvoice}, J=${other}`);
      
      // Check if this row is being tracked
      const rowKey = JSON.stringify(row.slice(0, 7));
      const trackedEntry = trackingData.find(tracked => tracked[0] === rowKey);
      
      if (trackedEntry && trackedEntry[1]) {
        const changeTime = new Date(trackedEntry[1]);
        const minutesAgo = Math.round((new Date() - changeTime) / 60000);
        Logger.log(`  ⏰ Tracked: Changed ${minutesAgo} minutes ago`);
        
        if (minutesAgo >= 5) {
          Logger.log(`  ✅ READY FOR PROCESSING (${minutesAgo} >= 5 minutes)`);
        } else {
          Logger.log(`  ⏳ WAITING (${5 - minutesAgo} minutes remaining)`);
        }
      } else {
        Logger.log(`  ❌ NOT BEING TRACKED - This is the problem!`);
      }
    }
  });
}

function checkTriggerStatus() {
  Logger.log('🔍 CHECKING TRIGGER STATUS');
  
  const triggers = ScriptApp.getProjectTriggers();
  const checkboxTriggers = triggers.filter(trigger => 
    trigger.getHandlerFunction() === 'processCheckboxChanges'
  );
  
  Logger.log(`Found ${checkboxTriggers.length} checkbox processing triggers`);
  
  if (checkboxTriggers.length === 0) {
    Logger.log('❌ PROBLEM FOUND: No checkbox processing triggers are set up!');
    Logger.log('💡 Run setupCheckboxTriggers() to fix this');
  } else {
    checkboxTriggers.forEach((trigger, index) => {
      Logger.log(`Trigger ${index + 1}:`);
      Logger.log(`  Function: ${trigger.getHandlerFunction()}`);
      Logger.log(`  Type: Time-based trigger (every 5 minutes)`);
    });
  }
}

function manuallyTrackCheckboxChange() {
  Logger.log('🔧 MANUALLY TRACKING CHECKBOX CHANGE FOR TESTING');
  
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
  
  // Find first row with checked boxes
  for (let i = 0; i < dataRows.length; i++) {
    const row = dataRows[i];
    const rowIndex = i + 2;
    
    if (!row[1]) continue; // Skip empty rows
    
    const hasCheckedBox = row[7] === true || row[8] === true || row[9] === true;
    
    if (hasCheckedBox) {
      Logger.log(`📝 Manually tracking checkbox change for row ${rowIndex} (${row[1]})`);
      
      // Create tracking entry
      let trackingSheet = ss.getSheetByName('_SortCheckboxTimes');
      
      if (!trackingSheet) {
        Logger.log('📋 Creating tracking sheet...');
        trackingSheet = ss.insertSheet('_SortCheckboxTimes');
        trackingSheet.hideSheet();
        trackingSheet.getRange(1, 1, 1, 4).setValues([['RowData', 'CheckboxChangeTime', 'RowIndex', 'CheckboxStates']]);
      }
      
      const currentTime = new Date();
      const rowKey = JSON.stringify(row.slice(0, 7));
      const checkboxStates = JSON.stringify(row.slice(7));
      
      // Check if entry already exists
      const existingData = trackingSheet.getDataRange().getValues().slice(1);
      const existingEntry = existingData.find(tracked => tracked[0] === rowKey);
      
      if (existingEntry) {
        Logger.log('📝 Updating existing tracking entry');
        const trackingRowIndex = existingData.indexOf(existingEntry) + 2;
        trackingSheet.getRange(trackingRowIndex, 2, 1, 3).setValues([[currentTime, rowIndex, checkboxStates]]);
      } else {
        Logger.log('📝 Creating new tracking entry');
        const lastRow = trackingSheet.getLastRow();
        trackingSheet.getRange(lastRow + 1, 1, 1, 4).setValues([[rowKey, currentTime, rowIndex, checkboxStates]]);
      }
      
      Logger.log(`✅ Tracked checkbox change at ${currentTime}`);
      Logger.log(`💡 This row should be processed in 5 minutes (at ${new Date(currentTime.getTime() + 5 * 60 * 1000)})`);
      
      return;
    }
  }
  
  Logger.log('❌ No rows with checked boxes found');
}

function forceProcessAfterDelay() {
  Logger.log('🔧 FORCING PROCESSING OF ROWS READY AFTER 5-MINUTE DELAY');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // This calls the regular processing function that respects the 5-minute delay
  processCheckboxChanges();
  
  Logger.log('✅ Processing complete - check logs above for results');
}

function resetAllCheckboxTracking() {
  Logger.log('🧹 RESETTING ALL CHECKBOX TRACKING');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trackingSheet = ss.getSheetByName('_SortCheckboxTimes');
  
  if (trackingSheet) {
    ss.deleteSheet(trackingSheet);
    Logger.log('✅ Deleted existing tracking sheet');
  } else {
    Logger.log('ℹ️ No tracking sheet found to delete');
  }
  
  Logger.log('💡 Next time checkboxes are changed, tracking will start fresh');
}