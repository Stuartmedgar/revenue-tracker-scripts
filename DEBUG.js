// ===============================================
// DEBUG.GS - Complete Debugging Functions for Student Engagement
// ===============================================

function checkCurrentTriggers() {
  Logger.log('=== CHECKING CURRENT TRIGGERS ===');
  
  const triggers = ScriptApp.getProjectTriggers();
  
  Logger.log(`Found ${triggers.length} active triggers:`);
  
  triggers.forEach((trigger, index) => {
    const handlerFunction = trigger.getHandlerFunction();
    const triggerType = trigger.getTriggerSource();
    const eventType = trigger.getEventType();
    
    Logger.log(`${index + 1}. Function: ${handlerFunction}`);
    Logger.log(`   Type: ${triggerType}`);
    Logger.log(`   Event: ${eventType}`);
    
    if (triggerType === ScriptApp.TriggerSource.CLOCK) {
      try {
        Logger.log(`   Frequency: Time-based trigger`);
      } catch (error) {
        Logger.log(`   Frequency: Unknown`);
      }
    }
    Logger.log('');
  });
  
  // Check for the specific triggers we need
  const dataProcessingTrigger = triggers.find(t => t.getHandlerFunction() === 'processDataEntries');
  const checkboxTrigger = triggers.find(t => t.getHandlerFunction() === 'processCheckboxChanges');
  const emailTrigger = triggers.find(t => t.getHandlerFunction() === 'sendWeeklyMonitoringEmail');
  
  Logger.log('📊 TRIGGER STATUS:');
  Logger.log(`✅ Data Processing (every 5 min): ${dataProcessingTrigger ? 'ACTIVE' : '❌ MISSING'}`);
  Logger.log(`✅ Checkbox Processing (every 5 min): ${checkboxTrigger ? 'ACTIVE' : '❌ MISSING'}`);
  Logger.log(`✅ Email Monitoring (weekly): ${emailTrigger ? 'ACTIVE' : '❌ MISSING'}`);
  
  if (!dataProcessingTrigger || !checkboxTrigger) {
    Logger.log('\n⚠️  MISSING TRIGGERS DETECTED!');
    Logger.log('Run initializeRevenueTracker() to set up missing triggers');
  }
}

function checkRecentActivity() {
  Logger.log('=== CHECKING RECENT ACTIVITY ===');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const now = new Date();
  
  Logger.log(`Current time: ${now}`);
  
  // Check Data Entry sheet for recent additions
  const dataEntrySheet = ss.getSheetByName('Data Entry');
  if (dataEntrySheet) {
    const dataRange = dataEntrySheet.getDataRange();
    if (dataRange.getNumRows() > 1) {
      Logger.log(`📝 Data Entry sheet has ${dataRange.getNumRows() - 1} rows waiting for processing`);
      
      // Show first few entries
      const allData = dataRange.getValues();
      const dataRows = allData.slice(1, 4); // Show first 3 rows
      dataRows.forEach((row, index) => {
        if (row[1]) { // If has a name
          Logger.log(`   ${index + 1}. ${row[1]} - £${row[5]} - ${row[3]}`);
        }
      });
    } else {
      Logger.log('📝 Data Entry sheet is empty');
    }
  } else {
    Logger.log('❌ Data Entry sheet not found');
  }
  
  // Check Sort sheet for checked boxes
  const sortSheet = ss.getSheetByName('Sort');
  if (sortSheet) {
    const dataRange = sortSheet.getDataRange();
    if (dataRange.getNumRows() > 1) {
      const allData = dataRange.getValues();
      const dataRows = allData.slice(1);
      
      let checkedBoxes = 0;
      let totalStudents = 0;
      
      dataRows.forEach((row, index) => {
        if (row[1]) { // If has a name
          totalStudents++;
          const manuallyEntered = row[7]; // Column H
          const employerInvoice = row[8]; // Column I  
          const other = row[9]; // Column J
          
          if (manuallyEntered === true || employerInvoice === true || other === true) {
            checkedBoxes++;
            Logger.log(`   ☑️ ${row[1]} has checked box`);
          }
        }
      });
      
      Logger.log(`📦 Sort sheet has ${totalStudents} students, ${checkedBoxes} with checked boxes`);
    } else {
      Logger.log('📦 Sort sheet is empty');
    }
  } else {
    Logger.log('❌ Sort sheet not found');
  }
  
  // Check Awaiting Invoice sheet
  const awaitingSheet = ss.getSheetByName('Awaiting Employer Invoice');
  if (awaitingSheet) {
    const dataRange = awaitingSheet.getDataRange();
    if (dataRange.getNumRows() > 1) {
      const allData = dataRange.getValues();
      const dataRows = allData.slice(1);
      
      let paidBoxes = 0;
      let totalInvoices = 0;
      
      dataRows.forEach((row, index) => {
        if (row[1]) { // If has a name
          totalInvoices++;
          const invoicePaid = row[7]; // Column H
          if (invoicePaid === true) {
            paidBoxes++;
            Logger.log(`   💰 ${row[1]} marked as paid`);
          }
        }
      });
      
      Logger.log(`💰 Awaiting Invoice sheet has ${totalInvoices} invoices, ${paidBoxes} marked as paid`);
    } else {
      Logger.log('💰 Awaiting Invoice sheet is empty');
    }
  } else {
    Logger.log('❌ Awaiting Invoice sheet not found');
  }
}

function auditMissingEngagementTransfers() {
  Logger.log('=== AUDITING STUDENTS IN MONTHLY SHEETS ===');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();
  
  // Get all monthly sheets
  const monthlySheets = allSheets.filter(sheet => isMonthlySheetName(sheet.getName()));
  
  Logger.log(`📊 Found ${monthlySheets.length} monthly sheets to check`);
  
  let totalEligible = 0;
  
  monthlySheets.forEach(sheet => {
    const sheetName = sheet.getName();
    Logger.log(`\n📋 Checking ${sheetName}...`);
    
    const dataRange = sheet.getDataRange();
    if (dataRange.getNumRows() <= 1) {
      Logger.log('  No data in this sheet');
      return;
    }
    
    const allData = dataRange.getValues();
    const headers = allData[0];
    const dataRows = allData.slice(1);
    
    const nameCol = headers.indexOf('Name');
    const sittingCol = headers.indexOf('Sitting');
    const actualPriceCol = headers.indexOf('Actual Price');
    
    if (nameCol === -1 || sittingCol === -1 || actualPriceCol === -1) {
      Logger.log('  ❌ Missing required columns');
      Logger.log(`  Headers found: ${headers.join(', ')}`);
      return;
    }
    
    let sheetEligible = 0;
    
    dataRows.forEach((row, index) => {
      const name = row[nameCol];
      const sitting = row[sittingCol];
      const actualPrice = row[actualPriceCol];
      
      if (name && sitting && actualPrice) {
        const isEligible = meetsEngagementCriteria(actualPrice);
        if (isEligible) {
          sheetEligible++;
          totalEligible++;
          Logger.log(`  ✅ ${name} (£${actualPrice}) - ${sitting} - SHOULD BE TRANSFERRED`);
        } else {
          Logger.log(`  ❌ ${name} (£${actualPrice}) - Does not meet criteria`);
        }
      }
    });
    
    Logger.log(`  📊 ${sheetEligible} eligible students in this sheet`);
  });
  
  Logger.log(`\n🎯 SUMMARY: Found ${totalEligible} students who SHOULD be in engagement spreadsheet`);
  
  if (totalEligible > 0) {
    Logger.log('\n🔍 Next step: Run debugEngagementTransfer() to check if we can access the engagement spreadsheet');
  } else {
    Logger.log('\n❓ No eligible students found. Check if students are reaching monthly sheets or meeting price criteria.');
  }
}

function debugEngagementTransfer() {
  Logger.log('=== DEBUGGING ENGAGEMENT SPREADSHEET ACCESS ===');
  
  try {
    // Test spreadsheet access
    const engagementSS = SpreadsheetApp.openById('1mhtp-bFZe-mFMKStBXE9ECgGj9T8K-xAZOJ511jpOH8');
    Logger.log('✅ Successfully accessed ATX Engagement spreadsheet');
    
    // List all available sheets
    const sheets = engagementSS.getSheets();
    Logger.log(`📊 Found ${sheets.length} sheets in ATX Engagement:`);
    sheets.forEach(sheet => {
      Logger.log(`  - "${sheet.getName()}"`);
    });
    
    // Test sitting normalization with common examples
    Logger.log('\n🧪 Testing sitting name normalization:');
    const testSittings = ['December 2025', 'Dec 2025', 'June 2025', 'Jun 2025', 'March 2026'];
    
    testSittings.forEach(sitting => {
      const normalized = normalizeSittingName(sitting);
      const targetSheet = engagementSS.getSheetByName(normalized);
      const exists = targetSheet ? '✅ EXISTS' : '❌ NOT FOUND';
      Logger.log(`  "${sitting}" → "${normalized}" ${exists}`);
    });
    
    // Test with actual data from monthly sheets
    Logger.log('\n🔍 Checking actual sitting names from monthly sheets:');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const allSheets = ss.getSheets();
    const monthlySheets = allSheets.filter(sheet => isMonthlySheetName(sheet.getName()));
    
    const uniqueSittings = new Set();
    
    monthlySheets.forEach(sheet => {
      const dataRange = sheet.getDataRange();
      if (dataRange.getNumRows() <= 1) return;
      
      const allData = dataRange.getValues();
      const headers = allData[0];
      const dataRows = allData.slice(1);
      
      const sittingCol = headers.indexOf('Sitting');
      if (sittingCol === -1) return;
      
      dataRows.forEach(row => {
        if (row[sittingCol]) {
          uniqueSittings.add(row[sittingCol]);
        }
      });
    });
    
    Logger.log(`Found ${uniqueSittings.size} unique sitting names in monthly sheets:`);
    Array.from(uniqueSittings).forEach(sitting => {
      const normalized = normalizeSittingName(sitting);
      const targetSheet = engagementSS.getSheetByName(normalized);
      const exists = targetSheet ? '✅ TARGET EXISTS' : '❌ NO TARGET SHEET';
      Logger.log(`  "${sitting}" → "${normalized}" ${exists}`);
    });
    
  } catch (error) {
    Logger.log(`❌ ERROR accessing ATX Engagement spreadsheet: ${error.toString()}`);
    Logger.log('Possible issues:');
    Logger.log('1. Spreadsheet ID might be wrong');
    Logger.log('2. Spreadsheet might not be shared with this script');
    Logger.log('3. Spreadsheet might have been deleted or moved');
  }
}

function testSingleEngagementTransfer() {
  Logger.log('=== TESTING SINGLE ENGAGEMENT TRANSFER ===');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Find a monthly sheet with data
  const allSheets = ss.getSheets();
  const monthlySheets = allSheets.filter(sheet => isMonthlySheetName(sheet.getName()));
  
  if (monthlySheets.length === 0) {
    Logger.log('❌ No monthly sheets found to test with');
    return;
  }
  
  // Find a sheet with eligible students
  let testStudent = null;
  let sourceSheet = null;
  
  for (const sheet of monthlySheets) {
    const dataRange = sheet.getDataRange();
    if (dataRange.getNumRows() <= 1) continue;
    
    const allData = dataRange.getValues();
    const headers = allData[0];
    const dataRows = allData.slice(1);
    
    const nameCol = headers.indexOf('Name');
    const sittingCol = headers.indexOf('Sitting');
    const actualPriceCol = headers.indexOf('Actual Price');
    
    if (nameCol === -1 || sittingCol === -1 || actualPriceCol === -1) continue;
    
    // Find first eligible student
    for (const row of dataRows) {
      const name = row[nameCol];
      const sitting = row[sittingCol];
      const actualPrice = row[actualPriceCol];
      
      if (name && sitting && actualPrice && meetsEngagementCriteria(actualPrice)) {
        testStudent = {
          name: name,
          sitting: sitting,
          actualPrice: actualPrice,
          course: getCourseTypeFromPrice(actualPrice)
        };
        sourceSheet = sheet.getName();
        break;
      }
    }
    if (testStudent) break;
  }
  
  if (!testStudent) {
    Logger.log('❌ No eligible students found in monthly sheets for testing');
    return;
  }
  
  Logger.log(`🧪 Testing transfer for: ${testStudent.name}`);
  Logger.log(`   Sitting: ${testStudent.sitting}`);
  Logger.log(`   Price: £${testStudent.actualPrice}`);
  Logger.log(`   Course: ${testStudent.course}`);
  Logger.log(`   Source sheet: ${sourceSheet}`);
  
  // Attempt the transfer
  Logger.log('\n🚀 Attempting transfer...');
  try {
    processStudentForEngagement(testStudent, sourceSheet);
    Logger.log('✅ Transfer function completed - check logs above for any issues');
  } catch (error) {
    Logger.log(`❌ Transfer failed: ${error.toString()}`);
  }
}

function quickSystemCheck() {
  Logger.log('=== QUICK SYSTEM HEALTH CHECK ===');
  
  // Check 1: Triggers
  const triggers = ScriptApp.getProjectTriggers();
  const dataProcessingTrigger = triggers.find(t => t.getHandlerFunction() === 'processDataEntries');
  const checkboxTrigger = triggers.find(t => t.getHandlerFunction() === 'processCheckboxChanges');
  
  Logger.log(`🔧 Triggers: ${dataProcessingTrigger && checkboxTrigger ? '✅ WORKING' : '❌ BROKEN'}`);
  
  // Check 2: Students waiting to be processed
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataEntrySheet = ss.getSheetByName('Data Entry');
  const waitingInDataEntry = dataEntrySheet && dataEntrySheet.getDataRange().getNumRows() > 1;
  
  Logger.log(`📝 Students in Data Entry: ${waitingInDataEntry ? '⚠️ YES - waiting to process' : '✅ None waiting'}`);
  
  // Check 3: Students in monthly sheets
  const allSheets = ss.getSheets();
  const monthlySheets = allSheets.filter(sheet => isMonthlySheetName(sheet.getName()));
  let totalInMonthly = 0;
  let eligibleForEngagement = 0;
  
  monthlySheets.forEach(sheet => {
    const dataRange = sheet.getDataRange();
    if (dataRange.getNumRows() <= 1) return;
    
    const allData = dataRange.getValues();
    const headers = allData[0];
    const dataRows = allData.slice(1);
    totalInMonthly += dataRows.length;
    
    const actualPriceCol = headers.indexOf('Actual Price');
    if (actualPriceCol !== -1) {
      dataRows.forEach(row => {
        if (row[actualPriceCol] && meetsEngagementCriteria(row[actualPriceCol])) {
          eligibleForEngagement++;
        }
      });
    }
  });
  
  Logger.log(`📊 Students in monthly sheets: ${totalInMonthly}`);
  Logger.log(`🎓 Eligible for engagement: ${eligibleForEngagement}`);
  
  // Check 4: Engagement spreadsheet access
  try {
    const engagementSS = SpreadsheetApp.openById('1mhtp-bFZe-mFMKStBXE9ECgGj9T8K-xAZOJ511jpOH8');
    Logger.log('🎯 Engagement spreadsheet: ✅ ACCESSIBLE');
  } catch (error) {
    Logger.log('🎯 Engagement spreadsheet: ❌ NOT ACCESSIBLE');
  }
  
  Logger.log('\n📋 SUMMARY:');
  if (!dataProcessingTrigger || !checkboxTrigger) {
    Logger.log('❌ MAIN ISSUE: Triggers not set up. Run initializeRevenueTracker()');
  } else if (eligibleForEngagement > 0) {
    Logger.log('❌ MAIN ISSUE: Students eligible but not transferring. Check engagement spreadsheet access.');
  } else if (totalInMonthly === 0) {
    Logger.log('❌ MAIN ISSUE: No students reaching monthly sheets. Check data processing.');
  } else {
    Logger.log('✅ System appears to be working. Students may not meet engagement criteria.');
  }
}
// ===============================================
// AWAITING INVOICE DIAGNOSTIC FUNCTIONS
// ===============================================

function diagnoseAwaitingInvoiceIssues() {
  Logger.log('🔍 DIAGNOSING AWAITING EMPLOYER INVOICE ISSUES');
  Logger.log('==============================================');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const awaitingSheet = ss.getSheetByName('Awaiting Employer Invoice');
  
  if (!awaitingSheet) {
    Logger.log('❌ Awaiting Employer Invoice sheet not found');
    return;
  }
  
  // Step 1: Check the data and checkbox values
  Logger.log('\n1️⃣ CHECKING DATA AND CHECKBOXES');
  Logger.log('===============================');
  
  const dataRange = awaitingSheet.getDataRange();
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
      Logger.log(`  Full Price: ${row[4]}`);
      Logger.log(`  Actual Price: ${row[5]}`);
      Logger.log(`  Invoice Paid (H): ${row[7]} (type: ${typeof row[7]})`);
      
      // Check if checkbox is actually a checkbox
      const cell = awaitingSheet.getRange(rowIndex, 8); // Column H
      const validation = cell.getDataValidation();
      Logger.log(`  Cell H validation: ${validation ? 'Has checkbox validation' : 'NO validation - this is the problem!'}`);
      
      if (row[7] === true) {
        Logger.log(`  ✅ This row should be processed!`);
        
        // Check how long it's been checked
        const checkTime = findCheckboxChangeTime(ss, row, 'AwaitingInvoice');
        if (checkTime) {
          const minutesAgo = Math.round((new Date() - checkTime) / 60000);
          Logger.log(`  ⏰ Checked ${minutesAgo} minutes ago`);
          if (minutesAgo >= 5) {
            Logger.log(`  ✅ Time requirement met (${minutesAgo} >= 5 minutes)`);
          } else {
            Logger.log(`  ⏳ Still waiting (${5 - minutesAgo} minutes remaining)`);
          }
        } else {
          Logger.log(`  ❌ No tracking found - this is likely the problem`);
        }
      }
    }
  });
  
  // Step 2: Check triggers
  Logger.log('\n2️⃣ CHECKING TRIGGERS');
  Logger.log('==================');
  
  const triggers = ScriptApp.getProjectTriggers();
  const relevantTriggers = triggers.filter(t => 
    t.getHandlerFunction() === 'processCheckboxChanges' ||
    t.getHandlerFunction() === 'processCheckboxChangesImproved'
  );
  
  Logger.log(`Found ${relevantTriggers.length} checkbox processing triggers`);
  
  if (relevantTriggers.length === 0) {
    Logger.log('❌ NO CHECKBOX TRIGGERS FOUND - This is the problem!');
    Logger.log('💡 Run setupCheckboxTriggers() or fixSortSheetCompletely() to fix');
  } else {
    relevantTriggers.forEach((trigger, index) => {
      Logger.log(`Trigger ${index + 1}: ${trigger.getHandlerFunction()}`);
    });
  }
  
  // Step 3: Check tracking sheet
  Logger.log('\n3️⃣ CHECKING TRACKING SHEET');
  Logger.log('=========================');
  
  const trackingSheet = ss.getSheetByName('_AwaitingInvoiceCheckboxTimes');
  
  if (trackingSheet) {
    const trackingData = trackingSheet.getDataRange().getValues();
    Logger.log(`Tracking entries: ${trackingData.length - 1}`);
    
    trackingData.forEach((row, index) => {
      if (index === 0) {
        Logger.log(`Headers: ${row.join(', ')}`);
      } else {
        const studentName = row[3] || 'Unknown';
        const changeTime = row[1];
        if (changeTime) {
          const minutesAgo = Math.round((new Date() - new Date(changeTime)) / 60000);
          Logger.log(`Entry ${index}: ${studentName}, ${minutesAgo} minutes ago`);
        } else {
          Logger.log(`Entry ${index}: ${studentName}, No time recorded`);
        }
      }
    });
  } else {
    Logger.log('❌ No tracking sheet found - this could be the problem');
    Logger.log('💡 The system needs _AwaitingInvoiceCheckboxTimes sheet to track timing');
  }
}

function findCheckboxChangeTime(ss, row, sheetName) {
  const trackingSheet = ss.getSheetByName(`_${sheetName}CheckboxTimes`);
  if (!trackingSheet) return null;
  
  const rowKey = JSON.stringify(row.slice(0, 7));
  const trackingData = trackingSheet.getDataRange().getValues().slice(1);
  
  const trackedEntry = trackingData.find(tracked => tracked[0] === rowKey);
  
  return trackedEntry && trackedEntry[1] ? new Date(trackedEntry[1]) : null;
}

function forceProcessAwaitingInvoiceNow() {
  Logger.log('🚀 FORCE PROCESSING AWAITING INVOICE (NO DELAY)');
  Logger.log('===============================================');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const awaitingSheet = ss.getSheetByName('Awaiting Employer Invoice');
  
  if (!awaitingSheet) {
    Logger.log('❌ Awaiting Employer Invoice sheet not found');
    return;
  }
  
  const dataRange = awaitingSheet.getDataRange();
  const allData = dataRange.getValues();
  const dataRows = allData.slice(1);
  
  Logger.log(`📊 Checking ${dataRows.length} rows for checked "Invoice Paid" boxes`);
  
  let processedCount = 0;
  
  // Process in reverse order to avoid index issues when deleting
  for (let i = dataRows.length - 1; i >= 0; i--) {
    const row = dataRows[i];
    const rowIndex = i + 2;
    
    // Skip empty rows
    if (!row[1]) {
      Logger.log(`⏸️ Row ${rowIndex}: Empty, skipping`);
      continue;
    }
    
    const invoicePaid = row[7]; // Column H
    
    Logger.log(`\n🔍 Row ${rowIndex} (${row[1]}):`);
    Logger.log(`  Invoice Paid: ${invoicePaid} (type: ${typeof invoicePaid})`);
    
    if (invoicePaid === true) {
      Logger.log(`  ✅ Processing: Invoice Paid`);
      
      try {
        // Update actual price to equal full price
        let originalData = row.slice(0, 7);
        const fullPrice = originalData[4];
        originalData[5] = fullPrice; // Set actual price = full price
        
        Logger.log(`  💰 Updated Actual Price from ${row[5]} to ${fullPrice}`);
        
        // Move to monthly sheet
        moveToMonthlySheetFromInvoice(ss, originalData);
        
        // Delete from awaiting sheet
        awaitingSheet.deleteRow(rowIndex);
        
        processedCount++;
        Logger.log(`  ✅ Success: Moved to monthly sheet with full payment`);
        
      } catch (error) {
        Logger.log(`  ❌ Error: ${error.toString()}`);
      }
    } else {
      Logger.log(`  ⏸️ Invoice not marked as paid, skipping`);
    }
  }
  
  Logger.log(`\n🎯 SUMMARY:`);
  Logger.log(`✅ Processed: ${processedCount} rows`);
  
  if (processedCount === 0) {
    Logger.log(`🤔 No rows were processed. Possible reasons:`);
    Logger.log(`  1. No checkboxes are actually checked (they might be text/boolean values)`);
    Logger.log(`  2. The checkboxes aren't properly formatted`);
    Logger.log(`  3. There are no records in the Awaiting Invoice sheet`);
  }
}

function fixAwaitingInvoiceCheckboxes() {
  Logger.log('🔧 FIXING AWAITING INVOICE CHECKBOXES');
  Logger.log('====================================');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const awaitingSheet = ss.getSheetByName('Awaiting Employer Invoice');
  
  if (!awaitingSheet) {
    Logger.log('❌ Awaiting Employer Invoice sheet not found');
    return;
  }
  
  const dataRange = awaitingSheet.getDataRange();
  const numRows = dataRange.getNumRows();
  
  if (numRows <= 1) {
    Logger.log('No data rows to fix');
    return;
  }
  
  Logger.log(`Fixing checkboxes for ${numRows - 1} data rows`);
  
  // Fix checkboxes in column H for all data rows
  for (let row = 2; row <= numRows; row++) {
    const nameCell = awaitingSheet.getRange(row, 2);
    if (!nameCell.getValue()) continue; // Skip empty rows
    
    const studentName = nameCell.getValue();
    Logger.log(`Fixing checkbox for row ${row} (${studentName})`);
    
    const cell = awaitingSheet.getRange(row, 8); // Column H
    const currentValue = cell.getValue();
    
    // Clear the cell and add proper checkbox
    cell.clearContent();
    cell.clearDataValidations();
    
    try {
      cell.insertCheckboxes();
      
      // If it was previously checked, set it back to checked
      if (currentValue === true || currentValue === 'TRUE' || currentValue === '✓') {
        cell.setValue(true);
        Logger.log(`  ✅ Fixed checkbox and restored checked state`);
      } else {
        Logger.log(`  ✅ Fixed checkbox (unchecked)`);
      }
    } catch (error) {
      // Fallback to data validation method
      const rule = SpreadsheetApp.newDataValidation()
        .requireCheckbox()
        .build();
      cell.setDataValidation(rule);
      
      if (currentValue === true || currentValue === 'TRUE' || currentValue === '✓') {
        cell.setValue(true);
        Logger.log(`  ⚠️ Used validation fallback and restored checked state`);
      } else {
        cell.setValue(false);
        Logger.log(`  ⚠️ Used validation fallback (unchecked)`);
      }
    }
  }
  
  Logger.log('✅ All checkboxes fixed in Awaiting Employer Invoice sheet');
}

function setupAwaitingInvoiceTriggers() {
  Logger.log('🔧 SETTING UP AWAITING INVOICE TRIGGERS');
  Logger.log('======================================');
  
  // This is usually part of the main checkbox system, but let's make sure
  setupCheckboxTriggers();
  
  Logger.log('✅ Checkbox triggers set up');
  Logger.log('💡 The system will now check every 5 minutes for checkbox changes');
}

function testAwaitingInvoiceSystem() {
  Logger.log('🧪 TESTING AWAITING INVOICE SYSTEM');
  Logger.log('==================================');
  
  // Step 1: Diagnose current state
  diagnoseAwaitingInvoiceIssues();
  
  // Step 2: Check if we can force process
  Logger.log('\n🚀 ATTEMPTING FORCE PROCESS...');
  forceProcessAwaitingInvoiceNow();
  
  Logger.log('\n✅ Test complete - check the logs above for issues');
}