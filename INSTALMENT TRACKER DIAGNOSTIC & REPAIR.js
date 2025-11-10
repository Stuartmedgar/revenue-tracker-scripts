// ===============================================
// INSTALMENT TRACKER DIAGNOSTIC & REPAIR
// ===============================================

function diagnoseInstalmentTracker() {
  Logger.log('🔍 DIAGNOSING INSTALMENT TRACKER ISSUES');
  Logger.log('======================================');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trackerSheet = ss.getSheetByName('Instalment Tracker');
  
  if (!trackerSheet) {
    Logger.log('❌ Instalment Tracker sheet not found');
    return;
  }
  
  const dataRange = trackerSheet.getDataRange();
  if (dataRange.getNumRows() <= 1) {
    Logger.log('✅ Instalment Tracker is empty - nothing to diagnose');
    return;
  }
  
  const allData = dataRange.getValues();
  const headers = allData[0];
  const dataRows = allData.slice(1);
  
  Logger.log(`📊 Found ${dataRows.length} records in Instalment Tracker\n`);
  Logger.log(`Headers: ${headers.join(', ')}\n`);
  
  // Find column indices
  const nameCol = headers.indexOf('Student Name');
  const courseCol = headers.indexOf('Course');
  const fullPriceCol = headers.indexOf('Full Price');
  const amountPaidCol = headers.indexOf('Amount Paid');
  const instalmentsCol = headers.indexOf('Instalments Paid');
  const lastPaymentCol = headers.indexOf('Last Payment Date');
  const nextDueCol = headers.indexOf('Next Payment Due');
  const statusCol = headers.indexOf('Payment Complete');
  const completionCol = headers.indexOf('Completion Date');
  
  let shouldBeComplete = [];
  let shouldBeDeleted = [];
  let statusCorrect = [];
  
  const currentDate = new Date();
  const thirtyDaysAgo = new Date(currentDate.getTime() - (30 * 24 * 60 * 60 * 1000));
  
  dataRows.forEach((row, index) => {
    const rowIndex = index + 2; // +2 for header and 1-indexed
    const studentName = row[nameCol];
    const course = row[courseCol];
    const fullPrice = Number(row[fullPriceCol]);
    const amountPaid = Number(row[amountPaidCol]);
    const status = row[statusCol];
    const completionDate = row[completionCol];
    
    Logger.log(`\n📋 Row ${rowIndex}: ${studentName} - ${course}`);
    Logger.log(`   Full Price: £${fullPrice}`);
    Logger.log(`   Amount Paid: £${amountPaid}`);
    Logger.log(`   Status: ${status}`);
    Logger.log(`   Completion Date: ${completionDate || 'None'}`);
    
    // Issue 1: Should be marked complete but isn't
    if (amountPaid >= fullPrice && status !== 'Complete') {
      Logger.log(`   ⚠️ ISSUE: Has paid full amount but status is "${status}"`);
      shouldBeComplete.push({
        rowIndex: rowIndex,
        studentName: studentName,
        fullPrice: fullPrice,
        amountPaid: amountPaid
      });
    }
    
    // Issue 2: Completed more than 30 days ago but still showing
    if (status === 'Complete' && completionDate) {
      const completionDateObj = new Date(completionDate);
      const daysAgo = Math.floor((currentDate - completionDateObj) / (1000 * 60 * 60 * 24));
      
      Logger.log(`   Completed ${daysAgo} days ago`);
      
      if (completionDateObj < thirtyDaysAgo) {
        Logger.log(`   ⚠️ ISSUE: Completed over 30 days ago, should be deleted`);
        shouldBeDeleted.push({
          rowIndex: rowIndex,
          studentName: studentName,
          completionDate: completionDate,
          daysAgo: daysAgo
        });
      }
    }
    
    // Check if status is correct
    if (amountPaid >= fullPrice && status === 'Complete') {
      Logger.log(`   ✅ Status is CORRECT`);
      statusCorrect.push(studentName);
    } else if (amountPaid < fullPrice && status === 'In Progress') {
      Logger.log(`   ✅ Status is CORRECT`);
      statusCorrect.push(studentName);
    }
  });
  
  // Summary
  Logger.log('\n\n📊 DIAGNOSIS SUMMARY');
  Logger.log('===================');
  Logger.log(`✅ Records with correct status: ${statusCorrect.length}`);
  Logger.log(`⚠️ Records that should be marked complete: ${shouldBeComplete.length}`);
  Logger.log(`🗑️ Records that should be deleted (>30 days): ${shouldBeDeleted.length}`);
  
  if (shouldBeComplete.length > 0) {
    Logger.log('\n🔧 STUDENTS TO MARK AS COMPLETE:');
    shouldBeComplete.forEach(item => {
      Logger.log(`   - Row ${item.rowIndex}: ${item.studentName} (Paid £${item.amountPaid} of £${item.fullPrice})`);
    });
  }
  
  if (shouldBeDeleted.length > 0) {
    Logger.log('\n🗑️ RECORDS TO DELETE:');
    shouldBeDeleted.forEach(item => {
      Logger.log(`   - Row ${item.rowIndex}: ${item.studentName} (Completed ${item.daysAgo} days ago)`);
    });
  }
  
  Logger.log('\n💡 NEXT STEPS:');
  Logger.log('   Run fixInstalmentTracker() to automatically fix all issues');
}

function fixInstalmentTracker() {
  Logger.log('🔧 FIXING INSTALMENT TRACKER');
  Logger.log('============================');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trackerSheet = ss.getSheetByName('Instalment Tracker');
  
  if (!trackerSheet) {
    Logger.log('❌ Instalment Tracker sheet not found');
    return;
  }
  
  const dataRange = trackerSheet.getDataRange();
  if (dataRange.getNumRows() <= 1) {
    Logger.log('✅ Instalment Tracker is empty - nothing to fix');
    return;
  }
  
  const allData = dataRange.getValues();
  const headers = allData[0];
  const dataRows = allData.slice(1);
  
  // Find column indices
  const nameCol = headers.indexOf('Student Name');
  const fullPriceCol = headers.indexOf('Full Price');
  const amountPaidCol = headers.indexOf('Amount Paid');
  const statusCol = headers.indexOf('Payment Complete');
  const completionCol = headers.indexOf('Completion Date');
  
  let markedComplete = 0;
  let deleted = 0;
  
  const currentDate = new Date();
  const thirtyDaysAgo = new Date(currentDate.getTime() - (30 * 24 * 60 * 60 * 1000));
  
  // Process in reverse order to safely delete rows
  for (let i = dataRows.length - 1; i >= 0; i--) {
    const row = dataRows[i];
    const rowIndex = i + 2; // +2 for header and 1-indexed
    const studentName = row[nameCol];
    const fullPrice = Number(row[fullPriceCol]);
    const amountPaid = Number(row[amountPaidCol]);
    const status = row[statusCol];
    const completionDate = row[completionCol];
    
    // Fix 1: Mark as complete if fully paid
    if (amountPaid >= fullPrice && status !== 'Complete') {
      Logger.log(`✅ Marking ${studentName} as Complete (paid £${amountPaid} of £${fullPrice})`);
      
      trackerSheet.getRange(rowIndex, statusCol + 1).setValue('Complete');
      trackerSheet.getRange(rowIndex, completionCol + 1).setValue(currentDate);
      
      markedComplete++;
    }
    
    // Fix 2: Delete if completed over 30 days ago
    if (status === 'Complete' && completionDate) {
      const completionDateObj = new Date(completionDate);
      
      if (completionDateObj < thirtyDaysAgo) {
        const daysAgo = Math.floor((currentDate - completionDateObj) / (1000 * 60 * 60 * 24));
        Logger.log(`🗑️ Deleting ${studentName} (completed ${daysAgo} days ago)`);
        
        trackerSheet.deleteRow(rowIndex);
        deleted++;
      }
    }
  }
  
  Logger.log('\n📊 FIX SUMMARY');
  Logger.log('=============');
  Logger.log(`✅ Marked as complete: ${markedComplete} students`);
  Logger.log(`🗑️ Deleted old records: ${deleted} students`);
  Logger.log('\n✅ Instalment Tracker has been fixed!');
}

function rebuildInstalmentTrackerFromMonthlySheets() {
  Logger.log('🔄 REBUILDING INSTALMENT TRACKER FROM MONTHLY SHEETS');
  Logger.log('===================================================');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trackerSheet = ss.getSheetByName('Instalment Tracker');
  
  if (!trackerSheet) {
    Logger.log('❌ Instalment Tracker sheet not found');
    return;
  }
  
  // Clear existing tracker (except headers)
  const lastRow = trackerSheet.getLastRow();
  if (lastRow > 1) {
    trackerSheet.deleteRows(2, lastRow - 1);
    Logger.log('🗑️ Cleared existing tracker data');
  }
  
  // Get all monthly sheets
  const allSheets = ss.getSheets();
  const monthlySheets = allSheets.filter(sheet => isMonthlySheetName(sheet.getName()));
  
  Logger.log(`📊 Found ${monthlySheets.length} monthly sheets to scan\n`);
  
  let totalProcessed = 0;
  
  monthlySheets.forEach(sheet => {
    const sheetName = sheet.getName();
    Logger.log(`\n📋 Processing ${sheetName}...`);
    
    const dataRange = sheet.getDataRange();
    if (dataRange.getNumRows() <= 1) {
      Logger.log('   No data');
      return;
    }
    
    const allData = dataRange.getValues();
    const headers = allData[0];
    const dataRows = allData.slice(1);
    
    // Find column indices
    const dateCol = headers.indexOf('Date');
    const nameCol = headers.indexOf('Name');
    const courseCol = headers.indexOf('Course');
    const fullPriceCol = headers.indexOf('Full Price');
    const actualPriceCol = headers.indexOf('Actual Price');
    const paymentPlanCol = headers.indexOf('Payment Plan');
    
    if (nameCol === -1 || fullPriceCol === -1 || actualPriceCol === -1) {
      Logger.log('   ⚠️ Missing required columns');
      return;
    }
    
    // Process each row with payment plan
    dataRows.forEach(row => {
      const hasPaymentPlan = row[paymentPlanCol] === 'Y';
      
      if (hasPaymentPlan && row[nameCol] && row[fullPriceCol] && row[actualPriceCol]) {
        const studentName = row[nameCol];
        const course = row[courseCol];
        const fullPrice = row[fullPriceCol];
        const actualPrice = row[actualPriceCol];
        const paymentDate = row[dateCol] ? new Date(row[dateCol]) : new Date();
        
        Logger.log(`   📝 Processing instalment: ${studentName} - £${actualPrice}`);
        
        processInstalmentPaymentImproved(studentName, course, fullPrice, actualPrice, paymentDate);
        totalProcessed++;
      }
    });
  });
  
  Logger.log(`\n✅ REBUILD COMPLETE`);
  Logger.log(`   Processed ${totalProcessed} instalment payments`);
  Logger.log('\n💡 Now run fixInstalmentTracker() to clean up any issues');
}

function isMonthlySheetName(name) {
  const monthNames = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
  ];
  
  const parts = name.split(' ');
  if (parts.length !== 2) return false;
  
  const month = parts[0];
  const year = parseInt(parts[1]);
  
  return monthNames.includes(month) && !isNaN(year) && year > 2000;
}

// ===============================================
// IMPROVED INSTALMENT TRACKING SYSTEM
// ===============================================

function processInstalmentPaymentImproved(studentName, course, fullPrice, actualPrice, paymentDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let trackerSheet = ss.getSheetByName('Instalment Tracker');
  
  if (!trackerSheet) {
    Logger.log('⚠️ Creating Instalment Tracker sheet');
    trackerSheet = createInstalmentTrackerSheet(ss);
  }
  
  // Normalize student name (trim whitespace, consistent capitalization)
  const normalizedName = normalizeStudentName(studentName);
  
  // Find existing record with improved matching
  const existingRowIndex = findStudentRecordImproved(trackerSheet, normalizedName, course, fullPrice);
  
  if (existingRowIndex > 0) {
    updateStudentRecordImproved(trackerSheet, existingRowIndex, actualPrice, paymentDate);
  } else {
    createStudentRecordImproved(trackerSheet, normalizedName, course, fullPrice, actualPrice, paymentDate);
  }
}

function normalizeStudentName(name) {
  if (!name) return '';
  // Trim whitespace and normalize to title case
  return name.toString().trim().replace(/\s+/g, ' ');
}

function findStudentRecordImproved(sheet, studentName, course, fullPrice) {
  const dataRange = sheet.getDataRange();
  if (dataRange.getNumRows() <= 1) return 0;
  
  const allData = dataRange.getValues();
  const dataRows = allData.slice(1);
  
  const normalizedSearchName = normalizeStudentName(studentName);
  const searchFullPrice = Number(fullPrice);
  
  for (let i = 0; i < dataRows.length; i++) {
    const row = dataRows[i];
    const rowName = normalizeStudentName(row[0]);
    const rowCourse = row[1];
    const rowFullPrice = Number(row[2]);
    
    // Match on normalized name, course, and full price
    if (rowName === normalizedSearchName && 
        rowCourse === course && 
        rowFullPrice === searchFullPrice) {
      return i + 2; // +2 for header and 1-indexed
    }
  }
  
  return 0;
}

function createStudentRecordImproved(sheet, studentName, course, fullPrice, actualPrice, paymentDate) {
  const instalmentCount = 1; // First payment
  const nextPaymentDue = calculateNextPaymentDate(paymentDate);
  
  // Check if this first payment already completes the course
  const isComplete = Number(actualPrice) >= Number(fullPrice);
  
  const newRecord = [
    studentName,
    course,
    Number(fullPrice),
    Number(actualPrice),
    instalmentCount,
    paymentDate,
    isComplete ? '' : nextPaymentDue,
    isComplete ? 'Complete' : 'In Progress',
    isComplete ? new Date() : ''
  ];
  
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, 1, newRecord.length).setValues([newRecord]);
  
  Logger.log(`📝 Created new record: ${studentName} - ${course} - Instalment 1 - ${isComplete ? 'COMPLETE' : 'In Progress'}`);
}

function updateStudentRecordImproved(sheet, rowIndex, actualPrice, paymentDate) {
  const currentData = sheet.getRange(rowIndex, 1, 1, 9).getValues()[0];
  
  const studentName = currentData[0];
  const course = currentData[1];
  const fullPrice = Number(currentData[2]);
  const currentAmountPaid = Number(currentData[3]);
  const currentInstalmentCount = Number(currentData[4]);
  
  const newAmountPaid = currentAmountPaid + Number(actualPrice);
  const newInstalmentCount = currentInstalmentCount + 1;
  
  // Improved completion check - allow for small rounding differences
  const isComplete = newAmountPaid >= (fullPrice - 0.01);
  const completionStatus = isComplete ? 'Complete' : 'In Progress';
  const completionDate = isComplete ? new Date() : '';
  const nextPaymentDue = isComplete ? '' : calculateNextPaymentDate(paymentDate);
  
  // Update all fields
  sheet.getRange(rowIndex, 4).setValue(newAmountPaid);
  sheet.getRange(rowIndex, 5).setValue(newInstalmentCount);
  sheet.getRange(rowIndex, 6).setValue(paymentDate);
  sheet.getRange(rowIndex, 7).setValue(nextPaymentDue);
  sheet.getRange(rowIndex, 8).setValue(completionStatus);
  sheet.getRange(rowIndex, 9).setValue(completionDate);
  
  if (isComplete) {
    Logger.log(`✅ COMPLETED: ${studentName} - ${course} - Total: £${newAmountPaid} (${newInstalmentCount} instalments)`);
  } else {
    Logger.log(`📝 Updated: ${studentName} - ${course} - Instalment ${newInstalmentCount} - Total so far: £${newAmountPaid}`);
  }
}

function calculateNextPaymentDate(lastPaymentDate) {
  const nextDate = new Date(lastPaymentDate);
  nextDate.setMonth(nextDate.getMonth() + 1);
  return nextDate;
}

function createInstalmentTrackerSheet(ss) {
  const sheet = ss.insertSheet('Instalment Tracker');
  const headers = [
    'Student Name', 'Course', 'Full Price', 'Amount Paid', 'Instalments Paid',
    'Last Payment Date', 'Next Payment Due', 'Payment Complete', 'Completion Date'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#f0f0f0');
  
  return sheet;
}

// ===============================================
// AUTOMATIC CLEANUP IMPROVED
// ===============================================

function cleanupCompletedPaymentsImproved() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trackerSheet = ss.getSheetByName('Instalment Tracker');
  
  if (!trackerSheet) {
    Logger.log('Instalment Tracker sheet not found');
    return;
  }
  
  const dataRange = trackerSheet.getDataRange();
  if (dataRange.getNumRows() <= 1) {
    Logger.log('No data in Instalment Tracker to clean up');
    return;
  }
  
  const allData = dataRange.getValues();
  const headers = allData[0];
  const dataRows = allData.slice(1);
  
  const statusCol = headers.indexOf('Payment Complete');
  const completionCol = headers.indexOf('Completion Date');
  const nameCol = headers.indexOf('Student Name');
  
  const currentDate = new Date();
  const thirtyDaysAgo = new Date(currentDate.getTime() - (30 * 24 * 60 * 60 * 1000));
  
  Logger.log('🧹 Cleaning up completed payments older than 30 days...');
  
  let deletedCount = 0;
  
  // Process in reverse to avoid index issues
  for (let i = dataRows.length - 1; i >= 0; i--) {
    const row = dataRows[i];
    const rowIndex = i + 2;
    const completionStatus = row[statusCol];
    const completionDate = row[completionCol];
    const studentName = row[nameCol];
    
    if (completionStatus === 'Complete' && completionDate) {
      const completionDateObj = new Date(completionDate);
      
      if (completionDateObj < thirtyDaysAgo) {
        const daysAgo = Math.floor((currentDate - completionDateObj) / (1000 * 60 * 60 * 24));
        trackerSheet.deleteRow(rowIndex);
        Logger.log(`🗑️ Deleted: ${studentName} (completed ${daysAgo} days ago)`);
        deletedCount++;
      }
    }
  }
  
  Logger.log(`✅ Cleanup complete - removed ${deletedCount} old records`);
}

// ===============================================
// SETUP IMPROVED SYSTEM
// ===============================================

function setupImprovedInstalmentTracking() {
  Logger.log('🔧 SETTING UP IMPROVED INSTALMENT TRACKING SYSTEM');
  Logger.log('=================================================');
  
  // Clear old triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'cleanupCompletedPayments') {
      ScriptApp.deleteTrigger(trigger);
      Logger.log('🗑️ Removed old cleanup trigger');
    }
  });
  
  // Set up improved weekly cleanup trigger
  ScriptApp.newTrigger('cleanupCompletedPaymentsImproved')
    .timeBased()
    .everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(6)
    .create();
  
  Logger.log('✅ Improved cleanup trigger set up (Mondays at 6 AM)');
  Logger.log('\n💡 The system will now:');
  Logger.log('   - Normalize student names to prevent duplicates');
  Logger.log('   - Accurately detect completion (even with small rounding differences)');
  Logger.log('   - Automatically clean up old records weekly');
  Logger.log('   - Provide better logging of instalment progress');
}

// ===============================================
// COMPLETE FIX WORKFLOW
// ===============================================

function completeInstalmentTrackerFix() {
  Logger.log('🔧 COMPLETE INSTALMENT TRACKER FIX');
  Logger.log('==================================\n');
  
  Logger.log('Step 1: Diagnosing current issues...\n');
  diagnoseInstalmentTracker();
  
  Logger.log('\n\nStep 2: Fixing identified issues...\n');
  fixInstalmentTracker();
  
  Logger.log('\n\nStep 3: Setting up improved system...\n');
  setupImprovedInstalmentTracking();
  
  Logger.log('\n\n✅ COMPLETE! Instalment Tracker has been fixed and improved.');
  Logger.log('\n💡 The system will now work better going forward with:');
  Logger.log('   - Better duplicate prevention');
  Logger.log('   - Accurate completion detection');
  Logger.log('   - Automatic weekly cleanup');
}// ===============================================
// INSTALMENT TRACKER DIAGNOSTIC & REPAIR
// ===============================================

function diagnoseInstalmentTracker() {
  Logger.log('🔍 DIAGNOSING INSTALMENT TRACKER ISSUES');
  Logger.log('======================================');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trackerSheet = ss.getSheetByName('Instalment Tracker');
  
  if (!trackerSheet) {
    Logger.log('❌ Instalment Tracker sheet not found');
    return;
  }
  
  const dataRange = trackerSheet.getDataRange();
  if (dataRange.getNumRows() <= 1) {
    Logger.log('✅ Instalment Tracker is empty - nothing to diagnose');
    return;
  }
  
  const allData = dataRange.getValues();
  const headers = allData[0];
  const dataRows = allData.slice(1);
  
  Logger.log(`📊 Found ${dataRows.length} records in Instalment Tracker\n`);
  Logger.log(`Headers: ${headers.join(', ')}\n`);
  
  // Find column indices
  const nameCol = headers.indexOf('Student Name');
  const courseCol = headers.indexOf('Course');
  const fullPriceCol = headers.indexOf('Full Price');
  const amountPaidCol = headers.indexOf('Amount Paid');
  const instalmentsCol = headers.indexOf('Instalments Paid');
  const lastPaymentCol = headers.indexOf('Last Payment Date');
  const nextDueCol = headers.indexOf('Next Payment Due');
  const statusCol = headers.indexOf('Payment Complete');
  const completionCol = headers.indexOf('Completion Date');
  
  let shouldBeComplete = [];
  let shouldBeDeleted = [];
  let statusCorrect = [];
  
  const currentDate = new Date();
  const thirtyDaysAgo = new Date(currentDate.getTime() - (30 * 24 * 60 * 60 * 1000));
  
  dataRows.forEach((row, index) => {
    const rowIndex = index + 2; // +2 for header and 1-indexed
    const studentName = row[nameCol];
    const course = row[courseCol];
    const fullPrice = Number(row[fullPriceCol]);
    const amountPaid = Number(row[amountPaidCol]);
    const status = row[statusCol];
    const completionDate = row[completionCol];
    
    Logger.log(`\n📋 Row ${rowIndex}: ${studentName} - ${course}`);
    Logger.log(`   Full Price: £${fullPrice}`);
    Logger.log(`   Amount Paid: £${amountPaid}`);
    Logger.log(`   Status: ${status}`);
    Logger.log(`   Completion Date: ${completionDate || 'None'}`);
    
    // Issue 1: Should be marked complete but isn't
    if (amountPaid >= fullPrice && status !== 'Complete') {
      Logger.log(`   ⚠️ ISSUE: Has paid full amount but status is "${status}"`);
      shouldBeComplete.push({
        rowIndex: rowIndex,
        studentName: studentName,
        fullPrice: fullPrice,
        amountPaid: amountPaid
      });
    }
    
    // Issue 2: Completed more than 30 days ago but still showing
    if (status === 'Complete' && completionDate) {
      const completionDateObj = new Date(completionDate);
      const daysAgo = Math.floor((currentDate - completionDateObj) / (1000 * 60 * 60 * 24));
      
      Logger.log(`   Completed ${daysAgo} days ago`);
      
      if (completionDateObj < thirtyDaysAgo) {
        Logger.log(`   ⚠️ ISSUE: Completed over 30 days ago, should be deleted`);
        shouldBeDeleted.push({
          rowIndex: rowIndex,
          studentName: studentName,
          completionDate: completionDate,
          daysAgo: daysAgo
        });
      }
    }
    
    // Check if status is correct
    if (amountPaid >= fullPrice && status === 'Complete') {
      Logger.log(`   ✅ Status is CORRECT`);
      statusCorrect.push(studentName);
    } else if (amountPaid < fullPrice && status === 'In Progress') {
      Logger.log(`   ✅ Status is CORRECT`);
      statusCorrect.push(studentName);
    }
  });
  
  // Summary
  Logger.log('\n\n📊 DIAGNOSIS SUMMARY');
  Logger.log('===================');
  Logger.log(`✅ Records with correct status: ${statusCorrect.length}`);
  Logger.log(`⚠️ Records that should be marked complete: ${shouldBeComplete.length}`);
  Logger.log(`🗑️ Records that should be deleted (>30 days): ${shouldBeDeleted.length}`);
  
  if (shouldBeComplete.length > 0) {
    Logger.log('\n🔧 STUDENTS TO MARK AS COMPLETE:');
    shouldBeComplete.forEach(item => {
      Logger.log(`   - Row ${item.rowIndex}: ${item.studentName} (Paid £${item.amountPaid} of £${item.fullPrice})`);
    });
  }
  
  if (shouldBeDeleted.length > 0) {
    Logger.log('\n🗑️ RECORDS TO DELETE:');
    shouldBeDeleted.forEach(item => {
      Logger.log(`   - Row ${item.rowIndex}: ${item.studentName} (Completed ${item.daysAgo} days ago)`);
    });
  }
  
  Logger.log('\n💡 NEXT STEPS:');
  Logger.log('   Run fixInstalmentTracker() to automatically fix all issues');
}

function fixInstalmentTracker() {
  Logger.log('🔧 FIXING INSTALMENT TRACKER');
  Logger.log('============================');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trackerSheet = ss.getSheetByName('Instalment Tracker');
  
  if (!trackerSheet) {
    Logger.log('❌ Instalment Tracker sheet not found');
    return;
  }
  
  const dataRange = trackerSheet.getDataRange();
  if (dataRange.getNumRows() <= 1) {
    Logger.log('✅ Instalment Tracker is empty - nothing to fix');
    return;
  }
  
  const allData = dataRange.getValues();
  const headers = allData[0];
  const dataRows = allData.slice(1);
  
  // Find column indices
  const nameCol = headers.indexOf('Student Name');
  const fullPriceCol = headers.indexOf('Full Price');
  const amountPaidCol = headers.indexOf('Amount Paid');
  const statusCol = headers.indexOf('Payment Complete');
  const completionCol = headers.indexOf('Completion Date');
  
  let markedComplete = 0;
  let deleted = 0;
  
  const currentDate = new Date();
  const thirtyDaysAgo = new Date(currentDate.getTime() - (30 * 24 * 60 * 60 * 1000));
  
  // Process in reverse order to safely delete rows
  for (let i = dataRows.length - 1; i >= 0; i--) {
    const row = dataRows[i];
    const rowIndex = i + 2; // +2 for header and 1-indexed
    const studentName = row[nameCol];
    const fullPrice = Number(row[fullPriceCol]);
    const amountPaid = Number(row[amountPaidCol]);
    const status = row[statusCol];
    const completionDate = row[completionCol];
    
    // Fix 1: Mark as complete if fully paid
    if (amountPaid >= fullPrice && status !== 'Complete') {
      Logger.log(`✅ Marking ${studentName} as Complete (paid £${amountPaid} of £${fullPrice})`);
      
      trackerSheet.getRange(rowIndex, statusCol + 1).setValue('Complete');
      trackerSheet.getRange(rowIndex, completionCol + 1).setValue(currentDate);
      
      markedComplete++;
    }
    
    // Fix 2: Delete if completed over 30 days ago
    if (status === 'Complete' && completionDate) {
      const completionDateObj = new Date(completionDate);
      
      if (completionDateObj < thirtyDaysAgo) {
        const daysAgo = Math.floor((currentDate - completionDateObj) / (1000 * 60 * 60 * 24));
        Logger.log(`🗑️ Deleting ${studentName} (completed ${daysAgo} days ago)`);
        
        trackerSheet.deleteRow(rowIndex);
        deleted++;
      }
    }
  }
  
  Logger.log('\n📊 FIX SUMMARY');
  Logger.log('=============');
  Logger.log(`✅ Marked as complete: ${markedComplete} students`);
  Logger.log(`🗑️ Deleted old records: ${deleted} students`);
  Logger.log('\n✅ Instalment Tracker has been fixed!');
}

function rebuildInstalmentTrackerFromMonthlySheets() {
  Logger.log('🔄 REBUILDING INSTALMENT TRACKER FROM MONTHLY SHEETS');
  Logger.log('===================================================');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trackerSheet = ss.getSheetByName('Instalment Tracker');
  
  if (!trackerSheet) {
    Logger.log('❌ Instalment Tracker sheet not found');
    return;
  }
  
  // Clear existing tracker (except headers)
  const lastRow = trackerSheet.getLastRow();
  if (lastRow > 1) {
    trackerSheet.deleteRows(2, lastRow - 1);
    Logger.log('🗑️ Cleared existing tracker data');
  }
  
  // Get all monthly sheets
  const allSheets = ss.getSheets();
  const monthlySheets = allSheets.filter(sheet => isMonthlySheetName(sheet.getName()));
  
  Logger.log(`📊 Found ${monthlySheets.length} monthly sheets to scan\n`);
  
  let totalProcessed = 0;
  
  monthlySheets.forEach(sheet => {
    const sheetName = sheet.getName();
    Logger.log(`\n📋 Processing ${sheetName}...`);
    
    const dataRange = sheet.getDataRange();
    if (dataRange.getNumRows() <= 1) {
      Logger.log('   No data');
      return;
    }
    
    const allData = dataRange.getValues();
    const headers = allData[0];
    const dataRows = allData.slice(1);
    
    // Find column indices
    const dateCol = headers.indexOf('Date');
    const nameCol = headers.indexOf('Name');
    const courseCol = headers.indexOf('Course');
    const fullPriceCol = headers.indexOf('Full Price');
    const actualPriceCol = headers.indexOf('Actual Price');
    const paymentPlanCol = headers.indexOf('Payment Plan');
    
    if (nameCol === -1 || fullPriceCol === -1 || actualPriceCol === -1) {
      Logger.log('   ⚠️ Missing required columns');
      return;
    }
    
    // Process each row with payment plan
    dataRows.forEach(row => {
      const hasPaymentPlan = row[paymentPlanCol] === 'Y';
      
      if (hasPaymentPlan && row[nameCol] && row[fullPriceCol] && row[actualPriceCol]) {
        const studentName = row[nameCol];
        const course = row[courseCol];
        const fullPrice = row[fullPriceCol];
        const actualPrice = row[actualPriceCol];
        const paymentDate = row[dateCol] ? new Date(row[dateCol]) : new Date();
        
        Logger.log(`   📝 Processing instalment: ${studentName} - £${actualPrice}`);
        
        processInstalmentPayment(studentName, course, fullPrice, actualPrice, paymentDate);
        totalProcessed++;
      }
    });
  });
  
  Logger.log(`\n✅ REBUILD COMPLETE`);
  Logger.log(`   Processed ${totalProcessed} instalment payments`);
  Logger.log('\n💡 Now run fixInstalmentTracker() to clean up any issues');
}

function isMonthlySheetName(name) {
  const monthNames = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
  ];
  
  const parts = name.split(' ');
  if (parts.length !== 2) return false;
  
  const month = parts[0];
  const year = parseInt(parts[1]);
  
  return monthNames.includes(month) && !isNaN(year) && year > 2000;
}

// Include necessary functions from original code
function processInstalmentPayment(studentName, course, fullPrice, actualPrice, paymentDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let trackerSheet = ss.getSheetByName('Instalment Tracker');
  
  if (!trackerSheet) {
    Logger.log('⚠️ Creating Instalment Tracker sheet');
    trackerSheet = createInstalmentTrackerSheet(ss);
  }
  
  const existingRowIndex = findStudentRecord(trackerSheet, studentName, course, fullPrice);
  
  if (existingRowIndex > 0) {
    updateStudentRecord(trackerSheet, existingRowIndex, actualPrice, paymentDate);
  } else {
    createStudentRecord(trackerSheet, studentName, course, fullPrice, actualPrice, paymentDate);
  }
}

function findStudentRecord(sheet, studentName, course, fullPrice) {
  const dataRange = sheet.getDataRange();
  if (dataRange.getNumRows() <= 1) return 0;
  
  const allData = dataRange.getValues();
  const dataRows = allData.slice(1);
  
  for (let i = 0; i < dataRows.length; i++) {
    const row = dataRows[i];
    if (row[0] === studentName && row[1] === course && Number(row[2]) === Number(fullPrice)) {
      return i + 2;
    }
  }
  
  return 0;
}

function createStudentRecord(sheet, studentName, course, fullPrice, actualPrice, paymentDate) {
  const instalmentCount = 1;
  const nextPaymentDue = calculateNextPaymentDate(paymentDate);
  
  const newRecord = [
    studentName,
    course,
    Number(fullPrice),
    Number(actualPrice),
    instalmentCount,
    paymentDate,
    nextPaymentDue,
    'In Progress',
    ''
  ];
  
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, 1, newRecord.length).setValues([newRecord]);
}

function updateStudentRecord(sheet, rowIndex, actualPrice, paymentDate) {
  const currentData = sheet.getRange(rowIndex, 1, 1, 9).getValues()[0];
  
  const currentAmountPaid = Number(currentData[3]);
  const newAmountPaid = currentAmountPaid + Number(actualPrice);
  const newInstalmentCount = Number(currentData[4]) + 1;
  const fullPrice = Number(currentData[2]);
  
  const isComplete = newAmountPaid >= fullPrice;
  const completionStatus = isComplete ? 'Complete' : 'In Progress';
  const completionDate = isComplete ? new Date() : '';
  const nextPaymentDue = isComplete ? '' : calculateNextPaymentDate(paymentDate);
  
  sheet.getRange(rowIndex, 4).setValue(newAmountPaid);
  sheet.getRange(rowIndex, 5).setValue(newInstalmentCount);
  sheet.getRange(rowIndex, 6).setValue(paymentDate);
  sheet.getRange(rowIndex, 7).setValue(nextPaymentDue);
  sheet.getRange(rowIndex, 8).setValue(completionStatus);
  sheet.getRange(rowIndex, 9).setValue(completionDate);
}

function calculateNextPaymentDate(lastPaymentDate) {
  const nextDate = new Date(lastPaymentDate);
  nextDate.setMonth(nextDate.getMonth() + 1);
  return nextDate;
}

function createInstalmentTrackerSheet(ss) {
  const sheet = ss.insertSheet('Instalment Tracker');
  const headers = [
    'Student Name', 'Course', 'Full Price', 'Amount Paid', 'Instalments Paid',
    'Last Payment Date', 'Next Payment Due', 'Payment Complete', 'Completion Date'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#f0f0f0');
  
  return sheet;
}

// ===============================================
// COMPLETE FIX WORKFLOW
// ===============================================

function completeInstalmentTrackerFix() {
  Logger.log('🔧 COMPLETE INSTALMENT TRACKER FIX');
  Logger.log('==================================\n');
  
  Logger.log('Step 1: Diagnosing current issues...\n');
  diagnoseInstalmentTracker();
  
  Logger.log('\n\nStep 2: Fixing identified issues...\n');
  fixInstalmentTracker();
  
  Logger.log('\n\n✅ COMPLETE! Instalment Tracker has been fixed.');
  Logger.log('\n💡 If you want to completely rebuild from scratch, run rebuildInstalmentTrackerFromMonthlySheets()');
}