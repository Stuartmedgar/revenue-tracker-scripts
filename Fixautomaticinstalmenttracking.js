// ===============================================
// FIX AUTOMATIC INSTALMENT TRACKING GOING FORWARD
// Run these functions in order to ensure new payments are tracked automatically
// ===============================================

/**
 * STEP 1: Verify the automatic system is set up correctly
 * Run this first to check if everything is in place
 */
function verifyInstalmentTrackingSetup() {
  Logger.log('🔍 VERIFYING INSTALMENT TRACKING SETUP');
  Logger.log('=====================================\n');
  
  // Check 1: Verify trigger exists
  Logger.log('1️⃣ CHECKING TRIGGERS:');
  const triggers = ScriptApp.getProjectTriggers();
  const dataProcessingTrigger = triggers.find(t => t.getHandlerFunction() === 'processDataEntries');
  
  if (dataProcessingTrigger) {
    Logger.log('✅ Data processing trigger EXISTS (runs every 5 minutes)');
  } else {
    Logger.log('❌ Data processing trigger MISSING');
    Logger.log('   This trigger should call processInstalmentPayment() automatically');
  }
  
  // Check 2: Verify Instalment Tracker sheet exists
  Logger.log('\n2️⃣ CHECKING INSTALMENT TRACKER SHEET:');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trackerSheet = ss.getSheetByName('Instalment Tracker');
  
  if (trackerSheet) {
    Logger.log('✅ Instalment Tracker sheet EXISTS');
    const rowCount = trackerSheet.getLastRow() - 1; // Exclude header
    Logger.log(`   Currently tracking ${rowCount} students`);
  } else {
    Logger.log('❌ Instalment Tracker sheet MISSING');
  }
  
  // Check 3: Test payment plan detection
  Logger.log('\n3️⃣ TESTING PAYMENT PLAN DETECTION:');
  
  const testCases = [
    { fullPrice: 1047, actualPrice: 397, expected: 'Platinum - Instalment 1 of 3' },
    { fullPrice: 997, actualPrice: 397, expected: 'Platinum - Instalment 1 of 3 (old pricing)' },
    { fullPrice: 647, actualPrice: 347, expected: 'Revision - Instalment 1 of 2' },
    { fullPrice: 597, actualPrice: 297, expected: 'Tuition - Instalment 1 of 2' }
  ];
  
  let allPassed = true;
  testCases.forEach(test => {
    const result = getPaymentPlanInfo(test.fullPrice, test.actualPrice);
    const passed = result.isPaymentPlan === true;
    const status = passed ? '✅' : '❌';
    Logger.log(`${status} £${test.actualPrice} of £${test.fullPrice}: ${result.course} - ${result.instalment || 'Full payment'}`);
    if (!passed) allPassed = false;
  });
  
  // Summary
  Logger.log('\n📋 SUMMARY:');
  if (!dataProcessingTrigger) {
    Logger.log('❌ PROBLEM: Missing automatic trigger');
    Logger.log('🔧 FIX: Run setupAutomaticInstalmentTracking()');
    return false;
  } else if (!trackerSheet) {
    Logger.log('⚠️ WARNING: Instalment Tracker sheet missing');
    Logger.log('   It will be created automatically on first payment');
    return true;
  } else if (!allPassed) {
    Logger.log('❌ PROBLEM: Payment plan detection failing');
    Logger.log('🔧 FIX: Check PaymentPlanDetector.js file');
    return false;
  } else {
    Logger.log('✅ System is set up correctly!');
    Logger.log('New payments should be tracked automatically');
    return true;
  }
}

/**
 * STEP 2: Set up automatic tracking
 * Run this if triggers are missing
 */
function setupAutomaticInstalmentTracking() {
  Logger.log('🔧 SETTING UP AUTOMATIC INSTALMENT TRACKING');
  Logger.log('==========================================\n');
  
  // Clear existing triggers first
  Logger.log('1️⃣ Clearing old triggers...');
  const triggers = ScriptApp.getProjectTriggers();
  let deletedCount = 0;
  
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'processDataEntries') {
      ScriptApp.deleteTrigger(trigger);
      deletedCount++;
    }
  });
  Logger.log(`   Deleted ${deletedCount} old trigger(s)`);
  
  // Set up fresh trigger
  Logger.log('\n2️⃣ Creating new trigger...');
  ScriptApp.newTrigger('processDataEntries')
    .timeBased()
    .everyMinutes(5)
    .create();
  Logger.log('   ✅ Created trigger to run every 5 minutes');
  
  // Set up cleanup trigger
  Logger.log('\n3️⃣ Setting up weekly cleanup...');
  const cleanupTriggers = triggers.filter(t => t.getHandlerFunction() === 'cleanupCompletedPayments');
  if (cleanupTriggers.length === 0) {
    ScriptApp.newTrigger('cleanupCompletedPayments')
      .timeBased()
      .everyWeeks(1)
      .onWeekDay(ScriptApp.WeekDay.MONDAY)
      .atHour(7)
      .create();
    Logger.log('   ✅ Created weekly cleanup trigger (Mondays at 7 AM)');
  } else {
    Logger.log('   ✅ Weekly cleanup trigger already exists');
  }
  
  Logger.log('\n✅ SETUP COMPLETE!');
  Logger.log('\n📋 HOW IT WORKS:');
  Logger.log('1. Every 5 minutes, processDataEntries() runs automatically');
  Logger.log('2. Students from Data Entry → Monthly sheets');
  Logger.log('3. Payment plan students → Instalment Tracker (automatic)');
  Logger.log('4. Completed payments >30 days deleted every Monday');
  
  Logger.log('\n💡 NEXT: Run testInstalmentTrackingFlow() to verify it works');
}

/**
 * STEP 3: Test the automatic flow
 * Run this to verify the system works end-to-end
 */
function testInstalmentTrackingFlow() {
  Logger.log('🧪 TESTING INSTALMENT TRACKING FLOW');
  Logger.log('==================================\n');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create a test payment
  const testPayment = {
    studentName: 'TEST STUDENT - Delete Me',
    course: 'Platinum',
    fullPrice: 1047,
    actualPrice: 397,
    paymentDate: new Date()
  };
  
  Logger.log(`Testing with: ${testPayment.studentName}`);
  Logger.log(`Payment: £${testPayment.actualPrice} of £${testPayment.fullPrice}\n`);
  
  // Test payment plan detection
  Logger.log('1️⃣ Testing payment plan detection...');
  const paymentInfo = getPaymentPlanInfo(testPayment.fullPrice, testPayment.actualPrice);
  Logger.log(`   Result: ${paymentInfo.course} - ${paymentInfo.isPaymentPlan ? 'Payment Plan' : 'Full Payment'}`);
  Logger.log(`   Instalment: ${paymentInfo.instalment || 'N/A'}`);
  
  if (!paymentInfo.isPaymentPlan) {
    Logger.log('   ❌ PROBLEM: Not detected as payment plan!');
    return false;
  }
  Logger.log('   ✅ Correctly detected as payment plan');
  
  // Test processInstalmentPayment
  Logger.log('\n2️⃣ Testing processInstalmentPayment()...');
  try {
    processInstalmentPayment(
      testPayment.studentName,
      testPayment.course,
      testPayment.fullPrice,
      testPayment.actualPrice,
      testPayment.paymentDate
    );
    Logger.log('   ✅ Function executed without errors');
  } catch (error) {
    Logger.log(`   ❌ ERROR: ${error.toString()}`);
    return false;
  }
  
  // Verify it was added to tracker
  Logger.log('\n3️⃣ Verifying student was added to tracker...');
  const trackerSheet = ss.getSheetByName('Instalment Tracker');
  if (!trackerSheet) {
    Logger.log('   ❌ Instalment Tracker sheet not found');
    return false;
  }
  
  const trackerData = trackerSheet.getDataRange().getValues();
  const testRow = trackerData.find(row => row[0] === testPayment.studentName);
  
  if (testRow) {
    Logger.log('   ✅ Student found in tracker!');
    Logger.log(`      Name: ${testRow[0]}`);
    Logger.log(`      Course: ${testRow[1]}`);
    Logger.log(`      Full Price: £${testRow[2]}`);
    Logger.log(`      Amount Paid: £${testRow[3]}`);
    Logger.log(`      Instalments: ${testRow[4]}`);
    Logger.log(`      Status: ${testRow[7]}`);
    
    // Clean up test entry
    Logger.log('\n4️⃣ Cleaning up test entry...');
    const rowIndex = trackerData.indexOf(testRow) + 1;
    trackerSheet.deleteRow(rowIndex);
    Logger.log('   ✅ Test entry deleted');
    
    Logger.log('\n✅ TEST PASSED!');
    Logger.log('The automatic tracking system is working correctly!');
    return true;
    
  } else {
    Logger.log('   ❌ Student NOT found in tracker');
    Logger.log('   processInstalmentPayment() is not working correctly');
    return false;
  }
}

/**
 * STEP 4: Catch up any missed payments from recent weeks
 * Run this to add students who were missed during the transition
 */
function catchUpMissedPayments() {
  Logger.log('🚨 CATCHING UP MISSED PAYMENTS');
  Logger.log('=============================\n');
  
  Logger.log('Scanning recent monthly sheets for missed payment plans...\n');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get current month and previous month
  const now = new Date();
  const monthNames = ['January', 'February', 'March', 'April', 'May', 'June',
                      'July', 'August', 'September', 'October', 'November', 'December'];
  
  const currentMonth = monthNames[now.getMonth()];
  const currentYear = now.getFullYear();
  
  const previousMonthIndex = now.getMonth() === 0 ? 11 : now.getMonth() - 1;
  const previousMonth = monthNames[previousMonthIndex];
  const previousYear = previousMonthIndex === 11 ? currentYear - 1 : currentYear;
  
  Logger.log(`Checking: ${previousMonth} ${previousYear} and ${currentMonth} ${currentYear}\n`);
  
  const sheetsToCheck = [
    `${previousMonth} ${previousYear}`,
    `${currentMonth} ${currentYear}`
  ];
  
  let totalAdded = 0;
  
  sheetsToCheck.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log(`⚠️ Sheet "${sheetName}" not found, skipping`);
      return;
    }
    
    Logger.log(`📋 Processing ${sheetName}...`);
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1);
    
    const nameCol = headers.indexOf('Name');
    const courseCol = headers.indexOf('Course');
    const fullPriceCol = headers.indexOf('Full Price');
    const actualPriceCol = headers.indexOf('Actual Price');
    const paymentPlanCol = headers.indexOf('Payment Plan');
    const dateCol = headers.indexOf('Date');
    
    if (nameCol === -1 || paymentPlanCol === -1) {
      Logger.log(`   Missing required columns, skipping`);
      return;
    }
    
    let sheetAddedCount = 0;
    
    rows.forEach(row => {
      if (row[paymentPlanCol] === 'Y' && row[nameCol]) {
        try {
          processInstalmentPayment(
            row[nameCol],
            row[courseCol],
            row[fullPriceCol],
            row[actualPriceCol],
            new Date(row[dateCol])
          );
          sheetAddedCount++;
        } catch (error) {
          Logger.log(`   ❌ Error processing ${row[nameCol]}: ${error.toString()}`);
        }
      }
    });
    
    Logger.log(`   ✅ Processed ${sheetAddedCount} payment plan(s)`);
    totalAdded += sheetAddedCount;
  });
  
  Logger.log(`\n✅ COMPLETE: Processed ${totalAdded} total payment plans`);
  Logger.log('\n💡 Run diagnoseNewPayments() to verify all students are now tracked');
}

/**
 * STEP 5: Monitor daily to ensure tracking is working
 * Run this each day to check if new payments are being tracked
 */
function monitorDailyInstalmentActivity() {
  Logger.log('📊 DAILY INSTALMENT MONITORING');
  Logger.log('==============================\n');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trackerSheet = ss.getSheetByName('Instalment Tracker');
  
  if (!trackerSheet) {
    Logger.log('❌ Instalment Tracker sheet not found');
    return;
  }
  
  const now = new Date();
  const yesterdayStart = new Date(now);
  yesterdayStart.setDate(yesterdayStart.getDate() - 1);
  yesterdayStart.setHours(0, 0, 0, 0);
  
  const yesterdayEnd = new Date(now);
  yesterdayEnd.setDate(yesterdayEnd.getDate() - 1);
  yesterdayEnd.setHours(23, 59, 59, 999);
  
  const todayStart = new Date(now);
  todayStart.setHours(0, 0, 0, 0);
  
  Logger.log(`Checking from ${yesterdayStart.toLocaleDateString()} to now\n`);
  
  // Check tracker for recent additions
  const trackerData = trackerSheet.getDataRange().getValues().slice(1);
  const recentInTracker = trackerData.filter(row => {
    if (!row[5]) return false;
    const lastPayment = new Date(row[5]);
    return lastPayment >= yesterdayStart;
  });
  
  Logger.log('🔹 IN TRACKER (last 24 hours):');
  if (recentInTracker.length === 0) {
    Logger.log('   No new students added\n');
  } else {
    recentInTracker.forEach(row => {
      Logger.log(`   ✅ ${row[0]} - £${row[3]} of £${row[2]} - ${row[7]}`);
    });
    Logger.log('');
  }
  
  // Check monthly sheets for recent payment plans
  const currentMonth = now.toLocaleString('en-US', { month: 'long' });
  const currentYear = now.getFullYear();
  const monthlySheet = ss.getSheetByName(`${currentMonth} ${currentYear}`);
  
  if (monthlySheet) {
    const data = monthlySheet.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1);
    
    const nameCol = headers.indexOf('Name');
    const paymentPlanCol = headers.indexOf('Payment Plan');
    const dateCol = headers.indexOf('Date');
    const actualPriceCol = headers.indexOf('Actual Price');
    
    const recentPaymentPlans = rows.filter(row => {
      if (row[paymentPlanCol] !== 'Y') return false;
      if (!row[dateCol]) return false;
      const paymentDate = new Date(row[dateCol]);
      return paymentDate >= yesterdayStart;
    });
    
    Logger.log(`🔹 IN ${currentMonth.toUpperCase()} SHEET (last 24 hours):`);
    if (recentPaymentPlans.length === 0) {
      Logger.log('   No payment plans found\n');
    } else {
      recentPaymentPlans.forEach(row => {
        const inTracker = trackerData.some(t => t[0] === row[nameCol]);
        const status = inTracker ? '✅' : '❌';
        Logger.log(`   ${status} ${row[nameCol]} - £${row[actualPriceCol]} - ${inTracker ? 'In tracker' : 'NOT IN TRACKER'}`);
      });
      Logger.log('');
    }
    
    // Check for discrepancies
    const missingFromTracker = recentPaymentPlans.filter(row => {
      return !trackerData.some(t => t[0] === row[nameCol]);
    });
    
    if (missingFromTracker.length > 0) {
      Logger.log('🚨 ALERT: Payment plans NOT in tracker:');
      missingFromTracker.forEach(row => {
        Logger.log(`   ❌ ${row[nameCol]}`);
      });
      Logger.log('\n🔧 FIX: Run catchUpMissedPayments()');
    } else if (recentPaymentPlans.length > 0) {
      Logger.log('✅ All recent payment plans are being tracked correctly');
    }
  }
}

/**
 * ONE-TIME SETUP: Run this once to get everything working
 */
function completeInstalmentTrackingSetup() {
  Logger.log('🚀 COMPLETE INSTALMENT TRACKING SETUP');
  Logger.log('====================================\n');
  
  Logger.log('STEP 1: Verifying current setup...\n');
  const setupOK = verifyInstalmentTrackingSetup();
  
  if (!setupOK) {
    Logger.log('\nSTEP 2: Setting up automatic tracking...\n');
    setupAutomaticInstalmentTracking();
  }
  
  Logger.log('\n\nSTEP 3: Testing the flow...\n');
  const testPassed = testInstalmentTrackingFlow();
  
  if (!testPassed) {
    Logger.log('\n❌ SETUP FAILED - Check errors above');
    return;
  }
  
  Logger.log('\n\nSTEP 4: Catching up missed payments...\n');
  catchUpMissedPayments();
  
  Logger.log('\n\n' + '='.repeat(80));
  Logger.log('✅ SETUP COMPLETE!');
  Logger.log('='.repeat(80));
  Logger.log('\n📋 WHAT TO DO NEXT:');
  Logger.log('1. Run monitorDailyInstalmentActivity() each day to verify tracking');
  Logger.log('2. Check logs for any "NOT IN TRACKER" alerts');
  Logger.log('3. If you see alerts, run catchUpMissedPayments()');
  Logger.log('\n💡 The system will now automatically track all new payment plans!');
}