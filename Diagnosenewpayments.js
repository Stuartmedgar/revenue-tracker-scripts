// ===============================================
// DIAGNOSE NEW PAYMENTS - Find Missing Students in Instalment Tracker
// ===============================================

function diagnoseNewPayments() {
  Logger.log('🔍 DIAGNOSING WHY NEW STUDENTS AREN\'T BEING TRACKED');
  Logger.log('===================================================\n');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trackerSheet = ss.getSheetByName('Instalment Tracker');
  const allSheets = ss.getSheets();
  
  // Get list of students already in tracker
  const trackedStudents = new Set();
  if (trackerSheet) {
    const trackerData = trackerSheet.getDataRange().getValues().slice(1);
    trackerData.forEach(row => {
      if (row[0]) trackedStudents.add(row[0]);
    });
  }
  
  Logger.log(`📊 Currently tracking ${trackedStudents.size} students\n`);
  
  // Find monthly sheets from last 3 months
  const now = new Date();
  const threeMonthsAgo = new Date(now.getTime() - (90 * 24 * 60 * 60 * 1000));
  
  const monthlySheets = allSheets.filter(sheet => {
    const name = sheet.getName();
    const match = name.match(/^(\w+) (\d{4})$/);
    if (!match) return false;
    
    const monthNames = ['January', 'February', 'March', 'April', 'May', 'June',
                        'July', 'August', 'September', 'October', 'November', 'December'];
    const monthIndex = monthNames.indexOf(match[1]);
    const year = parseInt(match[2]);
    
    if (monthIndex === -1) return false;
    
    const sheetDate = new Date(year, monthIndex, 1);
    return sheetDate >= threeMonthsAgo;
  });
  
  Logger.log(`🔎 Checking ${monthlySheets.length} recent monthly sheets...\n`);
  
  let foundPaymentPlans = 0;
  let missingFromTracker = [];
  
  monthlySheets.forEach(sheet => {
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1);
    
    const nameCol = headers.indexOf('Name');
    const paymentPlanCol = headers.indexOf('Payment Plan');
    const actualPriceCol = headers.indexOf('Actual Price');
    const fullPriceCol = headers.indexOf('Full Price');
    const dateCol = headers.indexOf('Date');
    
    if (nameCol === -1 || paymentPlanCol === -1) {
      Logger.log(`⚠️ Skipping ${sheet.getName()} - missing columns`);
      return;
    }
    
    rows.forEach(row => {
      const hasPaymentPlan = row[paymentPlanCol] === 'Y';
      const studentName = row[nameCol];
      
      if (hasPaymentPlan && studentName) {
        foundPaymentPlans++;
        
        if (!trackedStudents.has(studentName)) {
          missingFromTracker.push({
            name: studentName,
            sheet: sheet.getName(),
            actualPrice: row[actualPriceCol],
            fullPrice: row[fullPriceCol],
            date: row[dateCol]
          });
        }
      }
    });
  });
  
  Logger.log(`📋 FINDINGS:`);
  Logger.log(`✅ Found ${foundPaymentPlans} payment plan payments in recent sheets`);
  Logger.log(`❌ Missing from tracker: ${missingFromTracker.length} students\n`);
  
  if (missingFromTracker.length > 0) {
    Logger.log(`🚨 STUDENTS NOT IN TRACKER BUT SHOULD BE:\n`);
    missingFromTracker.forEach((student, i) => {
      Logger.log(`${i + 1}. ${student.name}`);
      Logger.log(`   Sheet: ${student.sheet}`);
      Logger.log(`   Payment: £${student.actualPrice} of £${student.fullPrice}`);
      Logger.log(`   Date: ${new Date(student.date).toDateString()}`);
      Logger.log('');
    });
    
    Logger.log('\n💡 POSSIBLE CAUSES:');
    Logger.log('1. processInstalmentPayment() not being called when data moves to monthly sheets');
    Logger.log('2. Price mismatch between old pricing (£997) and new pricing (£1047)');
    Logger.log('3. Automatic trigger not running properly');
    Logger.log('4. Error occurring in processInstalmentPayment() function');
    
    Logger.log('\n🔧 RECOMMENDED FIX:');
    Logger.log('Run updateInstalmentTrackerFromMonthlySheets() to add all missing students');
  } else {
    Logger.log('✅ All payment plan students are being tracked correctly!');
    Logger.log('✅ No action needed - tracker is up to date');
  }
  
  Logger.log('\n\n' + '='.repeat(80));
  Logger.log('DIAGNOSTIC COMPLETE');
}

// ===============================================
// CHECK SPECIFIC STUDENT
// ===============================================

function checkSpecificStudent(studentName) {
  Logger.log(`🔍 CHECKING PAYMENT HISTORY FOR: ${studentName}`);
  Logger.log('==========================================\n');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();
  const monthlySheets = allSheets.filter(sheet => {
    const name = sheet.getName();
    return /^(January|February|March|April|May|June|July|August|September|October|November|December) \d{4}$/.test(name);
  });
  
  Logger.log(`Searching ${monthlySheets.length} monthly sheets...\n`);
  
  const payments = [];
  
  monthlySheets.forEach(sheet => {
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1);
    
    const nameCol = headers.indexOf('Name');
    const actualPriceCol = headers.indexOf('Actual Price');
    const dateCol = headers.indexOf('Date');
    const fullPriceCol = headers.indexOf('Full Price');
    const paymentPlanCol = headers.indexOf('Payment Plan');
    
    if (nameCol === -1) return;
    
    rows.forEach(row => {
      if (row[nameCol] && row[nameCol].toString().toLowerCase().includes(studentName.toLowerCase())) {
        payments.push({
          sheet: sheet.getName(),
          name: row[nameCol],
          amount: row[actualPriceCol],
          fullPrice: row[fullPriceCol],
          paymentPlan: row[paymentPlanCol],
          date: row[dateCol]
        });
      }
    });
  });
  
  if (payments.length === 0) {
    Logger.log(`❌ No payments found for "${studentName}"`);
    return;
  }
  
  Logger.log(`✅ Found ${payments.length} payment(s):\n`);
  
  payments.sort((a, b) => new Date(a.date) - new Date(b.date));
  
  let total = 0;
  payments.forEach((p, i) => {
    total += Number(p.amount);
    Logger.log(`${i + 1}. ${p.sheet}`);
    Logger.log(`   Name: ${p.name}`);
    Logger.log(`   £${p.amount} on ${new Date(p.date).toDateString()}`);
    Logger.log(`   Full Price: £${p.fullPrice}`);
    Logger.log(`   Payment Plan: ${p.paymentPlan}`);
    Logger.log('');
  });
  
  Logger.log(`Total Paid: £${total}`);
  Logger.log(`Expected for Platinum: £997 (old) or £1047 (new)`);
  
  // Check tracker status
  const trackerSheet = ss.getSheetByName('Instalment Tracker');
  if (trackerSheet) {
    const trackerData = trackerSheet.getDataRange().getValues();
    const trackerRow = trackerData.find(row => 
      row[0] && row[0].toString().toLowerCase().includes(studentName.toLowerCase())
    );
    
    if (trackerRow) {
      Logger.log('\n✅ FOUND IN TRACKER:');
      Logger.log(`   Status: ${trackerRow[7]}`);
      Logger.log(`   Amount Paid: £${trackerRow[3]}`);
      Logger.log(`   Instalments: ${trackerRow[4]}`);
    } else {
      Logger.log('\n❌ NOT IN TRACKER');
    }
  }
  
  if (total < 997) {
    Logger.log('\n⚠️ INCOMPLETE - student has not paid full amount yet');
  } else if (total >= 997) {
    Logger.log('\n✅ COMPLETE - student has paid full amount');
  }
}

// ===============================================
// EXAMPLE USAGE
// ===============================================

function checkNinaWakulinska() {
  checkSpecificStudent('Nina Wakulinska');
}