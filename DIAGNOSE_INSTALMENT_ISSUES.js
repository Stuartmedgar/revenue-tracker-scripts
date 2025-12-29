// ===============================================
// DIAGNOSE INSTALMENT TRACKER ISSUES
// UPDATED: Platinum course now £1047 (was £997) with £397, £350, £300 installments
// ===============================================

function diagnoseInstalmentTrackerIssues() {
  Logger.log('🔍 DIAGNOSING INSTALMENT TRACKER ISSUES');
  Logger.log('=======================================\n');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trackerSheet = ss.getSheetByName('Instalment Tracker');
  
  if (!trackerSheet) {
    Logger.log('❌ Instalment Tracker sheet not found');
    return;
  }
  
  const dataRange = trackerSheet.getDataRange();
  const allData = dataRange.getValues();
  const headers = allData[0];
  const dataRows = allData.slice(1);
  
  Logger.log(`📊 Analyzing ${dataRows.length} records in Instalment Tracker\n`);
  
  const issues = [];
  
  dataRows.forEach((row, index) => {
    const studentName = row[0];
    const course = row[1];
    const fullPrice = Number(row[2]);
    const amountPaid = Number(row[3]);
    const instalmentsPaid = Number(row[4]);
    const status = row[7];
    
    if (!studentName) return;
    
    const issue = {
      rowNumber: index + 2,
      studentName: studentName,
      course: course,
      fullPrice: fullPrice,
      amountPaid: amountPaid,
      instalmentsPaid: instalmentsPaid,
      status: status,
      problems: []
    };
    
    // Check 1: First payment detection issue (300 or 350 as instalment 1)
    if ((amountPaid === 300 || amountPaid === 350) && instalmentsPaid === 1) {
      issue.problems.push(`❌ CRITICAL: Shows £${amountPaid} as instalment 1 (should be 2 or 3)`);
      issue.problems.push(`   Expected first payment: £${getExpectedFirstPayment(fullPrice)}`);
      issue.problems.push('   → Student likely has missing first payment');
    }
    
    // Check 2: Amount paid equals full price but shows multiple instalments
    if (amountPaid === fullPrice && instalmentsPaid > 1) {
      issue.problems.push(`❌ Shows ${instalmentsPaid} instalments but paid full price in one go`);
      issue.problems.push('   → Should show 1 instalment or full payment');
    }
    
    // Check 3: Amount paid less than expected for instalment number
    const expectedMinimum = getExpectedMinimumForInstalments(fullPrice, instalmentsPaid);
    if (amountPaid < expectedMinimum) {
      issue.problems.push(`⚠️ Amount paid (£${amountPaid}) less than expected minimum (£${expectedMinimum}) for ${instalmentsPaid} instalments`);
    }
    
    // Check 4: Course name mismatch
    const expectedCourse = getCourseFromPrice(fullPrice);
    if (expectedCourse && course !== expectedCourse) {
      issue.problems.push(`⚠️ Course shows as "${course}" but full price (£${fullPrice}) indicates "${expectedCourse}"`);
    }
    
    if (issue.problems.length > 0) {
      issues.push(issue);
    }
  });
  
  // Report findings
  Logger.log(`\n📋 FOUND ${issues.length} RECORDS WITH ISSUES:\n`);
  Logger.log('='.repeat(80));
  
  issues.forEach((issue, index) => {
    Logger.log(`\n${index + 1}. ${issue.studentName} (Row ${issue.rowNumber})`);
    Logger.log(`   Course: ${issue.course}, Full Price: £${issue.fullPrice}`);
    Logger.log(`   Amount Paid: £${issue.amountPaid}, Instalments: ${issue.instalmentsPaid}, Status: ${issue.status}`);
    Logger.log('   Problems:');
    issue.problems.forEach(problem => {
      Logger.log(`   ${problem}`);
    });
  });
  
  Logger.log('\n' + '='.repeat(80));
  Logger.log('\n🔎 NEXT STEP: Run searchMonthlySheetsMissingPayments() to find missing first payments');
}

function getExpectedFirstPayment(fullPrice) {
  switch (Number(fullPrice)) {
    case 1047: return 397; // UPDATED: Platinum
    case 822: return 522; // Tuition/Revision Plus
    case 647: return 347; // Revision
    case 597: return 297; // Tuition
    default: return '?';
  }
}

function getExpectedMinimumForInstalments(fullPrice, instalmentCount) {
  const full = Number(fullPrice);
  const count = Number(instalmentCount);
  
  // UPDATED: Platinum (1047): 397 + 350 + 300
  if (full === 1047) {
    if (count === 1) return 397;
    if (count === 2) return 747; // 397 + 350
    if (count === 3) return 1047; // 397 + 350 + 300
  }
  
  // Tuition/Revision Plus (822): 522 + 300
  if (full === 822) {
    if (count === 1) return 522;
    if (count === 2) return 822;
  }
  
  // Revision (647): 347 + 300
  if (full === 647) {
    if (count === 1) return 347;
    if (count === 2) return 647;
  }
  
  // Tuition (597): 297 + 300
  if (full === 597) {
    if (count === 1) return 297;
    if (count === 2) return 597;
  }
  
  return 0;
}

function searchMonthlySheetsMissingPayments() {
  Logger.log('🔍 SEARCHING MONTHLY SHEETS FOR MISSING FIRST PAYMENTS');
  Logger.log('=====================================================\n');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trackerSheet = ss.getSheetByName('Instalment Tracker');
  
  if (!trackerSheet) {
    Logger.log('❌ Instalment Tracker sheet not found');
    return;
  }
  
  // Get all students with 300 or 350 as first payment issue
  const dataRange = trackerSheet.getDataRange();
  const allData = dataRange.getValues();
  const dataRows = allData.slice(1);
  
  const problematicStudents = [];
  
  dataRows.forEach((row, index) => {
    const studentName = row[0];
    const amountPaid = Number(row[3]);
    const instalmentsPaid = Number(row[4]);
    const fullPrice = Number(row[2]);
    
    if (studentName && (amountPaid === 300 || amountPaid === 350) && instalmentsPaid === 1) {
      problematicStudents.push({
        name: studentName,
        fullPrice: fullPrice,
        expectedFirstPayment: getExpectedFirstPayment(fullPrice),
        rowNumber: index + 2
      });
    }
  });
  
  Logger.log(`📊 Found ${problematicStudents.length} students with potential missing first payments\n`);
  
  // Search monthly sheets for these students
  const allSheets = ss.getSheets();
  const monthlySheets = allSheets.filter(sheet => isMonthlySheetName(sheet.getName()));
  
  Logger.log(`🔎 Searching ${monthlySheets.length} monthly sheets...\n`);
  
  problematicStudents.forEach(student => {
    Logger.log(`\n${'='.repeat(80)}`);
    Logger.log(`🔍 Searching for: ${student.name}`);
    Logger.log(`   Expected first payment: £${student.expectedFirstPayment}`);
    Logger.log(`   Current tracker shows: £${student.amountPaid || 300} (instalment 1) ← INCORRECT`);
    
    let foundPayments = [];
    
    monthlySheets.forEach(sheet => {
      const sheetData = sheet.getDataRange().getValues();
      const sheetHeaders = sheetData[0];
      const sheetRows = sheetData.slice(1);
      
      const nameCol = sheetHeaders.indexOf('Name');
      const actualPriceCol = sheetHeaders.indexOf('Actual Price');
      const dateCol = sheetHeaders.indexOf('Date');
      
      if (nameCol === -1 || actualPriceCol === -1) return;
      
      sheetRows.forEach((row, rowIndex) => {
        if (row[nameCol] === student.name) {
          foundPayments.push({
            sheetName: sheet.getName(),
            amount: Number(row[actualPriceCol]),
            date: row[dateCol],
            rowNumber: rowIndex + 2
          });
        }
      });
    });
    
    if (foundPayments.length === 0) {
      Logger.log('   ❌ NO PAYMENTS FOUND in monthly sheets');
      Logger.log('   → Student might be in Data Entry or Sort sheet');
    } else {
      Logger.log(`   ✅ Found ${foundPayments.length} payment(s) in monthly sheets:`);
      
      // Sort by date
      foundPayments.sort((a, b) => new Date(a.date) - new Date(b.date));
      
      foundPayments.forEach((payment, index) => {
        const isExpectedFirst = payment.amount === student.expectedFirstPayment;
        const marker = isExpectedFirst ? '🎯 FIRST PAYMENT → ' : '     ';
        Logger.log(`   ${marker}${index + 1}. ${payment.sheetName} - £${payment.amount} on ${new Date(payment.date).toLocaleDateString()}`);
      });
      
      // Check if first payment exists
      const hasFirstPayment = foundPayments.some(p => p.amount === student.expectedFirstPayment);
      
      if (hasFirstPayment) {
        Logger.log('\n   ✅ SOLUTION: First payment EXISTS in monthly sheets');
        Logger.log(`   → Need to UPDATE Instalment Tracker row ${student.rowNumber}`);
        Logger.log(`   → Set Amount Paid from £${student.amountPaid || 300} to correct total`);
        Logger.log(`   → Set Instalments from 1 to correct number`);
      } else {
        Logger.log('\n   ❌ PROBLEM: First payment (£' + student.expectedFirstPayment + ') NOT FOUND');
        Logger.log('   → Student may have started payment plan but first payment is missing');
        Logger.log('   → Or first payment went through a different name/spelling');
      }
    }
  });
  
  Logger.log('\n' + '='.repeat(80));
  Logger.log('\n📋 SUMMARY: Run fixInstalmentTrackerData() to automatically fix the issues');
}

function fixInstalmentTrackerDataImproved() {
  Logger.log('🔧 FIXING INSTALMENT TRACKER DATA (IMPROVED VERSION)');
  Logger.log('===================================================\n');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trackerSheet = ss.getSheetByName('Instalment Tracker');
  
  if (!trackerSheet) {
    Logger.log('❌ Instalment Tracker sheet not found');
    return;
  }
  
  const dataRange = trackerSheet.getDataRange();
  const allData = dataRange.getValues();
  const headers = allData[0];
  const dataRows = allData.slice(1);
  
  const allSheets = ss.getSheets();
  const monthlySheets = allSheets.filter(sheet => isMonthlySheetName(sheet.getName()));
  
  let fixedCount = 0;
  let skippedCount = 0;
  
  dataRows.forEach((row, index) => {
    const rowNumber = index + 2;
    const studentName = row[0];
    const course = row[1];
    const fullPrice = Number(row[2]);
    const amountPaid = Number(row[3]);
    const instalmentsPaid = Number(row[4]);
    const status = row[7];
    
    if (!studentName) return;
    
    // ONLY FIX: Students showing £300 or £350 as instalment 1
    if ((amountPaid === 300 || amountPaid === 350) && instalmentsPaid === 1) {
      Logger.log(`\n🔧 Fixing: ${studentName} (Row ${rowNumber})`);
      Logger.log(`   Current: £${amountPaid} shown as instalment 1`);
      Logger.log(`   Expected first payment: £${getExpectedFirstPayment(fullPrice)}`);
      
      // Search for all payments in monthly sheets
      let allPayments = [];
      
      monthlySheets.forEach(sheet => {
        const sheetData = sheet.getDataRange().getValues();
        const sheetHeaders = sheetData[0];
        const sheetRows = sheetData.slice(1);
        
        const nameCol = sheetHeaders.indexOf('Name');
        const actualPriceCol = sheetHeaders.indexOf('Actual Price');
        const dateCol = sheetHeaders.indexOf('Date');
        
        if (nameCol === -1 || actualPriceCol === -1) return;
        
        sheetRows.forEach(r => {
          if (r[nameCol] === studentName) {
            allPayments.push({
              amount: Number(r[actualPriceCol]),
              date: new Date(r[dateCol])
            });
          }
        });
      });
      
      if (allPayments.length === 0) {
        Logger.log(`   ⚠️ No payments found in monthly sheets - skipping`);
        skippedCount++;
        return;
      }
      
      // Sort by date (oldest first)
      allPayments.sort((a, b) => a.date - b.date);
      
      // Calculate total amount and correct instalment count
      const totalPaid = allPayments.reduce((sum, p) => sum + p.amount, 0);
      const correctInstalmentCount = allPayments.length;
      
      Logger.log(`   ✅ Found ${allPayments.length} payment(s) in monthly sheets:`);
      allPayments.forEach((p, i) => {
        Logger.log(`      ${i + 1}. £${p.amount} on ${p.date.toLocaleDateString()}`);
      });
      
      Logger.log(`   📊 Updating tracker:`);
      Logger.log(`      Amount Paid: £${amountPaid} → £${totalPaid}`);
      Logger.log(`      Instalments: ${instalmentsPaid} → ${correctInstalmentCount}`);
      
      // Update the tracker
      trackerSheet.getRange(rowNumber, 4).setValue(totalPaid); // Amount Paid
      trackerSheet.getRange(rowNumber, 5).setValue(correctInstalmentCount); // Instalments Paid
      
      // Get the last payment date (most recent)
      const lastPaymentDate = allPayments[allPayments.length - 1].date;
      trackerSheet.getRange(rowNumber, 6).setValue(lastPaymentDate); // Last Payment Date
      
      // Update status if now complete
      if (totalPaid >= fullPrice) {
        trackerSheet.getRange(rowNumber, 7).setValue(''); // Clear Next Payment Due
        trackerSheet.getRange(rowNumber, 8).setValue('Complete'); // Payment Complete
        trackerSheet.getRange(rowNumber, 9).setValue(new Date()); // Completion Date
        Logger.log(`      Status: In Progress → Complete ✅`);
      } else {
        // Calculate next payment due (1 month after last payment)
        const nextDue = new Date(lastPaymentDate);
        nextDue.setMonth(nextDue.getMonth() + 1);
        trackerSheet.getRange(rowNumber, 7).setValue(nextDue);
        Logger.log(`      Next Payment Due: ${nextDue.toLocaleDateString()}`);
      }
      
      fixedCount++;
      Logger.log('   ✅ Fixed!');
    }
    
    // SKIP: Students who paid full price via instalments (they're correct)
    else if (amountPaid === fullPrice && instalmentsPaid > 1) {
      Logger.log(`\n⏭️ Skipping: ${studentName} (Row ${rowNumber})`);
      Logger.log(`   Reason: Paid £${fullPrice} in ${instalmentsPaid} instalments`);
      Logger.log(`   This is CORRECT - they made multiple instalment payments before tracker existed`);
      skippedCount++;
    }
  });
  
  Logger.log(`\n\n${'='.repeat(80)}`);
  Logger.log('📋 SUMMARY:');
  Logger.log(`✅ Fixed: ${fixedCount} records`);
  Logger.log(`⏭️ Skipped: ${skippedCount} records (already correct)`);
  Logger.log(`\n🎉 Instalment Tracker updated successfully!`);
  
  if (fixedCount > 0) {
    Logger.log('\n💡 NEXT STEPS:');
    Logger.log('1. Check your Instalment Tracker sheet to verify the updates');
    Logger.log('2. Investigate any remaining issues separately');
  }
}