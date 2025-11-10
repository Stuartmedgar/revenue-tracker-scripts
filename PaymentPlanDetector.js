// ===============================================
// PAYMENTPLANDETECTOR.GS - Payment Plan Detection Logic
// ===============================================

function detectPaymentPlan(fullPrice, actualPrice) {
  const full = Number(fullPrice);
  const actual = Number(actualPrice);
  
  // Platinum Course (997)
  if (full === 997) {
    if (actual === 997) {
      return { isPaymentPlan: false, instalment: '', course: 'Platinum' };
    } else if (actual === 397) {
      return { isPaymentPlan: true, instalment: '1 of 3', course: 'Platinum' };
    } else if (actual === 300) {
      return { isPaymentPlan: true, instalment: '2 or 3 of 3', course: 'Platinum' };
    }
  }
  
  // Revision Course (647)
  if (full === 647) {
    if (actual === 647) {
      return { isPaymentPlan: false, instalment: '', course: 'Revision' };
    } else if (actual === 347) {
      return { isPaymentPlan: true, instalment: '1 of 2', course: 'Revision' };
    } else if (actual === 300) {
      return { isPaymentPlan: true, instalment: '2 of 2', course: 'Revision' };
    }
  }
  
  // Tuition Course (597)
  if (full === 597) {
    if (actual === 597) {
      return { isPaymentPlan: false, instalment: '', course: 'Tuition' };
    } else if (actual === 297) {
      return { isPaymentPlan: true, instalment: '1 of 2', course: 'Tuition' };
    } else if (actual === 300) {
      return { isPaymentPlan: true, instalment: '2 of 2', course: 'Tuition' };
    }
  }
  
  // Default case - not a recognized payment pattern
  return { isPaymentPlan: false, instalment: '', course: getCourseFromPrice(full) };
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

function getPaymentPlanInfo(fullPrice, actualPrice) {
  // This function is used by other parts of the system
  return detectPaymentPlan(fullPrice, actualPrice);
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
    Logger.log(`  Result: ${result.course} - Payment Plan: ${result.isPaymentPlan ? 'Y' : 'N'} - ${result.instalment}`);
    Logger.log(`  Expected: ${test.expected}`);
    Logger.log('');
  });
  
  Logger.log('Payment plan detection test complete');
}