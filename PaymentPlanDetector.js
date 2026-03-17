// ===============================================
// PAYMENTPLANDETECTOR.GS - Payment Plan Detection Logic
// UPDATED: Platinum course now £1047 (was £997) with £397, £350, £300 installments
// UPDATED: Added Tuition/Revision Plus (822) support
// UPDATED: Added legacy £997 Platinum support for backward compatibility
// UPDATED: Added type coercion safety (handles text numbers from sheets)
// ===============================================

function detectPaymentPlan(fullPrice, actualPrice) {
  // Coerce to number and round to 2dp to guard against text values or
  // floating-point noise (e.g. "1047" string, or 299.9999 from Stripe)
  const full   = Math.round(Number(fullPrice)   * 100) / 100;
  const actual = Math.round(Number(actualPrice) * 100) / 100;

  // Platinum Course (1047) - current pricing
  if (full === 1047) {
    if (actual === 1047) return { isPaymentPlan: false, instalment: '',       course: 'Platinum' };
    if (actual === 397)  return { isPaymentPlan: true,  instalment: '1 of 3', course: 'Platinum' };
    if (actual === 350)  return { isPaymentPlan: true,  instalment: '2 of 3', course: 'Platinum' };
    if (actual === 300)  return { isPaymentPlan: true,  instalment: '3 of 3', course: 'Platinum' };
  }

  // Platinum Course (997) - legacy pricing, kept for backward compatibility
  if (full === 997) {
    if (actual === 997)  return { isPaymentPlan: false, instalment: '',       course: 'Platinum' };
    if (actual === 397)  return { isPaymentPlan: true,  instalment: '1 of 3', course: 'Platinum' };
    if (actual === 300)  return { isPaymentPlan: true,  instalment: '2 of 3', course: 'Platinum' };
    // Note: £997 plan had variable 3rd instalment (300); 2 of 3 and 3 of 3 both = 300
    // so we can only reliably detect instalment 1 — 2nd/3rd marked generically below
  }

  // Tuition/Revision Plus Course (822)
  if (full === 822) {
    if (actual === 822)  return { isPaymentPlan: false, instalment: '',       course: 'Tuition/Revision Plus' };
    if (actual === 522)  return { isPaymentPlan: true,  instalment: '1 of 2', course: 'Tuition/Revision Plus' };
    if (actual === 300)  return { isPaymentPlan: true,  instalment: '2 of 2', course: 'Tuition/Revision Plus' };
  }

  // Revision Course (647)
  if (full === 647) {
    if (actual === 647)  return { isPaymentPlan: false, instalment: '',       course: 'Revision' };
    if (actual === 347)  return { isPaymentPlan: true,  instalment: '1 of 2', course: 'Revision' };
    if (actual === 300)  return { isPaymentPlan: true,  instalment: '2 of 2', course: 'Revision' };
  }

  // Tuition Course (597)
  if (full === 597) {
    if (actual === 597)  return { isPaymentPlan: false, instalment: '',       course: 'Tuition' };
    if (actual === 297)  return { isPaymentPlan: true,  instalment: '1 of 2', course: 'Tuition' };
    if (actual === 300)  return { isPaymentPlan: true,  instalment: '2 of 2', course: 'Tuition' };
  }

  // Default — not a recognised payment pattern
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

  const fullPriceCol   = headers.indexOf('Full Price');
  const actualPriceCol = headers.indexOf('Actual Price');
  const paymentPlanCol = headers.indexOf('Payment Plan');
  const instalmentCol  = headers.indexOf('Instalment');

  if (fullPriceCol === -1 || actualPriceCol === -1 || paymentPlanCol === -1 || instalmentCol === -1) {
    Logger.log(`${sheet.getName()} missing required columns`);
    return;
  }

  Logger.log(`Updating ${dataRows.length} rows in ${sheet.getName()}`);

  dataRows.forEach((row, index) => {
    const rowNumber  = index + 2;
    const fullPrice  = row[fullPriceCol];
    const actualPrice = row[actualPriceCol];

    if (fullPrice && actualPrice) {
      const paymentInfo      = detectPaymentPlan(fullPrice, actualPrice);
      const paymentPlanValue = paymentInfo.isPaymentPlan ? 'Y' : '';
      sheet.getRange(rowNumber, paymentPlanCol + 1).setValue(paymentPlanValue);
      sheet.getRange(rowNumber, instalmentCol  + 1).setValue(paymentInfo.instalment);

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

  const testCases = [
    // Current Platinum pricing
    { full: 1047, actual: 1047, expected: 'Platinum - Full payment' },
    { full: 1047, actual: 397,  expected: 'Platinum - Instalment 1 of 3' },
    { full: 1047, actual: 350,  expected: 'Platinum - Instalment 2 of 3' },
    { full: 1047, actual: 300,  expected: 'Platinum - Instalment 3 of 3' },
    // Legacy Platinum pricing
    { full: 997,  actual: 997,  expected: 'Platinum (legacy) - Full payment' },
    { full: 997,  actual: 397,  expected: 'Platinum (legacy) - Instalment 1 of 3' },
    { full: 997,  actual: 300,  expected: 'Platinum (legacy) - Instalment 2/3 of 3' },
    // Tuition/Revision Plus
    { full: 822,  actual: 822,  expected: 'Tuition/Revision Plus - Full payment' },
    { full: 822,  actual: 522,  expected: 'Tuition/Revision Plus - Instalment 1 of 2' },
    { full: 822,  actual: 300,  expected: 'Tuition/Revision Plus - Instalment 2 of 2' },
    // Revision
    { full: 647,  actual: 647,  expected: 'Revision - Full payment' },
    { full: 647,  actual: 347,  expected: 'Revision - Instalment 1 of 2' },
    { full: 647,  actual: 300,  expected: 'Revision - Instalment 2 of 2' },
    // Tuition
    { full: 597,  actual: 597,  expected: 'Tuition - Full payment' },
    { full: 597,  actual: 297,  expected: 'Tuition - Instalment 1 of 2' },
    { full: 597,  actual: 300,  expected: 'Tuition - Instalment 2 of 2' },
    // Type coercion safety
    { full: '1047', actual: '397', expected: 'Handles string inputs correctly' },
  ];

  testCases.forEach(test => {
    const result = detectPaymentPlan(test.full, test.actual);
    Logger.log(`Full: ${test.full}, Actual: ${test.actual}`);
    Logger.log(`  Result: ${result.course} - Payment Plan: ${result.isPaymentPlan ? 'Y' : 'N'} - Instalment: "${result.instalment}"`);
    Logger.log(`  Expected: ${test.expected}`);
    Logger.log('');
  });

  Logger.log('Payment plan detection test complete');
}