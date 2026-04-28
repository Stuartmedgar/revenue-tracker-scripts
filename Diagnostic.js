function debugMarch2026() {
  // Check Revenue side
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const revSheet = ss.getSheetByName('March 2026');
  Logger.log('=== REVENUE AUTOMATED (March 2026) ===');
  for (let row = 2; row <= Math.min(6, revSheet.getLastRow()); row++) {
    const name = revSheet.getRange(row, 2).getValue();  // Column B
    const amount = revSheet.getRange(row, 6).getValue(); // Column F
    Logger.log(`Row ${row}: Name="${name}" | Amount=${amount}`);
  }

  // Check Payment Rec side
  const paymentSS = SpreadsheetApp.openById('1vi9n0mTXxGTGAppecDk04krBqgqvGf4ZNTvVlhseyFY');
  const paySheet = paymentSS.getSheetByName('Mar26');
  Logger.log('=== PAYMENT RECONCILIATION (Mar26) ===');
  for (let row = 15; row <= Math.min(20, paySheet.getLastRow()); row++) {
    const name = paySheet.getRange(row, 8).getValue();   // Column H
    const amount = paySheet.getRange(row, 19).getValue(); // Column S
    Logger.log(`Row ${row}: Name="${name}" | Amount=${amount}`);
  }
}