// ===============================================
// SHEETCREATORS.GS - Sheet Creation Functions (Updated)
// ===============================================

function createFailedOrdersSheet(ss) {
  const sheet = ss.insertSheet('Failed orders');
  const headers = ['Date', 'Name', 'Course', 'Sitting', 'Full Price', 'Actual Price', 'Order Type'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Format headers
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#f0f0f0');
  
  // Force Sitting column to text format
  const sittingColumn = sheet.getRange('D:D');
  sittingColumn.setNumberFormat('@');
  
  // Position the tab correctly
  positionFailedOrdersTab(ss, sheet);
  
  Logger.log('Created Failed orders sheet with text formatting');
  return sheet;
}

function createMonthlySheet(ss, sheetName) {
  const sheet = ss.insertSheet(sheetName);
  const headers = [
    'Date', 'Name', 'Course', 'Sitting', 'Full Price', 'Actual Price', 'Order Type',
    'FME Fee', 'Stripe Fee', 'Actual Stripe Fee', 'Expected Income', 
    'Payment Plan', 'Instalment', 'Comment'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Format headers
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#f0f0f0');
  
  // Force Sitting column to text format
  const sittingColumn = sheet.getRange('D:D');
  sittingColumn.setNumberFormat('@');
  
  Logger.log(`Created monthly sheet: ${sheetName} with text formatting`);
  return sheet;
}

function createSortingSheet(ss) {
  const sheet = ss.insertSheet('Sort');
  const headers = [
    'Date', 'Name', 'Course', 'Sitting', 'Full Price', 'Actual Price', 'Order Type',
    'Manually Entered', 'Employer Invoice', 'Other'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Format headers
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#f0f0f0');
  
  // Force Sitting column to text format
  const sittingColumn = sheet.getRange('D:D');
  sittingColumn.setNumberFormat('@');
  
  // Position the tab correctly
  positionSortingTab(ss, sheet);
  
  Logger.log('Created Sorting sheet with text formatting');
  return sheet;
}