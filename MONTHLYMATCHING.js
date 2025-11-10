// ===============================================
// MONTHLY MATCHING SYSTEM - Complete with Improved Name+Amount Matching
// ===============================================

function setupMonthlyMatching(month, year) {
  try {
    Logger.log(`Setting up monthly matching for ${month} ${year}...`);
    const revenueSheetName = `${month} ${year}`;
    const paymentSheetName = getShortMonthName(month) + year.toString().slice(-2);
    
    Logger.log(`Looking for sheets: "${revenueSheetName}" and "${paymentSheetName}"`);
    
    // Setup Revenue sheet (this spreadsheet)
    setupRevenueSheetMatching(revenueSheetName, paymentSheetName);
    
    // Setup Payment Reconciliation sheet (other spreadsheet)
    setupPaymentSheetMatching(revenueSheetName, paymentSheetName);
    
    // Setup costs transfer
    setupCostsTransfer(revenueSheetName, paymentSheetName);
    
    Logger.log(`Monthly matching setup complete for ${month} ${year}!`);
  } catch (error) {
    Logger.log(`Error setting up monthly matching: ${error.toString()}`);
  }
}

function setupRevenueSheetMatching(revenueSheetName, paymentSheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const revenueSheet = ss.getSheetByName(revenueSheetName);
  
  if (!revenueSheet) {
    Logger.log(`Revenue sheet "${revenueSheetName}" not found`);
    return;
  }
  
  Logger.log(`Setting up improved matching formulas in ${revenueSheetName}...`);
  
  // Add headers if they don't exist
  addMatchingHeaders(revenueSheet);
  
  // Add formulas to existing data rows
  const lastRow = revenueSheet.getLastRow();
  if (lastRow > 1) {
    // Use the improved Name+Amount matching formulas
    addRevenueMatchingFormulas(revenueSheet, paymentSheetName, lastRow);
  }
  
  // Add conditional formatting
  addRevenueConditionalFormatting(revenueSheet, lastRow);
  
  Logger.log(`✅ Revenue sheet setup complete with Name+Amount matching`);
}

function setupPaymentSheetMatching(revenueSheetName, paymentSheetName) {
  // Payment Reconciliation spreadsheet ID
  const PAYMENT_SPREADSHEET_ID = "1vi9n0mTXxGTGAppecDk04krBqgqvGf4ZNTvVlhseyFY";
  
  try {
    const paymentSpreadsheet = SpreadsheetApp.openById(PAYMENT_SPREADSHEET_ID);
    const paymentSheet = paymentSpreadsheet.getSheetByName(paymentSheetName);
    
    if (!paymentSheet) {
      Logger.log(`Payment sheet "${paymentSheetName}" not found - please create it first`);
      return;
    }
    
    Logger.log(`Setting up matching formulas in ${paymentSheetName}...`);
    
    // Find the last row with data
    const lastRow = paymentSheet.getLastRow();
    
    // Add match formula to column AL for all data rows (starting row 15)
    if (lastRow >= 15) {
      addPaymentMatchingFormulas(paymentSheet, revenueSheetName, lastRow);
    }
    
    // Add conditional formatting with special rules
    addPaymentConditionalFormattingWithSpecialRules(paymentSheet, lastRow);
    
    Logger.log(`Payment sheet setup complete`);
  } catch (error) {
    Logger.log(`Error accessing payment sheet: ${error.toString()}`);
    Logger.log(`Make sure you have access to the Payment Reconciliation spreadsheet`);
  }
}

// ===============================================
// IMPROVED MATCHING FORMULAS - NAME + AMOUNT
// ===============================================

function addRevenueMatchingFormulas(revenueSheet, paymentSheetName, lastRow) {
  const PAYMENT_SPREADSHEET_ID = "1vi9n0mTXxGTGAppecDk04krBqgqvGf4ZNTvVlhseyFY";

  // Clear existing formulas first
  revenueSheet.getRange(2, 15, lastRow - 1, 1).clearContent(); // Clear column O
  revenueSheet.getRange(2, 10, lastRow - 1, 1).clearContent(); // Clear column J

  // Add formulas to all data rows
  for (let row = 2; row <= lastRow; row++) {
    // Match formula for column O - matches both name AND amount
    const matchFormula = `=IF(AND(B${row}<>"",F${row}<>""),IF(COUNTIFS(IMPORTRANGE("${PAYMENT_SPREADSHEET_ID}","${paymentSheetName}!H:H"),B${row},IMPORTRANGE("${PAYMENT_SPREADSHEET_ID}","${paymentSheetName}!S:S"),F${row})>0,"Match",""),"")`;
    revenueSheet.getRange(row, 15).setFormula(matchFormula);

    // IMPROVED Stripe fee formula for column J - now matches BOTH name AND amount
    const stripeFeeFormula = `=IF(O${row}="Match",IFERROR(INDEX(FILTER(IMPORTRANGE("${PAYMENT_SPREADSHEET_ID}","${paymentSheetName}!V:V"),IMPORTRANGE("${PAYMENT_SPREADSHEET_ID}","${paymentSheetName}!H:H")=B${row},IMPORTRANGE("${PAYMENT_SPREADSHEET_ID}","${paymentSheetName}!S:S")=F${row}),1),""),"")`;
    revenueSheet.getRange(row, 10).setFormula(stripeFeeFormula);
  }

  Logger.log(`Added improved Name+Amount matching formulas to ${lastRow - 1} rows in revenue sheet`);
}

function addPaymentMatchingFormulas(paymentSheet, revenueSheetName, lastRow) {
  const REVENUE_SPREADSHEET_ID = "18tFthS9ibeoOHpc8yuiKFfkMW6SD8eamc-enRyrZiKc";
  
  // Clear existing formulas first
  paymentSheet.getRange(15, 38, lastRow - 14, 1).clearContent();
  
  // Add formulas to all data rows (starting from row 15)
  for (let row = 15; row <= lastRow; row++) {
    const matchFormula = `=IF(AND(H${row}<>"",S${row}<>""),IF(COUNTIFS(IMPORTRANGE("${REVENUE_SPREADSHEET_ID}","'${revenueSheetName}'!B:B"),H${row},IMPORTRANGE("${REVENUE_SPREADSHEET_ID}","'${revenueSheetName}'!F:F"),S${row})>0,"Match",""),"")`;
    paymentSheet.getRange(row, 38).setFormula(matchFormula);
  }
  
  Logger.log(`Added formulas to ${lastRow - 14} rows in payment sheet`);
}

function addMatchingHeaders(revenueSheet) {
  // Check if headers exist in columns J and O
  const headerRow = revenueSheet.getRange(1, 1, 1, 15).getValues()[0];
  
  // Add "Actual Stripe Fee" header in column J if not exists
  if (!headerRow[9] || headerRow[9] !== "Actual Stripe Fee") {
    revenueSheet.getRange(1, 10).setValue("Actual Stripe Fee");
  }
  
  // Add "Match" header in column O if not exists
  if (!headerRow[14] || headerRow[14] !== "Match") {
    revenueSheet.getRange(1, 15).setValue("Match");
  }
  
  // Format headers
  const headerRange = revenueSheet.getRange(1, 10, 1, 6);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#f0f0f0');
}

function addRevenueConditionalFormatting(revenueSheet, lastRow) {
  if (lastRow <= 1) return;
  
  const dataRange = revenueSheet.getRange(2, 1, lastRow - 1, 15);
  
  // Clear existing conditional formatting rules
  revenueSheet.setConditionalFormatRules([]);
  
  // Pink rule (no match)
  const pinkRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($B2<>"",$F2<>"",$O2<>"Match")')
    .setBackground('#ffcdd2')
    .setRanges([dataRange])
    .build();
  
  // Green rule (match)
  const greenRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$O2="Match"')
    .setBackground('#c8e6c9')
    .setRanges([dataRange])
    .build();
  
  // Apply rules
  revenueSheet.setConditionalFormatRules([pinkRule, greenRule]);
  
  Logger.log(`Applied conditional formatting to revenue sheet`);
}

function addPaymentConditionalFormattingWithSpecialRules(paymentSheet, lastRow) {
  if (lastRow < 15) return;
  
  const dataRange = paymentSheet.getRange(15, 1, lastRow - 14, 38);
  
  // Clear existing conditional formatting rules
  paymentSheet.setConditionalFormatRules([]);
  
  // RED rule (highest priority) - 0 in both R and S columns - needs FME fee check
  const redRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($R15=0,$S15=0)')
    .setBackground('#ff5252')
    .setBold(true)
    .setRanges([dataRange])
    .build();
  
  // LIGHT BLUE rule - 0 in column S only (exam sitting changes)
  const lightBlueRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($S15=0,$R15<>0)')
    .setBackground('#e1f5fe')
    .setRanges([dataRange])
    .build();
  
  // GREEN rule - Has match
  const greenRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$AL15="Match"')
    .setBackground('#c8e6c9')
    .setRanges([dataRange])
    .build();
  
  // PINK rule (lowest priority) - No match, has values
  const pinkRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$AL15<>"Match"')
    .setBackground('#ffcdd2')
    .setRanges([dataRange])
    .build();
  
  // Apply rules in priority order (first rule has highest priority)
  paymentSheet.setConditionalFormatRules([redRule, lightBlueRule, greenRule, pinkRule]);
  
  Logger.log(`Applied conditional formatting with special rules to payment sheet`);
}

// ===============================================
// COSTS TRANSFER FUNCTIONALITY
// ===============================================

function setupCostsTransfer(revenueSheetName, paymentSheetName) {
  Logger.log(`Setting up costs transfer from ${paymentSheetName} to ${revenueSheetName}...`);
  
  const PAYMENT_SPREADSHEET_ID = "1vi9n0mTXxGTGAppecDk04krBqgqvGf4ZNTvVlhseyFY";
  
  try {
    // Get costs data from payment sheet
    const costsData = extractCostsFromPaymentSheet(PAYMENT_SPREADSHEET_ID, paymentSheetName);
    
    if (costsData.length > 0) {
      // Add costs to revenue sheet
      addCostsToRevenueSheet(revenueSheetName, costsData);
      Logger.log(`Transferred ${costsData.length} cost entries`);
    } else {
      Logger.log(`No costs found in ${paymentSheetName}`);
    }
  } catch (error) {
    Logger.log(`Error setting up costs transfer: ${error.toString()}`);
  }
}

function extractCostsFromPaymentSheet(paymentSpreadsheetId, paymentSheetName) {
  const costsData = [];
  
  try {
    const paymentSpreadsheet = SpreadsheetApp.openById(paymentSpreadsheetId);
    const paymentSheet = paymentSpreadsheet.getSheetByName(paymentSheetName);
    
    if (!paymentSheet) {
      Logger.log(`Payment sheet ${paymentSheetName} not found`);
      return costsData;
    }
    
    const lastRow = paymentSheet.getLastRow();
    if (lastRow < 15) {
      Logger.log(`No data rows in payment sheet`);
      return costsData;
    }
    
    // Check each data row for costs in columns Z, AA, AB, AC, AD (columns 26, 27, 28, 29, 30)
    for (let row = 15; row <= lastRow; row++) {
      const rowData = paymentSheet.getRange(row, 1, 1, 30).getValues()[0];
      const description = rowData[14]; // Column O (index 14)
      const costZ = rowData[25]; // Column Z (index 25)
      const costAA = rowData[26]; // Column AA (index 26)
      const costAB = rowData[27]; // Column AB (index 27)
      const costAC = rowData[28]; // Column AC (index 28)
      const costAD = rowData[29]; // Column AD (index 29)
      
      // Check each cost column
      const costColumns = [
        { value: costZ, column: 'Z' },
        { value: costAA, column: 'AA' },
        { value: costAB, column: 'AB' },
        { value: costAC, column: 'AC' },
        { value: costAD, column: 'AD' }
      ];
      
      costColumns.forEach(cost => {
        if (cost.value && cost.value !== '' && cost.value !== 0) {
          costsData.push({
            what: description || `Cost from row ${row}`,
            cost: cost.value,
            sourceColumn: cost.column,
            sourceRow: row
          });
        }
      });
    }
    
    Logger.log(`Found ${costsData.length} costs in payment sheet`);
    return costsData;
  } catch (error) {
    Logger.log(`Error extracting costs: ${error.toString()}`);
    return costsData;
  }
}

function addCostsToRevenueSheet(revenueSheetName, costsData) {
  if (costsData.length === 0) return;
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const revenueSheet = ss.getSheetByName(revenueSheetName);
  
  if (!revenueSheet) {
    Logger.log(`Revenue sheet ${revenueSheetName} not found`);
    return;
  }
  
  // Find the last row with data
  const lastRow = revenueSheet.getLastRow();
  
  // Check if "Costs" section already exists
  const existingData = revenueSheet.getRange(1, 1, lastRow, 2).getValues();
  let costsRowExists = false;
  
  for (let i = 0; i < existingData.length; i++) {
    if (existingData[i][0] === 'Costs') {
      costsRowExists = true;
      Logger.log(`Costs section already exists - updating it`);
      break;
    }
  }
  
  let startRow;
  
  if (costsRowExists) {
    // Find where to insert new costs (after existing costs table)
    let costsEndRow = lastRow;
    for (let i = 0; i < existingData.length; i++) {
      if (existingData[i][0] === 'Costs') {
        // Find the end of the costs table
        for (let j = i + 2; j < existingData.length; j++) {
          if (existingData[j][0] === '' && existingData[j][1] === '') {
            costsEndRow = j + 1;
            break;
          }
        }
        break;
      }
    }
    startRow = costsEndRow + 1;
  } else {
    // Create new costs section
    startRow = lastRow + 3; // 3 rows below last filled row
    
    // Add "Costs" heading
    revenueSheet.getRange(startRow, 1).setValue('Costs');
    revenueSheet.getRange(startRow, 1).setFontWeight('bold');
    revenueSheet.getRange(startRow, 1).setFontSize(12);
    startRow += 1;
    
    // Add table headers
    revenueSheet.getRange(startRow, 1).setValue('What');
    revenueSheet.getRange(startRow, 2).setValue('Cost');
    
    // Format headers
    const headerRange = revenueSheet.getRange(startRow, 1, 1, 2);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#f0f0f0');
    startRow += 1;
  }
  
  // Add cost data
  const costsTableData = costsData.map(cost => [cost.what, cost.cost]);
  revenueSheet.getRange(startRow, 1, costsTableData.length, 2).setValues(costsTableData);
  
  // Format the costs column as currency
  const costsRange = revenueSheet.getRange(startRow, 2, costsTableData.length, 1);
  costsRange.setNumberFormat('£#,##0.00');
  
  Logger.log(`Added ${costsTableData.length} cost entries to ${revenueSheetName}`);
}

// ===============================================
// UTILITY FUNCTIONS
// ===============================================

function getShortMonthName(fullMonthName) {
  const monthMap = {
    'January': 'Jan',
    'February': 'Feb', 
    'March': 'Mar',
    'April': 'Apr',
    'May': 'May',
    'June': 'Jun',
    'July': 'Jul',
    'August': 'Aug',
    'September': 'Sep',
    'October': 'Oct',
    'November': 'Nov',
    'December': 'Dec'
  };
  return monthMap[fullMonthName] || fullMonthName.substring(0, 3);
}

// ===============================================
// UPDATE FUNCTIONS FOR EXISTING SHEETS
// ===============================================

function updateExistingMonthlySheetFormulas(monthName, year) {
  Logger.log(`Updating ${monthName} ${year} with improved Name+Amount matching...`);
  
  const revenueSheetName = `${monthName} ${year}`;
  const paymentSheetName = getShortMonthName(monthName) + year.toString().slice(-2);
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const revenueSheet = ss.getSheetByName(revenueSheetName);
  
  if (!revenueSheet) {
    Logger.log(`❌ Revenue sheet "${revenueSheetName}" not found`);
    return;
  }
  
  const lastRow = revenueSheet.getLastRow();
  
  if (lastRow > 1) {
    // Use the improved formula
    addRevenueMatchingFormulas(revenueSheet, paymentSheetName, lastRow);
    Logger.log(`✅ Updated ${revenueSheetName} with improved Name+Amount matching formulas`);
  } else {
    Logger.log(`❌ No data in ${revenueSheetName} to update`);
  }
}

// ===============================================
// TEST AND HELPER FUNCTIONS
// ===============================================

function testMonthlyMatching() {
  Logger.log('Testing monthly matching setup for July 2025...');
  setupMonthlyMatching('July', 2025);
  Logger.log('Test complete - check both spreadsheets!');
}

function setupCurrentMonthMatching() {
  const now = new Date();
  const monthName = getMonthName(now.getMonth());
  const year = now.getFullYear();
  Logger.log(`Setting up matching for current month: ${monthName} ${year}`);
  setupMonthlyMatching(monthName, year);
}

function setupJune2025() {
  setupMonthlyMatching('June', 2025);
}

function setupJuly2025() {
  setupMonthlyMatching('July', 2025);
}

function setupAugust2025() {
  setupMonthlyMatching('August', 2025);
}

function setupSeptember2025() {
  setupMonthlyMatching('September', 2025);
}

// Convenient wrapper functions for updating existing sheets
function updateJuly2025WithNameAmountMatching() {
  updateExistingMonthlySheetFormulas('July', 2025);
}

function updateAugust2025WithNameAmountMatching() {
  updateExistingMonthlySheetFormulas('August', 2025);
}

function updateJune2025WithNameAmountMatching() {
  updateExistingMonthlySheetFormulas('June', 2025);
}

// Function to test the improved matching on a specific month
function testImprovedMatching(monthName, year) {
  Logger.log(`=== TESTING IMPROVED MATCHING FOR ${monthName} ${year} ===`);
  
  const revenueSheetName = `${monthName} ${year}`;
  const paymentSheetName = getShortMonthName(monthName) + year.toString().slice(-2);
  
  Logger.log(`Revenue Sheet: ${revenueSheetName}`);
  Logger.log(`Payment Sheet: ${paymentSheetName}`);
  
  // Update the formulas
  updateExistingMonthlySheetFormulas(monthName, year);
  
  Logger.log('✅ Test complete - check your monthly sheet for improved matching!');
  Logger.log('The Stripe fees should now show the correct values from matching Name+Amount records');
}

// Debug function to see what the formula is actually doing
function debugNameAmountMatching(monthName, year) {
  Logger.log(`=== DEBUGGING NAME+AMOUNT MATCHING FOR ${monthName} ${year} ===`);
  
  const revenueSheetName = `${monthName} ${year}`;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const revenueSheet = ss.getSheetByName(revenueSheetName);
  
  if (!revenueSheet) {
    Logger.log(`❌ Revenue sheet "${revenueSheetName}" not found`);
    return;
  }
  
  const lastRow = revenueSheet.getLastRow();
  if (lastRow <= 1) {
    Logger.log(`❌ No data in ${revenueSheetName}`);
    return;
  }
  
  // Check a few rows to see what the formulas are returning
  for (let row = 2; row <= Math.min(5, lastRow); row++) {
    const name = revenueSheet.getRange(row, 2).getValue(); // Column B
    const amount = revenueSheet.getRange(row, 6).getValue(); // Column F  
    const match = revenueSheet.getRange(row, 15).getValue(); // Column O
    const stripeFee = revenueSheet.getRange(row, 10).getValue(); // Column J
    
    Logger.log(`Row ${row}:`);
    Logger.log(`  Name: ${name}`);
    Logger.log(`  Amount: £${amount}`);
    Logger.log(`  Match Status: ${match}`);
    Logger.log(`  Stripe Fee: £${stripeFee}`);
    Logger.log('');
  }
}

function debugCostsExtraction() {
  Logger.log('=== DEBUGGING COSTS EXTRACTION ===');
  const costsData = extractCostsFromPaymentSheet("1vi9n0mTXxGTGAppecDk04krBqgqvGf4ZNTvVlhseyFY", "Jul25");
  Logger.log(`Found ${costsData.length} costs:`);
  costsData.forEach((cost, index) => {
    Logger.log(`${index + 1}. What: "${cost.what}", Cost: ${cost.cost}, Source: Column ${cost.sourceColumn}, Row ${cost.sourceRow}`);
  });
}

function listAvailableSheets() {
  Logger.log('=== Available sheets in Revenue Automated ===');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  sheets.forEach(sheet => {
    Logger.log(`- ${sheet.getName()}`);
  });
  
  Logger.log('\n=== Available sheets in Payment Reconciliation ===');
  try {
    const paymentSS = SpreadsheetApp.openById("1vi9n0mTXxGTGAppecDk04krBqgqvGf4ZNTvVlhseyFY");
    const paymentSheets = paymentSS.getSheets();
    paymentSheets.forEach(sheet => {
      Logger.log(`- ${sheet.getName()}`);
    });
  } catch (error) {
    Logger.log(`Error accessing payment spreadsheet: ${error.toString()}`);
  }
}