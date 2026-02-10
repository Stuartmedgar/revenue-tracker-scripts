// ===============================================
// MONTHLY MATCHING SYSTEM - Complete with Improved Name+Amount Matching + FME VALIDATION
// Enhanced version with FME fee checking on BOTH sheets
// ===============================================

function setupMonthlyMatching(month, year, includeFMECheck = true) {
  try {
    Logger.log(`Setting up monthly matching for ${month} ${year}...`);
    const revenueSheetName = `${month} ${year}`;
    const paymentSheetName = getShortMonthName(month) + year.toString().slice(-2);
    
    Logger.log(`Looking for sheets: "${revenueSheetName}" and "${paymentSheetName}"`);
    
    // Setup Revenue sheet (this spreadsheet) with optional FME checking
    setupRevenueSheetMatching(revenueSheetName, paymentSheetName, includeFMECheck);
    
    // Setup Payment Reconciliation sheet (other spreadsheet) with FME checking
    setupPaymentSheetMatching(revenueSheetName, paymentSheetName, includeFMECheck);
    
    // Setup costs transfer
    setupCostsTransfer(revenueSheetName, paymentSheetName);
    
    const fmeStatus = includeFMECheck ? ' with FME validation' : '';
    Logger.log(`Monthly matching setup complete for ${month} ${year}${fmeStatus}!`);
  } catch (error) {
    Logger.log(`Error setting up monthly matching: ${error.toString()}`);
  }
}

function setupRevenueSheetMatching(revenueSheetName, paymentSheetName, includeFMECheck = true) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const revenueSheet = ss.getSheetByName(revenueSheetName);
  
  if (!revenueSheet) {
    Logger.log(`Revenue sheet "${revenueSheetName}" not found`);
    return;
  }
  
  const fmeStatus = includeFMECheck ? ' with FME validation' : '';
  Logger.log(`Setting up improved matching formulas in ${revenueSheetName}${fmeStatus}...`);
  
  // Add headers (includes FME columns if requested)
  addMatchingHeaders(revenueSheet, includeFMECheck);
  
  // Add formulas to existing data rows
  const lastRow = revenueSheet.getLastRow();
  if (lastRow > 1) {
    // Use the improved Name+Amount matching formulas
    addRevenueMatchingFormulas(revenueSheet, paymentSheetName, lastRow, includeFMECheck);
  }
  
  // Add conditional formatting
  addRevenueConditionalFormatting(revenueSheet, lastRow, includeFMECheck);
  
  const checkStatus = includeFMECheck ? 'Name+Amount matching and FME validation' : 'Name+Amount matching';
  Logger.log(`✅ Revenue sheet setup complete with ${checkStatus}`);
}

function setupPaymentSheetMatching(revenueSheetName, paymentSheetName, includeFMECheck = true) {
  // Payment Reconciliation spreadsheet ID
  const PAYMENT_SPREADSHEET_ID = "1vi9n0mTXxGTGAppecDk04krBqgqvGf4ZNTvVlhseyFY";
  
  try {
    const paymentSpreadsheet = SpreadsheetApp.openById(PAYMENT_SPREADSHEET_ID);
    const paymentSheet = paymentSpreadsheet.getSheetByName(paymentSheetName);
    
    if (!paymentSheet) {
      Logger.log(`Payment sheet "${paymentSheetName}" not found - please create it first`);
      return;
    }
    
    const fmeStatus = includeFMECheck ? ' with FME validation' : '';
    Logger.log(`Setting up matching formulas in ${paymentSheetName}${fmeStatus}...`);
    
    // Find the last row with data
    const lastRow = paymentSheet.getLastRow();
    
    // Add match formula to column AL for all data rows (starting row 15)
    if (lastRow >= 15) {
      addPaymentMatchingFormulas(paymentSheet, revenueSheetName, lastRow, includeFMECheck);
    }
    
    // Add conditional formatting with special rules (including FME if enabled)
    addPaymentConditionalFormattingWithSpecialRules(paymentSheet, lastRow, includeFMECheck);
    
    const checkStatus = includeFMECheck ? 'with FME validation' : 'standard';
    Logger.log(`Payment sheet setup complete ${checkStatus}`);
  } catch (error) {
    Logger.log(`Error accessing payment sheet: ${error.toString()}`);
    Logger.log(`Make sure you have access to the Payment Reconciliation spreadsheet`);
  }
}

// ===============================================
// IMPROVED MATCHING FORMULAS - NAME + AMOUNT (+ OPTIONAL FME)
// ===============================================

function addRevenueMatchingFormulas(revenueSheet, paymentSheetName, lastRow, includeFMECheck = true) {
  const PAYMENT_SPREADSHEET_ID = "1vi9n0mTXxGTGAppecDk04krBqgqvGf4ZNTvVlhseyFY";

  // Clear existing formulas first
  revenueSheet.getRange(2, 15, lastRow - 1, 1).clearContent(); // Clear column O (Match)
  revenueSheet.getRange(2, 10, lastRow - 1, 1).clearContent(); // Clear column J (Stripe Fee)
  
  if (includeFMECheck) {
    revenueSheet.getRange(2, 16, lastRow - 1, 2).clearContent(); // Clear columns P, Q (FME Check, Actual FME)
  }

  // Add formulas to all data rows
  for (let row = 2; row <= lastRow; row++) {
    // Column O - Match formula (matches both name AND amount)
    const matchFormula = `=IF(AND(B${row}<>"",F${row}<>""),IF(COUNTIFS(IMPORTRANGE("${PAYMENT_SPREADSHEET_ID}","${paymentSheetName}!H:H"),B${row},IMPORTRANGE("${PAYMENT_SPREADSHEET_ID}","${paymentSheetName}!S:S"),F${row})>0,"Match",""),"")`;
    revenueSheet.getRange(row, 15).setFormula(matchFormula);

    // Column J - Stripe fee formula (matches BOTH name AND amount)
    const stripeFeeFormula = `=IF(O${row}="Match",IFERROR(INDEX(FILTER(IMPORTRANGE("${PAYMENT_SPREADSHEET_ID}","${paymentSheetName}!V:V"),IMPORTRANGE("${PAYMENT_SPREADSHEET_ID}","${paymentSheetName}!H:H")=B${row},IMPORTRANGE("${PAYMENT_SPREADSHEET_ID}","${paymentSheetName}!S:S")=F${row}),1),""),"")`;
    revenueSheet.getRange(row, 10).setFormula(stripeFeeFormula);
    
    // FME VALIDATION FORMULAS (if enabled)
    if (includeFMECheck) {
      // Column Q - Actual FME from Payment Reconciliation (column Y)
      const actualFMEFormula = `=IF(O${row}="Match",IFERROR(INDEX(FILTER(IMPORTRANGE("${PAYMENT_SPREADSHEET_ID}","${paymentSheetName}!Y:Y"),IMPORTRANGE("${PAYMENT_SPREADSHEET_ID}","${paymentSheetName}!H:H")=B${row},IMPORTRANGE("${PAYMENT_SPREADSHEET_ID}","${paymentSheetName}!S:S")=F${row}),1),""),"")`;
      revenueSheet.getRange(row, 17).setFormula(actualFMEFormula);
      
      // Column P - FME Check
      // Logic: Both empty/0 = OK, otherwise compare and show difference if mismatch
      const fmeCheckFormula = `=IF(O${row}="Match",IF(AND(OR(H${row}=0,H${row}=""),OR(Q${row}=0,Q${row}="")),"OK",IF(ABS(IF(ISBLANK(H${row}),0,H${row})-IF(ISBLANK(Q${row}),0,Q${row}))<0.01,"OK",TEXT(IF(ISBLANK(H${row}),0,H${row})-IF(ISBLANK(Q${row}),0,Q${row}),"£0.00"))),"")`;
      revenueSheet.getRange(row, 16).setFormula(fmeCheckFormula);
    }
  }

  const formulaCount = includeFMECheck ? 'Name+Amount matching with FME validation' : 'Name+Amount matching';
  Logger.log(`Added ${formulaCount} formulas to ${lastRow - 1} rows in revenue sheet`);
}

function addPaymentMatchingFormulas(paymentSheet, revenueSheetName, lastRow, includeFMECheck = true) {
  const REVENUE_SPREADSHEET_ID = "18tFthS9ibeoOHpc8yuiKFfkMW6SD8eamc-enRyrZiKc";
  
  // Clear existing formulas first
  paymentSheet.getRange(15, 38, lastRow - 14, 1).clearContent(); // Column AL (Match)
  
  if (includeFMECheck) {
    paymentSheet.getRange(15, 39, lastRow - 14, 2).clearContent(); // Columns AM, AN (FME Check, Calculated FME)
  }
  
  // Add formulas to all data rows (starting from row 15)
  for (let row = 15; row <= lastRow; row++) {
    // Column AL - Match formula
    const matchFormula = `=IF(AND(H${row}<>"",S${row}<>""),IF(COUNTIFS(IMPORTRANGE("${REVENUE_SPREADSHEET_ID}","'${revenueSheetName}'!B:B"),H${row},IMPORTRANGE("${REVENUE_SPREADSHEET_ID}","'${revenueSheetName}'!F:F"),S${row})>0,"Match",""),"")`;
    paymentSheet.getRange(row, 38).setFormula(matchFormula);
    
    // FME VALIDATION FORMULAS (if enabled)
    if (includeFMECheck) {
      // Column AN - Calculated FME from Revenue Automated (column H)
      const calculatedFMEFormula = `=IF(AL${row}="Match",IFERROR(INDEX(FILTER(IMPORTRANGE("${REVENUE_SPREADSHEET_ID}","'${revenueSheetName}'!H:H"),IMPORTRANGE("${REVENUE_SPREADSHEET_ID}","'${revenueSheetName}'!B:B")=H${row},IMPORTRANGE("${REVENUE_SPREADSHEET_ID}","'${revenueSheetName}'!F:F")=S${row}),1),""),"")`;
      paymentSheet.getRange(row, 40).setFormula(calculatedFMEFormula);
      
      // Column AM - FME Check (comparing actual Y with calculated from Revenue)
      // Logic: Both empty/0 = OK, otherwise compare and show difference if mismatch
      const fmeCheckFormula = `=IF(AL${row}="Match",IF(AND(OR(Y${row}=0,Y${row}=""),OR(AN${row}=0,AN${row}="")),"OK",IF(ABS(IF(ISBLANK(Y${row}),0,Y${row})-IF(ISBLANK(AN${row}),0,AN${row}))<0.01,"OK",TEXT(IF(ISBLANK(Y${row}),0,Y${row})-IF(ISBLANK(AN${row}),0,AN${row}),"£0.00"))),"")`;
      paymentSheet.getRange(row, 39).setFormula(fmeCheckFormula);
    }
  }
  
  const formulaCount = includeFMECheck ? 'with FME validation' : 'standard';
  Logger.log(`Added formulas ${formulaCount} to ${lastRow - 14} rows in payment sheet`);
}

function addMatchingHeaders(revenueSheet, includeFMECheck = true) {
  // Check if headers exist
  const headerRow = revenueSheet.getRange(1, 1, 1, 17).getValues()[0];
  
  // Add "Actual Stripe Fee" header in column J if not exists
  if (!headerRow[9] || headerRow[9] !== "Actual Stripe Fee") {
    revenueSheet.getRange(1, 10).setValue("Actual Stripe Fee");
  }
  
  // Add "Match" header in column O if not exists
  if (!headerRow[14] || headerRow[14] !== "Match") {
    revenueSheet.getRange(1, 15).setValue("Match");
  }
  
  // Add FME validation headers if requested
  if (includeFMECheck) {
    // Column P - FME Check
    if (!headerRow[15] || headerRow[15] !== "FME Check") {
      revenueSheet.getRange(1, 16).setValue("FME Check");
    }
    
    // Column Q - Actual FME
    if (!headerRow[16] || headerRow[16] !== "Actual FME") {
      revenueSheet.getRange(1, 17).setValue("Actual FME");
    }
  }
  
  // Format headers
  const headerWidth = includeFMECheck ? 8 : 6;
  const headerRange = revenueSheet.getRange(1, 10, 1, headerWidth);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#f0f0f0');
  
  if (includeFMECheck) {
    Logger.log('Added matching headers including FME validation columns');
  }
}

function addRevenueConditionalFormatting(revenueSheet, lastRow, includeFMECheck = true) {
  if (lastRow <= 1) return;
  
  const dataRange = revenueSheet.getRange(2, 1, lastRow - 1, 17);
  
  // Clear existing conditional formatting rules
  revenueSheet.setConditionalFormatRules([]);
  
  const rules = [];
  
  // ORANGE rule (highest priority) - FME mismatch (only if FME checking is enabled)
  if (includeFMECheck) {
    const orangeRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($O2="Match",$P2<>"OK",$P2<>"")')
      .setBackground('#ffe0b2')
      .setBold(true)
      .setRanges([dataRange])
      .build();
    rules.push(orangeRule);
  }
  
  // PINK rule - No match
  const pinkRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($B2<>"",$F2<>"",$O2<>"Match")')
    .setBackground('#ffcdd2')
    .setRanges([dataRange])
    .build();
  rules.push(pinkRule);
  
  // DARK GREEN rule (lowest priority) - Match (and FME OK if checking enabled)
  const greenFormula = includeFMECheck 
    ? '=AND($O2="Match",OR($P2="OK",$P2=""))'
    : '=$O2="Match"';
  
  const greenRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(greenFormula)
    .setBackground('#2e7d32')  // DARKER GREEN - easier to distinguish from pink
    .setRanges([dataRange])
    .build();
  rules.push(greenRule);
  
  // Apply rules in priority order
  revenueSheet.setConditionalFormatRules(rules);
  
  const formatStatus = includeFMECheck ? 'with FME mismatch detection (orange)' : 'standard';
  Logger.log(`Applied conditional formatting ${formatStatus} to revenue sheet with DARKER GREEN`);
}

function addPaymentConditionalFormattingWithSpecialRules(paymentSheet, lastRow, includeFMECheck = true) {
  if (lastRow < 15) return;
  
  const dataRange = paymentSheet.getRange(15, 1, lastRow - 14, 40);
  
  // Clear existing conditional formatting rules
  paymentSheet.setConditionalFormatRules([]);
  
  const rules = [];
  
  // ORANGE rule (highest priority if FME checking enabled) - FME mismatch
  if (includeFMECheck) {
    const orangeRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($AL15="Match",$AM15<>"OK",$AM15<>"")')
      .setBackground('#ffe0b2')
      .setBold(true)
      .setRanges([dataRange])
      .build();
    rules.push(orangeRule);
  }
  
  // RED rule - 0 in both R and S columns (failed orders)
  const redRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($R15=0,$S15=0)')
    .setBackground('#ff5252')
    .setBold(true)
    .setRanges([dataRange])
    .build();
  rules.push(redRule);
  
  // LIGHT BLUE rule - 0 in column S only (exam sitting changes)
  const lightBlueRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($S15=0,$R15<>0)')
    .setBackground('#e1f5fe')
    .setRanges([dataRange])
    .build();
  rules.push(lightBlueRule);
  
  // DARK GREEN rule - Has match (and FME OK if checking enabled)
  const greenFormula = includeFMECheck
    ? '=AND($AL15="Match",OR($AM15="OK",$AM15=""))'
    : '=$AL15="Match"';
  
  const greenRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(greenFormula)
    .setBackground('#2e7d32')  // DARKER GREEN - easier to distinguish from pink
    .setRanges([dataRange])
    .build();
  rules.push(greenRule);
  
  // PINK rule (lowest priority) - No match, has values
  const pinkRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$AL15<>"Match"')
    .setBackground('#ffcdd2')
    .setRanges([dataRange])
    .build();
  rules.push(pinkRule);
  
  // Apply rules in priority order (first rule has highest priority)
  paymentSheet.setConditionalFormatRules(rules);
  
  const fmeStatus = includeFMECheck ? 'with FME validation' : '';
  Logger.log(`Applied conditional formatting with special rules ${fmeStatus} to payment sheet with DARKER GREEN`);
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

function updateExistingMonthlySheetFormulas(monthName, year, includeFMECheck = true) {
  const fmeStatus = includeFMECheck ? ' with FME validation' : '';
  Logger.log(`Updating ${monthName} ${year}${fmeStatus}...`);
  
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
    // Add headers
    addMatchingHeaders(revenueSheet, includeFMECheck);
    
    // Add formulas
    addRevenueMatchingFormulas(revenueSheet, paymentSheetName, lastRow, includeFMECheck);
    
    // Add conditional formatting
    addRevenueConditionalFormatting(revenueSheet, lastRow, includeFMECheck);
    
    const updateStatus = includeFMECheck ? 'Name+Amount matching with FME validation' : 'Name+Amount matching';
    Logger.log(`✅ Updated ${revenueSheetName} with ${updateStatus}`);
  } else {
    Logger.log(`❌ No data in ${revenueSheetName} to update`);
  }
  
  // Also update Payment Reconciliation sheet
  try {
    const PAYMENT_SPREADSHEET_ID = "1vi9n0mTXxGTGAppecDk04krBqgqvGf4ZNTvVlhseyFY";
    const paymentSpreadsheet = SpreadsheetApp.openById(PAYMENT_SPREADSHEET_ID);
    const paymentSheet = paymentSpreadsheet.getSheetByName(paymentSheetName);
    
    if (paymentSheet) {
      const paymentLastRow = paymentSheet.getLastRow();
      if (paymentLastRow >= 15) {
        addPaymentMatchingFormulas(paymentSheet, revenueSheetName, paymentLastRow, includeFMECheck);
        addPaymentConditionalFormattingWithSpecialRules(paymentSheet, paymentLastRow, includeFMECheck);
        Logger.log(`✅ Updated ${paymentSheetName} in Payment Reconciliation with ${updateStatus}`);
      }
    }
  } catch (error) {
    Logger.log(`Note: Could not update Payment Reconciliation sheet: ${error.toString()}`);
  }
}

// ===============================================
// FME DIAGNOSTIC FUNCTIONS
// ===============================================

function diagnoseFMEMismatches(monthName, year) {
  Logger.log(`🔍 DIAGNOSING FME MISMATCHES FOR ${monthName} ${year}`);
  Logger.log('='.repeat(60));
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const revenueSheetName = `${monthName} ${year}`;
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
  
  const allData = revenueSheet.getRange(1, 1, lastRow, 17).getValues();
  const headers = allData[0];
  const dataRows = allData.slice(1);
  
  const nameCol = 1; // Column B
  const actualPriceCol = 5; // Column F
  const calculatedFMECol = 7; // Column H
  const matchCol = 14; // Column O
  const fmeCheckCol = 15; // Column P
  const actualFMECol = 16; // Column Q
  
  let mismatches = [];
  let matches = 0;
  
  dataRows.forEach((row, index) => {
    const rowNumber = index + 2;
    const name = row[nameCol];
    const match = row[matchCol];
    const fmeCheck = row[fmeCheckCol];
    
    if (!name) return;
    
    if (match === "Match") {
      if (fmeCheck === "OK" || fmeCheck === "") {
        matches++;
      } else {
        // This is a mismatch - anything that's not "OK" or empty
        const calculatedFME = row[calculatedFMECol];
        const actualFME = row[actualFMECol];
        const difference = fmeCheck; // Already formatted as "£X.XX"
        
        mismatches.push({
          rowNumber: rowNumber,
          name: name,
          amount: row[actualPriceCol],
          calculatedFME: calculatedFME,
          actualFME: actualFME,
          difference: difference
        });
      }
    }
  });
  
  Logger.log(`\n📊 SUMMARY:`);
  Logger.log(`✅ FME Matches: ${matches}`);
  Logger.log(`⚠️ FME Mismatches: ${mismatches.length}`);
  
  if (mismatches.length > 0) {
    Logger.log(`\n🔴 FME MISMATCHES FOUND:\n`);
    mismatches.forEach((item, index) => {
      Logger.log(`${index + 1}. Row ${item.rowNumber}: ${item.name}`);
      Logger.log(`   Amount: £${item.amount}`);
      Logger.log(`   Calculated FME: £${item.calculatedFME}`);
      Logger.log(`   Actual FME: £${item.actualFME}`);
      Logger.log(`   Difference: ${item.difference}`);
      Logger.log('');
    });
    
    Logger.log(`💡 These rows will be highlighted in ORANGE in your sheet`);
  } else {
    Logger.log(`\n✅ No FME mismatches found!`);
  }
}

function explainFMELogic() {
  Logger.log('=== FME VALIDATION LOGIC EXPLANATION ===\n');
  
  Logger.log('The FME Check works on BOTH sheets:\n');
  
  Logger.log('📊 REVENUE AUTOMATED (Column P):');
  Logger.log('  Compares calculated FME (H) vs actual FME from Payment Rec (Y)\n');
  
  Logger.log('📊 PAYMENT RECONCILIATION (Column AM):');
  Logger.log('  Compares actual FME (Y) vs calculated FME from Revenue (H)\n');
  
  Logger.log('Logic for both sheets:\n');
  
  Logger.log('1. If BOTH sides are empty or 0:');
  Logger.log('   → Shows "OK" ✅ (Green)\n');
  
  Logger.log('2. If ONE side has a value and the other is 0/empty:');
  Logger.log('   → Shows "£X.XX" (the difference)');
  Logger.log('   → Row is highlighted in ORANGE ⚠️\n');
  
  Logger.log('3. If BOTH have values and match (within 1p):');
  Logger.log('   → Shows "OK" ✅ (Green)\n');
  
  Logger.log('4. If BOTH have values but DON\'T match:');
  Logger.log('   → Shows "£X.XX" (the difference)');
  Logger.log('   → Row is highlighted in ORANGE ⚠️\n');
  
  Logger.log('EXAMPLES:\n');
  Logger.log('Revenue H | Payment Y | Result        | Color');
  Logger.log('----------|-----------|---------------|-------');
  Logger.log('0 or ""   | 0 or ""   | OK            | Green');
  Logger.log('104.70    | 104.70    | OK            | Green');
  Logger.log('104.69    | 104.70    | OK            | Green (within 1p)');
  Logger.log('0 or ""   | 104.70    | £104.70       | ORANGE ⚠️');
  Logger.log('104.70    | 0 or ""   | £-104.70      | ORANGE ⚠️');
  Logger.log('100.00    | 104.70    | £4.70         | ORANGE ⚠️');
  Logger.log('110.00    | 104.70    | £-5.30        | ORANGE ⚠️\n');
  
  Logger.log('New Payment Reconciliation columns:');
  Logger.log('  Column AM: FME Check');
  Logger.log('  Column AN: Calculated FME (from Revenue Automated)');
}

// ===============================================
// TEST AND HELPER FUNCTIONS
// ===============================================

function testMonthlyMatching() {
  Logger.log('Testing monthly matching setup for July 2025...');
  setupMonthlyMatching('July', 2025);
  Logger.log('Test complete - check both spreadsheets!');
}

function testMonthlyMatchingWithoutFME() {
  Logger.log('Testing monthly matching WITHOUT FME validation...');
  setupMonthlyMatching('July', 2025, false);
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

function setupOctober2025() {
  setupMonthlyMatching('October', 2025);
}

function setupDecember2025() {
  setupMonthlyMatching('December', 2025);
}

function setupJanuary2025() {
  setupMonthlyMatching('January', 2025);
}

// Convenient wrapper functions for updating existing sheets
function updateDecember2025() {
  updateExistingMonthlySheetFormulas('December', 2025);
}

function diagnoseDecember2025FME() {
  diagnoseFMEMismatches('December', 2025);
}

function updateJanuary2025() {
  updateExistingMonthlySheetFormulas('January', 2025);
}

function updateJanuary2025WithoutFME() {
  updateExistingMonthlySheetFormulas('January', 2025, false);
}

function diagnoseJanuary2025FME() {
  diagnoseFMEMismatches('January', 2025);
}

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
    const fmeCheck = revenueSheet.getRange(row, 16).getValue(); // Column P
    
    Logger.log(`Row ${row}:`);
    Logger.log(`  Name: ${name}`);
    Logger.log(`  Amount: £${amount}`);
    Logger.log(`  Match Status: ${match}`);
    Logger.log(`  Stripe Fee: £${stripeFee}`);
    Logger.log(`  FME Check: ${fmeCheck}`);
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