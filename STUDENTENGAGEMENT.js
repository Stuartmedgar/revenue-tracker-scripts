// ===============================================
// STUDENTENGAGEMENT.GS - Complete Student Engagement with Tuition/Revision Plus Support
// UPDATED: Added £822 and £522 to engagement criteria
// ===============================================

const ATX_ENGAGEMENT_SPREADSHEET_ID = "1mhtp-bFZe-mFMKStBXE9ECgGj9T8K-xAZOJ511jpOH8";

// ===============================================
// MAIN TRANSFER FUNCTIONS
// ===============================================

function processStudentForEngagement(studentData, monthlySheetName) {
  try {
    Logger.log(`Processing student for engagement transfer: ${studentData.name}`);

    // Check if student meets criteria
    if (!meetsEngagementCriteria(studentData.actualPrice)) {
      Logger.log(`Student ${studentData.name} does not meet engagement criteria (price: ${studentData.actualPrice})`);
      return;
    }

    // Get course type from price
    const courseType = getCourseTypeFromPrice(studentData.actualPrice);

    // Normalize sitting name for sheet matching
    const targetSheetName = normalizeSittingName(studentData.sitting);

    if (!targetSheetName) {
      Logger.log(`Could not normalize sitting name: "${studentData.sitting}"`);
      return;
    }

    // Transfer to ATX Engagement spreadsheet
    transferStudentToEngagement(studentData.name, courseType, targetSheetName);

    Logger.log(`Successfully processed ${studentData.name} for ${targetSheetName} - ${courseType}`);

  } catch (error) {
    Logger.log(`Error processing student for engagement: ${error.toString()}`);
  }
}

function meetsEngagementCriteria(actualPrice) {
  const price = Number(actualPrice);

  // UPDATED: Added 822 and 522 for Tuition/Revision Plus
  // Full course payments
  if (price === 997 || price === 822 || price === 647 || price === 597) {
    return true;
  }

  // First instalment payments
  if (price === 522 || price === 397 || price === 347 || price === 297) {
    return true;
  }

  return false;
}

function getCourseTypeFromPrice(actualPrice) {
  const price = Number(actualPrice);

  // Platinum (997 full, 397 first instalment)
  if (price === 997 || price === 397) {
    return 'Platinum';
  }

  // UPDATED: Tuition/Revision Plus (822 full, 522 first instalment)
  if (price === 822 || price === 522) {
    return 'Tuition/Revision Plus';
  }

  // Revision (647 full, 347 first instalment)
  if (price === 647 || price === 347) {
    return 'Revision';
  }

  // Tuition (597 full, 297 first instalment)
  if (price === 597 || price === 297) {
    return 'Tuition';
  }

  return '';
}

function normalizeSittingName(sitting) {
  if (!sitting) return null;

  const sittingStr = sitting.toString().trim();

  // Handle various date formats and normalize to "Month YYYY" format
  // Expected format is already "December 2025" but handle variations

  // Remove extra spaces and normalize case
  let normalized = sittingStr.replace(/\s+/g, ' ').trim();

  // Convert month abbreviations to full names
  const monthMap = {
    'Jan': 'January', 'Feb': 'February', 'Mar': 'March', 'Apr': 'April',
    'May': 'May', 'Jun': 'June', 'Jul': 'July', 'Aug': 'August',
    'Sep': 'September', 'Oct': 'October', 'Nov': 'November', 'Dec': 'December'
  };

  // Handle formats like "Dec 2025" -> "December 2025"
  for (const [abbrev, full] of Object.entries(monthMap)) {
    const regex = new RegExp(`\\b${abbrev}\\b`, 'i');
    if (regex.test(normalized)) {
      normalized = normalized.replace(regex, full);
      break;
    }
  }

  // Handle formats like "December2025" -> "December 2025"
  normalized = normalized.replace(/([A-Za-z]+)(\d{4})/, '$1 $2');

  // Handle formats with dashes like "December-2025" -> "December 2025"
  normalized = normalized.replace(/([A-Za-z]+)-(\d{4})/, '$1 $2');

  // Capitalize first letter of month
  normalized = normalized.charAt(0).toUpperCase() + normalized.slice(1).toLowerCase();

  // Final format check - should be "Month YYYY"
  const formatMatch = normalized.match(/^([A-Z][a-z]+)\s+(\d{4})$/);
  if (!formatMatch) {
    Logger.log(`Warning: Could not parse sitting format: "${sittingStr}" -> "${normalized}"`);
    return null;
  }

  const [, month, year] = formatMatch;
  const finalName = month + ' ' + year;

  Logger.log(`Normalized sitting: "${sittingStr}" -> "${finalName}"`);
  return finalName;
}

function transferStudentToEngagement(studentName, courseType, targetSheetName) {
  try {
    const engagementSS = SpreadsheetApp.openById(ATX_ENGAGEMENT_SPREADSHEET_ID);
    let targetSheet = engagementSS.getSheetByName(targetSheetName);

    if (!targetSheet) {
      Logger.log(`Target sheet "${targetSheetName}" not found in ATX Engagement spreadsheet`);
      Logger.log(`Available sheets: ${engagementSS.getSheets().map(s => s.getName()).join(', ')}`);
      return;
    }

    // Check for duplicates BEFORE finding insertion point
    const isDuplicate = checkForDuplicate(targetSheet, studentName);

    // Find where to insert the new student (before Deferrals section)
    const insertRow = findFirstAvailableRow(targetSheet);

    // INSERT a new row at the correct position (this pushes Deferrals down)
    targetSheet.insertRowBefore(insertRow);

    // Add student name to column B of the newly inserted row
    targetSheet.getRange(insertRow, 2).setValue(studentName);

    // Set course type dropdown in column C
    setCourseTypeDropdown(targetSheet, insertRow, courseType);

    // Highlight duplicate in red if needed
    if (isDuplicate) {
      const rowRange = targetSheet.getRange(insertRow, 1, 1, targetSheet.getLastColumn());
      rowRange.setBackground('#ffcdd2'); // Light red background
      Logger.log(`Added duplicate student ${studentName} to ${targetSheetName} and highlighted in red`);
    } else {
      Logger.log(`Added student ${studentName} to ${targetSheetName} - ${courseType}`);
    }

  } catch (error) {
    Logger.log(`Error transferring student to engagement sheet: ${error.toString()}`);
  }
}

function checkForDuplicate(sheet, studentName) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 9) return false; // No data below headers

  // Check column B (names) from row 9 onwards, but stop at Deferrals section
  const columnBData = sheet.getRange(9, 2, lastRow - 8, 1).getValues();

  for (let i = 0; i < columnBData.length; i++) {
    const cellValue = columnBData[i][0];

    // Stop checking when we hit the Deferrals section
    if (cellValue && cellValue.toString().toLowerCase().includes('deferral')) {
      break;
    }

    // Check for duplicate student name
    if (cellValue === studentName) {
      return true;
    }
  }

  return false;
}

function findFirstAvailableRow(sheet) {
  const lastRow = sheet.getLastRow();

  // Start checking from row 9 (first data row after headers in row 8)
  let searchStartRow = Math.max(9, 1);

  // Get all data in column B (names) to analyze the sheet structure
  const columnBData = sheet.getRange(searchStartRow, 2, lastRow - searchStartRow + 1, 1).getValues();

  // Find the last student row before any "Deferrals" section
  let lastStudentRow = searchStartRow - 1; // Start before first data row

  for (let i = 0; i < columnBData.length; i++) {
    const rowIndex = searchStartRow + i;
    const cellValue = columnBData[i][0];

    // If we hit "Deferrals" heading, stop looking - we found where students should end
    if (cellValue && cellValue.toString().toLowerCase().includes('deferral')) {
      Logger.log(`Found "Deferrals" section at row ${rowIndex}, students should be added before this`);
      break;
    }

    // If this row has a student name (non-empty), update last student row
    if (cellValue && cellValue.toString().trim() !== '') {
      lastStudentRow = rowIndex;
    }
  }

  // The next available row is right after the last student
  const nextAvailableRow = lastStudentRow + 1;

  Logger.log(`Last student found at row ${lastStudentRow}, inserting new student at row ${nextAvailableRow}`);
  return nextAvailableRow;
}

function setCourseTypeDropdown(sheet, row, courseType) {
  const cell = sheet.getRange(row, 3); // Column C

  // Set the value (this should work with existing dropdown validation)
  cell.setValue(courseType);

  Logger.log(`Set course type to ${courseType} in row ${row}`);
}

// ===============================================
// MANUAL PROCESSING FUNCTIONS
// ===============================================

function processMonthlySheetForEngagement(monthName, year) {
  try {
    Logger.log(`Processing ${monthName} ${year} for student engagement transfers...`);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = `${monthName} ${year}`;
    const monthlySheet = ss.getSheetByName(sheetName);

    if (!monthlySheet) {
      Logger.log(`Monthly sheet "${sheetName}" not found`);
      return;
    }

    const dataRange = monthlySheet.getDataRange();
    if (dataRange.getNumRows() <= 1) {
      Logger.log(`No data in ${sheetName}`);
      return;
    }

    const allData = dataRange.getValues();
    const headers = allData[0];
    const dataRows = allData.slice(1);

    // Find column indices
    const nameCol = headers.indexOf('Name');
    const sittingCol = headers.indexOf('Sitting');
    const actualPriceCol = headers.indexOf('Actual Price');
    const courseCol = headers.indexOf('Course');

    if (nameCol === -1 || sittingCol === -1 || actualPriceCol === -1) {
      Logger.log(`Missing required columns in ${sheetName}`);
      return;
    }

    let processedCount = 0;

    // Process each row
    dataRows.forEach((row, index) => {
      const studentData = {
        name: row[nameCol],
        sitting: row[sittingCol],
        actualPrice: row[actualPriceCol],
        course: row[courseCol] || getCourseTypeFromPrice(row[actualPriceCol])
      };

      if (studentData.name && studentData.sitting && studentData.actualPrice) {
        if (meetsEngagementCriteria(studentData.actualPrice)) {
          processStudentForEngagement(studentData, sheetName);
          processedCount++;
        }
      }
    });

    Logger.log(`Processed ${processedCount} students from ${sheetName} for engagement transfer`);

  } catch (error) {
    Logger.log(`Error processing monthly sheet for engagement: ${error.toString()}`);
  }
}

// ===============================================
// CONVENIENT WRAPPER FUNCTIONS
// ===============================================

function processJuly2025ForEngagement() {
  processMonthlySheetForEngagement('July', 2025);
}

function processJune2025ForEngagement() {
  processMonthlySheetForEngagement('June', 2025);
}

function processAugust2025ForEngagement() {
  processMonthlySheetForEngagement('August', 2025);
}

function testEngagementTransfer() {
  Logger.log('=== TESTING ENGAGEMENT TRANSFER ===');

  // Test normalization
  const testSittings = ['December 2025', 'Dec 2025', 'December-2025', 'december 2025'];
  testSittings.forEach(sitting => {
    const normalized = normalizeSittingName(sitting);
    Logger.log(`"${sitting}" -> "${normalized}"`);
  });

  // UPDATED: Test criteria with new prices
  const testPrices = [997, 822, 647, 597, 522, 397, 347, 297, 300, 500];
  testPrices.forEach(price => {
    const meets = meetsEngagementCriteria(price);
    const course = getCourseTypeFromPrice(price);
    Logger.log(`Price ${price}: Meets criteria: ${meets}, Course: ${course}`);
  });
}

function debugEngagementSheets() {
  Logger.log('=== ATX ENGAGEMENT SPREADSHEET SHEETS ===');

  try {
    const engagementSS = SpreadsheetApp.openById(ATX_ENGAGEMENT_SPREADSHEET_ID);
    const sheets = engagementSS.getSheets();

    Logger.log(`Found ${sheets.length} sheets:`);
    sheets.forEach(sheet => {
      Logger.log(`- ${sheet.getName()}`);
    });

  } catch (error) {
    Logger.log(`Error accessing ATX Engagement spreadsheet: ${error.toString()}`);
  }
}

// ===============================================
// TESTING AND DEBUGGING FUNCTIONS
// ===============================================

function testRowPlacement() {
  Logger.log('=== TESTING IMPROVED ROW PLACEMENT ===');

  try {
    const engagementSS = SpreadsheetApp.openById(ATX_ENGAGEMENT_SPREADSHEET_ID);
    const testSheet = engagementSS.getSheetByName('December 2025');

    if (!testSheet) {
      Logger.log('December 2025 sheet not found for testing');
      return;
    }

    Logger.log('🔍 Analyzing December 2025 sheet structure:');

    // Get all data in column B to see the structure
    const lastRow = testSheet.getLastRow();
    const columnBData = testSheet.getRange(1, 2, lastRow, 1).getValues();

    let foundDeferrals = false;
    let deferralsRow = -1;
    let lastStudentRow = -1;

    for (let i = 0; i < columnBData.length; i++) {
      const rowIndex = i + 1;
      const cellValue = columnBData[i][0];

      if (cellValue) {
        Logger.log(`Row ${rowIndex}: "${cellValue}"`);

        if (cellValue.toString().toLowerCase().includes('deferral')) {
          foundDeferrals = true;
          deferralsRow = rowIndex;
          Logger.log(`  ^^^ DEFERRALS SECTION FOUND`);
        } else if (rowIndex >= 9 && !foundDeferrals) {
          // This looks like a student name
          lastStudentRow = rowIndex;
        }
      }
    }

    Logger.log(`\n📊 Analysis Results:`);
    Logger.log(`Last student row: ${lastStudentRow}`);
    Logger.log(`Deferrals section at: ${foundDeferrals ? deferralsRow : 'Not found'}`);

    const nextRow = findFirstAvailableRow(testSheet);
    Logger.log(`Next student should be inserted at row: ${nextRow}`);

    if (foundDeferrals && nextRow >= deferralsRow) {
      Logger.log('❌ WARNING: New student would be placed after Deferrals - this is the problem!');
    } else {
      Logger.log('✅ New student will be placed correctly before Deferrals section');
    }

  } catch (error) {
    Logger.log(`Error testing row placement: ${error.toString()}`);
  }
}

function testImprovedTransfer() {
  Logger.log('=== TESTING IMPROVED TRANSFER PLACEMENT ===');

  const testStudent = {
    name: 'Test Student - Row Placement Fix',
    sitting: 'December 2025',
    actualPrice: 397,
    course: 'Platinum'
  };

  Logger.log(`Testing with: ${testStudent.name}`);
  Logger.log('This should be placed BEFORE the Deferrals section');

  processStudentForEngagement(testStudent, 'Test Sheet');

  Logger.log('✅ Test complete - check the December 2025 sheet to see if placement is correct');
}

function transferAllEligibleStudentsToEngagement() {
  Logger.log('=== TRANSFERRING ALL ELIGIBLE STUDENTS TO ENGAGEMENT ===');
  Logger.log('⚠️ This will attempt to transfer ALL eligible students from monthly sheets');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();
  const monthlySheets = allSheets.filter(sheet => isMonthlySheetName(sheet.getName()));

  let totalTransferred = 0;

  monthlySheets.forEach(sheet => {
    const sheetName = sheet.getName();
    Logger.log(`\n📋 Processing ${sheetName}...`);

    const dataRange = sheet.getDataRange();
    if (dataRange.getNumRows() <= 1) return;

    const allData = dataRange.getValues();
    const headers = allData[0];
    const dataRows = allData.slice(1);

    const nameCol = headers.indexOf('Name');
    const sittingCol = headers.indexOf('Sitting');
    const actualPriceCol = headers.indexOf('Actual Price');

    if (nameCol === -1 || sittingCol === -1 || actualPriceCol === -1) return;

    dataRows.forEach((row, index) => {
      const name = row[nameCol];
      const sitting = row[sittingCol];
      const actualPrice = row[actualPriceCol];

      if (name && sitting && actualPrice && meetsEngagementCriteria(actualPrice)) {
        const studentData = {
          name: name,
          sitting: sitting,
          actualPrice: actualPrice,
          course: getCourseTypeFromPrice(actualPrice)
        };

        Logger.log(`  🎓 Transferring: ${name} (£${actualPrice})`);
        processStudentForEngagement(studentData, sheetName);
        totalTransferred++;
      }
    });
  });

  Logger.log(`\n✅ COMPLETE: Attempted to transfer ${totalTransferred} students to engagement spreadsheet`);
}