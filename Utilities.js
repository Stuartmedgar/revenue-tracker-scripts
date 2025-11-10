// ===============================================
// UTILITIES.GS - Helper Functions
// ===============================================

function getMonthName(monthIndex) {
  const months = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
  ];
  return months[monthIndex];
}

function deleteRowFromDataEntry(sheet, rowIndex) {
  try {
    sheet.deleteRow(rowIndex);
    Logger.log(`Deleted row ${rowIndex} from Data Entry sheet`);
  } catch (error) {
    Logger.log(`Error deleting row ${rowIndex}: ${error.toString()}`);
  }
}

// ===============================================
// TAB POSITIONING FUNCTIONS
// ===============================================

function positionSortingTab(ss, sheet) {
  // Position after Data Entry (index 1)
  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(2);
}

function positionAwaitingInvoiceTab(ss, sheet) {
  // Position after Sort (index 2), before Failed orders
  const sortSheet = ss.getSheetByName('Sort');
  const targetPosition = sortSheet ? 3 : 2;
  
  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(targetPosition);
}

function positionInstalmentTrackerTab(ss, sheet) {
  // Position after Awaiting Employer Invoice, before Failed orders
  const awaitingSheet = ss.getSheetByName('Awaiting Employer Invoice');
  const targetPosition = awaitingSheet ? 4 : 3;
  
  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(targetPosition);
}

function positionFailedOrdersTab(ss, sheet) {
  // Position after Instalment Tracker (index 4)
  // If Instalment Tracker doesn't exist, position after Awaiting Employer Invoice
  const instalmentSheet = ss.getSheetByName('Instalment Tracker');
  const awaitingSheet = ss.getSheetByName('Awaiting Employer Invoice');
  
  let targetPosition;
  if (instalmentSheet) {
    targetPosition = 5;
  } else if (awaitingSheet) {
    targetPosition = 4;
  } else {
    targetPosition = 3;
  }
  
  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(targetPosition);
}

function positionMonthlySheetTab(ss, sheet, date) {
  // Find position for monthly sheet - should be after Instalment Tracker
  // and ordered by date (newest month closest to Instalment Tracker)
  
  const allSheets = ss.getSheets();
  const monthlySheets = [];
  let instalmentTrackerPosition = -1;
  
  // Find existing monthly sheets and Instalment Tracker position
  allSheets.forEach((s, index) => {
    const name = s.getName();
    if (name === 'Instalment Tracker') {
      instalmentTrackerPosition = index;
    } else if (isMonthlySheetName(name)) {
      monthlySheets.push({
        sheet: s,
        position: index,
        date: parseMonthlySheetDate(name)
      });
    }
  });
  
  // Sort monthly sheets by date (newest first)
  monthlySheets.sort((a, b) => b.date.getTime() - a.date.getTime());
  
  // Find correct position for new sheet
  let targetPosition;
  if (instalmentTrackerPosition === -1) {
    // No Instalment Tracker yet, position at end
    targetPosition = allSheets.length;
  } else {
    // Find where this sheet should go among monthly sheets
    let insertIndex = 0;
    for (let i = 0; i < monthlySheets.length; i++) {
      if (date.getTime() > monthlySheets[i].date.getTime()) {
        break;
      }
      insertIndex = i + 1;
    }
    
    if (insertIndex === 0) {
      // This is the newest month, position right after Instalment Tracker
      targetPosition = instalmentTrackerPosition + 2;
    } else {
      // Position after the appropriate monthly sheet
      targetPosition = monthlySheets[insertIndex - 1].position + 2;
    }
  }
  
  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(targetPosition);
}

function isMonthlySheetName(name) {
  // Check if name matches pattern "Month Year" (e.g., "January 2025")
  const monthNames = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
  ];
  
  const parts = name.split(' ');
  if (parts.length !== 2) return false;
  
  const month = parts[0];
  const year = parseInt(parts[1]);
  
  return monthNames.includes(month) && !isNaN(year) && year > 2000;
}

function parseMonthlySheetDate(name) {
  const parts = name.split(' ');
  const monthName = parts[0];
  const year = parseInt(parts[1]);
  
  const monthNames = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
  ];
  
  const monthIndex = monthNames.indexOf(monthName);
  return new Date(year, monthIndex, 1);
}