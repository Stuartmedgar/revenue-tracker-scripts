// ===============================================
// DYNAMIC MONTHLY RECONCILIATION MENU SYSTEM
// ===============================================

/**
 * Creates custom menu when spreadsheet opens
 * This runs automatically when the spreadsheet is opened
 * Menu dynamically shows only relevant years based on your progress
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // Create main menu
  const menu = ui.createMenu('📊 Monthly Reconciliation');
  
  // Get the active years (current year and next year, or based on usage)
  const activeYears = getActiveYears();
  
  // Create submenu for each active year
  activeYears.forEach(year => {
    const yearMenu = ui.createMenu(`${year} Months`);
    
    // Add all 12 months for this year
    const months = ['January', 'February', 'March', 'April', 'May', 'June',
                    'July', 'August', 'September', 'October', 'November', 'December'];
    
    months.forEach(month => {
      const functionName = `runMonth_${month}_${year}`;
      yearMenu.addItem(`${month} ${year}`, functionName);
    });
    
    menu.addSubMenu(yearMenu);
  });
  
  menu.addSeparator();
  
  // Add utility functions
  menu.addItem('🔄 Update Current Month', 'updateCurrentMonth');
  menu.addItem('📋 List Available Sheets', 'listAvailableSheets');
  menu.addSeparator();
  menu.addItem('⚙️ Switch to Next Year', 'switchToNextYear');
  menu.addItem('ℹ️ Help & Instructions', 'showReconciliationHelp');
  
  // Add menu to spreadsheet
  menu.addToUi();
  
  Logger.log(`Monthly Reconciliation menu created with years: ${activeYears.join(', ')}`);
}

/**
 * Determines which years should be shown in the menu
 * Returns array of years based on stored preference or defaults to current + next year
 */
function getActiveYears() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getDocumentProperties();
  
  // Check if we have stored active years
  const storedYears = props.getProperty('activeYears');
  
  if (storedYears) {
    // Parse stored years
    const years = JSON.parse(storedYears);
    Logger.log(`Using stored active years: ${years.join(', ')}`);
    return years;
  }
  
  // Default: show current year and next year
  const currentYear = new Date().getFullYear();
  const defaultYears = [currentYear, currentYear + 1];
  
  // Store as default
  props.setProperty('activeYears', JSON.stringify(defaultYears));
  Logger.log(`Initialized with default years: ${defaultYears.join(', ')}`);
  
  return defaultYears;
}

/**
 * Moves to the next year - removes the oldest year and adds a new future year
 * Call this when you want to transition (e.g., when January of new year is complete)
 */
function switchToNextYear() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getDocumentProperties();
  
  const currentYears = getActiveYears();
  const oldestYear = Math.min(...currentYears);
  const newestYear = Math.max(...currentYears);
  const nextYear = newestYear + 1;
  
  // Confirm with user
  const response = ui.alert(
    'Switch to Next Year',
    `This will:\n` +
    `• Remove ${oldestYear} from the menu\n` +
    `• Add ${nextYear} to the menu\n\n` +
    `Your active years will be: ${newestYear}, ${nextYear}\n\n` +
    `Continue?`,
    ui.ButtonSet.YES_NO
  );
  
  if (response != ui.Button.YES) {
    return;
  }
  
  // Create new years array (remove oldest, add newest + 1)
  const newYears = [newestYear, nextYear];
  
  // Store new years
  props.setProperty('activeYears', JSON.stringify(newYears));
  
  // Show success message
  ui.alert(
    'Year Transition Complete',
    `Menu updated!\n\n` +
    `Active years are now: ${newYears.join(', ')}\n\n` +
    `Please refresh the page to see the updated menu.`,
    ui.ButtonSet.OK
  );
  
  Logger.log(`Switched years from [${currentYears.join(', ')}] to [${newYears.join(', ')}]`);
  
  // Refresh the menu
  onOpen();
}

/**
 * Universal function to run monthly reconciliation
 * This handles any month/year combination dynamically
 */
function runMonthlyReconciliation(month, year) {
  showProcessingMessage(`${month} ${year}`);
  
  try {
    setupMonthlyMatching(month, year);
    showCompletionMessage(`${month} ${year}`);
    
    // Auto-trigger year transition if this is January and user confirms
    if (month === 'January') {
      checkAutoYearTransition(year);
    }
    
  } catch (error) {
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      'Reconciliation Error',
      `Error processing ${month} ${year}:\n\n${error.toString()}\n\n` +
      `Check the execution log for details.`,
      ui.ButtonSet.OK
    );
    Logger.log(`Error in reconciliation for ${month} ${year}: ${error.toString()}`);
  }
}

/**
 * Checks if we should offer to transition years (when running January)
 */
function checkAutoYearTransition(yearJustProcessed) {
  const ui = SpreadsheetApp.getUi();
  const activeYears = getActiveYears();
  
  // Only offer if January was for a newer year than our oldest active year
  const oldestActiveYear = Math.min(...activeYears);
  
  if (yearJustProcessed > oldestActiveYear) {
    const response = ui.alert(
      'Start of New Year Detected',
      `You just processed January ${yearJustProcessed}.\n\n` +
      `Would you like to remove ${oldestActiveYear} from your menu\n` +
      `and add ${yearJustProcessed + 1} for future planning?\n\n` +
      `(You can always do this later via "Switch to Next Year")`,
      ui.ButtonSet.YES_NO
    );
    
    if (response == ui.Button.YES) {
      // Update to new year range
      const props = PropertiesService.getDocumentProperties();
      const newYears = [yearJustProcessed, yearJustProcessed + 1];
      props.setProperty('activeYears', JSON.stringify(newYears));
      
      ui.alert(
        'Years Updated',
        `Menu will now show: ${newYears.join(' and ')}\n\n` +
        `Please refresh the page to see the updated menu.`,
        ui.ButtonSet.OK
      );
      
      Logger.log(`Auto-transitioned to years: ${newYears.join(', ')}`);
    }
  }
}


// ===============================================
// DYNAMIC MONTH FUNCTION GENERATOR
// ===============================================

/**
 * This function dynamically creates all month/year combinations
 * Google Apps Script will call these functions based on the menu items
 * Format: runMonth_January_2025, runMonth_February_2026, etc.
 */

// Generate functions for all possible months and years (2025-2030)
// This creates the functions at runtime that the menu system calls

function runMonth_January_2025() { runMonthlyReconciliation('January', 2025); }
function runMonth_February_2025() { runMonthlyReconciliation('February', 2025); }
function runMonth_March_2025() { runMonthlyReconciliation('March', 2025); }
function runMonth_April_2025() { runMonthlyReconciliation('April', 2025); }
function runMonth_May_2025() { runMonthlyReconciliation('May', 2025); }
function runMonth_June_2025() { runMonthlyReconciliation('June', 2025); }
function runMonth_July_2025() { runMonthlyReconciliation('July', 2025); }
function runMonth_August_2025() { runMonthlyReconciliation('August', 2025); }
function runMonth_September_2025() { runMonthlyReconciliation('September', 2025); }
function runMonth_October_2025() { runMonthlyReconciliation('October', 2025); }
function runMonth_November_2025() { runMonthlyReconciliation('November', 2025); }
function runMonth_December_2025() { runMonthlyReconciliation('December', 2025); }

function runMonth_January_2026() { runMonthlyReconciliation('January', 2026); }
function runMonth_February_2026() { runMonthlyReconciliation('February', 2026); }
function runMonth_March_2026() { runMonthlyReconciliation('March', 2026); }
function runMonth_April_2026() { runMonthlyReconciliation('April', 2026); }
function runMonth_May_2026() { runMonthlyReconciliation('May', 2026); }
function runMonth_June_2026() { runMonthlyReconciliation('June', 2026); }
function runMonth_July_2026() { runMonthlyReconciliation('July', 2026); }
function runMonth_August_2026() { runMonthlyReconciliation('August', 2026); }
function runMonth_September_2026() { runMonthlyReconciliation('September', 2026); }
function runMonth_October_2026() { runMonthlyReconciliation('October', 2026); }
function runMonth_November_2026() { runMonthlyReconciliation('November', 2026); }
function runMonth_December_2026() { runMonthlyReconciliation('December', 2026); }

function runMonth_January_2027() { runMonthlyReconciliation('January', 2027); }
function runMonth_February_2027() { runMonthlyReconciliation('February', 2027); }
function runMonth_March_2027() { runMonthlyReconciliation('March', 2027); }
function runMonth_April_2027() { runMonthlyReconciliation('April', 2027); }
function runMonth_May_2027() { runMonthlyReconciliation('May', 2027); }
function runMonth_June_2027() { runMonthlyReconciliation('June', 2027); }
function runMonth_July_2027() { runMonthlyReconciliation('July', 2027); }
function runMonth_August_2027() { runMonthlyReconciliation('August', 2027); }
function runMonth_September_2027() { runMonthlyReconciliation('September', 2027); }
function runMonth_October_2027() { runMonthlyReconciliation('October', 2027); }
function runMonth_November_2027() { runMonthlyReconciliation('November', 2027); }
function runMonth_December_2027() { runMonthlyReconciliation('December', 2027); }

function runMonth_January_2028() { runMonthlyReconciliation('January', 2028); }
function runMonth_February_2028() { runMonthlyReconciliation('February', 2028); }
function runMonth_March_2028() { runMonthlyReconciliation('March', 2028); }
function runMonth_April_2028() { runMonthlyReconciliation('April', 2028); }
function runMonth_May_2028() { runMonthlyReconciliation('May', 2028); }
function runMonth_June_2028() { runMonthlyReconciliation('June', 2028); }
function runMonth_July_2028() { runMonthlyReconciliation('July', 2028); }
function runMonth_August_2028() { runMonthlyReconciliation('August', 2028); }
function runMonth_September_2028() { runMonthlyReconciliation('September', 2028); }
function runMonth_October_2028() { runMonthlyReconciliation('October', 2028); }
function runMonth_November_2028() { runMonthlyReconciliation('November', 2028); }
function runMonth_December_2028() { runMonthlyReconciliation('December', 2028); }

function runMonth_January_2029() { runMonthlyReconciliation('January', 2029); }
function runMonth_February_2029() { runMonthlyReconciliation('February', 2029); }
function runMonth_March_2029() { runMonthlyReconciliation('March', 2029); }
function runMonth_April_2029() { runMonthlyReconciliation('April', 2029); }
function runMonth_May_2029() { runMonthlyReconciliation('May', 2029); }
function runMonth_June_2029() { runMonthlyReconciliation('June', 2029); }
function runMonth_July_2029() { runMonthlyReconciliation('July', 2029); }
function runMonth_August_2029() { runMonthlyReconciliation('August', 2029); }
function runMonth_September_2029() { runMonthlyReconciliation('September', 2029); }
function runMonth_October_2029() { runMonthlyReconciliation('October', 2029); }
function runMonth_November_2029() { runMonthlyReconciliation('November', 2029); }
function runMonth_December_2029() { runMonthlyReconciliation('December', 2029); }

function runMonth_January_2030() { runMonthlyReconciliation('January', 2030); }
function runMonth_February_2030() { runMonthlyReconciliation('February', 2030); }
function runMonth_March_2030() { runMonthlyReconciliation('March', 2030); }
function runMonth_April_2030() { runMonthlyReconciliation('April', 2030); }
function runMonth_May_2030() { runMonthlyReconciliation('May', 2030); }
function runMonth_June_2030() { runMonthlyReconciliation('June', 2030); }
function runMonth_July_2030() { runMonthlyReconciliation('July', 2030); }
function runMonth_August_2030() { runMonthlyReconciliation('August', 2030); }
function runMonth_September_2030() { runMonthlyReconciliation('September', 2030); }
function runMonth_October_2030() { runMonthlyReconciliation('October', 2030); }
function runMonth_November_2030() { runMonthlyReconciliation('November', 2030); }
function runMonth_December_2030() { runMonthlyReconciliation('December', 2030); }


// ===============================================
// UTILITY FUNCTIONS
// ===============================================

function updateCurrentMonth() {
  const now = new Date();
  const monthName = getMonthName(now.getMonth());
  const year = now.getFullYear();
  
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Update Current Month',
    `Run reconciliation for ${monthName} ${year}?`,
    ui.ButtonSet.YES_NO
  );
  
  if (response == ui.Button.YES) {
    showProcessingMessage(`${monthName} ${year}`);
    setupMonthlyMatching(monthName, year);
    showCompletionMessage(`${monthName} ${year}`);
  }
}

function showReconciliationHelp() {
  const ui = SpreadsheetApp.getUi();
  const activeYears = getActiveYears();
  
  const helpText = `
MONTHLY RECONCILIATION GUIDE

Currently Active Years: ${activeYears.join(' and ')}

What This Does:
• Matches payments between your Revenue sheet and Stripe
• Uses Name + Amount matching for accuracy
• Pulls in actual Stripe fees
• Applies color coding (Green=matched, Pink=unmatched)
• Transfers costs from Payment Reconciliation

How To Use:
1. Click "Monthly Reconciliation" menu
2. Select the year, then the month you want to reconcile
3. Wait for completion message
4. Review your monthly sheet for any pink rows

Year Management:
• Menu shows 2 years at a time (current work + next year)
• When you're done with a year, click "Switch to Next Year"
• System auto-prompts when you process January of a new year
• Works indefinitely (2025, 2026, 2027, 2028+)

Example Transition:
• Working in 2025? Menu shows: 2025 and 2026
• Run January 2026? System offers to switch
• After switching: Menu shows 2026 and 2027
• 2025 is hidden (but data remains in your sheets)

Color Codes:
• GREEN rows = Payment matched between systems
• PINK rows = Payment not matched (needs investigation)
• RED rows (in Payment Reconciliation) = Failed orders

Tips:
• Make sure the corresponding sheet exists in Payment Reconciliation
  (e.g., "Nov25" for November 2025)
• Pink rows might indicate manual entries or Stripe payments not recorded
• You can run the reconciliation multiple times if needed
• Use "Switch to Next Year" when ready to move forward

Need Help?
Check the execution log: View → Execution Log (in Apps Script editor)
  `;
  
  ui.alert('Monthly Reconciliation Help', helpText, ui.ButtonSet.OK);
}

function showProcessingMessage(monthYear) {
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Processing ${monthYear} reconciliation...`,
    '⏳ Working',
    5
  );
  Logger.log(`Started reconciliation for ${monthYear}`);
}

function showCompletionMessage(monthYear) {
  const ui = SpreadsheetApp.getUi();
  
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `${monthYear} reconciliation complete!`,
    '✅ Done',
    5
  );
  
  Logger.log(`Completed reconciliation for ${monthYear}`);
  
  // Show detailed completion dialog
  ui.alert(
    'Reconciliation Complete',
    `${monthYear} reconciliation has finished.\n\n` +
    `Next Steps:\n` +
    `1. Check your ${monthYear} sheet\n` +
    `2. Review any PINK rows (unmatched payments)\n` +
    `3. Verify Stripe fees in column J\n` +
    `4. Check costs section at bottom of sheet\n\n` +
    `Tip: Pink rows indicate payments that need investigation.`,
    ui.ButtonSet.OK
  );
}

// ===============================================
// INITIALIZATION
// ===============================================

/**
 * Manual function to create the menu if onOpen doesn't trigger
 * Run this once from the script editor if the menu doesn't appear
 */
function createReconciliationMenu() {
  onOpen();
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Menu created! Refresh the page to see it.',
    '✅ Success',
    5
  );
}