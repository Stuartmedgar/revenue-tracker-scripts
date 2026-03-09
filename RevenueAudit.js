// ===============================================
// REVENUE AUDIT SHEET - Dec 2025 to Mar 2026
// INDEPENDENT VERSION - does NOT use Instalment Tracker
// Calculates total paid by scanning ALL monthly sheets for each student by name
// ===============================================

function createRevenueAuditSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Remove existing audit sheet if it exists
  const existingSheet = ss.getSheetByName('Revenue Audit Dec25-Mar26');
  if (existingSheet) {
    ss.deleteSheet(existingSheet);
    Logger.log('Deleted existing Revenue Audit sheet');
  }
  
  // Create new audit sheet
  const auditSheet = ss.insertSheet('Revenue Audit Dec25-Mar26');
  Logger.log('Created Revenue Audit sheet');
  
  // The 4 months we want to list students FROM
  const auditMonths = [
    'December 2025',
    'January 2026',
    'February 2026',
    'March 2026'
  ];
  
  // Monthly sheet column layout (A-N = 14 cols):
  // A:Date  B:Name  C:Course  D:Sitting  E:Full Price  F:Actual Price  G:Order Type
  // H:FME Fee  I:Stripe Fee  J:Actual Stripe Fee  K:Expected Income
  // L:Payment Plan  M:Instalment  N:Comment
  
  // Audit sheet: columns A-N same as monthly, then:
  // O: Source Month  P: Payments Found (breakdown)  Q: Total Amount Paid  R: Notes
  const auditHeaders = [
    'Date',               // A  (col 1)
    'Name',               // B  (col 2)
    'Course',             // C  (col 3)
    'Sitting',            // D  (col 4)
    'Full Price',         // E  (col 5)
    'Actual Price',       // F  (col 6)
    'Order Type',         // G  (col 7)
    'FME Fee',            // H  (col 8)
    'Stripe Fee',         // I  (col 9)
    'Actual Stripe Fee',  // J  (col 10)
    'Expected Income',    // K  (col 11)
    'Payment Plan',       // L  (col 12)
    'Instalment',         // M  (col 13)
    'Comment',            // N  (col 14)
    'Source Month',       // O  (col 15)
    'Payment Breakdown',  // P  (col 16)
    'Total Amount Paid',  // Q  (col 17)
    'Notes'               // R  (col 18)
  ];
  
  // Set headers
  auditSheet.getRange(1, 1, 1, auditHeaders.length).setValues([auditHeaders]);
  
  // Format header row
  const headerRange = auditSheet.getRange(1, 1, 1, auditHeaders.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#37474f');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontSize(10);
  
  // -------------------------------------------------------
  // STEP 1: Build a payment lookup map from ALL monthly sheets
  // Key: student name (lowercase), Value: array of {amount, date, sheetName}
  // This is completely independent of the Instalment Tracker
  // -------------------------------------------------------
  const allPaymentsMap = buildAllPaymentsMap(ss);
  Logger.log(`Built payment map covering ${Object.keys(allPaymentsMap).length} unique student names`);
  
  // -------------------------------------------------------
  // STEP 2: Collect all student rows from the 4 audit months
  // -------------------------------------------------------
  let allRows = [];
  
  auditMonths.forEach(monthName => {
    const monthSheet = ss.getSheetByName(monthName);
    if (!monthSheet) {
      Logger.log(`⚠️ Sheet "${monthName}" not found - skipping`);
      return;
    }
    
    const dataRange = monthSheet.getDataRange();
    if (dataRange.getNumRows() <= 1) {
      Logger.log(`Sheet "${monthName}" has no data rows`);
      return;
    }
    
    const allData = dataRange.getValues();
    const headers = allData[0];
    const dataRows = allData.slice(1);
    
    const nameCol    = headers.indexOf('Name');
    const fullPriceCol  = headers.indexOf('Full Price');
    const actualPriceCol = headers.indexOf('Actual Price');
    const paymentPlanCol = headers.indexOf('Payment Plan');
    
    let studentCount = 0;
    
    dataRows.forEach(row => {
      const nameVal = row[nameCol];
      
      // Skip blank rows and section headings like "Deferrals"
      if (!nameVal || nameVal.toString().trim() === '') return;
      if (nameVal.toString().toLowerCase().includes('deferral')) return;
      
      const studentName  = nameVal.toString().trim();
      const fullPrice    = Number(row[fullPriceCol]) || 0;
      const isPaymentPlan = row[paymentPlanCol] === 'Y';
      
      // Get the first 14 columns (A-N) - pad if needed
      const monthlyData = row.slice(0, 14);
      while (monthlyData.length < 14) monthlyData.push('');
      
      // -------------------------------------------------------
      // STEP 3: Calculate total paid from the payments map
      // For full payers: only one entry exists (their actual price = full price)
      // For payment plan students: sum ALL entries with this name
      // -------------------------------------------------------
      const nameKey = studentName.toLowerCase();
      const paymentsForStudent = allPaymentsMap[nameKey] || [];
      
      let totalAmountPaid = 0;
      let paymentBreakdown = '';
      let lastPaymentDate = null;
      
      if (!isPaymentPlan) {
        // Full payer - single payment, use actual price directly
        totalAmountPaid = Number(row[actualPriceCol]) || 0;
        paymentBreakdown = `Full payment: £${totalAmountPaid.toFixed(2)}`;
        lastPaymentDate = row[0] ? new Date(row[0]) : null; // Column A = Date
      } else {
        // Payment plan - sum all payments found for this student across all sheets
        if (paymentsForStudent.length > 0) {
          // Filter to only payment plan entries (actual price != full price) 
          // but also include any full-price payments in case they later paid in full
          const relevantPayments = paymentsForStudent.filter(p => 
            Math.abs(p.fullPrice - fullPrice) < 0.01 || p.fullPrice === 0
          );
          
          if (relevantPayments.length > 0) {
            totalAmountPaid = relevantPayments.reduce((sum, p) => sum + p.amount, 0);
            
            // Build breakdown string
            const sortedPayments = relevantPayments.sort((a, b) => a.date - b.date);
            paymentBreakdown = sortedPayments.map(p => 
              `£${p.amount.toFixed(2)} (${Utilities.formatDate(p.date, Session.getScriptTimeZone(), 'dd/MM/yy')} - ${p.sheetName})`
            ).join(' | ');
            
            // Get last payment date
            lastPaymentDate = sortedPayments[sortedPayments.length - 1].date;
          } else {
            // Fallback: use all entries for this student regardless of full price
            totalAmountPaid = paymentsForStudent.reduce((sum, p) => sum + p.amount, 0);
            paymentBreakdown = paymentsForStudent.sort((a,b) => a.date - b.date).map(p => 
              `£${p.amount.toFixed(2)} (${Utilities.formatDate(p.date, Session.getScriptTimeZone(), 'dd/MM/yy')})`
            ).join(' | ');
            lastPaymentDate = paymentsForStudent.sort((a,b) => a.date - b.date).pop()?.date || null;
          }
        } else {
          // No payments found in map at all - use the actual price from this row as minimum
          totalAmountPaid = Number(row[actualPriceCol]) || 0;
          paymentBreakdown = `Only entry found: £${totalAmountPaid.toFixed(2)} (${monthName})`;
          lastPaymentDate = row[0] ? new Date(row[0]) : null;
          Logger.log(`⚠️ No payment map entries for "${studentName}" - using row actual price`);
        }
      }
      
      // -------------------------------------------------------
      // STEP 4: Build the notes / next payment due comment
      // -------------------------------------------------------
      const isPaid = Math.abs(totalAmountPaid - fullPrice) < 0.01;
      let notes = '';
      
      if (!isPaid && isPaymentPlan) {
        const remaining = fullPrice - totalAmountPaid;
        
        // Estimate next payment due = 1 month after last payment
        if (lastPaymentDate && !isNaN(lastPaymentDate.getTime())) {
          const nextDue = new Date(lastPaymentDate);
          nextDue.setMonth(nextDue.getMonth() + 1);
          const formattedNextDue = Utilities.formatDate(nextDue, Session.getScriptTimeZone(), 'dd/MM/yyyy');
          notes = `Next payment due: ${formattedNextDue} | £${remaining.toFixed(2)} remaining`;
        } else {
          notes = `£${remaining.toFixed(2)} remaining - next due date unknown`;
        }
      } else if (!isPaid && !isPaymentPlan) {
        notes = `Full payment expected (£${fullPrice.toFixed(2)}) but only £${totalAmountPaid.toFixed(2)} found - check records`;
      }
      // If fully paid, leave notes blank (green row speaks for itself)
      
      // Build the full audit row
      const auditRow = [
        ...monthlyData,    // A-N  (cols 1-14)
        monthName,         // O    (col 15) - source month
        paymentBreakdown,  // P    (col 16) - payment breakdown
        totalAmountPaid,   // Q    (col 17) - total amount paid
        notes              // R    (col 18) - comments
      ];
      
      allRows.push({
        data: auditRow,
        studentName,
        fullPrice,
        totalAmountPaid,
        isPaid
      });
      
      studentCount++;
    });
    
    Logger.log(`Loaded ${studentCount} students from ${monthName}`);
  });
  
  Logger.log(`Total students collected: ${allRows.length}`);
  
  if (allRows.length === 0) {
    Logger.log('❌ No data found - check that the monthly sheets exist and have data');
    auditSheet.getRange(2, 1).setValue('No student data found in the specified monthly sheets. Check sheet names match exactly.');
    return;
  }
  
  // -------------------------------------------------------
  // STEP 5: Write all rows to the audit sheet at once
  // -------------------------------------------------------
  const dataToWrite = allRows.map(r => r.data);
  auditSheet.getRange(2, 1, dataToWrite.length, auditHeaders.length).setValues(dataToWrite);
  
  // -------------------------------------------------------
  // STEP 6: Apply red/green row colouring
  // -------------------------------------------------------
  allRows.forEach((rowInfo, index) => {
    const rowNumber = index + 2;
    const rowRange = auditSheet.getRange(rowNumber, 1, 1, auditHeaders.length);
    
    if (rowInfo.isPaid) {
      rowRange.setBackground('#c8e6c9'); // Light green
      rowRange.setFontColor('#1b5e20');  // Dark green text
    } else {
      rowRange.setBackground('#ffcdd2'); // Light red
      rowRange.setFontColor('#b71c1c');  // Dark red text
    }
  });
  
  // -------------------------------------------------------
  // STEP 7: Format columns and add summary totals
  // -------------------------------------------------------
  formatAuditSheetColumns(auditSheet, allRows.length);
  addAuditTotalsRow(auditSheet, allRows.length, allRows);
  
  auditSheet.setFrozenRows(1);
  auditSheet.autoResizeColumns(1, auditHeaders.length);
  auditSheet.setColumnWidth(16, 320); // P - Payment Breakdown
  auditSheet.setColumnWidth(18, 300); // R - Notes
  
  const paidCount = allRows.filter(r => r.isPaid).length;
  const unpaidCount = allRows.length - paidCount;
  Logger.log(`✅ Revenue Audit sheet created successfully`);
  Logger.log(`   Total students: ${allRows.length}`);
  Logger.log(`   Green (fully paid): ${paidCount}`);
  Logger.log(`   Red (outstanding): ${unpaidCount}`);
  Logger.log(`   NOTE: Total paid calculated by scanning ALL monthly sheets - independent of Instalment Tracker`);
}

// ===============================================
// CORE INDEPENDENT PAYMENT LOOKUP
// Scans every monthly sheet in the workbook and builds
// a map of { studentNameLowercase: [{amount, date, sheetName, fullPrice}] }
// ===============================================

function buildAllPaymentsMap(ss) {
  const paymentsMap = {};
  const allSheets = ss.getSheets();
  
  // Month names to identify monthly sheets
  const monthNames = ['January','February','March','April','May','June',
                      'July','August','September','October','November','December'];
  
  allSheets.forEach(sheet => {
    const sheetName = sheet.getName();
    
    // Only scan sheets that look like monthly sheets (e.g. "December 2025")
    const parts = sheetName.split(' ');
    const allowedSheets = [
    'December 2025', 'January 2026', 'February 2026', 'March 2026'
  ];
  if (!allowedSheets.includes(sheetName)) return;
    
    const dataRange = sheet.getDataRange();
    if (dataRange.getNumRows() <= 1) return;
    
    const allData = dataRange.getValues();
    const headers = allData[0];
    const dataRows = allData.slice(1);
    
    const nameCol       = headers.indexOf('Name');
    const actualPriceCol = headers.indexOf('Actual Price');
    const fullPriceCol  = headers.indexOf('Full Price');
    const dateCol       = headers.indexOf('Date');
    
    if (nameCol === -1 || actualPriceCol === -1) return;
    
    dataRows.forEach(row => {
      const nameVal = row[nameCol];
      if (!nameVal || nameVal.toString().trim() === '') return;
      if (nameVal.toString().toLowerCase().includes('deferral')) return;
      
      const studentName = nameVal.toString().trim();
      const nameKey     = studentName.toLowerCase();
      const amount      = Number(row[actualPriceCol]) || 0;
      const fullPrice   = Number(row[fullPriceCol]) || 0;
      const dateVal     = row[dateCol];
      const date        = dateVal ? new Date(dateVal) : new Date(0);
      
      if (amount <= 0) return; // Skip zero/blank amounts
      
      if (!paymentsMap[nameKey]) {
        paymentsMap[nameKey] = [];
      }
      
      paymentsMap[nameKey].push({ amount, date, sheetName, fullPrice });
    });
  });
  
  return paymentsMap;
}

// ===============================================
// FORMATTING HELPERS
// ===============================================

function formatAuditSheetColumns(sheet, dataRowCount) {
  if (dataRowCount <= 0) return;
  
  // Date column A
  sheet.getRange(2, 1, dataRowCount, 1).setNumberFormat('dd/MM/yyyy');
  
  // Currency columns: E=5, F=6, H=8, I=9, J=10, K=11, Q=17
  [5, 6, 8, 9, 10, 11, 17].forEach(col => {
    sheet.getRange(2, col, dataRowCount, 1).setNumberFormat('£#,##0.00');
  });
  
  // Make Q and R header cells stand out with distinct colours
  sheet.getRange(1, 17).setBackground('#1565c0'); // Blue - Total Amount Paid
  sheet.getRange(1, 18).setBackground('#4a148c'); // Purple - Notes
}

function addAuditTotalsRow(sheet, dataRowCount, allRows) {
  const totalsRow = dataRowCount + 3;
  
  const totalFullPrice = allRows.reduce((sum, r) => sum + r.fullPrice, 0);
  const totalPaid      = allRows.reduce((sum, r) => sum + r.totalAmountPaid, 0);
  const outstanding    = totalFullPrice - totalPaid;
  const paidCount      = allRows.filter(r => r.isPaid).length;
  const unpaidCount    = allRows.length - paidCount;
  
  sheet.getRange(totalsRow, 1).setValue('TOTALS / SUMMARY');
  sheet.getRange(totalsRow, 5).setValue(totalFullPrice);
  sheet.getRange(totalsRow, 17).setValue(totalPaid);
  sheet.getRange(totalsRow, 18).setValue(`Outstanding: £${outstanding.toFixed(2)}`);
  
  const summaryRow = totalsRow + 1;
  sheet.getRange(summaryRow, 1).setValue(
    `${allRows.length} students total  |  ${paidCount} fully paid (green)  |  ${unpaidCount} outstanding (red)  |  Outstanding balance: £${outstanding.toFixed(2)}`
  );
  sheet.getRange(summaryRow, 1, 1, 10).merge();
  
  // Format totals
  [sheet.getRange(totalsRow, 1), sheet.getRange(summaryRow, 1)].forEach(r => r.setFontWeight('bold'));
  [sheet.getRange(totalsRow, 5), sheet.getRange(totalsRow, 17)].forEach(r => {
    r.setNumberFormat('£#,##0.00');
    r.setFontWeight('bold');
  });
  sheet.getRange(totalsRow, 18).setFontWeight('bold');
  sheet.getRange(summaryRow, 1).setFontStyle('italic');
}