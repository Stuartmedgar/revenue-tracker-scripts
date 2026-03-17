// ===============================================
// INSTALMENTTRACKER.GS
// Clean, reliable instalment tracking.
//
// DESIGN PRINCIPLES:
//   - Tracker shows ACTIVE students only (still owe money)
//   - When fully paid → row deleted immediately (monthly sheets = permanent record)
//   - rebuildInstalmentTracker() can be run any time to resync from scratch
//   - No guessing instalment numbers — we sum ALL payments found per student+fullPrice
//
// COLUMNS: A=Student Name, B=Course, C=Full Price, D=Amount Paid,
//          E=Amount Remaining, F=Instalments Paid, G=Last Payment Date,
//          H=Next Payment Due
//
// PRICING REFERENCE:
//   Platinum (1047):           397 → 350 → 300
//   Tuition/Revision Plus (822): 522 → 300
//   Revision (647):            347 → 300
//   Tuition (597):             297 → 300
// ===============================================

// -------------------------------------------------------
// SHEET SETUP
// -------------------------------------------------------

function getOrCreateInstalmentTrackerSheet(ss) {
  let sheet = ss.getSheetByName('Instalment Tracker');
  if (!sheet) {
    sheet = ss.insertSheet('Instalment Tracker');
    _writeInstalmentTrackerHeaders(sheet);
    positionInstalmentTrackerTab(ss, sheet);
    Logger.log('Created new Instalment Tracker sheet');
  }
  return sheet;
}

function _writeInstalmentTrackerHeaders(sheet) {
  const headers = [
    'Student Name',     // A
    'Course',           // B
    'Full Price',       // C
    'Amount Paid',      // D
    'Amount Remaining', // E
    'Instalments Paid', // F
    'Last Payment Date',// G
    'Next Payment Due'  // H
  ];
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#37474f');
  headerRange.setFontColor('#ffffff');
  sheet.setFrozenRows(1);
}

// -------------------------------------------------------
// REBUILD — scan all monthly sheets, write active students
// -------------------------------------------------------

/**
 * Master rebuild function. Safe to run at any time.
 * Clears the tracker and rewrites it from the monthly sheets.
 * Only writes students who still have a balance outstanding.
 */
function rebuildInstalmentTracker() {
  Logger.log('🔄 REBUILDING INSTALMENT TRACKER FROM MONTHLY SHEETS');
  Logger.log('=====================================================');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trackerSheet = getOrCreateInstalmentTrackerSheet(ss);

  // Clear all data rows (keep header)
  const lastRow = trackerSheet.getLastRow();
  if (lastRow > 1) {
    trackerSheet.deleteRows(2, lastRow - 1);
  }

  // Build payment map from ALL monthly sheets
  const paymentMap = _buildPaymentMapAllSheets(ss);
  Logger.log(`Payment map built: ${Object.keys(paymentMap).length} student+course combinations found`);

  // Write active (unpaid) students to tracker
  let activeCount = 0;
  let completedCount = 0;
  const rows = [];

  Object.values(paymentMap).forEach(entry => {
    const remaining = entry.fullPrice - entry.totalPaid;

    if (remaining > 0.01) {
      // Still owes money → include in tracker
      const nextDue = _calcNextPaymentDate(entry.lastPaymentDate);
      rows.push([
        entry.studentName,
        entry.course,
        entry.fullPrice,
        entry.totalPaid,
        remaining,
        entry.instalmentCount,
        entry.lastPaymentDate,
        nextDue
      ]);
      activeCount++;
    } else {
      completedCount++;
    }
  });

  // Sort by Next Payment Due (soonest first)
  rows.sort((a, b) => {
    const da = a[7] instanceof Date ? a[7] : new Date(a[7]);
    const db = b[7] instanceof Date ? b[7] : new Date(b[7]);
    return da - db;
  });

  if (rows.length > 0) {
    trackerSheet.getRange(2, 1, rows.length, 8).setValues(rows);
    _formatTrackerDataRows(trackerSheet, 2, rows.length);
  }

  Logger.log(`\n✅ REBUILD COMPLETE`);
  Logger.log(`   Active (outstanding balance): ${activeCount}`);
  Logger.log(`   Already fully paid (not shown): ${completedCount}`);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Instalment Tracker rebuilt. ${activeCount} active students.`,
    '✅ Done', 6
  );
}

// -------------------------------------------------------
// PROCESS SINGLE PAYMENT — called automatically when data
// moves to a monthly sheet (via processInstalmentPayment)
// -------------------------------------------------------

/**
 * Called from DataProcessor.js / moveToMonthlySheet when
 * paymentInfo.isPaymentPlan === true.
 *
 * Finds or creates the tracker row for this student+course+fullPrice,
 * updates totals, then deletes the row if now fully paid.
 */
function processInstalmentPayment(studentName, course, fullPrice, actualPrice, paymentDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trackerSheet = getOrCreateInstalmentTrackerSheet(ss);

  const full   = Number(fullPrice);
  const actual = Number(actualPrice);
  const name   = studentName ? studentName.toString().trim() : '';

  if (!name || !full || !actual) {
    Logger.log(`processInstalmentPayment: skipping - missing data (name="${name}", full=${full}, actual=${actual})`);
    return;
  }

  // Re-derive total from monthly sheets so we are always accurate,
  // not relying on incremental state that can drift.
  const allPayments = _getAllPaymentsForStudent(ss, name, full);

  if (allPayments.length === 0) {
    // Shouldn't happen (we were just called after writing to a monthly sheet)
    // but fall back to the single payment value.
    allPayments.push({ amount: actual, date: paymentDate instanceof Date ? paymentDate : new Date(paymentDate) });
  }

  const totalPaid       = allPayments.reduce((s, p) => s + p.amount, 0);
  const remaining       = full - totalPaid;
  const instalmentCount = allPayments.length;
  const sortedPayments  = allPayments.slice().sort((a, b) => a.date - b.date);
  const lastDate        = sortedPayments[sortedPayments.length - 1].date;

  // Find existing row in tracker
  const existingRow = _findTrackerRow(trackerSheet, name, full);

  if (remaining <= 0.01) {
    // Fully paid → remove from tracker (monthly sheets are the record)
    if (existingRow > 0) {
      trackerSheet.deleteRow(existingRow);
      Logger.log(`✅ COMPLETED & REMOVED: ${name} - ${course} (paid £${totalPaid.toFixed(2)} of £${full})`);
    } else {
      Logger.log(`✅ COMPLETED (not in tracker): ${name} - £${totalPaid.toFixed(2)} of £${full}`);
    }
    return;
  }

  // Still owes money → upsert
  const nextDue = _calcNextPaymentDate(lastDate);
  const rowData = [name, course, full, totalPaid, remaining, instalmentCount, lastDate, nextDue];

  if (existingRow > 0) {
    trackerSheet.getRange(existingRow, 1, 1, rowData.length).setValues([rowData]);
    _formatTrackerDataRows(trackerSheet, existingRow, 1);
    Logger.log(`Updated tracker: ${name} - paid £${totalPaid.toFixed(2)}, £${remaining.toFixed(2)} remaining`);
  } else {
    const newRow = trackerSheet.getLastRow() + 1;
    trackerSheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);
    _formatTrackerDataRows(trackerSheet, newRow, 1);
    Logger.log(`Added to tracker: ${name} - ${course} - £${totalPaid.toFixed(2)} of £${full}`);
  }
}

// -------------------------------------------------------
// INTERNAL HELPERS
// -------------------------------------------------------

/**
 * Scans ALL monthly sheets for accurate payment totals, but only returns
 * students who had at least one payment within the last 6 months.
 *
 * This ensures:
 *   - totalPaid is always correct even if instalment 1 was >6 months ago
 *   - The tracker stays focused on recently active students
 *
 * Only includes entries where Payment Plan = "Y".
 */
function _buildPaymentMapAllSheets(ss) {
  const map = {};
  const allSheets = ss.getSheets();
  const cutoff = _getSixMonthCutoff();

  allSheets.forEach(sheet => {
    if (!isMonthlySheetName(sheet.getName())) return;
    // No cutoff filter here — scan every sheet for accurate totals

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return;

    const headers = data[0];
    const nameCol        = headers.indexOf('Name');
    const courseCol      = headers.indexOf('Course');
    const fullPriceCol   = headers.indexOf('Full Price');
    const actualPriceCol = headers.indexOf('Actual Price');
    const paymentPlanCol = headers.indexOf('Payment Plan');
    const dateCol        = headers.indexOf('Date');

    if (nameCol === -1 || actualPriceCol === -1 || fullPriceCol === -1) return;

    const sheetIsRecent = _isSheetWithinCutoff(sheet.getName(), cutoff);

    data.slice(1).forEach(row => {
      if (row[paymentPlanCol] !== 'Y') return;

      const rawName = row[nameCol];
      if (!rawName || rawName.toString().trim() === '') return;

      const studentName = rawName.toString().trim();
      const full        = Number(row[fullPriceCol]);
      const actual      = Number(row[actualPriceCol]);
      const course      = row[courseCol] ? row[courseCol].toString() : getCourseFromPrice(full);
      const date        = row[dateCol] ? new Date(row[dateCol]) : new Date();

      if (!full || !actual) return;

      const key = `${studentName.toLowerCase()}|||${full}`;

      if (!map[key]) {
        map[key] = {
          studentName,
          course,
          fullPrice: full,
          totalPaid: 0,
          instalmentCount: 0,
          lastPaymentDate: date,
          hasRecentPayment: false  // gate: only include if true after scanning all sheets
        };
      }

      map[key].totalPaid       += actual;
      map[key].instalmentCount += 1;

      if (date > map[key].lastPaymentDate) {
        map[key].lastPaymentDate = date;
      }

      // Flag if this payment falls within the last 6 months
      if (sheetIsRecent) {
        map[key].hasRecentPayment = true;
      }
    });
  });

  // Filter: only return students with at least one payment in the last 6 months
  Object.keys(map).forEach(key => {
    if (!map[key].hasRecentPayment) {
      Logger.log(`Excluding (no recent activity): ${map[key].studentName} - last payment ${map[key].lastPaymentDate.toDateString()}`);
      delete map[key];
    }
  });

  return map;
}

/**
 * Returns all payment plan payments for a specific student+fullPrice
 * across all monthly sheets.
 */
function _getAllPaymentsForStudent(ss, studentName, fullPrice) {
  const payments = [];
  const nameLower = studentName.toLowerCase();
  const full = Number(fullPrice);
  // No cutoff filter — scan all sheets so totalPaid is always accurate

  ss.getSheets().forEach(sheet => {
    if (!isMonthlySheetName(sheet.getName())) return;

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return;

    const headers = data[0];
    const nameCol        = headers.indexOf('Name');
    const fullPriceCol   = headers.indexOf('Full Price');
    const actualPriceCol = headers.indexOf('Actual Price');
    const paymentPlanCol = headers.indexOf('Payment Plan');
    const dateCol        = headers.indexOf('Date');

    if (nameCol === -1 || actualPriceCol === -1) return;

    data.slice(1).forEach(row => {
      if (row[paymentPlanCol] !== 'Y') return;
      const rowName = row[nameCol] ? row[nameCol].toString().trim() : '';
      if (rowName.toLowerCase() !== nameLower) return;
      if (Math.abs(Number(row[fullPriceCol]) - full) > 0.01) return;

      const amount = Number(row[actualPriceCol]);
      const date   = row[dateCol] ? new Date(row[dateCol]) : new Date();
      if (amount > 0) payments.push({ amount, date });
    });
  });

  return payments;
}

/**
 * Returns a Date 6 months ago from today (start of that month).
 */
function _getSixMonthCutoff() {
  const d = new Date();
  d.setMonth(d.getMonth() - 6);
  d.setDate(1);
  d.setHours(0, 0, 0, 0);
  return d;
}

/**
 * Returns true if the sheet name (e.g. "Jan26", "March 2026") represents
 * a month on or after the cutoff date.
 * Handles both short format (Jan26) and long format (January 2026).
 */
function _isSheetWithinCutoff(sheetName, cutoff) {
  const monthMap = {
    jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5,
    jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11,
    january: 0, february: 1, march: 2, april: 3, june: 5,
    july: 6, august: 7, september: 8, october: 9, november: 10, december: 11
  };

  const name = sheetName.trim().toLowerCase();

  // Short format: Jan26, Feb26 etc.
  const shortMatch = name.match(/^([a-z]+)(\d{2})$/);
  if (shortMatch) {
    const month = monthMap[shortMatch[1]];
    const year  = 2000 + parseInt(shortMatch[2], 10);
    if (month === undefined || isNaN(year)) return false;
    return new Date(year, month, 1) >= cutoff;
  }

  // Long format: January 2026, March 2026 etc.
  const longMatch = name.match(/^([a-z]+)\s+(\d{4})$/);
  if (longMatch) {
    const month = monthMap[longMatch[1]];
    const year  = parseInt(longMatch[2], 10);
    if (month === undefined || isNaN(year)) return false;
    return new Date(year, month, 1) >= cutoff;
  }

  return false;
}

/**
 * Finds the row index (1-based) for a student+fullPrice in the tracker.
 * Returns 0 if not found.
 */
function _findTrackerRow(sheet, studentName, fullPrice) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 0;

  const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  const nameLower = studentName.toLowerCase();
  const full = Number(fullPrice);

  for (let i = 0; i < data.length; i++) {
    const rowName  = data[i][0] ? data[i][0].toString().trim().toLowerCase() : '';
    const rowFull  = Number(data[i][2]);
    if (rowName === nameLower && Math.abs(rowFull - full) < 0.01) {
      return i + 2; // +2: header row + 0-index
    }
  }
  return 0;
}

/**
 * Returns a Date one month after the given date.
 */
function _calcNextPaymentDate(lastPaymentDate) {
  const d = lastPaymentDate instanceof Date ? new Date(lastPaymentDate) : new Date(lastPaymentDate);
  if (isNaN(d.getTime())) return '';
  d.setMonth(d.getMonth() + 1);
  return d;
}

/**
 * Formats data rows: currency for price columns, date for date columns.
 */
function _formatTrackerDataRows(sheet, startRow, numRows) {
  if (numRows <= 0) return;
  sheet.getRange(startRow, 3, numRows, 3).setNumberFormat('£#,##0.00'); // C, D, E
  sheet.getRange(startRow, 7, numRows, 2).setNumberFormat('dd/mm/yyyy'); // G, H
}

// -------------------------------------------------------
// DIAGNOSTIC — run to check current state
// -------------------------------------------------------

function diagnoseInstalmentTracker() {
  Logger.log('🔍 INSTALMENT TRACKER DIAGNOSIS');
  Logger.log('================================\n');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trackerSheet = ss.getSheetByName('Instalment Tracker');

  if (!trackerSheet || trackerSheet.getLastRow() <= 1) {
    Logger.log('Tracker is empty or missing. Run rebuildInstalmentTracker() to populate it.');
    return;
  }

  const data = trackerSheet.getDataRange().getValues().slice(1);
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  let overdueCount = 0;

  data.forEach((row, i) => {
    const name      = row[0];
    const course    = row[1];
    const full      = Number(row[2]);
    const paid      = Number(row[3]);
    const remaining = Number(row[4]);
    const count     = Number(row[5]);
    const lastDate  = row[6] ? new Date(row[6]) : null;
    const nextDue   = row[7] ? new Date(row[7]) : null;

    const isOverdue = nextDue && nextDue < today;
    if (isOverdue) overdueCount++;

    Logger.log(`${i + 1}. ${name} (${course})`);
    Logger.log(`   Paid: £${paid.toFixed(2)} / £${full.toFixed(2)}   Remaining: £${remaining.toFixed(2)}   Instalments: ${count}`);
    Logger.log(`   Last payment: ${lastDate ? lastDate.toDateString() : 'unknown'}   Next due: ${nextDue ? nextDue.toDateString() : 'unknown'}${isOverdue ? '  ⚠️ OVERDUE' : ''}`);
  });

  Logger.log(`\n📊 Summary: ${data.length} active students, ${overdueCount} overdue`);
  Logger.log('\n💡 Run rebuildInstalmentTracker() if anything looks wrong.');
}

// -------------------------------------------------------
// MANUAL TOOLS — for fixing individual students
// -------------------------------------------------------

/**
 * Diagnostic: explains exactly why a specific student is missing from the tracker.
 * Set STUDENT_NAME to the student you are investigating, then run from the editor.
 * Check Logs for output.
 */
function diagnoseStudentMissing() {
  const STUDENT_NAME = 'Enter Student Name Here'; // ← change this
  _diagnoseStudent(STUDENT_NAME);
}

function _diagnoseStudent(studentName) {
  Logger.log('🔍 DIAGNOSING: "' + studentName + '"');
  Logger.log('==================================================');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cutoff = _getSixMonthCutoff();
  const nameLower = studentName.trim().toLowerCase();
  const allSheets = ss.getSheets();

  let totalFound = 0;
  let recentFound = 0;
  const matchesBySheet = [];

  allSheets.forEach(sheet => {
    if (!isMonthlySheetName(sheet.getName())) return;

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return;

    const headers        = data[0];
    const nameCol        = headers.indexOf('Name');
    const fullPriceCol   = headers.indexOf('Full Price');
    const actualPriceCol = headers.indexOf('Actual Price');
    const paymentPlanCol = headers.indexOf('Payment Plan');
    const dateCol        = headers.indexOf('Date');

    if (nameCol === -1) return;

    const sheetIsRecent = _isSheetWithinCutoff(sheet.getName(), cutoff);

    data.slice(1).forEach((row, i) => {
      const rowName = row[nameCol] ? row[nameCol].toString().trim() : '';
      const exactMatch   = rowName.toLowerCase() === nameLower;
      const partialMatch = rowName.toLowerCase().includes(nameLower) || nameLower.includes(rowName.toLowerCase());

      if (!exactMatch && !partialMatch) return;

      const payPlan = row[paymentPlanCol];
      const full    = Number(row[fullPriceCol]);
      const actual  = Number(row[actualPriceCol]);
      const date    = row[dateCol] ? new Date(row[dateCol]) : null;

      totalFound++;
      if (sheetIsRecent) recentFound++;

      matchesBySheet.push({ sheet: sheet.getName(), rowNum: i + 2, storedName: rowName,
        exactMatch, payPlan, full, actual, date, sheetIsRecent });
    });
  });

  if (totalFound === 0) {
    Logger.log('❌ NOT FOUND in any monthly sheet.');
    Logger.log('   Check the exact name spelling in the monthly sheets.');
    return;
  }

  Logger.log('\nFound ' + totalFound + ' row(s) across all sheets (' + recentFound + ' within last 6 months):\n');

  matchesBySheet.forEach(m => {
    const flags = [];
    if (!m.exactMatch)     flags.push('⚠️ NAME MISMATCH (partial match only)');
    if (m.payPlan !== 'Y') flags.push('⚠️ Payment Plan = "' + m.payPlan + '" (expected "Y")');
    if (!m.full)           flags.push('⚠️ Full Price is blank/zero');
    if (!m.actual)         flags.push('⚠️ Actual Price is blank/zero');
    if (!m.sheetIsRecent)  flags.push('⚠️ Sheet is outside 6-month window');

    const status = flags.length === 0 ? '✅ Should be picked up' : '❌ Will be skipped';
    Logger.log('  Sheet: ' + m.sheet + '  Row: ' + m.rowNum + '  StoredName: "' + m.storedName + '"  ' + status);
    Logger.log('    PayPlan=' + m.payPlan + '  Full=£' + m.full + '  Actual=£' + m.actual +
      '  Date=' + (m.date ? m.date.toDateString() : 'blank') + '  RecentSheet=' + m.sheetIsRecent);
    flags.forEach(f => Logger.log('    ' + f));
    Logger.log('');
  });

  Logger.log('--- DIAGNOSIS ---');
  const payPlanIssue = matchesBySheet.some(m => m.payPlan !== 'Y');
  const nameIssue    = matchesBySheet.every(m => !m.exactMatch);
  const recencyIssue = recentFound === 0;
  const priceIssue   = matchesBySheet.some(m => !m.full || !m.actual);

  if (nameIssue)    Logger.log('🔴 Name mismatch — no exact match. Check spelling in monthly sheets.');
  if (payPlanIssue) Logger.log('🔴 Payment Plan not set to "Y" — rebuild skips these rows.');
  if (recencyIssue) Logger.log('🔴 All payments outside 6-month window — use manuallyAddToInstalmentTracker().');
  if (priceIssue)   Logger.log('🔴 Full Price or Actual Price blank/zero on one or more rows.');
  if (!nameIssue && !payPlanIssue && !recencyIssue && !priceIssue) {
    Logger.log('✅ No obvious issues found. Try running rebuildInstalmentTracker() again.');
  }
  Logger.log('\n💡 Use manuallyAddToInstalmentTracker() to force-add this student if needed.');
}

/**
 * Manually add or update a single student in the tracker.
 * Use when a student is correct in the monthly sheets but rebuild is not picking them up.
 *
 * Fill in the three constants below and run from the Apps Script editor.
 */
function manuallyAddToInstalmentTracker() {
  // ── FILL THESE IN ──────────────────────────────────
  const STUDENT_NAME = 'Enter Student Name Here'; // exact name as in monthly sheets
  const COURSE       = 'Platinum';                // e.g. Platinum, Revision, Tuition
  const FULL_PRICE   = 1047;                      // e.g. 1047, 822, 647, 597
  // ───────────────────────────────────────────────────

  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const name = STUDENT_NAME.trim();
  const full = Number(FULL_PRICE);

  if (name === 'Enter Student Name Here' || !full) {
    Logger.log('❌ Please fill in STUDENT_NAME and FULL_PRICE before running.');
    return;
  }

  Logger.log('Manual add: scanning all sheets for "' + name + '" with full price £' + full + '...');

  const allPayments = _getAllPaymentsForStudent(ss, name, full);

  if (allPayments.length === 0) {
    Logger.log('❌ No payments found for "' + name + '" with full price £' + full + '.');
    Logger.log('   Check name spelling and full price match exactly what is in the monthly sheets.');
    Logger.log('   Run diagnoseStudentMissing() for a detailed breakdown.');
    return;
  }

  const totalPaid = allPayments.reduce((s, p) => s + p.amount, 0);
  const remaining = full - totalPaid;

  if (remaining <= 0.01) {
    Logger.log('✅ "' + name + '" has already paid in full (£' + totalPaid.toFixed(2) + ' of £' + full + '). Not adding to tracker.');
    return;
  }

  const sorted   = allPayments.slice().sort((a, b) => a.date - b.date);
  const lastDate = sorted[sorted.length - 1].date;

  processInstalmentPayment(name, COURSE, full, allPayments[allPayments.length - 1].amount, lastDate);

  const nextDue = _calcNextPaymentDate(lastDate);
  Logger.log('✅ Done. "' + name + '" added/updated in tracker.');
  Logger.log('   Paid: £' + totalPaid.toFixed(2) + '  Remaining: £' + remaining.toFixed(2) + '  Instalments: ' + allPayments.length);
  Logger.log('   Next payment due: ' + (nextDue instanceof Date ? nextDue.toDateString() : 'unknown'));
}

// -------------------------------------------------------
// LEGACY STUBS — keep these so existing call sites don't break
// -------------------------------------------------------

// Called from Code.js setupInstalmentTracking
function setupInstalmentTracking() {
  // Weekly cleanup no longer needed (tracker auto-deletes completed students)
  // but keep the trigger for backward compat in case other scripts reference it.
  Logger.log('setupInstalmentTracking: nothing to set up (tracker self-maintains)');
}

// Called from Code.js updateInstalmentTrackerFromMonthlySheets
function updateInstalmentTrackerFromMonthlySheets() {
  rebuildInstalmentTracker();
}

// Called from Code.js testInstalmentTracking
function debugInstalmentTracker() {
  diagnoseInstalmentTracker();
}

// Kept for backward compat — processInstalmentPayment is the main entry point
function cleanupCompletedPayments() {
  Logger.log('cleanupCompletedPayments: not needed in new system — tracker auto-deletes completed students');
  Logger.log('Run rebuildInstalmentTracker() if you need a full resync.');
}

// -------------------------------------------------------
// STUDENT PAYMENT HISTORY — lookup any student's full record
// -------------------------------------------------------

/**
 * Prints a full chronological payment history for a specific student
 * across all monthly sheets. Works for both instalment and full-payment students.
 *
 * Set STUDENT_NAME below and run from the Apps Script editor, then check Logs.
 */
function getStudentPaymentHistory() {
  const STUDENT_NAME = 'Palvi Halai'; // ← change this
  _printStudentPaymentHistory(STUDENT_NAME);
}

function _printStudentPaymentHistory(studentName) {
  Logger.log('📋 PAYMENT HISTORY: "' + studentName + '"');
  Logger.log('='.repeat(50));

  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const nameLower = studentName.trim().toLowerCase();
  const allSheets = ss.getSheets();
  const records   = [];

  allSheets.forEach(sheet => {
    if (!isMonthlySheetName(sheet.getName())) return;

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return;

    const headers        = data[0];
    const nameCol        = headers.indexOf('Name');
    const dateCol        = headers.indexOf('Date');
    const courseCol      = headers.indexOf('Course');
    const sittingCol     = headers.indexOf('Sitting');
    const fullPriceCol   = headers.indexOf('Full Price');
    const actualPriceCol = headers.indexOf('Actual Price');
    const payPlanCol     = headers.indexOf('Payment Plan');
    const instalmentCol  = headers.indexOf('Instalment');

    if (nameCol === -1) return;

    data.slice(1).forEach((row, i) => {
      const rowName = row[nameCol] ? row[nameCol].toString().trim() : '';
      if (rowName.toLowerCase() !== nameLower) return;

      records.push({
        sheet:      sheet.getName(),
        date:       row[dateCol]        ? new Date(row[dateCol]) : null,
        course:     row[courseCol]      ? row[courseCol].toString() : '',
        sitting:    row[sittingCol]     ? row[sittingCol].toString() : '',
        fullPrice:  Number(row[fullPriceCol])   || 0,
        actual:     Number(row[actualPriceCol]) || 0,
        payPlan:    row[payPlanCol]     ? row[payPlanCol].toString() : '',
        instalment: row[instalmentCol] ? row[instalmentCol].toString() : '',
        rowNum:     i + 2
      });
    });
  });

  if (records.length === 0) {
    Logger.log('❌ No records found for "' + studentName + '" in any monthly sheet.');
    Logger.log('   Check the exact name spelling.');
    return;
  }

  // Sort chronologically
  records.sort((a, b) => (a.date || 0) - (b.date || 0));

  // Group by fullPrice (in case student has multiple courses)
  const groups = {};
  records.forEach(r => {
    const key = r.fullPrice || 'unknown';
    if (!groups[key]) groups[key] = [];
    groups[key].push(r);
  });

  Object.entries(groups).forEach(([fullPrice, payments]) => {
    const course     = payments[0].course || getCourseFromPrice(fullPrice);
    const totalPaid  = payments.reduce((s, p) => s + p.actual, 0);
    const remaining  = Number(fullPrice) - totalPaid;

    Logger.log('\n── ' + course + ' (Full price: £' + fullPrice + ') ──');
    Logger.log('   Payments found: ' + payments.length);

    payments.forEach((p, i) => {
      const dateStr     = p.date ? p.date.toLocaleDateString('en-GB') : 'unknown date';
      const planLabel   = p.payPlan === 'Y'
        ? ' [Instalment ' + (p.instalment || '?') + ']'
        : ' [Full payment]';
      Logger.log('   ' + (i + 1) + '. £' + p.actual.toFixed(2) + '  ' + dateStr + planLabel + '  (Sheet: ' + p.sheet + ', Row ' + p.rowNum + ')');
    });

    Logger.log('   ─────────────────────────────');
    Logger.log('   Total paid:  £' + totalPaid.toFixed(2));
    Logger.log('   Full price:  £' + fullPrice);
    if (remaining > 0.01) {
      Logger.log('   ⚠️  Still owed: £' + remaining.toFixed(2));
    } else {
      Logger.log('   ✅  Fully paid');
    }
  });

  Logger.log('\n📊 Total records found: ' + records.length + ' across ' +
    [...new Set(records.map(r => r.sheet))].length + ' sheet(s)');
}