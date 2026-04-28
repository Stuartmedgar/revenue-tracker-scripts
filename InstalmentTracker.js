// ===============================================
// INSTALMENTTRACKER.GS
// Clean, reliable instalment tracking.
//
// DESIGN PRINCIPLES:
//   - Active students shown at top (still owe money)
//   - Completed students shown below a "COMPLETED" header row (greyed out)
//   - rebuildInstalmentTracker() can be run any time to resync from scratch
//   - No guessing instalment numbers — we sum ALL payments found per student+fullPrice
//   - Full payment history shown per student (P1, P2, P3 date + amount columns)
//
// COLUMNS (13 total):
//   A = Student Name
//   B = Course
//   C = Full Price
//   D = Amount Paid
//   E = Amount Remaining
//   F = Next Payment Due  (blank for completed)
//   G = P1 Date
//   H = P1 Amount
//   I = P2 Date
//   J = P2 Amount
//   K = P3 Date
//   L = P3 Amount
//   M = Completion Date   (blank for active)
//
// PRICING REFERENCE:
//   Platinum (1047):             397 → 350 → 300
//   Tuition/Revision Plus (822): 522 → 300
//   Revision (647):              347 → 300
//   Tuition (597):               297 → 300
// ===============================================

var TRACKER_COL_COUNT = 13; // A–M

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
    'Student Name',      // A
    'Course',            // B
    'Full Price',        // C
    'Amount Paid',       // D
    'Amount Remaining',  // E
    'Next Payment Due',  // F
    'P1 Date',           // G
    'P1 Amount',         // H
    'P2 Date',           // I
    'P2 Amount',         // J
    'P3 Date',           // K
    'P3 Amount',         // L
    'Completion Date'    // M
  ];
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#37474f');
  headerRange.setFontColor('#ffffff');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 180);  // Name
  sheet.setColumnWidth(2, 140);  // Course
  sheet.setColumnWidth(3, 90);   // Full Price
  sheet.setColumnWidth(4, 100);  // Amount Paid
  sheet.setColumnWidth(5, 120);  // Amount Remaining
  sheet.setColumnWidth(6, 120);  // Next Payment Due
  sheet.setColumnWidth(7, 90);   // P1 Date
  sheet.setColumnWidth(8, 90);   // P1 Amount
  sheet.setColumnWidth(9, 90);   // P2 Date
  sheet.setColumnWidth(10, 90);  // P2 Amount
  sheet.setColumnWidth(11, 90);  // P3 Date
  sheet.setColumnWidth(12, 90);  // P3 Amount
  sheet.setColumnWidth(13, 120); // Completion Date
}

// -------------------------------------------------------
// REBUILD — scan all monthly sheets, write both sections
// -------------------------------------------------------

/**
 * Master rebuild function. Safe to run at any time.
 * Section 1: Active students (outstanding balance), sorted by Next Due soonest first.
 * Section 2: "COMPLETED" header row, then completed students sorted by completion date desc.
 */
function rebuildInstalmentTracker() {
  Logger.log('🔄 REBUILDING INSTALMENT TRACKER FROM MONTHLY SHEETS');
  Logger.log('=====================================================');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trackerSheet = getOrCreateInstalmentTrackerSheet(ss);

  // Clear all data rows (keep header row 1)
  const lastRow = trackerSheet.getLastRow();
  if (lastRow > 1) {
    trackerSheet.deleteRows(2, lastRow - 1);
  }

  // Build payment map from ALL monthly sheets
  const paymentMap = _buildPaymentMapAllSheets(ss);
  Logger.log('Payment map built: ' + Object.keys(paymentMap).length + ' student+course combinations found');

  const activeRows    = [];
  const completedRows = [];

  Object.values(paymentMap).forEach(function(entry) {
    const remaining = entry.fullPrice - entry.totalPaid;
    if (remaining > 0.01) {
      const nextDue = _calcNextPaymentDate(entry.lastPaymentDate);
      activeRows.push(_buildTrackerRow(entry, nextDue, null));
    } else {
      // Completion date = date of last payment
      completedRows.push(_buildTrackerRow(entry, '', entry.lastPaymentDate));
    }
  });

  // Sort active: soonest next payment first
  activeRows.sort(function(a, b) {
    const da = a[5] instanceof Date ? a[5] : new Date(a[5]);
    const db = b[5] instanceof Date ? b[5] : new Date(b[5]);
    return da - db;
  });

  // Sort completed: most recently completed first
  completedRows.sort(function(a, b) {
    const da = a[12] instanceof Date ? a[12] : new Date(a[12]);
    const db = b[12] instanceof Date ? b[12] : new Date(b[12]);
    return db - da;
  });

  let nextWriteRow = 2;

  // Write active section
  if (activeRows.length > 0) {
    trackerSheet.getRange(nextWriteRow, 1, activeRows.length, TRACKER_COL_COUNT).setValues(activeRows);
    _formatTrackerDataRows(trackerSheet, nextWriteRow, activeRows.length);
    _applyActiveRowColours(trackerSheet, nextWriteRow, activeRows.length);
    nextWriteRow += activeRows.length;
  }

  // Write COMPLETED separator row
  const separatorRange = trackerSheet.getRange(nextWriteRow, 1, 1, TRACKER_COL_COUNT);
  separatorRange.merge();
  separatorRange.setValue('✅ COMPLETED');
  separatorRange.setFontWeight('bold');
  separatorRange.setFontSize(11);
  separatorRange.setBackground('#1b5e20');
  separatorRange.setFontColor('#ffffff');
  separatorRange.setHorizontalAlignment('center');
  nextWriteRow++;

  // Write completed section
  if (completedRows.length > 0) {
    trackerSheet.getRange(nextWriteRow, 1, completedRows.length, TRACKER_COL_COUNT).setValues(completedRows);
    _formatTrackerDataRows(trackerSheet, nextWriteRow, completedRows.length);
    _applyCompletedRowColours(trackerSheet, nextWriteRow, completedRows.length);
  }

  Logger.log('\n✅ REBUILD COMPLETE');
  Logger.log('   Active (outstanding balance): ' + activeRows.length);
  Logger.log('   Completed: ' + completedRows.length);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Tracker rebuilt. ' + activeRows.length + ' active, ' + completedRows.length + ' completed.',
    '✅ Done', 6
  );
}

/**
 * Builds a 13-element row array.
 * Pass nextDue='' and completionDate=Date for completed students.
 * Pass nextDue=Date and completionDate=null for active students.
 */
function _buildTrackerRow(entry, nextDue, completionDate) {
  const payments  = entry.payments || [];
  const remaining = Math.max(0, entry.fullPrice - entry.totalPaid);

  return [
    entry.studentName,                         // A
    entry.course,                              // B
    entry.fullPrice,                           // C
    entry.totalPaid,                           // D
    remaining,                                 // E
    nextDue || '',                             // F — Next Payment Due (blank if completed)
    payments[0] ? payments[0].date   : '',     // G — P1 Date
    payments[0] ? payments[0].amount : '',     // H — P1 Amount
    payments[1] ? payments[1].date   : '',     // I — P2 Date
    payments[1] ? payments[1].amount : '',     // J — P2 Amount
    payments[2] ? payments[2].date   : '',     // K — P3 Date
    payments[2] ? payments[2].amount : '',     // L — P3 Amount
    completionDate || ''                       // M — Completion Date (blank if active)
  ];
}

// -------------------------------------------------------
// PROCESS SINGLE PAYMENT — called automatically when data
// moves to a monthly sheet
// -------------------------------------------------------

/**
 * Called from DataProcessor.js / moveToMonthlySheet when
 * paymentInfo.isPaymentPlan === true.
 *
 * Re-derives all payments from monthly sheets (always accurate).
 * If still active → upserts in the active section.
 * If now complete → moves to completed section below the separator.
 */
function processInstalmentPayment(studentName, course, fullPrice, actualPrice, paymentDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trackerSheet = getOrCreateInstalmentTrackerSheet(ss);

  const full   = Number(fullPrice);
  const actual = Number(actualPrice);
  const name   = studentName ? studentName.toString().trim() : '';

  if (!name || !full || !actual) {
    Logger.log('processInstalmentPayment: skipping - missing data (name="' + name + '", full=' + full + ', actual=' + actual + ')');
    return;
  }

  // Always re-derive totals from monthly sheets so state can't drift
  const allPayments = _getAllPaymentsForStudent(ss, name, full);

  if (allPayments.length === 0) {
    allPayments.push({ amount: actual, date: paymentDate instanceof Date ? paymentDate : new Date(paymentDate) });
  }

  allPayments.sort(function(a, b) { return a.date - b.date; });

  const totalPaid = allPayments.reduce(function(s, p) { return s + p.amount; }, 0);
  const remaining = full - totalPaid;
  const lastDate  = allPayments[allPayments.length - 1].date;

  const entry = { studentName: name, course: course, fullPrice: full, totalPaid: totalPaid, payments: allPayments };

  // Find the student anywhere in the sheet (active or completed section)
  const existingRow = _findTrackerRowAnywhere(trackerSheet, name, full);

  if (remaining <= 0.01) {
    // Student is now complete — remove from wherever they are, insert into completed section
    if (existingRow > 0) {
      trackerSheet.deleteRow(existingRow);
    }

    const completedRow = _buildTrackerRow(entry, '', lastDate);
    const separatorRow = _findSeparatorRow(trackerSheet);
    const insertAt     = separatorRow > 0 ? separatorRow + 1 : trackerSheet.getLastRow() + 1;

    trackerSheet.insertRowBefore(insertAt);
    trackerSheet.getRange(insertAt, 1, 1, TRACKER_COL_COUNT).setValues([completedRow]);
    _formatTrackerDataRows(trackerSheet, insertAt, 1);
    _applyCompletedRowColours(trackerSheet, insertAt, 1);
    Logger.log('✅ COMPLETED: ' + name + ' — moved to completed section (paid £' + totalPaid.toFixed(2) + ' of £' + full + ')');

  } else {
    // Student still active — upsert in active section
    const nextDue = _calcNextPaymentDate(lastDate);
    const rowData = _buildTrackerRow(entry, nextDue, null);

    if (existingRow > 0) {
      trackerSheet.getRange(existingRow, 1, 1, TRACKER_COL_COUNT).setValues([rowData]);
      _formatTrackerDataRows(trackerSheet, existingRow, 1);
      _applyActiveRowColours(trackerSheet, existingRow, 1);
      Logger.log('Updated tracker: ' + name + ' — paid £' + totalPaid.toFixed(2) + ', £' + remaining.toFixed(2) + ' remaining');
    } else {
      // New student — insert before separator so they stay in the active section
      const separatorRow = _findSeparatorRow(trackerSheet);
      const insertAt     = separatorRow > 0 ? separatorRow : trackerSheet.getLastRow() + 1;
      trackerSheet.insertRowBefore(insertAt);
      trackerSheet.getRange(insertAt, 1, 1, TRACKER_COL_COUNT).setValues([rowData]);
      _formatTrackerDataRows(trackerSheet, insertAt, 1);
      _applyActiveRowColours(trackerSheet, insertAt, 1);
      Logger.log('Added to tracker: ' + name + ' — ' + course + ' — £' + totalPaid.toFixed(2) + ' of £' + full);
    }
  }
}

// -------------------------------------------------------
// PAYMENT MAP — scans all monthly sheets
// -------------------------------------------------------

/**
 * Scans ALL monthly sheets for accurate payment totals.
 * Returns ALL instalment students (active AND completed)
 * provided they had at least one payment in the last 6 months.
 */
function _buildPaymentMapAllSheets(ss) {
  const map     = {};
  const allSheets = ss.getSheets();
  const cutoff  = _getSixMonthCutoff();

  allSheets.forEach(function(sheet) {
    if (!isMonthlySheetName(sheet.getName())) return;

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return;

    const headers        = data[0];
    const nameCol        = headers.indexOf('Name');
    const courseCol      = headers.indexOf('Course');
    const fullPriceCol   = headers.indexOf('Full Price');
    const actualPriceCol = headers.indexOf('Actual Price');
    const paymentPlanCol = headers.indexOf('Payment Plan');
    const dateCol        = headers.indexOf('Date');

    if (nameCol === -1 || actualPriceCol === -1 || fullPriceCol === -1) return;

    const sheetIsRecent = _isSheetWithinCutoff(sheet.getName(), cutoff);

    data.slice(1).forEach(function(row) {
      const rawName = row[nameCol];
      if (!rawName || rawName.toString().trim() === '') return;

      const studentName = rawName.toString().trim();
      const full        = Number(row[fullPriceCol]);
      const actual      = Number(row[actualPriceCol]);

      // Treat as instalment if actual paid is less than full price
      if (!full || !actual || actual >= full) return;
      const course      = row[courseCol] ? row[courseCol].toString() : getCourseFromPrice(full);
      const date        = row[dateCol] ? new Date(row[dateCol]) : new Date();

      if (!full || !actual) return;

      const key = studentName.toLowerCase() + '|||' + full;

      if (!map[key]) {
        map[key] = {
          studentName:      studentName,
          course:           course,
          fullPrice:        full,
          totalPaid:        0,
          instalmentCount:  0,
          lastPaymentDate:  date,
          hasRecentPayment: false,
          payments:         []
        };
      }

      map[key].totalPaid       += actual;
      map[key].instalmentCount += 1;
      map[key].payments.push({ amount: actual, date: date });

      if (date > map[key].lastPaymentDate) {
        map[key].lastPaymentDate = date;
      }

      if (sheetIsRecent) {
        map[key].hasRecentPayment = true;
      }
    });
  });

  // Sort each student's payments chronologically
  Object.values(map).forEach(function(entry) {
    entry.payments.sort(function(a, b) { return a.date - b.date; });
  });

  // Filter: only keep students with at least one payment in the last 6 months
  Object.keys(map).forEach(function(key) {
    if (!map[key].hasRecentPayment) {
      Logger.log('Excluding (no recent activity): ' + map[key].studentName);
      delete map[key];
    }
  });

  return map;
}

/**
 * Returns all payment plan payments for a specific student+fullPrice
 * across all monthly sheets (no cutoff — full history for accuracy).
 */
function _getAllPaymentsForStudent(ss, studentName, fullPrice) {
  const payments  = [];
  const nameLower = studentName.toLowerCase();
  const full      = Number(fullPrice);

  ss.getSheets().forEach(function(sheet) {
    if (!isMonthlySheetName(sheet.getName())) return;

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return;

    const headers        = data[0];
    const nameCol        = headers.indexOf('Name');
    const fullPriceCol   = headers.indexOf('Full Price');
    const actualPriceCol = headers.indexOf('Actual Price');
    const paymentPlanCol = headers.indexOf('Payment Plan');
    const dateCol        = headers.indexOf('Date');

    if (nameCol === -1 || actualPriceCol === -1) return;

    data.slice(1).forEach(function(row) {
      const rowName = row[nameCol] ? row[nameCol].toString().trim() : '';
      if (rowName.toLowerCase() !== nameLower) return;
      if (Math.abs(Number(row[fullPriceCol]) - full) > 0.01) return;

      const amount = Number(row[actualPriceCol]);
      // Treat as instalment if actual paid is less than full price
      if (amount >= full) return;
      const date   = row[dateCol] ? new Date(row[dateCol]) : new Date();
      if (amount > 0) payments.push({ amount: amount, date: date });
    });
  });

  return payments;
}

// -------------------------------------------------------
// INTERNAL HELPERS
// -------------------------------------------------------

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
 * Returns true if the sheet name represents a month on or after the cutoff.
 * Handles short format (Jan26) and long format (January 2026).
 */
function _isSheetWithinCutoff(sheetName, cutoff) {
  const monthMap = {
    jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5,
    jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11,
    january: 0, february: 1, march: 2, april: 3, june: 5,
    july: 6, august: 7, september: 8, october: 9, november: 10, december: 11
  };

  const name = sheetName.trim().toLowerCase();

  const shortMatch = name.match(/^([a-z]+)(\d{2})$/);
  if (shortMatch) {
    const month = monthMap[shortMatch[1]];
    const year  = 2000 + parseInt(shortMatch[2], 10);
    if (month === undefined || isNaN(year)) return false;
    return new Date(year, month, 1) >= cutoff;
  }

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
 * Finds the 1-based row index for a student+fullPrice anywhere in the tracker
 * (searches both active and completed sections, skips separator row).
 * Returns 0 if not found.
 */
function _findTrackerRowAnywhere(sheet, studentName, fullPrice) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 0;

  const data      = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  const nameLower = studentName.toLowerCase();
  const full      = Number(fullPrice);

  for (let i = 0; i < data.length; i++) {
    const cellA = data[i][0] ? data[i][0].toString().trim() : '';
    if (cellA === '✅ COMPLETED') continue; // skip separator
    const rowFull = Number(data[i][2]);
    if (cellA.toLowerCase() === nameLower && Math.abs(rowFull - full) < 0.01) {
      return i + 2;
    }
  }
  return 0;
}

/**
 * Finds the 1-based row index of the "✅ COMPLETED" separator row.
 * Returns 0 if not found.
 */
function _findSeparatorRow(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 0;

  const colA = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < colA.length; i++) {
    if (colA[i][0] && colA[i][0].toString().trim() === '✅ COMPLETED') {
      return i + 2;
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
 * Formats currency and date columns for the given row range.
 */
function _formatTrackerDataRows(sheet, startRow, numRows) {
  if (numRows <= 0) return;
  // Currency: Full Price (C=3), Amount Paid (D=4), Remaining (E=5)
  sheet.getRange(startRow, 3, numRows, 3).setNumberFormat('£#,##0.00');
  // Currency: P1 Amount (H=8), P2 Amount (J=10), P3 Amount (L=12)
  sheet.getRange(startRow, 8, numRows, 1).setNumberFormat('£#,##0.00');
  sheet.getRange(startRow, 10, numRows, 1).setNumberFormat('£#,##0.00');
  sheet.getRange(startRow, 12, numRows, 1).setNumberFormat('£#,##0.00');
  // Dates: Next Due (F=6), P1 (G=7), P2 (I=9), P3 (K=11), Completion (M=13)
  sheet.getRange(startRow, 6, numRows, 1).setNumberFormat('dd/mm/yyyy');
  sheet.getRange(startRow, 7, numRows, 1).setNumberFormat('dd/mm/yyyy');
  sheet.getRange(startRow, 9, numRows, 1).setNumberFormat('dd/mm/yyyy');
  sheet.getRange(startRow, 11, numRows, 1).setNumberFormat('dd/mm/yyyy');
  sheet.getRange(startRow, 13, numRows, 1).setNumberFormat('dd/mm/yyyy');
}

/**
 * Green / amber / red colouring for active students based on Next Payment Due (col F).
 *   Green  = next due in the future
 *   Amber  = overdue ≤30 days
 *   Red    = overdue >30 days
 */
function _applyActiveRowColours(sheet, startRow, numRows) {
  if (numRows <= 0) return;
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const thirtyDaysAgo = new Date(today);
  thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);

  for (let i = 0; i < numRows; i++) {
    const rowIndex   = startRow + i;
    const nextDueVal = sheet.getRange(rowIndex, 6).getValue();
    const rowRange   = sheet.getRange(rowIndex, 1, 1, TRACKER_COL_COUNT);

    if (!nextDueVal || !(nextDueVal instanceof Date)) {
      rowRange.setBackground(null);
      rowRange.setFontColor('#000000');
      continue;
    }

    const nextDue = new Date(nextDueVal);
    nextDue.setHours(0, 0, 0, 0);

    if (nextDue >= today) {
      rowRange.setBackground('#e8f5e9'); // Light green
      rowRange.setFontColor('#1b5e20');
    } else if (nextDue >= thirtyDaysAgo) {
      rowRange.setBackground('#fff3e0'); // Amber
      rowRange.setFontColor('#e65100');
    } else {
      rowRange.setBackground('#ffebee'); // Light red
      rowRange.setFontColor('#b71c1c');
    }
  }
}

/**
 * Grey styling for completed students.
 */
function _applyCompletedRowColours(sheet, startRow, numRows) {
  if (numRows <= 0) return;
  sheet.getRange(startRow, 1, numRows, TRACKER_COL_COUNT)
    .setBackground('#f5f5f5')
    .setFontColor('#757575');
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

  const data  = trackerSheet.getDataRange().getValues().slice(1);
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  let activeCount = 0, completedCount = 0, overdueCount = 0;
  let inCompletedSection = false;

  data.forEach(function(row) {
    const name = row[0] ? row[0].toString().trim() : '';

    if (name === '✅ COMPLETED') {
      inCompletedSection = true;
      Logger.log('--- COMPLETED SECTION ---\n');
      return;
    }
    if (!name) return;

    const course         = row[1];
    const full           = Number(row[2]);
    const paid           = Number(row[3]);
    const remaining      = Number(row[4]);
    const nextDue        = row[5]  ? new Date(row[5])  : null;
    const p1Date         = row[6]  ? new Date(row[6])  : null;
    const p1Amount       = row[7]  ? Number(row[7])    : null;
    const p2Date         = row[8]  ? new Date(row[8])  : null;
    const p2Amount       = row[9]  ? Number(row[9])    : null;
    const p3Date         = row[10] ? new Date(row[10]) : null;
    const p3Amount       = row[11] ? Number(row[11])   : null;
    const completionDate = row[12] ? new Date(row[12]) : null;

    if (inCompletedSection) {
      completedCount++;
      Logger.log('✅ ' + name + ' — ' + course);
      Logger.log('   Full: £' + full + '  |  Completed: ' + (completionDate ? completionDate.toLocaleDateString('en-GB') : 'unknown'));
      if (p1Date && p1Amount) Logger.log('   P1: £' + p1Amount + ' on ' + p1Date.toLocaleDateString('en-GB'));
      if (p2Date && p2Amount) Logger.log('   P2: £' + p2Amount + ' on ' + p2Date.toLocaleDateString('en-GB'));
      if (p3Date && p3Amount) Logger.log('   P3: £' + p3Amount + ' on ' + p3Date.toLocaleDateString('en-GB'));
    } else {
      activeCount++;
      const isOverdue = nextDue && nextDue < today;
      if (isOverdue) overdueCount++;
      Logger.log(activeCount + '. ' + name + ' — ' + course);
      Logger.log('   Full: £' + full + '  Paid: £' + paid + '  Remaining: £' + remaining);
      if (p1Date && p1Amount) Logger.log('   P1: £' + p1Amount + ' on ' + p1Date.toLocaleDateString('en-GB'));
      if (p2Date && p2Amount) Logger.log('   P2: £' + p2Amount + ' on ' + p2Date.toLocaleDateString('en-GB'));
      if (p3Date && p3Amount) Logger.log('   P3: £' + p3Amount + ' on ' + p3Date.toLocaleDateString('en-GB'));
      Logger.log('   Next Due: ' + (nextDue ? nextDue.toLocaleDateString('en-GB') : 'unknown') + (isOverdue ? ' ⚠️ OVERDUE' : ' ✅'));
    }
    Logger.log('');
  });

  Logger.log('📊 Active: ' + activeCount + '  |  Overdue: ' + overdueCount + '  |  Completed: ' + completedCount);
}

// -------------------------------------------------------
// LEGACY STUBS — keep so existing call sites don't break
// -------------------------------------------------------

function setupInstalmentTracking() {
  Logger.log('setupInstalmentTracking: nothing to set up (tracker self-maintains)');
}

function updateInstalmentTrackerFromMonthlySheets() {
  rebuildInstalmentTracker();
}

function debugInstalmentTracker() {
  diagnoseInstalmentTracker();
}

function cleanupCompletedPayments() {
  Logger.log('cleanupCompletedPayments: completed students now shown in tracker below separator.');
  Logger.log('Run rebuildInstalmentTracker() if you need a full resync.');
}

// -------------------------------------------------------
// MANUAL ADD — force-add a student the rebuild missed
// -------------------------------------------------------

/**
 * Fill in the constants below and run from the Apps Script editor.
 */
function manuallyAddToInstalmentTracker() {
  const STUDENT_NAME = 'Enter Student Name Here';
  const COURSE       = 'Platinum';
  const FULL_PRICE   = 1047;

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
    return;
  }

  processInstalmentPayment(name, COURSE, full, allPayments[allPayments.length - 1].amount, allPayments[allPayments.length - 1].date);
  Logger.log('✅ Done. "' + name + '" added/updated in tracker.');
}

// -------------------------------------------------------
// STUDENT PAYMENT HISTORY — full chronological lookup
// -------------------------------------------------------

/**
 * Prints a full payment history for a specific student across all monthly sheets.
 * Set STUDENT_NAME and FULL_PRICE below and run from the Apps Script editor.
 */
function printStudentPaymentHistory() {
  const STUDENT_NAME = 'Enter Student Name Here';
  const FULL_PRICE   = 1047;

  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const payments = _getAllPaymentsForStudent(ss, STUDENT_NAME.trim(), Number(FULL_PRICE));

  if (payments.length === 0) {
    Logger.log('No payment plan payments found for "' + STUDENT_NAME + '" with full price £' + FULL_PRICE + '.');
    return;
  }

  payments.sort(function(a, b) { return a.date - b.date; });
  Logger.log('Payment history for ' + STUDENT_NAME + ' (Full Price: £' + FULL_PRICE + '):');
  payments.forEach(function(p, i) {
    Logger.log('  P' + (i + 1) + ': £' + p.amount + ' on ' + p.date.toLocaleDateString('en-GB'));
  });
  const total = payments.reduce(function(s, p) { return s + p.amount; }, 0);
  Logger.log('  Total paid: £' + total.toFixed(2) + '  |  Remaining: £' + (Number(FULL_PRICE) - total).toFixed(2));
}