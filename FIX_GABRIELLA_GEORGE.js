// ===============================================
// FIX GABRIELLA GEORGE - FIND PAYMENT WITH NAME VARIATION
// ===============================================

function findGabriellaGeorgePayment() {
  Logger.log('🔍 SEARCHING FOR GABRIELLA GEORGE\'S £397 PAYMENT');
  Logger.log('================================================\n');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const septemberSheet = ss.getSheetByName('September 2025');
  
  if (!septemberSheet) {
    Logger.log('❌ September 2025 sheet not found');
    return;
  }
  
  const dataRange = septemberSheet.getDataRange();
  const allData = dataRange.getValues();
  const headers = allData[0];
  const dataRows = allData.slice(1);
  
  const nameCol = headers.indexOf('Name');
  const actualPriceCol = headers.indexOf('Actual Price');
  const dateCol = headers.indexOf('Date');
  
  Logger.log('🔎 Looking for £397 Platinum payments in September 2025...\n');
  
  let found397Payments = [];
  
  dataRows.forEach((row, index) => {
    const name = row[nameCol];
    const amount = Number(row[actualPriceCol]);
    const date = row[dateCol];
    
    if (amount === 397) {
      found397Payments.push({
        name: name,
        amount: amount,
        date: date,
        rowNumber: index + 2
      });
    }
  });
  
  Logger.log(`✅ Found ${found397Payments.length} payment(s) of £397 in September:\n`);
  
  found397Payments.forEach((payment, index) => {
    Logger.log(`${index + 1}. "${payment.name}" - £${payment.amount} on ${new Date(payment.date).toLocaleDateString()} (Row ${payment.rowNumber})`);
  });
  
  // Look for names similar to "Gabriella George"
  Logger.log('\n🎯 Looking for names similar to "Gabriella George":\n');
  
  found397Payments.forEach(payment => {
    const nameLower = payment.name.toLowerCase();
    if (nameLower.includes('gabriel') || nameLower.includes('george')) {
      Logger.log(`⭐ POSSIBLE MATCH: "${payment.name}"`);
      Logger.log(`   This could be Gabriella George's first payment!`);
    }
  });
  
  Logger.log('\n💡 Once you confirm which spelling it is, run fixGabriellaGeorgeRecord() to update the tracker');
}

function fixGabriellaGeorgeRecord() {
  Logger.log('🔧 FIXING GABRIELLA GEORGE\'S INSTALMENT TRACKER RECORD');
  Logger.log('======================================================\n');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Search for ALL payments with name variations of Gabriella George
  const allSheets = ss.getSheets();
  const monthlySheets = allSheets.filter(sheet => isMonthlySheetName(sheet.getName()));
  
  Logger.log('🔎 Searching all monthly sheets for name variations...\n');
  
  // Common variations to search for
  const nameVariations = [
    'Gabriella George',
    'gabriella george',
    'Gabriela George',
    'Gabriella Georg',
    'Gabriela Georg'
    // Add more variations if needed
  ];
  
  let allPayments = [];
  
  monthlySheets.forEach(sheet => {
    const sheetData = sheet.getDataRange().getValues();
    const sheetHeaders = sheetData[0];
    const sheetRows = sheetData.slice(1);
    
    const nameCol = sheetHeaders.indexOf('Name');
    const actualPriceCol = sheetHeaders.indexOf('Actual Price');
    const dateCol = sheetHeaders.indexOf('Date');
    
    if (nameCol === -1 || actualPriceCol === -1) return;
    
    sheetRows.forEach(row => {
      const name = row[nameCol];
      const amount = Number(row[actualPriceCol]);
      const date = new Date(row[dateCol]);
      
      // Check if name contains "gabriel" and "george" (case insensitive)
      if (name && typeof name === 'string') {
        const nameLower = name.toLowerCase();
        if ((nameLower.includes('gabriel') && nameLower.includes('george')) || 
            name === 'Gabriella George') {
          allPayments.push({
            sheetName: sheet.getName(),
            name: name,
            amount: amount,
            date: date
          });
        }
      }
    });
  });
  
  if (allPayments.length === 0) {
    Logger.log('❌ No payments found with name variations of Gabriella George');
    return;
  }
  
  // Sort by date
  allPayments.sort((a, b) => a.date - b.date);
  
  Logger.log(`✅ Found ${allPayments.length} payment(s) for Gabriella George (with variations):\n`);
  
  allPayments.forEach((p, i) => {
    Logger.log(`${i + 1}. ${p.sheetName} - "${p.name}" - £${p.amount} on ${p.date.toLocaleDateString()}`);
  });
  
  // Calculate totals
  const totalPaid = allPayments.reduce((sum, p) => sum + p.amount, 0);
  const instalmentCount = allPayments.length;
  
  Logger.log(`\n📊 TOTALS:`);
  Logger.log(`   Total Amount Paid: £${totalPaid}`);
  Logger.log(`   Number of Instalments: ${instalmentCount}`);
  
  // Update the Instalment Tracker
  const trackerSheet = ss.getSheetByName('Instalment Tracker');
  if (!trackerSheet) {
    Logger.log('\n❌ Instalment Tracker sheet not found');
    return;
  }
  
  // Find Gabriella George's row in the tracker
  const trackerData = trackerSheet.getDataRange().getValues();
  const trackerRows = trackerData.slice(1);
  
  let gabriellaRow = -1;
  trackerRows.forEach((row, index) => {
    if (row[0] && row[0].toLowerCase().includes('gabriel') && row[0].toLowerCase().includes('george')) {
      gabriellaRow = index + 2; // +2 for header and 1-indexed
    }
  });
  
  if (gabriellaRow === -1) {
    Logger.log('\n❌ Gabriella George not found in Instalment Tracker');
    return;
  }
  
  Logger.log(`\n🔧 Updating Instalment Tracker (Row ${gabriellaRow}):`);
  Logger.log(`   Amount Paid: £600 → £${totalPaid}`);
  Logger.log(`   Instalments: 2 → ${instalmentCount}`);
  
  // Update the tracker
  trackerSheet.getRange(gabriellaRow, 4).setValue(totalPaid); // Amount Paid
  trackerSheet.getRange(gabriellaRow, 5).setValue(instalmentCount); // Instalments Paid
  
  // Get last payment date
  const lastPaymentDate = allPayments[allPayments.length - 1].date;
  trackerSheet.getRange(gabriellaRow, 6).setValue(lastPaymentDate); // Last Payment Date
  
  // Check if complete
  const fullPrice = 997; // Platinum
  if (totalPaid >= fullPrice) {
    trackerSheet.getRange(gabriellaRow, 7).setValue(''); // Clear Next Payment Due
    trackerSheet.getRange(gabriellaRow, 8).setValue('Complete'); // Payment Complete
    trackerSheet.getRange(gabriellaRow, 9).setValue(new Date()); // Completion Date
    Logger.log(`   Status: In Progress → Complete ✅`);
  } else {
    const nextDue = new Date(lastPaymentDate);
    nextDue.setMonth(nextDue.getMonth() + 1);
    trackerSheet.getRange(gabriellaRow, 7).setValue(nextDue);
    Logger.log(`   Next Payment Due: ${nextDue.toLocaleDateString()}`);
    Logger.log(`   Status: Still In Progress (£${fullPrice - totalPaid} remaining)`);
  }
  
  Logger.log('\n✅ Gabriella George\'s record updated successfully!');
}

// Quick manual fix if you know the exact alternate spelling
function fixGabriellaGeorgeManual(alternateSpelling) {
  Logger.log(`🔧 MANUAL FIX FOR GABRIELLA GEORGE`);
  Logger.log(`   Searching for payments under: "${alternateSpelling}"\n`);
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();
  const monthlySheets = allSheets.filter(sheet => isMonthlySheetName(sheet.getName()));
  
  let allPayments = [];
  
  // Search for both "Gabriella George" and the alternate spelling
  monthlySheets.forEach(sheet => {
    const sheetData = sheet.getDataRange().getValues();
    const sheetHeaders = sheetData[0];
    const sheetRows = sheetData.slice(1);
    
    const nameCol = sheetHeaders.indexOf('Name');
    const actualPriceCol = sheetHeaders.indexOf('Actual Price');
    const dateCol = sheetHeaders.indexOf('Date');
    
    if (nameCol === -1 || actualPriceCol === -1) return;
    
    sheetRows.forEach(row => {
      const name = row[nameCol];
      if (name === 'Gabriella George' || name === alternateSpelling) {
        allPayments.push({
          sheetName: sheet.getName(),
          name: name,
          amount: Number(row[actualPriceCol]),
          date: new Date(row[dateCol])
        });
      }
    });
  });
  
  allPayments.sort((a, b) => a.date - b.date);
  
  Logger.log(`✅ Found ${allPayments.length} payments:\n`);
  allPayments.forEach((p, i) => {
    Logger.log(`${i + 1}. ${p.sheetName} - "${p.name}" - £${p.amount} on ${p.date.toLocaleDateString()}`);
  });
  
  const totalPaid = allPayments.reduce((sum, p) => sum + p.amount, 0);
  
  Logger.log(`\nTotal: £${totalPaid} in ${allPayments.length} instalments`);
  
  // Now update with the same logic as fixGabriellaGeorgeRecord()...
  // (Rest of update logic here)
}