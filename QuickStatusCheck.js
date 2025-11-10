function quickStatusCheck() {
  Logger.log('⚡ QUICK STATUS CHECK');
  Logger.log('====================');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Check key sheets quickly
  const sheets = ['Data Entry', 'Sort', 'Awaiting Employer Invoice'];
  
  sheets.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      const numRows = sheet.getDataRange().getNumRows();
      const dataRows = numRows > 1 ? numRows - 1 : 0;
      Logger.log(`${sheetName}: ${dataRows} records`);
      
      if (dataRows > 0 && dataRows <= 5) {
        // Show the data if there's not too much
        const allData = sheet.getDataRange().getValues();
        const headers = allData[0];
        const dataRows = allData.slice(1);
        
        dataRows.forEach((row, index) => {
          if (row[1]) { // Has name
            Logger.log(`  ${index + 1}. ${row[1]} - ${row[0]} - ${row[2] || 'No course'}`);
          }
        });
      }
    } else {
      Logger.log(`${sheetName}: Not found`);
    }
  });
  
  // Check recent monthly sheets
  Logger.log('\nRecent Monthly Sheets:');
  const monthNames = ['December', 'January', 'February'];
  const years = [2024, 2025];
  
  years.forEach(year => {
    monthNames.forEach(month => {
      const sheetName = `${month} ${year}`;
      const sheet = ss.getSheetByName(sheetName);
      if (sheet) {
        const numRows = sheet.getDataRange().getNumRows();
        const dataRows = numRows > 1 ? numRows - 1 : 0;
        if (dataRows > 0) {
          Logger.log(`${sheetName}: ${dataRows} records`);
        }
      }
    });
  });
}

function showLastFewActions() {
  Logger.log('📈 CHECKING WHAT JUST HAPPENED');
  Logger.log('==============================');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Look at recent data in monthly sheets
  const now = new Date();
  const oneHourAgo = new Date(now.getTime() - (60 * 60 * 1000));
  
  Logger.log('Looking for data added in the last hour...');
  
  const allSheets = ss.getSheets();
  
  allSheets.forEach(sheet => {
    const sheetName = sheet.getName();
    
    // Check if it's a monthly sheet or key sheet
    const isMonthlySheet = /^(January|February|March|April|May|June|July|August|September|October|November|December) \d{4}$/.test(sheetName);
    const isKeySheet = ['Awaiting Employer Invoice', 'Sort', 'Data Entry'].includes(sheetName);
    
    if (isMonthlySheet || isKeySheet) {
      const dataRange = sheet.getDataRange();
      if (dataRange.getNumRows() > 1) {
        const allData = dataRange.getValues();
        const dataRows = allData.slice(1);
        
        const recentData = dataRows.filter(row => {
          if (row[0] && row[1]) { // Has date and name
            const rowDate = new Date(row[0]);
            return rowDate >= oneHourAgo;
          }
          return false;
        });
        
        if (recentData.length > 0) {
          Logger.log(`\n🆕 ${sheetName}: ${recentData.length} recent records`);
          recentData.forEach(row => {
            Logger.log(`  - ${row[1]} (${new Date(row[0]).toLocaleString()})`);
          });
        }
      }
    }
  });
}