// ===============================================
// EMAILMONITORING.GS - Weekly Email Alert System
// ===============================================

function setupWeeklyEmailMonitoring() {
  // Clear existing email triggers
  clearEmailTriggers();
  
  // Set up weekly trigger for Monday at 7 AM UK time
  ScriptApp.newTrigger('sendWeeklyMonitoringEmail')
    .timeBased()
    .everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(7) // 7 AM
    .inTimezone('Europe/London') // UK time
    .create();
  
  Logger.log('Weekly email monitoring trigger set up for Mondays at 7 AM UK time');
}

function clearEmailTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'sendWeeklyMonitoringEmail') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function sendWeeklyMonitoringEmail() {
  try {
    Logger.log('Starting weekly monitoring email generation...');
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const currentDate = new Date();
    
    // Collect all monitoring data
    const unpaidInvoices = findUnpaidInvoices(ss, currentDate);
    const missedInstalments = findMissedInstalments(ss, currentDate);
    const incompletePayments = findIncompletePayments(ss, currentDate);
    
    // Only send email if there are issues to report
    if (unpaidInvoices.length > 0 || missedInstalments.length > 0 || incompletePayments.length > 0) {
      const emailBody = generateEmailBody(unpaidInvoices, missedInstalments, incompletePayments, currentDate);
      const subject = `Revenue Tracker Alert - ${formatDate(currentDate)}`;
      
      // Send to specific email address
      MailApp.sendEmail({
        to: 'stuartmedgarwork@gmail.com',
        subject: subject,
        htmlBody: emailBody
      });
      
      Logger.log(`Weekly monitoring email sent to stuartmedgarwork@gmail.com`);
      Logger.log(`Issues found: ${unpaidInvoices.length} unpaid invoices, ${missedInstalments.length} missed instalments, ${incompletePayments.length} incomplete payments`);
    } else {
      Logger.log('No issues found - no email sent');
    }
    
  } catch (error) {
    Logger.log('Error sending weekly monitoring email: ' + error.toString());
    
    // Send simple error notification
    try {
      MailApp.sendEmail({
        to: 'stuartmedgarwork@gmail.com',
        subject: 'Revenue Tracker - Email System Error',
        body: `There was an error generating your weekly revenue tracker report:\n\n${error.toString()}\n\nPlease check the system logs.`
      });
    } catch (emailError) {
      Logger.log('Failed to send error notification email: ' + emailError.toString());
    }
  }
}

function findUnpaidInvoices(ss, currentDate) {
  const awaitingSheet = ss.getSheetByName('Awaiting Employer Invoice');
  const unpaidInvoices = [];
  
  if (!awaitingSheet) {
    Logger.log('Awaiting Employer Invoice sheet not found');
    return unpaidInvoices;
  }
  
  const dataRange = awaitingSheet.getDataRange();
  if (dataRange.getNumRows() <= 1) return unpaidInvoices;
  
  const allData = dataRange.getValues();
  const headers = allData[0];
  const dataRows = allData.slice(1);
  
  const dateCol = headers.indexOf('Date');
  const nameCol = headers.indexOf('Name');
  const courseCol = headers.indexOf('Course');
  const fullPriceCol = headers.indexOf('Full Price');
  
  if (dateCol === -1 || nameCol === -1 || courseCol === -1 || fullPriceCol === -1) {
    Logger.log('Missing required columns in Awaiting Employer Invoice sheet');
    return unpaidInvoices;
  }
  
  const thirtyDaysAgo = new Date(currentDate.getTime() - (30 * 24 * 60 * 60 * 1000));
  
  dataRows.forEach(row => {
    const entryDate = new Date(row[dateCol]);
    
    if (entryDate < thirtyDaysAgo) {
      unpaidInvoices.push({
        name: row[nameCol],
        course: row[courseCol],
        amount: row[fullPriceCol],
        daysOverdue: Math.floor((currentDate - entryDate) / (1000 * 60 * 60 * 24)),
        date: entryDate
      });
    }
  });
  
  Logger.log(`Found ${unpaidInvoices.length} unpaid invoices over 30 days old`);
  return unpaidInvoices;
}

function findMissedInstalments(ss, currentDate) {
  const trackerSheet = ss.getSheetByName('Instalment Tracker');
  const missedInstalments = [];
  
  if (!trackerSheet) {
    Logger.log('Instalment Tracker sheet not found');
    return missedInstalments;
  }
  
  const dataRange = trackerSheet.getDataRange();
  if (dataRange.getNumRows() <= 1) return missedInstalments;
  
  const allData = dataRange.getValues();
  const headers = allData[0];
  const dataRows = allData.slice(1);
  
  const nameCol = headers.indexOf('Student Name');
  const courseCol = headers.indexOf('Course');
  const amountPaidCol = headers.indexOf('Amount Paid');
  const nextDueCol = headers.indexOf('Next Payment Due');
  const statusCol = headers.indexOf('Payment Complete');
  
  if (nameCol === -1 || courseCol === -1 || amountPaidCol === -1 || nextDueCol === -1 || statusCol === -1) {
    Logger.log('Missing required columns in Instalment Tracker sheet');
    return missedInstalments;
  }
  
  const sevenDaysAgo = new Date(currentDate.getTime() - (7 * 24 * 60 * 60 * 1000));
  
  dataRows.forEach(row => {
    const status = row[statusCol];
    const nextDueDate = row[nextDueCol];
    
    if (status === 'In Progress' && nextDueDate && new Date(nextDueDate) < sevenDaysAgo) {
      const daysOverdue = Math.floor((currentDate - new Date(nextDueDate)) / (1000 * 60 * 60 * 24));
      
      missedInstalments.push({
        name: row[nameCol],
        course: row[courseCol],
        amountPaid: row[amountPaidCol],
        nextDue: new Date(nextDueDate),
        daysOverdue: daysOverdue
      });
    }
  });
  
  Logger.log(`Found ${missedInstalments.length} missed instalments over 7 days late`);
  return missedInstalments;
}

function findIncompletePayments(ss, currentDate) {
  const trackerSheet = ss.getSheetByName('Instalment Tracker');
  const incompletePayments = [];
  
  if (!trackerSheet) {
    Logger.log('Instalment Tracker sheet not found');
    return incompletePayments;
  }
  
  const dataRange = trackerSheet.getDataRange();
  if (dataRange.getNumRows() <= 1) return incompletePayments;
  
  const allData = dataRange.getValues();
  const headers = allData[0];
  const dataRows = allData.slice(1);
  
  const nameCol = headers.indexOf('Student Name');
  const courseCol = headers.indexOf('Course');
  const fullPriceCol = headers.indexOf('Full Price');
  const amountPaidCol = headers.indexOf('Amount Paid');
  const statusCol = headers.indexOf('Payment Complete');
  const nextDueCol = headers.indexOf('Next Payment Due');
  
  if (nameCol === -1 || courseCol === -1 || fullPriceCol === -1 || 
      amountPaidCol === -1 || statusCol === -1 || nextDueCol === -1) {
    Logger.log('Missing required columns in Instalment Tracker sheet');
    return incompletePayments;
  }
  
  dataRows.forEach(row => {
    const status = row[statusCol];
    const fullPrice = Number(row[fullPriceCol]);
    const amountPaid = Number(row[amountPaidCol]);
    const nextDue = row[nextDueCol];
    
    // Students who have finished instalments (no next payment due) but haven't paid enough
    if (status === 'In Progress' && !nextDue && amountPaid < fullPrice) {
      const shortfall = fullPrice - amountPaid;
      
      incompletePayments.push({
        name: row[nameCol],
        course: row[courseCol],
        fullPrice: fullPrice,
        amountPaid: amountPaid,
        shortfall: shortfall
      });
    }
  });
  
  Logger.log(`Found ${incompletePayments.length} incomplete payments`);
  return incompletePayments;
}

function generateEmailBody(unpaidInvoices, missedInstalments, incompletePayments, currentDate) {
  let emailBody = `
    <html>
    <head>
      <style>
        body { font-family: Arial, sans-serif; }
        .header { background-color: #f44336; color: white; padding: 15px; text-align: center; }
        .section { margin: 20px 0; }
        .section h3 { color: #d32f2f; border-bottom: 2px solid #d32f2f; padding-bottom: 5px; }
        table { border-collapse: collapse; width: 100%; margin: 10px 0; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #f5f5f5; font-weight: bold; }
        .amount { text-align: right; }
        .days-overdue { font-weight: bold; color: #d32f2f; }
        .summary { background-color: #fff3e0; padding: 15px; border-left: 4px solid #ff9800; margin: 20px 0; }
      </style>
    </head>
    <body>
      <div class="header">
        <h2>📊 Revenue Tracker Weekly Alert</h2>
        <p>Report for ${formatDate(currentDate)}</p>
      </div>
  `;
  
  // Summary section
  const totalIssues = unpaidInvoices.length + missedInstalments.length + incompletePayments.length;
  emailBody += `
    <div class="summary">
      <h3>📋 Summary</h3>
      <p><strong>Total Issues Found:</strong> ${totalIssues}</p>
      <ul>
        <li><strong>Unpaid Invoices (30+ days):</strong> ${unpaidInvoices.length}</li>
        <li><strong>Missed Instalments (7+ days):</strong> ${missedInstalments.length}</li>
        <li><strong>Incomplete Payments:</strong> ${incompletePayments.length}</li>
      </ul>
    </div>
  `;
  
  // Unpaid invoices section
  if (unpaidInvoices.length > 0) {
    emailBody += `
      <div class="section">
        <h3>💰 Unpaid Invoices (30+ Days Old)</h3>
        <table>
          <tr>
            <th>Student Name</th>
            <th>Course</th>
            <th>Amount</th>
            <th>Days Overdue</th>
            <th>Original Date</th>
          </tr>
    `;
    
    unpaidInvoices.forEach(invoice => {
      emailBody += `
        <tr>
          <td>${invoice.name}</td>
          <td>${invoice.course}</td>
          <td class="amount">£${invoice.amount}</td>
          <td class="days-overdue">${invoice.daysOverdue} days</td>
          <td>${formatDate(invoice.date)}</td>
        </tr>
      `;
    });
    
    emailBody += '</table></div>';
  }
  
  // Missed instalments section
  if (missedInstalments.length > 0) {
    emailBody += `
      <div class="section">
        <h3>⏰ Missed Instalments (7+ Days Late)</h3>
        <table>
          <tr>
            <th>Student Name</th>
            <th>Course</th>
            <th>Amount Paid So Far</th>
            <th>Payment Was Due</th>
            <th>Days Overdue</th>
          </tr>
    `;
    
    missedInstalments.forEach(instalment => {
      emailBody += `
        <tr>
          <td>${instalment.name}</td>
          <td>${instalment.course}</td>
          <td class="amount">£${instalment.amountPaid}</td>
          <td>${formatDate(instalment.nextDue)}</td>
          <td class="days-overdue">${instalment.daysOverdue} days</td>
        </tr>
      `;
    });
    
    emailBody += '</table></div>';
  }
  
  // Incomplete payments section
  if (incompletePayments.length > 0) {
    emailBody += `
      <div class="section">
        <h3>⚠️ Incomplete Payments</h3>
        <p><em>Students who finished their instalment schedule but haven't paid the full amount.</em></p>
        <table>
          <tr>
            <th>Student Name</th>
            <th>Course</th>
            <th>Full Price</th>
            <th>Amount Paid</th>
            <th>Shortfall</th>
          </tr>
    `;
    
    incompletePayments.forEach(payment => {
      emailBody += `
        <tr>
          <td>${payment.name}</td>
          <td>${payment.course}</td>
          <td class="amount">£${payment.fullPrice}</td>
          <td class="amount">£${payment.amountPaid}</td>
          <td class="amount days-overdue">£${payment.shortfall}</td>
        </tr>
      `;
    });
    
    emailBody += '</table></div>';
  }
  
  emailBody += `
      <div class="section" style="text-align: center; color: #666; font-size: 12px; margin-top: 30px;">
        <p>📧 This email was generated automatically by your Revenue Tracker system.</p>
        <p>Report generated on ${formatDateTime(currentDate)}</p>
      </div>
    </body>
    </html>
  `;
  
  return emailBody;
}

function formatDate(date) {
  return date.toLocaleDateString('en-GB', { 
    day: 'numeric', 
    month: 'long', 
    year: 'numeric' 
  });
}

function formatDateTime(date) {
  return date.toLocaleString('en-GB', { 
    day: 'numeric', 
    month: 'long', 
    year: 'numeric',
    hour: '2-digit',
    minute: '2-digit'
  });
}

// ===============================================
// TEST FUNCTIONS
// ===============================================

function testWeeklyEmail() {
  Logger.log('Testing weekly email generation...');
  sendWeeklyMonitoringEmail();
  Logger.log('Test email sent - check your inbox!');
}

function previewWeeklyEmail() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentDate = new Date();
  
  const unpaidInvoices = findUnpaidInvoices(ss, currentDate);
  const missedInstalments = findMissedInstalments(ss, currentDate);
  const incompletePayments = findIncompletePayments(ss, currentDate);
  
  const emailBody = generateEmailBody(unpaidInvoices, missedInstalments, incompletePayments, currentDate);
  
  Logger.log('=== EMAIL PREVIEW ===');
  Logger.log(`Unpaid Invoices: ${unpaidInvoices.length}`);
  Logger.log(`Missed Instalments: ${missedInstalments.length}`);
  Logger.log(`Incomplete Payments: ${incompletePayments.length}`);
  Logger.log('Email body generated - ready to send');
  
  return emailBody;
}