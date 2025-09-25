function sendEmails() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const rows = sheet.getDataRange().getValues();
  
  const messageBodyTemplate = rows[1][3]; 

  for (let i = 1; i < rows.length; i++) { 
    const email = rows[i][0]; // column A: recipient
    const subjectLine = rows[i][1];   // column B: subject line
    const personName = rows[i][2];   // column C: name

    if (email && subjectLine) {
      const messageBody = messageBodyTemplate.replace("[personName]", personName || "there")
      GmailApp.sendEmail(email, subjectLine, messageBody);
    }
  }
}
