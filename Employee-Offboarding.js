/**
 * Google Apps Script: Automated Employee Offboarding Communications
 * 
 * Automatically dispatches exit procedure instructions to departing staff
 * enforcing data retention policies and device return protocols.
 */

function sendExitEmailDaily() {
  const sheetId = "<YOUR_SPREADSHEET_ID>";  
  const sheetName = "Exiting Staff"; 
  const emailColumn = 15;      
  const roleColumn = 16;       
  const dateSentColumn = 17;   
  
  const googleAccountProceduresUrl = "https://help.yourdomain.com/Exiting-Staff-Google-Account";
  const ipadProceduresUrl = "https://help.yourdomain.com/Exiting-Staff-iPad";
  const emailSubject = "Exit Procedures and Important Information";
  const fromEmail = "noreply@yourdomain.com"; 

  const ss = SpreadsheetApp.openById(sheetId);
  const sheet = ss.getSheetByName(sheetName);
  const lastRow = sheet.getLastRow();

  for (let i = 1; i <= lastRow; i++) {  
    const recipientEmail = sheet.getRange(i, emailColumn).getValue();
    const roleValue = sheet.getRange(i, roleColumn).getValue();
    const dateSentValue = sheet.getRange(i, dateSentColumn).getValue();

    if (recipientEmail && typeof recipientEmail === 'string' && recipientEmail.trim() !== "" &&
        roleValue && typeof roleValue === 'string' && roleValue.toLowerCase() === 'staff' &&
        !dateSentValue) {
      
      const emailBody = `
        <!DOCTYPE html>
        <html lang="en">
        <body>
          <div class="container">
            <p>Hello,</p>
            <p>The link below provides detailed instructions on managing your Google account contents before your last day. Some of these steps may take a significant amount of time depending on the volume of data you wish to transfer.</p>
            
            <p class="key-reminders"><b>Key Reminders:</b></p>
            <ul>
              <li>Your account will be disabled immediately after your exit interview and will no longer be accessible.</li>
              <li>Once disabled, accounts cannot be reopened to retrieve files.</li>
              <li>All Drive files will be permanently deleted after 24 months.</li>
            </ul>

            <p><b>Instructions:</b></p>
            <p><a href="${googleAccountProceduresUrl}">Exiting Staff Google Account Procedures</a></p>

            <p><b>If you have a district-issued device, please follow the link below:</b></p>
            <p><a href="${ipadProceduresUrl}">Exiting Staff Device Procedures</a></p>

            <p>Thank you for your contributions, and best wishes for success in your next adventure!</p>
          </div>
        </body>
        </html>
      `;

      try {
        MailApp.sendEmail({
          to: recipientEmail,
          subject: emailSubject,
          htmlBody: emailBody,
          from: fromEmail
        });
        
        // Mark as sent
        sheet.getRange(i, dateSentColumn).setValue(new Date());
      } catch (error) {
        Logger.log("Error sending email: " + error.message);
      }
    } 
  }
}
