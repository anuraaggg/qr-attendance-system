function authorizeOnce() {
  UrlFetchApp.fetch("https://www.google.com");
}


function doGet(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var id = e.parameter.id;

  if (!id) {
    return bigMessage("Invalid QR Code", "#e74c3c");
  }

  // Logged-in user email
  var userEmail = Session.getActiveUser().getEmail();
  var allowedEmail = "anuraagshnkr@gmail.com";

  if (!userEmail) {
    return bigMessage("Please sign in with your Google account", "#f39c12");
  }

  if (userEmail !== allowedEmail) {
    return bigMessage("Unauthorized Access", "#e74c3c");
  }

  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString() === id.toString()) {

      // Present column (I)
      if (data[i][8] === "YES") {
        return bigMessage("Attendance Already Marked", "#3498db");
      }

      // Mark attendance
      sheet.getRange(i + 1, 9).setValue("YES");
      var now = new Date();
      var formattedDateTime = Utilities.formatDate(
        now,
        Session.getScriptTimeZone(),
        "dd-MM-yy HH:mm"
      );

      sheet.getRange(i + 1, 12).setValue(formattedDateTime);



      return bigMessage("Attendance Marked Successfully", "#2ecc71");
    }
  }

  return bigMessage("ID Not Found", "#e74c3c");
}


function bigMessage(message, color) {
  var html = `
    <html>
      <head>
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
      </head>
      <body style="
        margin:0;
        height:100vh;
        display:flex;
        align-items:center;
        justify-content:center;
        background:#f4f6f8;
        font-family: Arial, sans-serif;
      ">
        <div style="
          text-align:center;
          padding:30px 40px;
          border-radius:12px;
          background:white;
          box-shadow:0 10px 25px rgba(0,0,0,0.15);
        ">
          <h1 style="
            color:${color};
            font-size:36px;
            margin:0;
          ">
            ${message}
          </h1>
        </div>
      </body>
    </html>
  `;

  return HtmlService.createHtmlOutput(html);
}


function sendQrEmail(e) {

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var row = e.range.getRow();

  // Skip header row
  if (row === 1) return;

  // Read all required fields
  var id        = sheet.getRange(row, 1).getValue(); // A
  var name      = sheet.getRange(row, 2).getValue(); // B
  var email     = sheet.getRange(row, 3).getValue(); // C
  var phone     = sheet.getRange(row, 4).getValue(); // D
  var college   = sheet.getRange(row, 5).getValue(); // E
  var city      = sheet.getRange(row, 6).getValue(); // F
  var state     = sheet.getRange(row, 7).getValue(); // G
  var teamCount = sheet.getRange(row, 8).getValue(); // H
  var mailSent  = sheet.getRange(row, 10).getValue(); // J

  // Stop if mail already sent
  if (mailSent === "YES") return;

  // Stop if ANY required field is missing
  if (
    !id || !name || !email || !phone ||
    !college || !city || !state || !teamCount
  ) {
    return;
  }

  // Generate QR URL
  var qrUrl =
    "https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=" +
    "https://script.google.com/macros/s/AKfycbyWFwCaBau1rcom9VXpQxkIp5ayxkMhOYEmpPa2j8YGOohT-zfwZf8rWkj58k09HwlOtQ/exec?id=" +
    id;

  // Put QR image formula in QR_Code column (I)
  var qrFormula =
    '=IMAGE("https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=' +
    'https://script.google.com/macros/s/AKfycbyWFwCaBau1rcom9VXpQxkIp5ayxkMhOYEmpPa2j8YGOohT-zfwZf8rWkj58k09HwlOtQ/exec?id=' +
    id + '")';

  sheet.getRange(row, 9).setFormula(qrFormula);

  // Format date as dd-mm-yyyy
  var today = new Date();
  var day = String(today.getDate()).padStart(2, '0');
  var month = String(today.getMonth() + 1).padStart(2, '0');
  var year = today.getFullYear();
  var formattedDate = day + "-" + month + "-" + year;

  var subject = "Your Attendance QR Code";

// Fetch QR as image blob
  var qrBlob = UrlFetchApp.fetch(qrUrl)
    .getBlob()
    .setName("attendance_qr.png");

  var htmlBody =
    "<p>Hi " + name + ",</p>" +
    "<p>This is your personal QR code for attendance.</p>" +
    "<p><b>Date Sent:</b> " + formattedDate + "</p>" +
    "<p>Please scan this QR during the event:</p>" +
    '<img src="cid:qrImage" width="220" height="220">' +
    "<p>Regards,<br>Event Team</p>";

  MailApp.sendEmail({
    to: email,
    subject: subject,
    htmlBody: htmlBody,
    inlineImages: {
      qrImage: qrBlob
    }
  });


  // Mark mail as sent
  sheet.getRange(row, 10).setValue("YES");
  sheet.getRange(row, 11).setValue(formattedDate);

}


