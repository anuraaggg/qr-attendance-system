// ============ CONFIGURATION ============
var CONFIG = {
  WEB_APP_URL: "https://script.google.com/macros/s/AKfycbz3j6SOlu8gc24eYwV5mUR8uMsuMIYdFeyNtPJr85abe9slInmGssV7JLB7rfqsK-IryQ/exec",
  ALLOWED_EMAILS: [
    "anuraagshnkr@gmail.com",
    "jefreyjose4605@gmail.com"
  ]
};

// ============ COLUMN INDICES (1-based) ============
var COLUMNS = {
  ID: 1,
  NAME: 2,
  EMAIL: 3,
  PHONE: 4,
  COLLEGE: 5,
  CITY: 6,
  STATE: 7,
  TEAM_COUNT: 8,
  QR_CODE: 9,
  MAIL_SENT: 10,
  DATE_SENT: 11,
  TIMESTAMP: 12
};


function authorizeOnce() {
  UrlFetchApp.fetch("https://www.google.com");
}

// ============ UTILITY FUNCTIONS ============

/**
 * Validate email format
 */
function isValidEmail(email) {
  var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}

/**
 * Safe trim and lowercase
 */
function normalizeEmail(email) {
  return (email || "").toString().trim().toLowerCase();
}

/**
 * Log errors and important events for debugging
 */
function logEvent(message, data) {
  var timestamp = new Date().toLocaleString();
  console.log("[" + timestamp + "] " + message);
  if (data) {
    console.log(JSON.stringify(data));
  }
}


/**
 * Main attendance marking endpoint
 */
function doGet(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var id = e.parameter.id;

    if (!id) {
      logEvent("Invalid QR Code: No ID provided");
      return bigMessage("Invalid QR Code", "#e74c3c");
    }

    // Logged-in user email
    var userEmail = normalizeEmail(Session.getActiveUser().getEmail());

    if (!userEmail) {
      logEvent("Unauthorized: No user email detected");
      return bigMessage("Unauthorized Access", "#f39c12");
    }

    // Check if user is authorized (case-insensitive)
    var isAuthorized = CONFIG.ALLOWED_EMAILS.some(function(email) {
      return normalizeEmail(email) === userEmail;
    });

    if (!isAuthorized) {
      logEvent("Unauthorized access attempt from: " + userEmail);
      return bigMessage("Unauthorized Access", "#e74c3c");
    }

    var data = sheet.getDataRange().getValues();

    // Find matching ID in sheet
    for (var i = 1; i < data.length; i++) {
      if (data[i][COLUMNS.ID - 1].toString().trim() === id.toString().trim()) {

        // Check if already marked (column I: QR_CODE/MAIL_SENT column)
        var presentStatus = (data[i][COLUMNS.MAIL_SENT - 1] || "").toString().trim().toUpperCase();
        if (presentStatus === "YES") {
          logEvent("Duplicate attendance attempt for ID: " + id);
          return bigMessage("Attendance Already Marked", "#3498db");
        }

        // Mark attendance
        sheet.getRange(i + 1, COLUMNS.MAIL_SENT).setValue("YES");
        var now = new Date();
        var formattedDateTime = Utilities.formatDate(
          now,
          Session.getScriptTimeZone(),
          "dd-MM-yy HH:mm"
        );

        sheet.getRange(i + 1, COLUMNS.TIMESTAMP).setValue(formattedDateTime);
        logEvent("Attendance marked for ID: " + id + " by: " + userEmail);

        return bigMessage("Attendance Marked Successfully", "#2ecc71");
      }
    }

    logEvent("ID not found in sheet: " + id);
    return bigMessage("ID Not Found", "#e74c3c");
  } catch (error) {
    logEvent("Error in doGet: " + error.toString());
    return bigMessage("System Error", "#e74c3c");
  }
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


/**
 * Send QR code email to participant
 */
function sendQrEmail(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var row = e.range.getRow();

    // Skip header row
    if (row === 1) return;

    // Read all required fields
    var id        = sheet.getRange(row, COLUMNS.ID).getValue();
    var name      = sheet.getRange(row, COLUMNS.NAME).getValue();
    var email     = sheet.getRange(row, COLUMNS.EMAIL).getValue();
    var phone     = sheet.getRange(row, COLUMNS.PHONE).getValue();
    var college   = sheet.getRange(row, COLUMNS.COLLEGE).getValue();
    var city      = sheet.getRange(row, COLUMNS.CITY).getValue();
    var state     = sheet.getRange(row, COLUMNS.STATE).getValue();
    var teamCount = sheet.getRange(row, COLUMNS.TEAM_COUNT).getValue();
    var mailSent  = sheet.getRange(row, COLUMNS.MAIL_SENT).getValue();

    // Stop if mail already sent
    if ((mailSent || "").toString().trim().toUpperCase() === "YES") {
      logEvent("Email already sent for row: " + row);
      return;
    }

    // Validate all required fields are present
    if (
      !id || !name || !email || !phone ||
      !college || !city || !state || !teamCount
    ) {
      logEvent("Incomplete data for row: " + row, {
        id: id, name: name, email: email, phone: phone,
        college: college, city: city, state: state, teamCount: teamCount
      });
      return;
    }

    // Validate email format
    email = normalizeEmail(email);
    if (!isValidEmail(email)) {
      logEvent("Invalid email format for row: " + row + " - " + email);
      return;
    }

    // Generate QR URL (encode the web app URL + ID)
    var qrData = CONFIG.WEB_APP_URL + "?id=" + encodeURIComponent(id);
    var qrUrl =
      "https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=" +
      encodeURIComponent(qrData);

    // Put QR image formula in QR_Code column
    var qrFormula =
      '=IMAGE("https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=' +
      encodeURIComponent(qrData) + '")';

    sheet.getRange(row, COLUMNS.QR_CODE).setFormula(qrFormula);

    // Format date as dd-mm-yyyy
    var today = new Date();
    var day = String(today.getDate()).padStart(2, '0');
    var month = String(today.getMonth() + 1).padStart(2, '0');
    var year = today.getFullYear();
    var formattedDate = day + "-" + month + "-" + year;

    var subject = "Your Attendance QR Code";

    // Fetch QR as image blob
    var response = UrlFetchApp.fetch(qrUrl, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) {
      logEvent("Failed to fetch QR code for row: " + row + " - HTTP " + response.getResponseCode());
      return;
    }

    var qrBlob = response.getBlob().setName("attendance_qr.png");

    var htmlBody =
      "<p>Hi " + name + ",</p>" +
      "<p>This is your personal QR code for attendance.</p>" +
      "<p><b>Date Sent:</b> " + formattedDate + "</p>" +
      "<p>Please scan this QR during the event:</p>" +
      '<img src="cid:qrImage" width="220" height="220" style="border: 1px solid #ccc; padding: 10px;">' +
      "<p>Regards,<br>Event Team</p>";

    MailApp.sendEmail({
      to: email,
      subject: subject,
      htmlBody: htmlBody,
      inlineImages: {
        qrImage: qrBlob
      }
    });

    // Mark mail as sent and log date
    sheet.getRange(row, COLUMNS.MAIL_SENT).setValue("YES");
    sheet.getRange(row, COLUMNS.DATE_SENT).setValue(formattedDate);
    logEvent("Email sent to: " + email + " for row: " + row);
  } catch (error) {
    logEvent("Error in sendQrEmail for row: " + e.range.getRow() + " - " + error.toString());
  }
}