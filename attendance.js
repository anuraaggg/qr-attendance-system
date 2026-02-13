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
 * Main endpoint - Routes to different functions based on action parameter
 */
function doGet(e) {
  try {
    var action = e.parameter.action || "scan";

    // Route to appropriate handler
    if (action === "dashboard") {
      return dashboardView();
    } else if (action === "report") {
      return reportView();
    } else if (action === "reset" && e.parameter.id) {
      return resetAttendance(e.parameter.id);
    } else if (action === "export") {
      return exportToCSV();
    } else {
      // Default: QR code scan
      return handleQRScan(e);
    }
  } catch (error) {
    logEvent("Error in doGet: " + error.toString());
    return bigMessage("System Error", "#e74c3c");
  }
}

/**
 * Handle QR code scanning
 */
function handleQRScan(e) {
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
    logEvent("Error in handleQRScan: " + error.toString());
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

// ============ ADMIN FUNCTIONS ============

/**
 * Dashboard view - Shows attendance statistics
 */
function dashboardView() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
  var total = data.length - 1; // Exclude header
  var attended = 0;
  var pending = 0;
  
  for (var i = 1; i < data.length; i++) {
    var status = (data[i][COLUMNS.MAIL_SENT - 1] || "").toString().trim().toUpperCase();
    if (status === "YES" && (data[i][COLUMNS.TIMESTAMP - 1] || "").toString().trim() !== "") {
      attended++;
    } else {
      pending++;
    }
  }
  
  var attendanceRate = total > 0 ? Math.round((attended / total) * 100) : 0;
  
  var html = `
    <html>
      <head>
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <style>
          * { margin: 0; padding: 0; box-sizing: border-box; }
          body {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            min-height: 100vh;
            display: flex;
            align-items: center;
            justify-content: center;
            padding: 20px;
          }
          .container {
            background: white;
            border-radius: 12px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            padding: 40px;
            max-width: 600px;
            width: 100%;
          }
          h1 {
            color: #333;
            margin-bottom: 30px;
            text-align: center;
            font-size: 28px;
          }
          .stats {
            display: grid;
            grid-template-columns: 1fr 1fr 1fr;
            gap: 20px;
            margin-bottom: 30px;
          }
          .stat-card {
            background: #f8f9fa;
            padding: 20px;
            border-radius: 8px;
            text-align: center;
            border-left: 4px solid #667eea;
          }
          .stat-card.attended {
            border-left-color: #2ecc71;
          }
          .stat-card.pending {
            border-left-color: #f39c12;
          }
          .stat-card.rate {
            border-left-color: #3498db;
          }
          .stat-number {
            font-size: 32px;
            font-weight: bold;
            color: #333;
            margin: 10px 0;
          }
          .stat-label {
            color: #666;
            font-size: 14px;
            text-transform: uppercase;
            letter-spacing: 1px;
          }
          .button-group {
            display: flex;
            gap: 10px;
            flex-direction: column;
          }
          button {
            padding: 12px 20px;
            border: none;
            border-radius: 6px;
            font-size: 16px;
            cursor: pointer;
            transition: all 0.3s;
            font-weight: 600;
          }
          .btn-primary {
            background: #667eea;
            color: white;
          }
          .btn-primary:hover {
            background: #5568d3;
            transform: translateY(-2px);
            box-shadow: 0 5px 15px rgba(102, 126, 234, 0.4);
          }
          .btn-secondary {
            background: #e74c3c;
            color: white;
          }
          .btn-secondary:hover {
            background: #c0392b;
          }
          .footer {
            text-align: center;
            color: #999;
            font-size: 12px;
            margin-top: 20px;
          }
        </style>
      </head>
      <body>
        <div class="container">
          <h1>📊 Attendance Dashboard</h1>
          
          <div class="stats">
            <div class="stat-card">
              <div class="stat-label">Total Participants</div>
              <div class="stat-number">${total}</div>
            </div>
            <div class="stat-card attended">
              <div class="stat-label">Attended</div>
              <div class="stat-number">${attended}</div>
            </div>
            <div class="stat-card pending">
              <div class="stat-label">Pending</div>
              <div class="stat-number">${pending}</div>
            </div>
          </div>
          
          <div class="stat-card rate" style="text-align: center; margin-bottom: 30px;">
            <div class="stat-label">Attendance Rate</div>
            <div class="stat-number" style="color: #3498db;">${attendanceRate}%</div>
          </div>
          
          <div class="button-group">
            <button class="btn-primary" onclick="window.location.href=window.location.href.split('?')[0] + '?action=report';">
              View Detailed Report
            </button>
            <button class="btn-primary" onclick="window.location.href=window.location.href.split('?')[0] + '?action=export';">
              ⬇ Download CSV
            </button>
            <button class="btn-secondary" onclick="alert('Please refresh from Google Sheets');">
              Back to Apps Script
            </button>
          </div>
          
          <div class="footer">Last updated: ${new Date().toLocaleString()}</div>
        </div>
      </body>
    </html>
  `;
  
  return HtmlService.createHtmlOutput(html);
}

/**
 * Report view - Shows detailed attendance list
 */
function reportView() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
  var rows = [];
  var attended = 0;
  
  for (var i = 1; i < data.length; i++) {
    var id = data[i][COLUMNS.ID - 1] || "";
    var name = data[i][COLUMNS.NAME - 1] || "";
    var email = data[i][COLUMNS.EMAIL - 1] || "";
    var status = (data[i][COLUMNS.MAIL_SENT - 1] || "").toString().trim().toUpperCase();
    var timestamp = data[i][COLUMNS.TIMESTAMP - 1] || "Not marked";
    var isPresent = status === "YES" && timestamp !== "Not marked";
    var statusBadge = isPresent ? "✓ Present" : "○ Absent";
    var statusColor = isPresent ? "#2ecc71" : "#e74c3c";
    
    if (isPresent) {
      attended++;
    }
    
    rows.push({
      id: id,
      name: name,
      email: email,
      status: statusBadge,
      statusColor: statusColor,
      timestamp: timestamp,
      isPresent: isPresent
    });
  }
  
  var total = data.length - 1;
  var attendanceRate = total > 0 ? Math.round((attended / total) * 100) : 0;
  
  var rowsJSON = JSON.stringify(rows);
  
  var tableRows = rows.map(function(row, index) {
    return `
      <tr data-id="${row.id}" data-name="${row.name}" data-email="${row.email}" data-present="${row.isPresent}" data-index="${index}">
        <td>${row.id}</td>
        <td>${row.name}</td>
        <td><span style="color: ${row.statusColor}; font-weight: bold;">${row.status}</span></td>
        <td>${row.timestamp}</td>
      </tr>
    `;
  }).join("");
  
  var html = `
    <html>
      <head>
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <style>
          * { margin: 0; padding: 0; box-sizing: border-box; }
          body {
            background: #f5f7fa;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            padding: 20px;
          }
          .container {
            max-width: 1000px;
            margin: 0 auto;
            background: white;
            border-radius: 12px;
            box-shadow: 0 10px 30px rgba(0,0,0,0.1);
            overflow: hidden;
          }
          .header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 30px;
          }
          .header h1 {
            font-size: 28px;
            margin-bottom: 10px;
          }
          .header p {
            opacity: 0.9;
            font-size: 16px;
          }
          .filter-section {
            background: #f8f9fa;
            padding: 20px 30px;
            border-bottom: 1px solid #e0e0e0;
            display: flex;
            gap: 15px;
            flex-wrap: wrap;
            align-items: center;
          }
          .filter-group {
            flex: 1;
            min-width: 200px;
          }
          .filter-group label {
            display: block;
            font-size: 12px;
            font-weight: 600;
            color: #666;
            margin-bottom: 5px;
            text-transform: uppercase;
          }
          .filter-group input,
          .filter-group select {
            width: 100%;
            padding: 10px 12px;
            border: 1px solid #ddd;
            border-radius: 6px;
            font-size: 14px;
            font-family: inherit;
          }
          .filter-group input:focus,
          .filter-group select:focus {
            outline: none;
            border-color: #667eea;
            box-shadow: 0 0 0 3px rgba(102, 126, 234, 0.1);
          }
          .clear-btn {
            padding: 10px 15px;
            background: #e74c3c;
            color: white;
            border: none;
            border-radius: 6px;
            cursor: pointer;
            font-weight: 600;
            transition: all 0.3s;
          }
          .clear-btn:hover {
            background: #c0392b;
          }
          .result-count {
            padding: 10px 30px;
            background: #f8f9fa;
            font-size: 14px;
            color: #666;
            border-bottom: 1px solid #e0e0e0;
          }
          table {
            width: 100%;
            border-collapse: collapse;
          }
          th {
            background: #f8f9fa;
            padding: 15px;
            text-align: left;
            font-weight: 600;
            color: #333;
            border-bottom: 2px solid #e0e0e0;
          }
          td {
            padding: 12px 15px;
            border-bottom: 1px solid #e0e0e0;
          }
          tr:hover {
            background: #f8f9fa;
          }
          tr.hidden {
            display: none;
          }
          .no-results {
            text-align: center;
            padding: 40px 20px;
            color: #999;
            font-size: 16px;
          }
          .footer-stats {
            background: #f8f9fa;
            padding: 20px 30px;
            border-top: 1px solid #e0e0e0;
            display: flex;
            justify-content: space-around;
            flex-wrap: wrap;
          }
          .footer-stat {
            text-align: center;
          }
          .footer-stat-label {
            color: #666;
            font-size: 12px;
            text-transform: uppercase;
            margin-bottom: 5px;
          }
          .footer-stat-value {
            font-size: 24px;
            font-weight: bold;
            color: #333;
          }
        </style>
        <script>
          var allRows = ${rowsJSON};
          
          function filterTable() {
            var searchText = document.getElementById('searchInput').value.toLowerCase();
            var statusFilter = document.getElementById('statusFilter').value;
            var tableRows = document.querySelectorAll('tbody tr');
            var visibleCount = 0;
            
            tableRows.forEach(function(row) {
              var id = row.getAttribute('data-id').toLowerCase();
              var name = row.getAttribute('data-name').toLowerCase();
              var email = row.getAttribute('data-email').toLowerCase();
              var isPresent = row.getAttribute('data-present') === 'true';
              
              var matchesSearch = id.includes(searchText) || name.includes(searchText) || email.includes(searchText);
              var matchesStatus = statusFilter === 'all' || (statusFilter === 'present' && isPresent) || (statusFilter === 'absent' && !isPresent);
              
              if (matchesSearch && matchesStatus) {
                row.classList.remove('hidden');
                visibleCount++;
              } else {
                row.classList.add('hidden');
              }
            });
            
            var noResults = document.getElementById('noResults');
            if (visibleCount === 0) {
              noResults.style.display = 'block';
            } else {
              noResults.style.display = 'none';
            }
            
            document.getElementById('resultCount').textContent = 'Showing ' + visibleCount + ' of ' + tableRows.length + ' records';
          }
          
          function clearFilters() {
            document.getElementById('searchInput').value = '';
            document.getElementById('statusFilter').value = 'all';
            filterTable();
          }
        </script>
      </head>
      <body>
        <div class="container">
          <div class="header">
            <h1>📋 Attendance Report</h1>
            <p>Generated on ${new Date().toLocaleString()}</p>
          </div>
          
          <div class="filter-section">
            <div class="filter-group" style="flex: 2;">
              <label for="searchInput">🔍 Search (ID, Name, or Email)</label>
              <input type="text" id="searchInput" placeholder="Type to search..." onkeyup="filterTable()">
            </div>
            <div class="filter-group">
              <label for="statusFilter">Status</label>
              <select id="statusFilter" onchange="filterTable()">
                <option value="all">All</option>
                <option value="present">✓ Present</option>
                <option value="absent">○ Absent</option>
              </select>
            </div>
            <button class="clear-btn" onclick="clearFilters()">Clear Filters</button>
          </div>
          
          <div class="result-count" id="resultCount">Showing ${rows.length} of ${rows.length} records</div>
          
          <table>
            <thead>
              <tr>
                <th>ID</th>
                <th>Name</th>
                <th>Status</th>
                <th>Timestamp</th>
              </tr>
            </thead>
            <tbody>
              ${tableRows}
            </tbody>
          </table>
          
          <div class="no-results" id="noResults" style="display: none;">
            No results found. Try adjusting your search or filters.
          </div>
          
          <div class="footer-stats">
            <div class="footer-stat">
              <div class="footer-stat-label">Total</div>
              <div class="footer-stat-value">${total}</div>
            </div>
            <div class="footer-stat">
              <div class="footer-stat-label">Attended</div>
              <div class="footer-stat-value" style="color: #2ecc71;">${attended}</div>
            </div>
            <div class="footer-stat">
              <div class="footer-stat-label">Absent</div>
              <div class="footer-stat-value" style="color: #e74c3c;">${total - attended}</div>
            </div>
            <div class="footer-stat">
              <div class="footer-stat-label">Rate</div>
              <div class="footer-stat-value" style="color: #3498db;">${attendanceRate}%</div>
            </div>
          </div>
          
          <div style="display: flex; gap: 10px; margin-top: 30px; padding: 0 30px; margin-bottom: 30px;">
            <button style="flex: 1; padding: 12px 20px; background: #667eea; color: white; border: none; border-radius: 6px; font-size: 16px; cursor: pointer; font-weight: 600; transition: all 0.3s;" onmouseover="this.style.background='#5568d3'" onmouseout="this.style.background='#667eea'" onclick="window.location.href=window.location.href.split('?')[0] + '?action=export';">
              ⬇ Download as CSV
            </button>
            <button style="flex: 1; padding: 12px 20px; background: #3498db; color: white; border: none; border-radius: 6px; font-size: 16px; cursor: pointer; font-weight: 600; transition: all 0.3s;" onmouseover="this.style.background='#2980b9'" onmouseout="this.style.background='#3498db'" onclick="window.location.href=window.location.href.split('?')[0] + '?action=dashboard';">
              ← Back to Dashboard
            </button>
          </div>
        </div>
      </body>
    </html>
  `;
  
  return HtmlService.createHtmlOutput(html);
}

/**
 * Export attendance to CSV format
 */
function exportToCSV() {
  try {
    var userEmail = normalizeEmail(Session.getActiveUser().getEmail());
    
    // Check if user is authorized
    var isAuthorized = CONFIG.ALLOWED_EMAILS.some(function(email) {
      return normalizeEmail(email) === userEmail;
    });
    
    if (!isAuthorized) {
      logEvent("Unauthorized export attempt from: " + userEmail);
      return bigMessage("Unauthorized - Admin access required", "#e74c3c");
    }
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = sheet.getDataRange().getValues();
    
    // CSV Headers
    var headers = ["ID", "Name", "Email", "Phone", "College", "City", "State", "Team Count", "Mail Sent", "Date Sent", "Attendance Status", "Timestamp"];
    var csvContent = [headers.map(function(h) { return '"' + h + '"'; }).join(",")];
    
    // Add data rows
    for (var i = 1; i < data.length; i++) {
      var rowData = [
        data[i][COLUMNS.ID - 1] || "",
        data[i][COLUMNS.NAME - 1] || "",
        data[i][COLUMNS.EMAIL - 1] || "",
        data[i][COLUMNS.PHONE - 1] || "",
        data[i][COLUMNS.COLLEGE - 1] || "",
        data[i][COLUMNS.CITY - 1] || "",
        data[i][COLUMNS.STATE - 1] || "",
        data[i][COLUMNS.TEAM_COUNT - 1] || "",
        data[i][COLUMNS.MAIL_SENT - 1] || "",
        data[i][COLUMNS.DATE_SENT - 1] || "",
        (data[i][COLUMNS.MAIL_SENT - 1] || "").toString().trim().toUpperCase() === "YES" && (data[i][COLUMNS.TIMESTAMP - 1] || "").toString().trim() !== "" ? "Present" : "Absent",
        data[i][COLUMNS.TIMESTAMP - 1] || ""
      ];
      
      // Escape quotes and wrap in quotes
      var escapedRow = rowData.map(function(cell) {
        var str = cell.toString();
        str = str.replace(/"/g, '""'); // Escape quotes by doubling them
        return '"' + str + '"';
      });
      
      csvContent.push(escapedRow.join(","));
    }
    
    var csv = csvContent.join("\n");
    
    // Generate filename with timestamp
    var now = new Date();
    var timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd_HH-mm-ss");
    var filename = "attendance_report_" + timestamp + ".csv";
    
    logEvent("CSV export generated: " + filename);
    
    // Return as downloadable file
    return ContentService
      .createTextOutput(csv)
      .setMimeType(ContentService.MimeType.CSV)
      .downloadAsFile(filename);
      
  } catch (error) {
    logEvent("Error in exportToCSV: " + error.toString());
    return bigMessage("Export Error", "#e74c3c");
  }
}

/**
 * Reset attendance for a specific ID (admin only)
 */
function resetAttendance(id) {
  try {
    var userEmail = normalizeEmail(Session.getActiveUser().getEmail());
    
    // Check if user is authorized
    var isAuthorized = CONFIG.ALLOWED_EMAILS.some(function(email) {
      return normalizeEmail(email) === userEmail;
    });
    
    if (!isAuthorized) {
      logEvent("Unauthorized reset attempt from: " + userEmail);
      return bigMessage("Unauthorized - Admin access required", "#e74c3c");
    }
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][COLUMNS.ID - 1].toString().trim() === id.toString().trim()) {
        sheet.getRange(i + 1, COLUMNS.MAIL_SENT).setValue("");
        sheet.getRange(i + 1, COLUMNS.TIMESTAMP).setValue("");
        logEvent("Attendance reset for ID: " + id + " by: " + userEmail);
        return bigMessage("Attendance Reset for ID: " + id, "#3498db");
      }
    }
    
    logEvent("ID not found for reset: " + id);
    return bigMessage("ID Not Found", "#e74c3c");
  } catch (error) {
    logEvent("Error in resetAttendance: " + error.toString());
    return bigMessage("System Error", "#e74c3c");
  }
}

/**
 * Resend QR codes to participants (run from Apps Script editor)
 * Useful if some emails didn't send
 */
function resendQRToMissing() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var resendCount = 0;
  
  for (var i = 1; i < data.length; i++) {
    var mailSent = (data[i][COLUMNS.MAIL_SENT - 1] || "").toString().trim().toUpperCase();
    var email = data[i][COLUMNS.EMAIL - 1] || "";
    
    // Resend if mail not marked as sent and email exists
    if (mailSent !== "YES" && email) {
      try {
        sheet.getRange(i + 1, COLUMNS.MAIL_SENT).setValue("");
        sendQrEmailManual(i + 1);
        resendCount++;
      } catch (error) {
        logEvent("Error resending QR for row: " + (i + 1) + " - " + error.toString());
      }
    }
  }
  
  logEvent("Resent QR codes to " + resendCount + " participants");
  return resendCount;
}

/**
 * Manual QR email send (for bulk operations)
 */
function sendQrEmailManual(rowNumber) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    var id        = sheet.getRange(rowNumber, COLUMNS.ID).getValue();
    var name      = sheet.getRange(rowNumber, COLUMNS.NAME).getValue();
    var email     = sheet.getRange(rowNumber, COLUMNS.EMAIL).getValue();
    var phone     = sheet.getRange(rowNumber, COLUMNS.PHONE).getValue();
    var college   = sheet.getRange(rowNumber, COLUMNS.COLLEGE).getValue();
    var city      = sheet.getRange(rowNumber, COLUMNS.CITY).getValue();
    var state     = sheet.getRange(rowNumber, COLUMNS.STATE).getValue();
    var teamCount = sheet.getRange(rowNumber, COLUMNS.TEAM_COUNT).getValue();

    if (!id || !name || !email || !phone || !college || !city || !state || !teamCount) {
      logEvent("Incomplete data for row: " + rowNumber);
      return;
    }

    email = normalizeEmail(email);
    if (!isValidEmail(email)) {
      logEvent("Invalid email format for row: " + rowNumber + " - " + email);
      return;
    }

    var qrData = CONFIG.WEB_APP_URL + "?id=" + encodeURIComponent(id);
    var qrUrl = "https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=" + encodeURIComponent(qrData);

    var qrFormula = '=IMAGE("https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=' + encodeURIComponent(qrData) + '")';
    sheet.getRange(rowNumber, COLUMNS.QR_CODE).setFormula(qrFormula);

    var today = new Date();
    var day = String(today.getDate()).padStart(2, '0');
    var month = String(today.getMonth() + 1).padStart(2, '0');
    var year = today.getFullYear();
    var formattedDate = day + "-" + month + "-" + year;

    var response = UrlFetchApp.fetch(qrUrl, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) {
      logEvent("Failed to fetch QR code for row: " + rowNumber);
      return;
    }

    var qrBlob = response.getBlob().setName("attendance_qr.png");

    var htmlBody = "<p>Hi " + name + ",</p>" +
      "<p>This is your personal QR code for attendance.</p>" +
      "<p><b>Date Sent:</b> " + formattedDate + "</p>" +
      "<p>Please scan this QR during the event:</p>" +
      '<img src="cid:qrImage" width="220" height="220" style="border: 1px solid #ccc; padding: 10px;">' +
      "<p>Regards,<br>Event Team</p>";

    MailApp.sendEmail({
      to: email,
      subject: "Your Attendance QR Code",
      htmlBody: htmlBody,
      inlineImages: {
        qrImage: qrBlob
      }
    });

    sheet.getRange(rowNumber, COLUMNS.MAIL_SENT).setValue("YES");
    sheet.getRange(rowNumber, COLUMNS.DATE_SENT).setValue(formattedDate);
    logEvent("Email resent to: " + email + " for row: " + rowNumber);
  } catch (error) {
    logEvent("Error in sendQrEmailManual for row: " + rowNumber + " - " + error.toString());
  }
}