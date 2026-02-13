# QR-Based Attendance System

Google Apps Script solution for automated QR-based attendance tracking with email delivery and admin authentication.

## Prerequisites

- A Google Account (free)
- Access to Google Sheets and Gmail
- A Google Drive folder for your scripts
- Basic familiarity with Google Apps Script interface

## Features

- Automatic QR code generation and email delivery per participant
- Secure attendance marking (admin-only Google account authentication)
- Duplicate attendance prevention
- Timestamp logging in `dd-MM-yy HH:mm` format
- Google Sheets as backend database
- **NEW:** Attendance dashboard with live statistics
- **NEW:** Detailed attendance report with search & filter
- **NEW:** Reset attendance functionality for corrections
- **NEW:** Bulk resend QR codes to participants
- **NEW:** Export attendance to CSV for analysis

## Tech Stack

- Google Apps Script
- Google Sheets
- Gmail Service (MailApp)
- QRServer API

## Project Structure

```
qr-attendance-system/
├── attendance.js          # Main Google Apps Script
└── README.md              # Documentation
```

**To use:** Copy contents of `attendance.js` into Google Apps Script editor (Extensions → Apps Script in Google Sheets).

## How It Works

1. Participants added to Google Sheet with their details
2. QR code auto-generated and emailed to each participant
3. Admin scans QR code during event (requires authorized Google login)
4. System validates authentication and marks attendance with timestamp

## Admin Features

### Dashboard
View real-time attendance statistics:
- Total participants
- Number attended
- Pending attendance count
- Overall attendance rate percentage

**Access:** In your web app URL, add `?action=dashboard`
```
https://your-web-app-url/exec?action=dashboard
```

### Detailed Report
View complete attendance list with participant details:
- ID, Name, Status (Present/Absent)
- Exact timestamp when attendance was marked
- Summary statistics (attended, absent, rate)

**Search & Filter:**
- Search by ID, Name, or Email (real-time filtering)
- Filter by Attendance Status:
  - All (default)
  - ✓ Present
  - ○ Absent
- "Clear Filters" button to reset and view all records

**Access:** In your web app URL, add `?action=report`
```
https://your-web-app-url/exec?action=report
```

### Reset Attendance
Manually reset attendance for a participant if marked incorrectly.

**From Apps Script Editor:** Run function `resetAttendance` with participant ID

**Via URL Parameter:**
```
https://your-web-app-url/exec?action=reset&id=PARTICIPANT_ID
```

### Resend QR Codes
Resend QR codes to participants who didn't receive them.

**To resend to all who didn't receive:**
1. Go to Apps Script Editor (Extensions → Apps Script)
2. Select function `resendQRToMissing` from dropdown
3. Click "Run" (▶ icon)
4. Authorize when prompted
5. Check logs for how many were resent

**Note:** Resets the MAIL_SENT status and sends fresh QR codes to participants without successful delivery

### Export to CSV
Download all attendance data as a CSV file for further analysis in Excel, Google Sheets, or other tools.

**Features:**
- Exports all participant information (ID, Name, Email, Phone, College, City, State, Team Count)
- Includes email delivery status and timestamps
- Automatically calculates attendance status (Present/Absent)
- Filename includes timestamp for easy organization

**Access:** 
- Click "Download CSV" button in Dashboard or Report view, OR
- Direct URL: `https://your-web-app-url/exec?action=export`

**CSV Columns:**
- ID, Name, Email, Phone, College, City, State, Team Count
- Mail Sent (YES/blank)
- Date Sent
- Attendance Status (Present/Absent)
- Timestamp (exact time marked)

## Setup

1. **Create Google Sheet** with columns: ID, Name, Email, Phone, College, City, State, Team Count, QR_Code, Mail_Sent, Date_Sent, Timestamp

2. **Open Apps Script** (Extensions → Apps Script) and paste `attendance.js`

3. **Configure Admin Emails** – Update `CONFIG.ALLOWED_EMAILS` in `attendance.js`

4. **Set Up Trigger** – Add trigger: Function `sendQrEmail`, Event source: From spreadsheet, Event type: On edit

5. **Deploy as Web App** – Deploy → New deployment → Web app
   - **Execute as:** Your user account (the account running the script)
   - **Who has access:** Anyone with a Google account
   
   ⚠️ **Important:** Must use "Anyone with a Google account" (not "Anyone") so that `Session.getActiveUser().getEmail()` can identify and authorize the scanner

6. **Update Web App URL** – Set your deployed web app URL in `CONFIG.WEB_APP_URL` inside `attendance.js` (single location)

7. **Grant Permissions** – Authorize access to Google Sheets, Gmail, and external services

## Security

- Attendance marking restricted to single authorized email
- Duplicate prevention built-in
- Google account authentication required
- No direct participant data editing via web app

## Quick Reference - URLs

Save these URLs for quick access:

| Feature | URL |
|---------|-----|
| Scan QR | `{WEB_APP_URL}?id=PARTICIPANT_ID` |
| Dashboard | `{WEB_APP_URL}?action=dashboard` |
| Report | `{WEB_APP_URL}?action=report` |
| Reset ID | `{WEB_APP_URL}?action=reset&id=PARTICIPANT_ID` |

*Replace `{WEB_APP_URL}` with your deployed web app URL*

## Use Cases

College events, workshops, hackathons, conferences, training programs

## Troubleshooting

### QR codes not sending
- Check trigger is set up correctly (Extensions → Triggers)
- Verify Gmail API permissions are granted
- Ensure participant email addresses are valid
- Manually run `sendQrEmail()` from Apps Script editor to test

### Attendance marking fails with authentication error
- Verify web app is deployed as "Anyone with a Google account" (not "Anyone")
- Ensure authorized admin email matches `CONFIG.ALLOWED_EMAILS`
- Confirm you're logged into Google with an authorized account

### Duplicate attendance entries
- Check if QR code was scanned multiple times
- Clear attendance records and rescan if needed

### "Need to authorize" message
- Go to Apps Script editor, select `sendQrEmail` function, and run it
- Google will prompt for authorization—grant all permissions

## Troubleshooting New Features

### Dashboard shows "0" for all statistics
- Ensure at least one participant row is in the sheet (after header)
- Dashboard should refresh automatically with correct data

### Attendance Report is empty
- Confirm participant has been added to the sheet
- Check that participant ID column is not empty

### Reset Attendance doesn't work
- Verify you're logged in with an authorized admin email
- Ensure participant ID matches exactly (including case sensitivity)
- Try resetting via Apps Script Editor instead of URL

### QR Codes not resending
- Run `resendQRToMissing()` from Apps Script Editor and check logs
- Verify email addresses are valid in participant sheet
- Ensure trigger is still set up (Extensions → Triggers)

## Limitations

- Single authorized scanner (extendable for multiple admins by adding emails to `CONFIG.ALLOWED_EMAILS`)
- Manual reset required for incorrectly marked attendance (now available via UI or URL)
- QR codes don't expire (security consideration)
- Dashboard and reports are read-only (for admin viewing only)
- Requires manual Google account login for attendance marking
- Limited to Google Sheets storage (performance may vary with large datasets)

## Future Improvements

- Token-based one-time QR codes
- QR expiration and time limits
- Multiple authorized scanners
- Analytics dashboard
- Export reports (PDF/CSV)
