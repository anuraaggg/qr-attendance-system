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

## Limitations

- Single authorized scanner (extendable for multiple admins)
- No built-in analytics or reporting interface
- QR codes don't expire (security consideration)
- Requires manual Google account login for attendance marking
- Limited to Google Sheets storage (performance may vary with large datasets)

## Future Improvements

- Token-based one-time QR codes
- QR expiration and time limits
- Multiple authorized scanners
- Analytics dashboard
- Export reports (PDF/CSV)
