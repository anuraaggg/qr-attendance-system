# QR-Based Attendance System

Google Apps Script solution for automated QR-based attendance tracking with email delivery and admin authentication.

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

## How It Works

1. Participants added to Google Sheet with their details
2. QR code auto-generated and emailed to each participant
3. Admin scans QR code during event (requires authorized Google login)
4. System validates authentication and marks attendance with timestamp

## Setup

1. **Create Google Sheet** with columns: ID, Name, Email, Phone, College, City, State, Team Count, QR_Code, Mail_Sent, Date_Sent, Timestamp

2. **Open Apps Script** (Extensions → Apps Script) and paste `attendance.js`

3. **Configure Admin Email** – Update `allowedEmail` variable in the script

4. **Set Up Trigger** – Add trigger: Function `sendQrEmail`, Event source: From spreadsheet, Event type: On change

5. **Deploy as Web App** – Deploy → New deployment → Web app
   - **Execute as:** Your user account (the account running the script)
   - **Who has access:** Anyone

6. **Update QR Generation URLs** – The QR codes need to point to your deployed web app. Replace the deployment URL in **2 locations**:
   - Line 122-124: In the `qrUrl` variable
   - Line 129-131: In the `qrFormula` variable
   
   Replace `https://script.google.com/macros/s/AKfycbyWFwCaBau1rcom9VXpQxkIp5ayxkMhOYEmpPa2j8YGOohT-zfwZf8rWkj58k09HwlOtQ/exec` with your actual deployed Web App URL

7. **Grant Permissions** – Authorize access to Google Sheets, Gmail, and external services

## Security

- Attendance marking restricted to single authorized email
- Duplicate prevention built-in
- Google account authentication required
- No direct participant data editing via web app

## Use Cases

College events, workshops, hackathons, conferences, training programs

## Future Improvements

- Token-based one-time QR codes
- QR expiration and time limits
- Multiple authorized scanners
- Analytics dashboard
- Export reports (PDF/CSV)
