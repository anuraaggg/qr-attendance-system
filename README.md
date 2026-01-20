# QR-Based Attendance System (Google Apps Script)

A serverless attendance management system built with Google Apps Script that generates unique QR codes for participants, automatically emails them, and enables secure attendance marking through QR code scanning with administrator authentication.

---

## 📋 Project Description

This project leverages Google Apps Script to create a streamlined attendance tracking system. Participants receive personalized QR codes via email, which can be scanned during events to mark attendance. The system uses Google account authentication to ensure only authorized administrators can mark attendance, while preventing duplicate entries and maintaining accurate timestamps in Google Sheets.

---

## ✨ Key Features

- **🎫 QR Code Generation Per Participant** – Unique QR codes created for each registered participant
- **📧 Automatic Email Delivery** – QR codes sent automatically via Gmail when participant data is entered
- **🔐 Admin-Only Authentication** – Attendance marking restricted to a single authorized Google account
- **🚫 Duplicate Prevention** – System prevents marking attendance multiple times for the same participant
- **⏰ Timestamp Logging** – Records attendance with formatted timestamps in `dd-MM-yy HH:mm` format
- **📊 Google Sheets Backend** – All participant data and attendance records stored in Google Sheets
- **📱 Mobile-Friendly Interface** – Responsive UI for attendance confirmation messages

---

## 🛠️ Tech Stack

- **Google Apps Script** – Server-side scripting and automation
- **Google Sheets** – Database for participant records and attendance logs
- **Gmail Service (MailApp)** – Automated email delivery system
- **QRServer API** – QR code generation service
- **HTML/CSS** – Custom UI for attendance confirmation feedback

---

## 🔄 How It Works

1. **Registration** – Participants are added to a Google Sheet with their details (ID, name, email, phone, college, city, state, team count)

2. **QR Code Generation** – When a new participant row is filled, a trigger automatically:
   - Generates a unique QR code containing the participant ID
   - Embeds the QR code image in the sheet
   - Sends the QR code to the participant's email

3. **Event Day Scanning** – During the event:
   - Admin scans the participant's QR code
   - The QR code opens a web app URL with the participant ID

4. **Authentication & Validation** – The system:
   - Verifies the scanner is logged in with the authorized Google account
   - Checks if attendance has already been marked
   - Validates the participant ID exists in the database

5. **Attendance Recording** – If all checks pass:
   - Marks attendance as "YES" in the sheet
   - Records a formatted timestamp
   - Displays a success confirmation message

---

## 📁 File Structure

```
qr-attendance-system/
│
├── attendance.js       # Main Apps Script file containing all logic
│                       # - doGet(): Web app handler for QR scanning
│                       # - sendQrEmail(): Trigger function for QR generation
│                       # - bigMessage(): UI response helper
│
└── README.md          # Project documentation
```

---

## 🚀 Setup Instructions

### 1. Create Google Sheet
Create a new Google Sheet with the following columns:

| Column | Header | Description |
|--------|--------|-------------|
| A | ID | Unique participant identifier |
| B | Name | Participant name |
| C | Email | Participant email address |
| D | Phone | Contact number |
| E | College | College/Institution name |
| F | City | City name |
| G | State | State name |
| H | Team Count | Number of team members |
| I | QR_Code | Auto-generated QR code image |
| J | Mail_Sent | Email status flag |
| K | Date_Sent | Date when email was sent |
| L | Timestamp | Attendance marking timestamp |

### 2. Open Apps Script Editor
- In your Google Sheet, go to **Extensions → Apps Script**
- Delete any existing code in the editor
- Paste the entire contents of `attendance.js`

### 3. Configure Authorized Email
- In `attendance.js`, locate line 13:
  ```javascript
  var allowedEmail = "anuraagshnkr@gmail.com";
  ```
- Replace with your authorized administrator email address

### 4. Set Up Trigger
- In the Apps Script editor, click the **clock icon** (Triggers) in the left sidebar
- Click **+ Add Trigger**
- Configure:
  - **Function**: `sendQrEmail`
  - **Event source**: From spreadsheet
  - **Event type**: On change
- Click **Save**

### 5. Deploy as Web App
- Click **Deploy → New deployment**
- Click the gear icon ⚙️ and select **Web app**
- Configure:
  - **Description**: QR Attendance Scanner
  - **Execute as**: Me
  - **Who has access**: Anyone
- Click **Deploy**
- Copy the **Web App URL**

### 6. Update QR Code URL
- In `attendance.js`, locate lines 122-124 and 129-131
- Replace the existing deployment URL with your new Web App URL in both locations

### 7. Grant Permissions
- When you first run the script or deploy, Google will request permissions
- Review and grant access to:
  - Google Sheets
  - Gmail (for sending emails)
  - External services (QRServer API)

---

## 🔒 Security Notes

- ✅ **Single Administrator Control** – Attendance marking restricted to one authorized email address
- ✅ **Duplicate Prevention** – QR codes cannot be reused once attendance is marked
- ✅ **Authentication Required** – Scanner must be logged into the authorized Google account
- ✅ **No Direct Data Editing** – Participants cannot modify attendance records via the web app
- ✅ **Session-Based Verification** – Uses Google's built-in session management

---

## 💡 Use Cases

- 🎓 **College Events** – Track student attendance for seminars, orientations, and ceremonies
- 🛠️ **Workshops** – Verify participant attendance for certification
- 💻 **Hackathons** – Manage team check-ins and attendance tracking
- 🎤 **Conferences** – Monitor session attendance and participant engagement
- 📚 **Training Programs** – Record attendance for compliance and reporting

---

## 🔮 Future Improvements

- [ ] **Token-Based One-Time QR Codes** – Generate time-limited QR codes for enhanced security
- [ ] **QR Expiration** – Automatically invalidate QR codes after the event window
- [ ] **Multiple Authorized Scanners** – Support for multiple admin accounts with role management
- [ ] **Scan Analytics Dashboard** – Real-time attendance statistics and visualization
- [ ] **SMS Integration** – Send QR codes via SMS for participants without email
- [ ] **Offline Mode** – Cache QR codes locally for scanning without internet
- [ ] **Export Reports** – Generate PDF/CSV attendance reports

---

## 📧 Contact

For questions or support, please contact the administrator at the email configured in the system.

---

## 📝 License

This project is open-source and available for educational and commercial use.

---

**Built with ❤️ using Google Apps Script**
