# ระบบลงเวลาออนไลน์ อบต.หัวนา

ระบบจัดการลงเวลาเข้า-ออกงานสำหรับองค์การบริหารส่วนตำบลหัวนา พร้อมแดशบอร์ดแอดมินและระบบแจ้งเตือน Telegram

## ✨ Features

- 📱 **Progressive Web App (PWA)** - รองรับการติดตั้งบนมือถือ
- ⏰ **ลงเวลาเข้า-ออกงาน** ผ่าน LINE LIFF
- 📍 **ระบุตำแหน่ง GPS** พร้อมแสดงชื่อสถานที่
- 👥 **แดชบอร์ดแอดมิน** - จัดการข้อมูลและดูสถิติ
- 📊 **ส่งออก Excel** - รายงานรายวัน/รายเดือน/ช่วงวันที่
- 🔔 **ระบบแจ้งเตือน Telegram** - แบ่งประเภทการแจ้งเตือน
- 🤖 **ลงเวลาออกอัตโนมัติ** - สำหรับผู้ที่ลืมลงเวลาออก
- 🛡️ **ระบบยกเว้น** - สำหรับยามกลางคืน
- 💾 **Google Sheets Integration** - เก็บข้อมูลใน Google Sheets

## 🏗️ Tech Stack

- **Backend:** Node.js, Express.js
- **Frontend:** HTML5, CSS3, JavaScript (Vanilla)
- **Database:** Google Sheets API
- **Authentication:** JWT
- **Notifications:** Telegram Bot API
- **Maps:** OpenStreetMap Nominatim API
- **Export:** ExcelJS
- **Deployment:** Render.com

## 📋 Prerequisites

1. **Google Cloud Project** พร้อม Service Account
2. **Google Sheets** สำหรับเก็บข้อมูล
3. **Telegram Bot** (optional - สำหรับการแจ้งเตือน)
4. **LINE Developers Account** สำหรับ LIFF
5. **Render.com Account** สำหรับ deployment

## 🚀 Quick Start

### 1. Clone Repository

```bash
git clone <repository-url>
cd newtime_tracker
```

### 2. Install Dependencies

```bash
npm install
```

### 3. Environment Setup

```bash
cp .env.example .env
# แก้ไขค่าใน .env ตามความเหมาะสม
```

### 4. Google Sheets Setup

```bash
node setup-telegram-sheet.js
```

### 5. Run Development Server

```bash
npm start
# หรือ
node server.js
```

## 📊 Google Sheets Structure

ระบบใช้ Google Sheets 4 แผ่น:

1. **MAIN** - บันทึกการลงเวลาทั้งหมด
2. **EMPLOYEES** - รายชื่อพนักงาน
3. **ON_WORK** - พนักงานที่กำลังทำงาน
4. **TELEGRAM_SETTINGS** - การตั้งค่าการแจ้งเตือน

## 🔔 Telegram Settings

ระบบรองรับการแจ้งเตือน 3 ประเภท:

- **clock_in_out** - แจ้งเตือนการลงเวลาเข้า-ออก
- **missed_checkout** - แจ้งเตือนลืมลงเวลาออก
- **system_alerts** - แจ้งเตือนระบบ

## 🛠️ Deployment

### Render.com Deployment

1. **Fork repository** ไปยัง GitHub account ของคุณ

2. **สร้าง Web Service ใน Render.com**
   - Connect GitHub repository
   - Build Command: `npm install`
   - Start Command: `npm start`

3. **ตั้งค่า Environment Variables**
   ```
   GOOGLE_SPREADSHEET_ID=your_spreadsheet_id
   GOOGLE_PRIVATE_KEY=your_private_key
   GOOGLE_CLIENT_EMAIL=your_service_account_email
   LIFF_ID=your_liff_id
   JWT_SECRET=your_jwt_secret
   RENDER_SERVICE_URL=https://your-app.onrender.com
   KEEP_ALIVE_ENABLED=true
   NODE_ENV=production
   ```

### Other Deployment Options

- **Heroku**
- **Railway**
- **Vercel**
- **Self-hosted VPS**

## 📱 Admin Panel

เข้าใช้งานแดชบอร์ดแอดมินที่: `https://your-domain.com/admin`

**Default Admin Account:**
- Username: `admin`
- Password: `khayai042315962`

### Admin Features

- 📊 **Dashboard** - สถิติการลงเวลาแบบเรียลไทม์
- 👥 **จัดการพนักงาน** - ดูรายชื่อและสถานะ
- 📈 **รายงานและส่งออก** - Excel reports
- ⚙️ **ตั้งค่าระบบ** - Emergency mode, cache management

## 🔧 Configuration

### 📱 Telegram Notifications (Fast Mode)

ระบบใช้การตั้งค่า Telegram แบบ `.env` เพื่อความเร็วสูงสุด:

```bash
# การแจ้งเตือนลงเวลาเข้า-ออก
TELEGRAM_CLOCK_IN_OUT_ENABLED=true
TELEGRAM_CLOCK_IN_OUT_BOT_TOKEN=your_bot_token
TELEGRAM_CLOCK_IN_OUT_CHAT_ID=your_chat_id

# การแจ้งเตือนระบบ (ลืมลงเวลาออก)
TELEGRAM_SYSTEM_ALERT_ENABLED=true
TELEGRAM_SYSTEM_ALERT_BOT_TOKEN=your_bot_token
TELEGRAM_SYSTEM_ALERT_CHAT_ID=your_chat_id
```

**Notification Types:**
- `clock_in_out` - แจ้งเตือนเมื่อลงเวลาเข้า-ออก
- `system_alert` - แจ้งเตือนระบบ (ลืมลงเวลาออก)
- `forgot_clock` - แจ้งเตือนเฉพาะลืมลงเวลา

### Auto Checkout Settings

```javascript
AUTO_CHECKOUT: {
  EXEMPT_EMPLOYEES: [
    '1017-เปรมชัย ทองสงคราม' // ยามกลางคืน
  ],
  CUTOFF_HOUR: 23,
  CUTOFF_MINUTE: 59
}
```

### Cache Settings

- **EMPLOYEES**: 10 minutes
- **ON_WORK**: 1 minute
- **MAIN**: 30 seconds
- **STATS**: 2 minutes
- **TELEGRAM_SETTINGS**: 5 minutes

## 🔒 Security

- JWT Authentication สำหรับ Admin Panel
- Rate limiting สำหรับ API calls
- Input validation และ sanitization
- CORS configuration
- Environment variables สำหรับ sensitive data

## 📝 API Endpoints

### Public API
- `POST /clock-in` - ลงเวลาเข้างาน
- `POST /clock-out` - ลงเวลาออกงาน
- `GET /employee-status/:name` - ตรวจสอบสถานะ
- `GET /employees` - รายชื่อพนักงาน

### Admin API (Authentication Required)
- `GET /api/admin/stats` - สถิติระบบ
- `GET /api/admin/export/:type` - ส่งออกรายงาน
- `GET /api/admin/telegram-settings` - การตั้งค่า Telegram
- `PUT /api/admin/telegram-settings/:id` - อัปเดตการตั้งค่า
- `POST /api/admin/telegram-settings/test` - ทดสอบการแจ้งเตือน

## 🐛 Troubleshooting

### Common Issues

1. **Google Sheets API Quota Exceeded**
   - ระบบจะใช้ cache และ emergency mode
   - ตรวจสอบ API usage ใน Google Cloud Console

2. **Telegram Notifications Not Working**
   - ตรวจสอบ Bot Token และ Chat ID
   - ใช้ test function ในแดชบอร์ด admin

3. **Excel Export Errors**
   - ตรวจสอบ memory limits
   - ลดช่วงวันที่ในการส่งออก

## 📞 Support

สำหรับการสนับสนุนและคำถาม:
- 📧 Email: [your-email]
- 💬 LINE: [your-line-id]

## 📄 License

This project is licensed under the MIT License.

## 🙏 Acknowledgments

- Google Sheets API
- Telegram Bot API
- OpenStreetMap Nominatim
- ExcelJS Library
- Moment.js
