# 🚀 คู่มือ Deploy Time Tracker บน Render.com

## 📋 Overview ระบบ

ระบบ Time Tracker ของคุณทำงานแบบนี้:

```
Frontend (HTML/LIFF) → Node.js (Render.com) → Google Sheets API
                           ↓
                    Google Apps Script ← Telegram Bot
                    (Maps & Notifications)
```

## 🔧 การเตรียมไฟล์

### 1. สร้างโครงสร้างโปรเจค
```
time-tracker-render/
├── server.js              # Main server file
├── package.json           # Dependencies
├── .env                   # Environment variables (สำหรับ local)
├── public/
│   └── index.html         # Frontend
└── README.md
```

### 2. ไฟล์ที่ต้องอัพโหลดไป GitHub

**ไม่ต้องอัพโหลด `.env`** - จะตั้งค่าใน Render dashboard แทน

## 🌐 ขั้นตอนการ Deploy บน Render.com

### Step 1: สร้าง GitHub Repository
1. สร้าง repo ใหม่บน GitHub
2. อัพโหลดไฟล์ทั้งหมด ยกเว้น `.env`
3. Push ไปยัง main branch

### Step 2: สร้าง Web Service บน Render
1. ไปที่ [render.com](https://render.com)
2. สร้างบัญชีหรือ login
3. กด "New" → "Web Service"
4. เชื่อมต่อ GitHub repository
5. ตั้งค่าดังนี้:

```
Name: time-tracker-huanna
Region: Singapore (ใกล้ไทยที่สุด)
Branch: main
Build Command: npm install
Start Command: npm start
Instance Type: Free
```

### Step 3: ตั้งค่า Environment Variables
ใน Render dashboard ไปที่ "Environment" และเพิ่มตัวแปรเหล่านี้:

```bash
# Google Sheets
GOOGLE_SPREADSHEET_ID=1EMg58MQw7rmdF3pRx6oHAuIZbTKcDbWqct0vrMhzD9Q
GOOGLE_CLIENT_EMAIL=time-tracker-node-js@gen-lang-client-0377192621.iam.gserviceaccount.com
GOOGLE_PRIVATE_KEY=-----BEGIN PRIVATE KEY-----\nMIIEvAIBADANBgkqhkiG9w0BAQEFAASCBKYwggSiAgEAAoIBAQC8LxxotfUh2ZZ0\n/ZmsF1Cn/G3hC+8iXDxBkujTUTCDkj7ObW0vDUVJW7uZsyzx4PXeka7esKTEZzhx\nS6vKhAGnLEGY9/mWdBQrZnVOQdjohvkwQG472sJsBs7GVuPey9EHbCea62C0yGAV\nVcn9L8+CsE1gB2rhqv2nedTcPo3suKrODa00Wz0AAkLfilACyM+wgPDd0GenyROV\njPl1xm4kyIvm9z4I9fLKpOE5zr1FdnViBmOcNr1f4rHONGGwMpC77jjHPfwC6f3a\ni0Z4GEiDze6kyJM6SVanr0eFpUVS7yyip4YGC54xQlSDnlLx6SZ8XqiuI66oVUMD\n/CMe8hp9AgMBAAECggEABdQ3vw6Tzz6cKHeKgQgf2XQ6OxRRjfDpdOaGC7WiGRE4\nnNBK54Azuzf6MaKZK8zaENDWZ9N05xiDaQ78/ULlgjYeugxEUOK7lTSRQaFMhLdZ\nlKMKRxRZnVsAoKgkWsxZZy90cpoD3tWuFDsaDJukg9nOK8FPEDpprPxbGY9eegyY\nlFCwYoB1nao0aWE1ez4S3RSv5hxJc/QNd3SFqbsLVJSNTFOv05JXt/yEkVE43OfW\nUD85khWmCEMYjbVDg2cmDTac5+jGIyGhgTg6Km5y+W8tsSrZfH7dm84Jpi8OdBLa\n3UcESrF3ZK3qsI5tnbytvJhhR4K+xXLT5mp29VlGAQKBgQD7o0/rOwlBisloiqAE\nr2ODAjGnZN7M4BNd53ayDGomk9SYGjUc3+LRIVId9uYj0AWmlaYEmSCUrSFz9cmQ\nB2edljazn6IaA43pHnBloEKMkrLw3tRDv4Py83ackKlXaB9VaWYNL8xbs3sppZah\niMdPqbIUAsYqXDkhNzrBYaxOvQKBgQC/cjX6EW008pivgduOnFeqnLYRYTZDOwkD\nrVTwnfwNGWrDAtxXOnTUMHtMvYMvgvjZ4qE2vpUCX2GMb4SkrkIQnG1adthosXQ2\nrUIMpVBi7IvXpRkNOqnifAWJYltwbElKeFMueDNlBtGWgRLEz364pgReNFDe0ZVH\nOSmgUAiWwQKBgGfdqgAzVwe5rJa9GX21k0KhJjOs/BXeq7/H6YNmgm43+Llrn96y\nPuIJeeaqYaYImDyBaoxdVEhqCfPeUPtlQwyV0zBjRLquGuZNTSF1e+KgLsIjh8QL\nCgC/I4dOYseUT9KmdZwdzaFQPRccpUc5uOMV7U47MuaLOH2QWW02zrOVAoGAGHHM\n0pFHEGupcz9xeVQdHXvFA5MWCp+PFxkar158wG9uYlgLKlgccrt+At5v0bE3dRqq\n2wKapCLpobTbiut1JAnVLKfgGf4OiKy2skapbPgnIvHBsR68cl7Dljco1cH92bRj\napuOdGfaew0gCGE2HP2VsTGc4daA6QczeXS+pAECgYARWd0Kl6cRTBh5PmygKmpw\na0oNvu2q6kAb7eWHqaHPQfc4qKBbkep15hOCH5cdb8SxlZ3GIoQOfgt9yrWIXwUE\nIBircSR1IX6DL+f9JaRxTYJ4Rr6WwwVvCy62R21+fqAyj1w9cbupWgb+/6Dr+SMU\nyiReJJ2lFAZyVrjF3RQUkA==\n-----END PRIVATE KEY-----\n

# Telegram
TELEGRAM_BOT_TOKEN=7610983723:AAEFXDbDlq5uTHeyID8Fc5XEmIUx-LT6rJM
TELEGRAM_CHAT_ID=7809169283

# Server
NODE_ENV=production
KEEP_ALIVE_ENABLED=true

# Render specific (อัพเดทตาม URL จริง)
RENDER_EXTERNAL_HOSTNAME=time-tracker-huanna.onrender.com
RENDER_SERVICE_URL=https://time-tracker-huanna.onrender.com

# Security
GSA_WEBHOOK_SECRET=huanna-secret-2024
WEBHOOK_SECRET_KEY=huanna-webhook-secret
```

**⚠️ สำคัญ:** 
- `RENDER_EXTERNAL_HOSTNAME` = ชื่อ app ที่คุณตั้งใน Render
- `RENDER_SERVICE_URL` = URL เต็มของ service

### Step 4: Deploy
1. กด "Create Web Service"
2. รอให้ build เสร็จ (ประมาณ 2-3 นาที)
3. เมื่อเสร็จจะได้ URL เช่น: `https://time-tracker-huanna.onrender.com`

## 🤖 ตั้งค่า Google Apps Script

### Step 1: สร้าง Google Apps Script Project
1. ไปที่ [script.google.com](https://script.google.com)
2. สร้าง New Project
3. วาง code จาก "Google Apps Script - Map Service"
4. อัพเดท CONFIG ใน code:

```javascript
const CONFIG = {
  TELEGRAM: {
    BOT_TOKEN: "7610983723:AAEFXDbDlq5uTHeyID8Fc5XEmIUx-LT6rJM",
    CHAT_ID: "-4587553843"  // หรือ Chat ID ของคุณ
  },
  RENDER_SERVICE: {
    BASE_URL: "https://time-tracker-huanna.onrender.com", // URL จริงของ Render
    WEBHOOK_SECRET: "huanna-secret-2024"
  }
};
```

### Step 2: Deploy GSA เป็น Web App
1. กด "Deploy" → "New deployment"
2. ตั้งค่า:
   - Type: Web app
   - Execute as: Me
   - Who has access: Anyone
3. กด "Deploy"
4. คัดลอก URL ที่ได้

### Step 3: อัพเดท Environment Variables ใน Render
เพิ่มตัวแปรนี้ใน Render:
```
GSA_MAP_WEBHOOK_URL=https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec
```

### Step 4: ตั้งค่า Keep-Alive Triggers
ใน Google Apps Script:
1. รัน function `setupTriggers()` ครั้งเดียว
2. จะสร้าง trigger ปิง Render ทุก 10 นาที

## ⚙️ อัพเดท Frontend

อัพเดท `index.html` ใน public folder:

```javascript
// เปลี่ยน apiUrl จาก localhost เป็น Render URL
var apiUrl = 'https://time-tracker-huanna.onrender.com/api';
```

## 🔧 การทดสอบระบบ

### ทดสอบ API Endpoints:
```bash
# Health check
curl https://time-tracker-huanna.onrender.com/api/health

# Ping
curl https://time-tracker-huanna.onrender.com/api/ping

# Test GSA
https://script.google.com/macros/s/YOUR_ID/exec?action=test
```

### ทดสอบ Keep-Alive:
1. เช็ค logs ใน Render dashboard
2. ดู console logs ใน Google Apps Script
3. เช็ค Telegram ว่ามีข้อความทดสอบไหม

## 📱 การใช้งาน

1. เปิด `https://time-tracker-huanna.onrender.com`
2. ระบบจะทำงานเหมือนเดิม แต่รันบน cloud
3. Keep-alive จะทำงานอัตโนมัติในเวลา 05:00-10:00 และ 15:00-20:00

## 🛠️ การแก้ไขปัญหา

### Render หลับ (503 Error):
- เช็คว่า Keep-alive triggers ทำงานไหม
- ดูใน GSA execution log
- ปิง manual: เข้า URL `/api/ping`

### Google Sheets ไม่เชื่อมต่อ:
- เช็ค GOOGLE_PRIVATE_KEY format
- ตรวจสอบ Service Account permissions

### Telegram ไม่ส่งข้อความ:
- ทดสอบ GSA ด้วย `?action=test`
- เช็ค Bot Token และ Chat ID

## 📈 Monitoring

### Render Dashboard:
- ดู Logs real-time
- เช็ค Memory/CPU usage
- ดู Deploy history

### Google Apps Script:
- ดู Execution transcript
- เช็ค Trigger history
- ดู Quota usage

## 💡 Tips สำหรับแผนฟรี

1. **Keep-Alive ทำงานเฉพาะเวลาทำงาน** - ประหยัด resources
2. **ใช้ GSA เป็น backup** - ถ้า Render หลับ GSA จะปิงให้ตื่น
3. **Monitor Quota** - Google Apps Script มี daily quota
4. **Cache data** - ลด API calls ไปยัง Google Sheets

ระบบนี้จะทำให้ Time Tracker ของคุณทำงานบน cloud ได้แบบ 24/7 แม้ใช้แผนฟรี! 🎉