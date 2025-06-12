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
