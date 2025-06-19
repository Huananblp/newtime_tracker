# 🚀 Time Tracker System v2.1.2 - Employee Validation Update

## 📋 Overview ระบบ

ระบบ Time Tracker สำหรับองค์การบริหารส่วนตำบลข่าใหญ่ พร้อมระบบตรวจสอบความถูกต้องของชื่อพนักงาน

```
Frontend (HTML/LIFF) → Node.js (Render.com) → Google Sheets API
        ↓                     ↓
Employee Validation    Location Service (OpenStreetMap)
        ↓                     ↓
Google Apps Script ← Telegram Bot (Optional)
(Maps & Notifications)
```

## 🆕 อัปเดตเวอร์ชัน 2.1.2 (20 มิถุนายน 2568)

### ✨ **ฟีเจอร์ใหม่: Employee Validation System**
- 🔒 **ตรวจสอบชื่อพนักงาน**: ลงเวลาได้เฉพาะชื่อที่มีในรายการเท่านั้น
- 🎯 **Real-time Validation**: แสดงสีเขียว/แดงทันทีเมื่อพิมพ์
- 🚫 **Block Invalid Names**: ป้องกันการลงเวลาด้วยชื่อไม่ถูกต้อง
- 📝 **Smart Error Messages**: แสดงข้อความแจ้งเตือนชัดเจน

### 🔧 การปรับปรุงเทคนิค
- ✅ เก็บรายชื่อพนักงานในตัวแปร `validEmployees[]`
- ✅ ตรวจสอบชื่อก่อนเรียก API ลงเวลา
- ✅ UI Feedback แบบ Real-time
- ✅ รองรับ Case-insensitive Validation

## 🎯 ความสามารถหลัก

### **1. ระบบลงเวลาทำงาน**
- ✅ ลงเวลาเข้า-ออกงาน ผ่าน LINE LIFF
- ✅ **ตรวจสอบชื่อพนักงานก่อนลงเวลา** 🆕
- ✅ ตรวจสอบสถานะการทำงาน
- ✅ บันทึกพิกัด GPS จริงจากผู้ใช้เท่านั้น
- ✅ แปลงพิกัดเป็นชื่อสถานที่อัตโนมัติ (OpenStreetMap API)
- ✅ บันทึกเวลาตาม Timezone ไทย (Asia/Bangkok)

### **2. ระบบตรวจสอบพนักงาน** 🆕
- � **Auto-complete**: แสดงรายชื่อพนักงานขณะพิมพ์
- ✅ **Validation**: ตรวจสอบชื่อกับฐานข้อมูล Google Sheets
- 🎨 **Visual Feedback**: สีเขียว (ถูกต้อง) / สีแดง (ผิด)
- 🚫 **Access Control**: ป้องกันการลงเวลาด้วยชื่อปลอม

### **3. การจัดการข้อมูล**
- ✅ เชื่อมต่อ Google Sheets แบบ Real-time
- ✅ ส่งออกรายงาน Excel (รายวัน/ช่วงวันที่)
- ✅ รายงานสรุปการเข้างาน (แยกตามวัน)
- ✅ คำนวณชั่วโมงทำงานอัตโนมัติ

### **4. ระบบแอดมิน**
- ✅ Panel ควบคุมแอดมิน (/admin/login)
- ✅ ดูรายชื่อพนักงานจาก Google Sheets
- ✅ ส่งออกรายงาน Excel
- ✅ ตรวจสอบสถานะการทำงาน

## 🔧 เทคโนโลยีที่ใช้

### **Backend (Node.js)**
- `express` 4.18.2 - Web Framework
- `google-spreadsheet` 4.1.1 - Google Sheets API
- `moment-timezone` 0.6.0 - จัดการเวลา
- `exceljs` 4.4.0 - ส่งออก Excel
- `node-fetch` - HTTP Client (OpenStreetMap API)
- `bcryptjs` 2.4.3 - เข้ารหัสรหัสผ่าน
- `jsonwebtoken` 9.0.2 - JWT Authentication

### **Frontend**
- `HTML5` + `Bootstrap 5`
- `LINE LIFF SDK`
- `SweetAlert2`
- `jQuery` + `jQuery UI` (Autocomplete)

### **External APIs**
- `Google Sheets API` - ฐานข้อมูล
- `OpenStreetMap Nominatim` - แปลงพิกัดเป็นชื่อสถานที่
- `LINE LIFF` - Authentication & Frontend

## 📊 โครงสร้าง Google Sheets

### **Sheet: MAIN**
- พนักงาน, LINE Name, รูปภาพ, เวลาเข้า, ข้อมูลเพิ่มเติม
- เวลาออก, พิกัดเข้า, ที่อยู่เข้า, พิกัดออก, ที่อยู่ออก, ชั่วโมงทำงาน

### **Sheet: EMPLOYEES** 🔑
- **รายชื่อพนักงานที่ได้รับอนุญาต** (ใช้สำหรับ Validation)
- ข้อมูลส่วนตัวและสิทธิ์การเข้าถึง

### **Sheet: ON WORK**
- ข้อมูลการทำงานปัจจุบัน (สำหรับตรวจสอบสถานะ)

## ⚙️ การกำหนดค่าสำคัญ

### **Environment Variables (.env)**
```env
GOOGLE_SPREADSHEET_ID=xxxxx
GOOGLE_CLIENT_EMAIL=xxxxx
GOOGLE_PRIVATE_KEY=xxxxx
LIFF_ID=xxxxx
TELEGRAM_BOT_TOKEN=xxxxx (optional)
TELEGRAM_CHAT_ID=xxxxx (optional)
```

### **การตั้งค่าเวลา**
- **Timezone:** `Asia/Bangkok`
- **เวลามาสาย:** หลัง 08:30
- **รูปแบบเวลา:** `YYYY-MM-DD HH:mm:ss`

### **การตั้งค่าการตรวจสอบพนักงาน** 🆕
- **Validation Method:** Case-insensitive string matching
- **Data Source:** Google Sheets EMPLOYEES tab
- **Real-time Check:** ทุกครั้งที่พิมพ์และก่อนลงเวลา
- **Fallback:** ข้อมูลตัวอย่างถ้าโหลดจาก Sheets ไม่ได้

## 🚀 การ Deploy
- **Platform:** Render.com
- **Port:** 3000 (auto-detect)
- **Keep-Alive:** เปิดใช้งาน
- **Auto-Deploy:** จาก Git Repository

## 🔐 ความปลอดภัย
- ✅ JWT Authentication สำหรับแอดมิน
- ✅ BCrypt สำหรับเข้ารหัสรหัสผ่าน
- ✅ Environment Variables สำหรับ API Keys
- ✅ CORS Configuration
- ✅ GPS Validation (ใช้ตำแหน่งจริงเท่านั้น)
- ✅ **Employee Name Validation** (ป้องกันการลงเวลาปลอม) 🆕

## 📱 การใช้งาน
1. **พนักงาน:** เข้าระบบผ่าน LINE LIFF → พิมพ์ชื่อ (ต้องถูกต้อง) → ลงเวลาเข้า-ออก
2. **แอดมิน:** เข้า `/admin/login` → ดูรายงาน/ส่งออก Excel → จัดการรายชื่อพนักงาน
3. **ข้อมูล:** อัปเดตใน Google Sheets แบบ Real-time

## 🎯 จุดเด่นเวอร์ชัน 2.1.2
- 🔒 **Employee Validation System** - ป้องกันการลงเวลาด้วยชื่อปลอม
- 🎨 **Smart UI Feedback** - แสดงสีเขียว/แดงตามความถูกต้อง
- ⚡ **Real-time Validation** - ตรวจสอบทันทีขณะพิมพ์
- 🌍 **ใช้ตำแหน่งจริงเท่านั้น** (ไม่ใช้ IP หรือ Default Location)
- 🕐 **เวลาไทยแม่นยำ** (ไม่มี UTC Offset Issue)
- 📍 **ชื่อสถานที่อัตโนมัติ** (แปลงจากพิกัด GPS)
- 📊 **รายงานครบถ้วน** (รายวัน + Excel Export)
- 🔄 **Real-time Sync** กับ Google Sheets

## 📝 การอัปเดตรายชื่อพนักงาน

### **สำหรับแอดมิน:**
1. เข้า Google Sheets ที่เชื่อมต่อกับระบบ
2. ไปที่ Tab "EMPLOYEES"
3. เพิ่ม/แก้ไข/ลบชื่อพนักงานในคอลัมน์ที่กำหนด
4. ระบบจะอัปเดตรายชื่ออัตโนมัติทันที

### **หมายเหตุสำคัญ:**
- ชื่อที่ไม่มีในรายการจะไม่สามารถลงเวลาได้
- การตรวจสอบไม่สนใจตัวพิมพ์เล็ก/ใหญ่
- ระบบรองรับชื่อภาษาไทยและอังกฤษ

## 📦 สำรองข้อมูล
- **เวอร์ชัน:** 2.1.2
- **วันที่สำรอง:** 20 มิถุนายน 2568
- **โฟลเดอร์สำรอง:** `timetracker_backup_v2.1.2_20250620_001653`

## 🛠️ การแก้ไขปัญหา

### **ปัญหาที่อาจพบ:**
1. **ชื่อถูกต้องแต่ลงเวลาไม่ได้**
   - ตรวจสอบการเชื่อมต่อ Google Sheets
   - ลองรีเฟรชหน้าเว็บ

2. **ไม่แสดงรายชื่อ Autocomplete**
   - ตรวจสอบ Console (F12) หาข้อผิดพลาด
   - ตรวจสอบการตั้งค่า Google Sheets API

3. **ชื่อแสดงสีแดงแต่เป็นชื่อที่ถูกต้อง**
   - ตรวจสอบการสะกดชื่อใน Google Sheets
   - ตรวจสอบช่องว่างหรืออักขระพิเศษ

**🎉 ระบบพร้อมใช้งานเต็มประสิทธิภาพพร้อมการตรวจสอบความปลอดภัยขั้นสูง!**
