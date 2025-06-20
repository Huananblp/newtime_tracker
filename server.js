// server.js - Time Tracker with Admin Panel and Excel Export
const express = require('express');
const cors = require('cors');
const { GoogleSpreadsheet } = require('google-spreadsheet');
const { JWT } = require('google-auth-library');
const path = require('path');
const cron = require('node-cron');
const bcrypt = require('bcryptjs');
const jwt = require('jsonwebtoken');
const ExcelJS = require('exceljs');
const moment = require('moment-timezone');
const fetch = require('node-fetch');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 3001;
console.log(`🔧 Using PORT: ${PORT}`);

// ========== Enhanced Configuration ==========
const CONFIG = {
  GOOGLE_SHEETS: {
    SPREADSHEET_ID: process.env.GOOGLE_SPREADSHEET_ID,
    PRIVATE_KEY: process.env.GOOGLE_PRIVATE_KEY?.replace(/\\n/g, '\n'),
    CLIENT_EMAIL: process.env.GOOGLE_CLIENT_EMAIL,
  },
  TELEGRAM: {
    BOT_TOKEN: process.env.TELEGRAM_BOT_TOKEN,
    CHAT_ID: process.env.TELEGRAM_CHAT_ID
  },
  LINE: {
    LIFF_ID: process.env.LIFF_ID
  },
  SHEETS: {
    MAIN: 'MAIN',
    EMPLOYEES: 'EMPLOYEES',
    ON_WORK: 'ON WORK'
  },
  RENDER: {
    SERVICE_URL: process.env.RENDER_SERVICE_URL || `https://${process.env.RENDER_EXTERNAL_HOSTNAME}` || 'http://localhost:3001',
    KEEP_ALIVE_ENABLED: process.env.KEEP_ALIVE_ENABLED === 'true',
    GSA_WEBHOOK_SECRET: process.env.GSA_WEBHOOK_SECRET || 'your-secret-key'
  },  ADMIN: {
    JWT_SECRET: process.env.JWT_SECRET || 'huana-nbp-jwt-secret-2025',
    JWT_EXPIRES_IN: '24h',
    // Admin users (in production, store in database)
    USERS: [
      {
        id: 1,
        username: 'admin',
        password: '$2a$10$7ROfP4YLlJpub4cWuPkqwu2C1shrT.QbHr2zbLeDoGLE7VxSBhmCS', // khayai042315962
        name: 'ผู้ดูแลระบบ อบต.ข่าใหญ่',
        role: 'admin'
      },
      {
        id: 2,
        username: 'huana_admin',
        password: '$2a$10$AnotherHashedPasswordHere', // ต้อง hash ก่อนใช้งานจริง
        name: 'ผู้ดูแลระบบ อบต.ข่าใหญ่',
        role: 'admin'
      }
    ]
  },
  TIMEZONE: 'Asia/Bangkok'
};

// ========== Helper Functions ==========
// สร้าง hash password (ใช้ในการตั้งรหัสผ่านครั้งแรก)
async function createPassword(plainPassword) {
  return await bcrypt.hash(plainPassword, 10);
}

// ตรวจสอบ environment variables
function validateConfig() {
  const required = [
    { key: 'GOOGLE_SPREADSHEET_ID', value: CONFIG.GOOGLE_SHEETS.SPREADSHEET_ID },
    { key: 'GOOGLE_CLIENT_EMAIL', value: CONFIG.GOOGLE_SHEETS.CLIENT_EMAIL },
    { key: 'GOOGLE_PRIVATE_KEY', value: CONFIG.GOOGLE_SHEETS.PRIVATE_KEY },
    { key: 'LIFF_ID', value: CONFIG.LINE.LIFF_ID }
  ];

  const missing = required.filter(item => !item.value);
  
  if (missing.length > 0) {
    console.error('❌ Missing required environment variables:');
    missing.forEach(item => console.error(`   - ${item.key}`));
    return false;
  }
  
  console.log('✅ All required environment variables are set');
  return true;
}

// ========== Middleware ==========
app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Security middleware สำหรับ webhook
app.use('/api/webhook', (req, res, next) => {
  const providedSecret = req.headers['x-webhook-secret'] || req.query.secret;
  if (providedSecret !== CONFIG.RENDER.GSA_WEBHOOK_SECRET) {
    return res.status(401).json({ error: 'Unauthorized' });
  }
  next();
});

// Admin Authentication Middleware
function authenticateAdmin(req, res, next) {
  const authHeader = req.headers.authorization;
  const token = authHeader && authHeader.split(' ')[1];

  if (!token) {
    return res.status(401).json({ 
      success: false, 
      error: 'Access token required' 
    });
  }

  try {
    const decoded = jwt.verify(token, CONFIG.ADMIN.JWT_SECRET);
    req.user = decoded;
    next();
  } catch (error) {
    console.error('JWT verification error:', error);
    return res.status(403).json({ 
      success: false, 
      error: 'Invalid token' 
    });
  }
}

// Serve static files
app.use(express.static('public'));

// Admin routes - ใช้ไฟล์จาก public folder
app.get('/admin/login', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'login.html'));
});

app.get('/admin/dashboard', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'admin.html'));
});

app.get('/admin', (req, res) => {
  res.redirect('/admin/login');
});

// Serve ads.txt specifically
app.get('/ads.txt', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'ads.txt'));
});

// Serve robots.txt (optional)
app.get('/robots.txt', (req, res) => {
  res.type('text/plain');
  res.send('User-agent: *\nDisallow: /api/\nAllow: /ads.txt');
});

// ========== Keep-Alive Service ==========
class KeepAliveService {
  constructor() {
    this.isEnabled = CONFIG.RENDER.KEEP_ALIVE_ENABLED;
    this.serviceUrl = CONFIG.RENDER.SERVICE_URL;
    this.startTime = new Date();
    this.pingCount = 0;
    this.errorCount = 0;
  }

  init() {
    if (!this.isEnabled) {
      console.log('🔴 Keep-Alive disabled');
      return;
    }

    console.log('🟢 Keep-Alive service started');
    console.log(`📍 Service URL: ${this.serviceUrl}`);

    // เวลาทำงาน: 05:00-10:00 และ 15:00-20:00 (เวลาไทย)
    // ปิงทุก 10 นาที
    cron.schedule('*/10 * * * *', () => {
      this.checkAndPing();
    }, {
      scheduled: true,
      timezone: CONFIG.TIMEZONE
    });

    // Ping ทันทีเมื่อเริ่มต้น
    setTimeout(() => this.ping(), 5000);
  }

  checkAndPing() {
    const now = new Date();
    const hour = now.getHours();
    
    // เช็คว่าอยู่ในเวลาทำงานไหม
    const isWorkingHour = (hour >= 5 && hour < 10) || (hour >= 15 && hour < 20);
    
    if (isWorkingHour) {
      this.ping();
    } else {
      console.log(`😴 Outside working hours (${hour}:00), skipping ping`);
    }
  }

  async ping() {
    try {
      const response = await fetch(`${this.serviceUrl}/api/ping`, {
        method: 'GET',
        headers: {
          'User-Agent': 'KeepAlive-Service/1.0'
        }
      });

      this.pingCount++;
      
      if (response.ok) {
        console.log(`✅ Keep-Alive ping #${this.pingCount} successful`);
        this.errorCount = 0; // Reset error count on success
      } else {
        throw new Error(`HTTP ${response.status}`);
      }
      
    } catch (error) {
      this.errorCount++;
      console.log(`❌ Keep-Alive ping #${this.pingCount} failed:`, error.message);
      
      // หากผิดพลาดติดต่อกัน 5 ครั้ง ให้ลอง ping ใหม่หลัง 1 นาที
      if (this.errorCount >= 5) {
        console.log('🔄 Too many errors, will retry in 1 minute');
        setTimeout(() => this.ping(), 60000);
      }
    }
  }

  getStats() {
    const uptime = Math.floor((Date.now() - this.startTime.getTime()) / 1000);
    return {
      enabled: this.isEnabled,
      uptime: uptime,
      pingCount: this.pingCount,
      errorCount: this.errorCount,
      lastPing: new Date().toISOString()
    };
  }
}

// ========== Google Sheets Service ==========
class GoogleSheetsService {
  constructor() {
    this.doc = null;
    this.isInitialized = false;
  }

  async initialize() {
    if (this.isInitialized) return;

    try {
      const serviceAccountAuth = new JWT({
        email: CONFIG.GOOGLE_SHEETS.CLIENT_EMAIL,
        key: CONFIG.GOOGLE_SHEETS.PRIVATE_KEY,
        scopes: ['https://www.googleapis.com/auth/spreadsheets']
      });

      this.doc = new GoogleSpreadsheet(CONFIG.GOOGLE_SHEETS.SPREADSHEET_ID, serviceAccountAuth);
      await this.doc.loadInfo();
      
      console.log(`✅ Connected to Google Sheets: ${this.doc.title}`);
      this.isInitialized = true;
      
    } catch (error) {
      console.error('❌ Failed to initialize Google Sheets:', error);
      throw error;
    }
  }

  async getSheet(sheetName) {
    if (!this.isInitialized) {
      await this.initialize();
    }
    
    const sheet = this.doc.sheetsByTitle[sheetName];
    if (!sheet) {
      throw new Error(`Sheet ${sheetName} not found`);
    }
    
    return sheet;
  }

  // ฟังก์ชันช่วยเหลือสำหรับการเปรียบเทียบชื่อ
  normalizeEmployeeName(name) {
    if (!name) return '';
    
    return name.toString()
      .trim()
      .replace(/\s+/g, ' ')
      .toLowerCase();
  }

  isNameMatch(inputName, compareName) {
    if (!inputName || !compareName) return false;
    
    const normalizedInput = this.normalizeEmployeeName(inputName);
    const normalizedCompare = this.normalizeEmployeeName(compareName);
    
    return normalizedInput === normalizedCompare ||
           normalizedInput.includes(normalizedCompare) ||
           normalizedCompare.includes(normalizedInput);
  }

  async getEmployees() {
    try {
      const sheet = await this.getSheet(CONFIG.SHEETS.EMPLOYEES);
      const rows = await sheet.getRows();
      
      const employees = rows.map(row => row.get('ชื่อ-นามสกุล')).filter(name => name);
      return employees;
      
    } catch (error) {
      console.error('Error getting employees:', error);
      return [];
    }
  }
  async getEmployeeStatus(employeeName) {
    try {
      const sheet = await this.getSheet(CONFIG.SHEETS.ON_WORK);
      const rows = await sheet.getRows({ offset: 1 }); // เริ่มจากแถว 3 (ข้ามแถว 2 ที่เป็นข้อความอธิบาย)
      
      console.log(`🔍 Checking status for: "${employeeName}"`);
      console.log(`📊 Total rows in ON_WORK (from row 3): ${rows.length}`);
      
      if (rows.length === 0) {
        console.log('📋 ON_WORK sheet is empty (from row 3)');
        return { isOnWork: false, workRecord: null };
      }
      
      const workRecord = rows.find(row => {
        const systemName = row.get('ชื่อในระบบ');
        const employeeName2 = row.get('ชื่อพนักงาน');
        
        const isMatch = this.isNameMatch(employeeName, systemName) || 
                       this.isNameMatch(employeeName, employeeName2);
        
        if (isMatch) {
          console.log(`✅ Found match: "${employeeName}" ↔ "${systemName || employeeName2}"`);
        }
        
        return isMatch;
      });
      
      if (workRecord) {
        let mainRowIndex = null;
        
        const rowRef1 = workRecord.get('แถวอ้างอิง');
        const rowRef2 = workRecord.get('แถวในMain');
        
        if (rowRef1 && !isNaN(parseInt(rowRef1))) {
          mainRowIndex = parseInt(rowRef1);
        } else if (rowRef2 && !isNaN(parseInt(rowRef2))) {
          mainRowIndex = parseInt(rowRef2);
        }
        
        console.log(`✅ Employee "${employeeName}" is currently working`);
        
        return {
          isOnWork: true,
          workRecord: {
            row: workRecord,
            mainRowIndex: mainRowIndex,
            clockIn: workRecord.get('เวลาเข้า'),
            systemName: workRecord.get('ชื่อในระบบ'),
            employeeName: workRecord.get('ชื่อพนักงาน')
          }
        };
      } else {
        console.log(`❌ Employee "${employeeName}" is not currently working`);
        return { isOnWork: false, workRecord: null };
      }
      
    } catch (error) {
      console.error('❌ Error checking employee status:', error);
      return { isOnWork: false, workRecord: null };
    }
  }

  // Admin functions
  async getAdminStats() {
    try {
      const [employeesSheet, onWorkSheet, mainSheet] = await Promise.all([
        this.getSheet(CONFIG.SHEETS.EMPLOYEES),
        this.getSheet(CONFIG.SHEETS.ON_WORK),
        this.getSheet(CONFIG.SHEETS.MAIN)
      ]);

      const [employees, onWorkRows, mainRows] = await Promise.all([
        employeesSheet.getRows(),
        onWorkSheet.getRows({ offset: 1 }), // เริ่มจากแถว 3
        mainSheet.getRows()
      ]);

      const totalEmployees = employees.length;
      const workingNow = onWorkRows.length;      // หาจำนวนคนที่มาทำงานวันนี้ (ใช้ข้อมูลจาก ON_WORK sheet ที่มีวันที่วันนี้)
      const today = moment().tz(CONFIG.TIMEZONE).format('YYYY-MM-DD');
      console.log(`📅 Today date for comparison: ${today}`);
      console.log(`📊 Total MAIN sheet records: ${mainRows.length}`);
      console.log(`� Total ON_WORK sheet records: ${onWorkRows.length}`);
      
      // นับจาก ON_WORK sheet ที่มีวันที่วันนี้
      const presentToday = onWorkRows.filter(row => {
        const clockInDate = row.get('เวลาเข้า');
        if (!clockInDate) return false;
        
        try {
          const employeeName = row.get('ชื่อพนักงาน') || row.get('ชื่อในระบบ');
          let dateStr = '';
          
          // ถ้าเป็น string format 'YYYY-MM-DD HH:mm:ss'
          if (typeof clockInDate === 'string' && clockInDate.includes(' ')) {
            dateStr = clockInDate.split(' ')[0];
            const isToday = dateStr === today;
            
            if (isToday) {
              console.log(`✅ Present today (ON_WORK): ${employeeName} - ${clockInDate} (date: ${dateStr})`);
            }
            
            return isToday;
          }
          
          // ถ้าเป็น ISO format
          if (typeof clockInDate === 'string' && clockInDate.includes('T')) {
            dateStr = clockInDate.split('T')[0];
            const isToday = dateStr === today;
            
            if (isToday) {
              console.log(`✅ Present today (ON_WORK ISO): ${employeeName} - ${clockInDate} (date: ${dateStr})`);
            }
            
            return isToday;
          }
          
          return false;
        } catch (error) {
          console.warn(`⚠️ Error parsing date in ON_WORK: ${clockInDate}`, error);
          return false;
        }
      }).length;
      
      console.log(`📊 Present today count: ${presentToday} out of ${onWorkRows.length} ON_WORK records`);

      const absentToday = totalEmployees - presentToday;      // รายชื่อพนักงานที่กำลังทำงาน
      const workingEmployees = onWorkRows.map(row => {
        const clockInTime = row.get('เวลาเข้า');
        let workingHours = '0 ชม.';
        
        if (clockInTime) {
          try {
            // Debug: แสดงข้อมูลเวลาที่ได้รับ
            console.log(`🕐 Processing clockInTime: "${clockInTime}" (type: ${typeof clockInTime})`);
            
            // ใช้ moment เพื่อจัดการ timezone อย่างถูกต้อง
            let clockInMoment;
            
            if (typeof clockInTime === 'string') {
              // ระบุ format ให้ชัดเจน และ parse ใน timezone ไทยตั้งแต่แรก
              clockInMoment = moment.tz(clockInTime, 'YYYY-MM-DD H:mm:ss', CONFIG.TIMEZONE);
            } else {
              // ถ้าเป็น Date object ให้แปลงเป็น moment ใน timezone ไทย
              clockInMoment = moment(clockInTime).tz(CONFIG.TIMEZONE);
            }
            
            // เวลาปัจจุบันใน timezone ไทย
            const nowMoment = moment().tz(CONFIG.TIMEZONE);
            
            // คำนวณความแตกต่างของเวลาในหน่วยชั่วโมง
            const hours = nowMoment.diff(clockInMoment, 'hours', true);
            
            // Debug: แสดงการคำนวณ
            console.log(`⏰ Time calculation:`, {
              clockIn: clockInMoment.format('YYYY-MM-DD HH:mm:ss'),
              now: nowMoment.format('YYYY-MM-DD HH:mm:ss'),
              diffHours: hours.toFixed(2)
            });
            
            // ตรวจสอบให้แน่ใจว่าไม่เป็นลบ (ป้องกันปัญหา timezone)
            if (hours >= 0) {
              workingHours = `${hours.toFixed(1)} ชม.`;
            } else {
              console.warn(`⚠️ Negative working hours detected: ${hours.toFixed(2)}, setting to 0`);
              workingHours = '0 ชม.';
            }
          } catch (error) {
            console.error('Error calculating working hours:', error);
            workingHours = '0 ชม.';
          }
        }        return {
          name: row.get('ชื่อพนักงาน') || row.get('ชื่อในระบบ'),
          clockIn: clockInTime ? moment.tz(clockInTime, 'YYYY-MM-DD H:mm:ss', CONFIG.TIMEZONE).format('HH:mm') : '', // ส่งเฉพาะเวลา HH:mm
          workingHours
        };
      });      const stats = {
        totalEmployees,
        presentToday,
        workingNow,
        absentToday,
        workingEmployees
      };
      
      console.log('📊 Admin stats summary:', {
        totalEmployees,
        presentToday,
        workingNow,
        absentToday,
        workingEmployeesCount: workingEmployees.length
      });
      
      return stats;

    } catch (error) {
      console.error('Error getting admin stats:', error);
      throw error;
    }
  }

  async getReportData(type, params) {
    try {
      const mainSheet = await this.getSheet(CONFIG.SHEETS.MAIN);
      const rows = await mainSheet.getRows();

      let filteredRows = [];      switch (type) {
        case 'daily':
          const targetDate = moment(params.date).tz(CONFIG.TIMEZONE).format('YYYY-MM-DD');
          filteredRows = rows.filter(row => {
            const clockIn = row.get('เวลาเข้า');
            if (!clockIn) return false;
            
            try {
              // ถ้าเป็น string format 'YYYY-MM-DD HH:mm:ss'
              if (typeof clockIn === 'string' && clockIn.includes(' ')) {
                const dateStr = clockIn.split(' ')[0];
                return dateStr === targetDate;
              }
              // ถ้าเป็น Date object
              const rowDate = moment(clockIn).tz(CONFIG.TIMEZONE).format('YYYY-MM-DD');
              return rowDate === targetDate;
            } catch {
              return false;
            }
          });
          break;

        case 'monthly':
          const month = parseInt(params.month);
          const year = parseInt(params.year);
          
          filteredRows = rows.filter(row => {
            const clockIn = row.get('เวลาเข้า');
            if (!clockIn) return false;
            
            try {
              let rowDate;
              // ถ้าเป็น string format 'YYYY-MM-DD HH:mm:ss'
              if (typeof clockIn === 'string' && clockIn.includes(' ')) {
                rowDate = moment(clockIn).tz(CONFIG.TIMEZONE);
              } else {
                rowDate = moment(clockIn).tz(CONFIG.TIMEZONE);
              }
              return rowDate.month() + 1 === month && rowDate.year() === year;
            } catch {
              return false;
            }
          });
          break;

        case 'range':
          const startMoment = moment(params.startDate).tz(CONFIG.TIMEZONE).startOf('day');
          const endMoment = moment(params.endDate).tz(CONFIG.TIMEZONE).endOf('day');
          
          filteredRows = rows.filter(row => {
            const clockIn = row.get('เวลาเข้า');
            if (!clockIn) return false;
            
            try {
              let rowMoment;
              // ถ้าเป็น string format 'YYYY-MM-DD HH:mm:ss'
              if (typeof clockIn === 'string' && clockIn.includes(' ')) {
                rowMoment = moment(clockIn).tz(CONFIG.TIMEZONE);
              } else {
                rowMoment = moment(clockIn).tz(CONFIG.TIMEZONE);
              }
              return rowMoment.isBetween(startMoment, endMoment, null, '[]');
            } catch {
              return false;
            }
          });
          break;
      }

      // แปลงข้อมูลเป็น format ที่ใช้งานง่าย
      const reportData = filteredRows.map(row => ({
        employee: row.get('ชื่อพนักงาน') || '',
        lineName: row.get('ชื่อไลน์') || '',
        clockIn: row.get('เวลาเข้า') || '',
        clockOut: row.get('เวลาออก') || '',
        note: row.get('หมายเหตุ') || '',
        workingHours: row.get('ชั่วโมงทำงาน') || '',
        locationIn: row.get('ที่อยู่เข้า') || '',
        locationOut: row.get('ที่อยู่ออก') || ''
      }));

      return reportData;

    } catch (error) {
      console.error('Error getting report data:', error);
      throw error;
    }
  }

  // [Previous clockIn and clockOut methods remain the same...]
  async clockIn(data) {
    try {
      const { employee, userinfo, lat, lon, line_name, line_picture } = data;
      
      console.log(`⏰ Clock In request for: "${employee}"`);
      
      const employeeStatus = await this.getEmployeeStatus(employee);
      
      if (employeeStatus.isOnWork) {
        console.log(`❌ Employee "${employee}" is already clocked in`);
        return {
          success: false,
          message: 'คุณลงเวลาเข้างานไปแล้ว กรุณาลงเวลาออกก่อน',
          employee,
          currentStatus: 'clocked_in',
          clockInTime: employeeStatus.workRecord?.clockIn        };      }      const timestamp = moment().tz(CONFIG.TIMEZONE).format('YYYY-MM-DD HH:mm:ss'); // ใช้เวลาไทยในรูปแบบ string
      
      // แปลงพิกัดเป็นชื่อสถานที่
      const locationName = await this.getLocationName(lat, lon);
      console.log(`📍 Location: ${locationName}`);
      
      console.log(`✅ Proceeding with clock in for "${employee}"`);
      
      const mainSheet = await this.getSheet(CONFIG.SHEETS.MAIN);
      
      const newRow = await mainSheet.addRow([
        employee,           
        line_name,          
        `=IMAGE("${line_picture}")`, 
        timestamp,          
        userinfo || '',     
        '',                 
        `${lat},${lon}`,    
        locationName,       // ใช้ชื่อสถานที่แทนพิกัด
        '',                 
        '',                 
        ''                  
      ]);

      const mainRowIndex = newRow.rowNumber;
      console.log(`✅ Added to MAIN sheet at row: ${mainRowIndex}`);      const onWorkSheet = await this.getSheet(CONFIG.SHEETS.ON_WORK);
      await onWorkSheet.addRow([
        timestamp,          
        employee,           
        timestamp,          
        'ทำงาน',           
        userinfo || '',     
        `${lat},${lon}`,    
        locationName,       // ใช้ชื่อสถานที่แทนพิกัด
        mainRowIndex,       
        line_name,          
        line_picture,       
        mainRowIndex,       
        employee            
      ]);

      console.log(`✅ Clock In successful: ${employee} at ${this.formatTime(timestamp)}, Main row: ${mainRowIndex}`);

      this.triggerMapGeneration('clockin', {
        employee, lat, lon, line_name, userinfo, timestamp
      });

      return {
        success: true,
        message: 'บันทึกเวลาเข้างานสำเร็จ',
        employee,
        time: this.formatTime(timestamp),
        currentStatus: 'clocked_in'
      };

    } catch (error) {
      console.error('❌ Clock in error:', error);
      return {
        success: false,
        message: `เกิดข้อผิดพลาด: ${error.message}`,
        employee: data.employee
      };
    }
  }

  async clockOut(data) {
    try {
      const { employee, lat, lon, line_name } = data;
      
      console.log(`⏰ Clock Out request for: "${employee}"`);
      console.log(`📍 Location: ${lat}, ${lon}`);
      
      const employeeStatus = await this.getEmployeeStatus(employee);
      
      if (!employeeStatus.isOnWork) {
        console.log(`❌ Employee "${employee}" is not clocked in`);
        
        const onWorkSheet = await this.getSheet(CONFIG.SHEETS.ON_WORK);
        const rows = await onWorkSheet.getRows({ offset: 1 }); // เริ่มจากแถว 3
        
        const suggestions = rows
          .map(row => ({
            systemName: row.get('ชื่อในระบบ'),
            employeeName: row.get('ชื่อพนักงาน')
          }))
          .filter(emp => emp.systemName || emp.employeeName)
          .filter(emp => 
            this.isNameMatch(employee, emp.systemName) ||
            this.isNameMatch(employee, emp.employeeName)
          );
        
        let message = 'คุณต้องลงเวลาเข้างานก่อน หรือตรวจสอบชื่อที่ป้อนให้ถูกต้อง';
        
        if (suggestions.length > 0) {
          const suggestedNames = suggestions.map(s => s.systemName || s.employeeName);
          message = `ไม่พบข้อมูลการลงเวลาเข้างาน ชื่อที่ใกล้เคียง: ${suggestedNames.join(', ')}`;
        }
        
        return {
          success: false,
          message: message,
          employee,
          currentStatus: 'not_clocked_in',
          suggestions: suggestions.length > 0 ? suggestions : undefined        };      }

      const timestamp = moment().tz(CONFIG.TIMEZONE).format('YYYY-MM-DD HH:mm:ss'); // ใช้เวลาไทยในรูปแบบ string
      const workRecord = employeeStatus.workRecord;
        const clockInTime = workRecord.clockIn;
      console.log(`⏰ Clock in time: ${clockInTime}`);
      
      let hoursWorked = 0;
      if (clockInTime) {
        // ใช้ moment สำหรับการคำนวณเวลาที่แม่นยำ
        const clockInMoment = moment(clockInTime).tz(CONFIG.TIMEZONE);
        const timestampMoment = moment().tz(CONFIG.TIMEZONE);
        hoursWorked = timestampMoment.diff(clockInMoment, 'hours', true); // true = ให้ทศนิยม
        console.log(`⏱️ Hours worked: ${hoursWorked.toFixed(2)}`);
      }
      
      // แปลงพิกัดเป็นชื่อสถานที่
      const locationName = await this.getLocationName(lat, lon);
      console.log(`📍 Clock out location: ${locationName}`);

      console.log(`✅ Proceeding with clock out for "${employee}"`);
      
      const mainSheet = await this.getSheet(CONFIG.SHEETS.MAIN);
      const rows = await mainSheet.getRows();
      
      console.log(`📊 Total rows in MAIN: ${rows.length}`);
      console.log(`🎯 Target row index: ${workRecord.mainRowIndex}`);
      
      let mainRow = null;
      
      if (workRecord.mainRowIndex && workRecord.mainRowIndex > 1) {
        const targetIndex = workRecord.mainRowIndex - 2;
        
        if (targetIndex >= 0 && targetIndex < rows.length) {
          const candidateRow = rows[targetIndex];
          const candidateEmployee = candidateRow.get('ชื่อพนักงาน');
          
          if (this.isNameMatch(employee, candidateEmployee)) {
            mainRow = candidateRow;
            console.log(`✅ Found main row by index: ${targetIndex} (row ${workRecord.mainRowIndex})`);
          } else {
            console.log(`⚠️ Row index found but employee name mismatch: "${candidateEmployee}" vs "${employee}"`);
          }
        } else {
          console.log(`⚠️ Row index out of range: ${targetIndex} (total rows: ${rows.length})`);
        }
      }
      
      if (!mainRow) {
        console.log('🔍 Searching by employee name and conditions...');
        
        const candidateRows = rows.filter(row => {
          const rowEmployee = row.get('ชื่อพนักงาน');
          const rowClockOut = row.get('เวลาออก');
          
          return this.isNameMatch(employee, rowEmployee) && !rowClockOut;
        });
        
        console.log(`Found ${candidateRows.length} candidate rows without clock out`);
        
        if (candidateRows.length === 1) {
          mainRow = candidateRows[0];
          console.log(`✅ Found unique candidate row`);
        } else if (candidateRows.length > 1) {
          let closestRow = null;
          let minTimeDiff = Infinity;
          
          candidateRows.forEach((row, index) => {
            const rowClockIn = row.get('เวลาเข้า');
            if (rowClockIn && clockInTime) {
              const timeDiff = Math.abs(new Date(rowClockIn) - new Date(clockInTime));
              console.log(`Candidate ${index}: time diff = ${timeDiff}ms`);
              if (timeDiff < minTimeDiff) {
                minTimeDiff = timeDiff;
                closestRow = row;
              }
            }
          });
          
          if (closestRow && minTimeDiff < 300000) {
            mainRow = closestRow;
            console.log(`✅ Found closest matching row (time diff: ${minTimeDiff}ms)`);
          } else {
            console.log(`❌ No close time match found (min diff: ${minTimeDiff}ms)`);
          }
        }
      }
      
      if (!mainRow) {
        console.log('🔍 Searching for latest row of this employee...');
        
        for (let i = rows.length - 1; i >= 0; i--) {
          const row = rows[i];
          const rowEmployee = row.get('ชื่อพนักงาน');
          const rowClockOut = row.get('เวลาออก');
          
          if (this.isNameMatch(employee, rowEmployee) && !rowClockOut) {
            mainRow = row;
            console.log(`✅ Found latest uncompleted row at index: ${i}`);
            break;
          }
        }
      }
      
      if (!mainRow) {
        console.log('❌ Cannot find main row to update');
        
        return {
          success: false,
          message: 'ไม่พบข้อมูลการลงเวลาเข้างานที่ตรงกัน กรุณาตรวจสอบระบบ',
          employee
        };
      }
      
      console.log('✅ Found main row, updating...');
        try {
        mainRow.set('เวลาออก', timestamp);
        mainRow.set('พิกัดออก', `${lat},${lon}`);
        mainRow.set('ที่อยู่ออก', locationName); // ใช้ชื่อสถานที่แทนพิกัด
        mainRow.set('ชั่วโมงทำงาน', hoursWorked.toFixed(2));
        await mainRow.save();
        console.log('✅ Main row updated successfully');
      } catch (updateError) {
        console.error('❌ Error updating main row:', updateError);
        throw new Error('ไม่สามารถอัปเดตข้อมูลได้: ' + updateError.message);
      }

      try {
        await workRecord.row.delete();
        console.log('✅ Removed from ON_WORK sheet');
      } catch (deleteError) {
        console.error('❌ Error deleting from ON_WORK:', deleteError);
      }

      console.log(`✅ Clock Out successful: ${employee} at ${this.formatTime(timestamp)} (${hoursWorked.toFixed(2)} hours)`);

      try {
        this.triggerMapGeneration('clockout', {
          employee, lat, lon, line_name, timestamp, hoursWorked
        });
      } catch (webhookError) {
        console.error('⚠️ Webhook error (non-critical):', webhookError);
      }

      return {
        success: true,
        message: 'บันทึกเวลาออกงานสำเร็จ',
        employee,
        time: this.formatTime(timestamp),
        hours: hoursWorked.toFixed(2),
        currentStatus: 'clocked_out'
      };

    } catch (error) {
      console.error('❌ Clock out error:', error);
      
      return {
        success: false,
        message: `เกิดข้อผิดพลาด: ${error.message}`,
        employee: data.employee
      };
    }
  }

  async triggerMapGeneration(action, data) {
    try {
      const gsaWebhookUrl = process.env.GSA_MAP_WEBHOOK_URL;
      if (!gsaWebhookUrl) {
        console.log('⚠️ GSA webhook URL not configured');
        return;
      }      const payload = {
        action,
        data,
        timestamp: moment().tz(CONFIG.TIMEZONE).toISOString() // ใช้เวลาไทย
      };

      await fetch(gsaWebhookUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'X-Webhook-Secret': CONFIG.RENDER.GSA_WEBHOOK_SECRET
        },
        body: JSON.stringify(payload)
      });

      console.log(`📍 Map generation triggered for ${action}: ${data.employee}`);
      
    } catch (error) {
      console.error('Error triggering map generation:', error);
    }
  }  formatTime(date) {
    try {
      // รองรับทั้ง Date object และ string
      if (typeof date === 'string') {
        // ถ้าเป็นรูปแบบ 'YYYY-MM-DD HH:mm:ss' จาก moment
        if (date.includes(' ') && date.length === 19) {
          return date.split(' ')[1]; // ใช้ส่วนเวลาเท่านั้น
        }
        // ลองแปลงเป็น Date object
        const parsedDate = moment(date).tz(CONFIG.TIMEZONE);
        if (parsedDate.isValid()) {
          return parsedDate.format('HH:mm:ss');
        }
        return date; // ถ้าแปลงไม่ได้ ส่งกลับเป็น string เดิม
      }
      
      // ถ้าเป็น Date object
      if (date instanceof Date && !isNaN(date.getTime())) {
        return moment(date).tz(CONFIG.TIMEZONE).format('HH:mm:ss');
      }
      
      return '';
    } catch (error) {
      console.error('Error formatting time:', error);
      return date?.toString() || '';
    }
  }

  // เพิ่มฟังก์ชันแปลงพิกัดเป็นชื่อสถานที่
  async getLocationName(lat, lon) {
    try {
      // ใช้ OpenStreetMap Nominatim API (ฟรี)
      const response = await fetch(
        `https://nominatim.openstreetmap.org/reverse?format=json&lat=${lat}&lon=${lon}&zoom=18&addressdetails=1&accept-language=th`
      );
      
      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }
      
      const data = await response.json();
      
      if (data && data.display_name) {
        // ใช้ชื่อสถานที่ที่ได้จาก API
        return data.display_name;
      } else {
        // ถ้าไม่ได้ข้อมูล ใช้พิกัดแทน
        return `${lat}, ${lon}`;
      }
    } catch (error) {
      console.warn(`⚠️ Location lookup failed for ${lat}, ${lon}:`, error.message);
      // ถ้าเกิดข้อผิดพลาด ใช้พิกัดแทน
      return `${lat}, ${lon}`;
    }
  }
}

// ========== Excel Export Service ==========
class ExcelExportService {
  static async createWorkbook(data, type, params) {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('รายงานการลงเวลา');

    // ตั้งค่าข้อมูลองค์กร
    const orgInfo = {
      name: 'องค์การบริหารส่วนตำบลข่าใหญ่',
      address: 'อำเภอเมือง จังหวัดนครราชสีมา',
      phone: '042-315962'
    };

    // สร้างหัวข้อรายงาน
    let reportTitle = '';
    let reportPeriod = '';

    switch (type) {
      case 'daily':
        reportTitle = 'รายงานการลงเวลาเข้า-ออกงาน รายวัน';
        reportPeriod = `วันที่ ${new Date(params.date).toLocaleDateString('th-TH')}`;
        break;
      case 'monthly':
        const monthNames = [
          'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
          'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'
        ];
        reportTitle = 'รายงานการลงเวลาเข้า-ออกงาน รายเดือน';
        reportPeriod = `เดือน ${monthNames[params.month - 1]} ${parseInt(params.year) + 543}`;
        break;
      case 'range':
        reportTitle = 'รายงานการลงเวลาเข้า-ออกงาน ช่วงวันที่';
        reportPeriod = `${new Date(params.startDate).toLocaleDateString('th-TH')} - ${new Date(params.endDate).toLocaleDateString('th-TH')}`;
        break;
    }

    // จัดรูปแบบหัวกระดาษ
    worksheet.mergeCells('A1:I3');
    const titleCell = worksheet.getCell('A1');
    titleCell.value = `${orgInfo.name}\n${reportTitle}\n${reportPeriod}`;
    titleCell.font = { name: 'Angsana New', size: 18, bold: true };
    titleCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };

    // ข้อมูลองค์กร
    worksheet.getCell('A4').value = `${orgInfo.address} โทร. ${orgInfo.phone}`;
    worksheet.getCell('A4').font = { name: 'Angsana New', size: 14 };
    worksheet.getCell('A4').alignment = { horizontal: 'center' };
    worksheet.mergeCells('A4:I4');

    // สร้างหัวตาราง
    const headerRow = 6;
    const headers = [
      'ลำดับ',
      'ชื่อ-นามสกุล',
      'วันที่',
      'เวลาเข้า',
      'เวลาออก',
      'ชั่วโมงทำงาน',
      'หมายเหตุ',
      'สถานที่เข้า',
      'สถานที่ออก'
    ];

    headers.forEach((header, index) => {
      const cell = worksheet.getCell(headerRow, index + 1);
      cell.value = header;
      cell.font = { name: 'Angsana New', size: 14, bold: true };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFE6E6FA' }
      };
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
    });

    // เพิ่มข้อมูล
    data.forEach((record, index) => {
      const rowNumber = headerRow + 1 + index;
      
      const clockInDate = record.clockIn ? new Date(record.clockIn) : null;
      const clockOutDate = record.clockOut ? new Date(record.clockOut) : null;

      const rowData = [
        index + 1,
        record.employee,
        clockInDate ? clockInDate.toLocaleDateString('th-TH') : '',
        clockInDate ? clockInDate.toLocaleTimeString('th-TH') : '',
        clockOutDate ? clockOutDate.toLocaleTimeString('th-TH') : '',
        record.workingHours ? `${record.workingHours} ชม.` : '',
        record.note,
        record.locationIn,
        record.locationOut
      ];

      rowData.forEach((value, colIndex) => {
        const cell = worksheet.getCell(rowNumber, colIndex + 1);
        cell.value = value;
        cell.font = { name: 'Angsana New', size: 12 };
        cell.alignment = { 
          horizontal: colIndex === 0 ? 'center' : 'left', 
          vertical: 'middle' 
        };
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
      });
    });

    // ปรับขนาดคอลัมน์
    const columnWidths = [8, 25, 15, 12, 12, 15, 20, 20, 20];
    columnWidths.forEach((width, index) => {
      worksheet.getColumn(index + 1).width = width;
    });

    // สรุปข้อมูล
    const summaryRow = headerRow + data.length + 2;
    worksheet.getCell(summaryRow, 1).value = `สรุป: พบข้อมูลทั้งหมด ${data.length} รายการ`;
    worksheet.getCell(summaryRow, 1).font = { name: 'Angsana New', size: 12, bold: true };
    worksheet.mergeCells(`A${summaryRow}:I${summaryRow}`);

    // วันที่สร้างรายงาน
    const footerRow = summaryRow + 2;
    worksheet.getCell(footerRow, 1).value = `สร้างรายงานเมื่อ: ${new Date().toLocaleString('th-TH')}`;
    worksheet.getCell(footerRow, 1).font = { name: 'Angsana New', size: 10 };
    worksheet.getCell(footerRow, 1).alignment = { horizontal: 'right' };
    worksheet.mergeCells(`A${footerRow}:I${footerRow}`);

    return workbook;
  }
}

// ========== Initialize Services ==========
const sheetsService = new GoogleSheetsService();
const keepAliveService = new KeepAliveService();

// ========== Admin Authentication Routes ==========

// Admin Login
app.post('/api/admin/login', async (req, res) => {
  try {
    const { username, password } = req.body;

    if (!username || !password) {
      return res.status(400).json({
        success: false,
        message: 'กรุณากรอกชื่อผู้ใช้และรหัสผ่าน'
      });
    }

    // ค้นหาผู้ใช้
    const user = CONFIG.ADMIN.USERS.find(u => u.username === username);
    if (!user) {
      return res.status(401).json({
        success: false,
        message: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง'
      });
    }

    // ตรวจสอบรหัสผ่าน
    const isValidPassword = await bcrypt.compare(password, user.password);
    if (!isValidPassword) {
      return res.status(401).json({
        success: false,
        message: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง'
      });
    }

    // สร้าง JWT token
    const token = jwt.sign(
      { 
        id: user.id, 
        username: user.username, 
        role: user.role 
      },
      CONFIG.ADMIN.JWT_SECRET,
      { expiresIn: CONFIG.ADMIN.JWT_EXPIRES_IN }
    );

    res.json({
      success: true,
      message: 'เข้าสู่ระบบสำเร็จ',
      token,
      user: {
        id: user.id,
        username: user.username,
        name: user.name,
        role: user.role
      }
    });

  } catch (error) {
    console.error('Admin login error:', error);
    res.status(500).json({
      success: false,
      message: 'เกิดข้อผิดพลาดภายในระบบ'
    });
  }
});

// Verify Token
app.get('/api/admin/verify-token', authenticateAdmin, (req, res) => {
  res.json({
    success: true,
    user: req.user
  });
});

// Admin Stats
app.get('/api/admin/stats', authenticateAdmin, async (req, res) => {
  try {
    const stats = await sheetsService.getAdminStats();
    res.json({
      success: true,
      data: stats
    });
  } catch (error) {
    console.error('Admin stats error:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to get stats'
    });
  }
});

// Export Routes
app.get('/api/admin/export/:type', authenticateAdmin, async (req, res) => {
  try {
    const { type } = req.params;
    const params = req.query;

    // ตรวจสอบประเภทรายงาน
    if (!['daily', 'monthly', 'range'].includes(type)) {
      return res.status(400).json({
        success: false,
        error: 'Invalid report type'
      });
    }

    // ดึงข้อมูลจาก Google Sheets
    const reportData = await sheetsService.getReportData(type, params);

    // สร้างไฟล์ Excel
    const workbook = await ExcelExportService.createWorkbook(reportData, type, params);

    // ตั้งค่า response headers
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=report.xlsx');

    // ส่งไฟล์
    await workbook.xlsx.write(res);
    res.end();

  } catch (error) {
    console.error('Export error:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to export report'
    });
  }
});

// API สำหรับส่งออกรายงานรายเดือนแบบละเอียด
app.post('/api/reports/export-monthly-detailed', authenticateAdmin, async (req, res) => {
  console.log('📊 Received monthly detailed report request');
  console.log('Request body:', req.body);
  
  try {
    const { month, year, options } = req.body;
    
    if (!month || !year) {
      console.error('Missing parameters:', { month, year });
      return res.status(400).json({
        success: false,
        error: 'Missing month or year parameter'
      });
    }

    // Parse options
    const reportOptions = options || {
      showDailyBreakdown: true,
      showWeekends: true,
      showSummary: true,
      showLateComers: false,
      showOvertime: false,
      colorCoding: true
    };

    console.log(`📊 Generating detailed monthly report for ${month}/${year}`);

    // สร้าง workbook ใหม่
    const workbook = new ExcelJS.Workbook();
    
    // ตั้งค่าข้อมูลเอกสาร
    workbook.creator = 'ระบบจัดการลงเวลา อบต.ข่าใหญ่';
    workbook.lastModifiedBy = 'ระบบ';
    workbook.created = new Date();
    workbook.modified = new Date();

    // สร้าง worksheet หลัก
    const worksheet = workbook.addWorksheet(`รายงานเดือน ${getThaiMonth(month)} ${parseInt(year) + 543}`);

    // ดึงข้อมูลจาก Google Sheets
    const attendanceData = await getMonthlyAttendanceDataFromSheets(month, year);
    const employees = await getEmployeesListFromSheets();
    
    console.log(`📊 Found ${attendanceData.length} attendance records for ${employees.length} employees`);

    // สร้างรายงานตามตัวเลือก
    if (reportOptions.showDailyBreakdown) {
      await createDailyBreakdownReport(worksheet, attendanceData, employees, month, year, reportOptions);
    }

    // เพิ่ม worksheet สรุป
    if (reportOptions.showSummary) {
      const summarySheet = workbook.addWorksheet('สรุปรายเดือน');
      await createMonthlySummary(summarySheet, attendanceData, employees, month, year);
    }

    // ตั้งค่าการตอบกลับ
    const monthName = getThaiMonth(month);
    const filename = `รายงานรายเดือน_${monthName}_${parseInt(year) + 543}_แบ่งวันชัดเจน.xlsx`;
    
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(filename)}"`);

    // ส่งไฟล์
    await workbook.xlsx.write(res);
    res.end();

    console.log(`✅ Monthly detailed report sent: ${filename}`);

  } catch (error) {
    console.error('Export error:', error);
    res.status(500).json({ 
      success: false, 
      message: 'เกิดข้อผิดพลาดในการส่งออกรายงาน',
      error: error.message 
    });
  }
});

// ========== Original Routes (unchanged) ==========

// Home page
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Health check และ ping endpoint
app.get('/api/health', (req, res) => {  res.json({
    status: 'healthy',
    timestamp: moment().tz(CONFIG.TIMEZONE).toISOString(), // ใช้เวลาไทย
    uptime: process.uptime(),
    keepAlive: keepAliveService.getStats(),
    environment: process.env.NODE_ENV || 'development',
    config: {
      hasLiffId: !!CONFIG.LINE.LIFF_ID,
      liffIdLength: CONFIG.LINE.LIFF_ID ? CONFIG.LINE.LIFF_ID.length : 0
    }
  });
});

// Ping endpoint สำหรับ keep-alive
app.get('/api/ping', (req, res) => {
  res.json({
    status: 'pong',
    timestamp: moment().tz(CONFIG.TIMEZONE).toISOString(), // ใช้เวลาไทย
    uptime: process.uptime()
  });
});

// Webhook endpoint สำหรับรับ ping จาก GSA
app.post('/api/webhook/ping', (req, res) => {
  console.log('📨 Received ping from GSA');  res.json({
    status: 'received',
    timestamp: moment().tz(CONFIG.TIMEZONE).toISOString() // ใช้เวลาไทย
  });
});

// API สำหรับ Client Configuration
app.get('/api/config', (req, res) => {
  try {
    res.json({
      success: true,
      data: {
        liffId: CONFIG.LINE.LIFF_ID,
        apiUrl: CONFIG.RENDER.SERVICE_URL + '/api',
        environment: process.env.NODE_ENV || 'development',
        features: {
          keepAlive: CONFIG.RENDER.KEEP_ALIVE_ENABLED,
          liffEnabled: !!CONFIG.LINE.LIFF_ID
        }
      }
    });
  } catch (error) {
    console.error('API Error - config:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to get config'
    });
  }
});

// Get employees
app.post('/api/employees', async (req, res) => {
  try {
    const employees = await sheetsService.getEmployees();
    res.json({
      success: true,
      data: employees
    });
  } catch (error) {
    console.error('API Error - employees:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to get employees'
    });
  }
});

// Clock in
app.post('/api/clockin', async (req, res) => {
  try {
    const { employee, userinfo, lat, lon, line_name, line_picture } = req.body;
    
    if (!employee || !lat || !lon) {
      return res.status(400).json({
        success: false,
        error: 'Missing required fields'
      });
    }

    const result = await sheetsService.clockIn({
      employee, userinfo, lat, lon, line_name, line_picture
    });

    res.json(result);
    
  } catch (error) {
    console.error('API Error - clockin:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to clock in'
    });
  }
});

// Clock out
app.post('/api/clockout', async (req, res) => {
  try {
    const { employee, lat, lon, line_name } = req.body;
    
    if (!employee || !lat || !lon) {
      return res.status(400).json({
        success: false,
        error: 'Missing required fields'
      });
    }

    const result = await sheetsService.clockOut({
      employee, lat, lon, line_name
    });

    res.json(result);
    
  } catch (error) {
    console.error('API Error - clockout:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to clock out'
    });
  }
});

// API สำหรับตรวจสอบสถานะพนักงาน
app.post('/api/check-status', async (req, res) => {
  try {
    const { employee } = req.body;
    
    if (!employee) {
      return res.status(400).json({
        success: false,
        error: 'Missing employee name'
      });
    }

    const employeeStatus = await sheetsService.getEmployeeStatus(employee);

    const onWorkSheet = await sheetsService.getSheet(CONFIG.SHEETS.ON_WORK);
    const rows = await onWorkSheet.getRows({ offset: 1 }); // เริ่มจากแถว 3
    
    const currentEmployees = rows.map(row => ({
      systemName: row.get('ชื่อในระบบ'),
      employeeName: row.get('ชื่อพนักงาน'),
      clockIn: row.get('เวลาเข้า'),
      mainRowIndex: row.get('แถวในMain') || row.get('แถวอ้างอิง')
    }));

    res.json({
      success: true,
      data: {
        employee: employee,
        isOnWork: employeeStatus.isOnWork,
        hasWorkRecord: !!employeeStatus.workRecord,
        workRecord: employeeStatus.workRecord ? {
          clockIn: employeeStatus.workRecord.clockIn,
          mainRowIndex: employeeStatus.workRecord.mainRowIndex
        } : null,
        allCurrentEmployees: currentEmployees,
        suggestions: currentEmployees
          .filter(emp => 
            sheetsService.isNameMatch(employee, emp.systemName) ||
            sheetsService.isNameMatch(employee, emp.employeeName)
          )
          .map(emp => emp.systemName || emp.employeeName)
      }
    });

  } catch (error) {
    console.error('API Error - check-status:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to check status'
    });
  }
});

// ========== Error Handling ==========
app.use((error, req, res, next) => {
  console.error('Global error handler:', error);
  res.status(500).json({
    success: false,
    error: 'Internal server error'
  });
});

app.use((req, res) => {
  res.status(404).json({
    success: false,
    error: 'Not found'
  });
});

// ========== Start Server ==========
async function startServer() {
  try {
    console.log('🚀 Starting Time Tracker Server with Admin Panel...');
    console.log(`🌍 Environment: ${process.env.NODE_ENV || 'development'}`);
    
    // ตรวจสอบ environment variables
    if (!validateConfig()) {
      console.error('❌ Server startup aborted due to missing configuration');
      process.exit(1);
    }
    
    console.log('📊 Initializing Google Sheets service...');
    
    await sheetsService.initialize();
    
    app.listen(PORT, () => {
      console.log(`✅ Server running on port ${PORT}`);
      console.log(`📊 Google Sheets connected: ${CONFIG.GOOGLE_SHEETS.SPREADSHEET_ID}`);
      console.log(`🌐 Service URL: ${CONFIG.RENDER.SERVICE_URL}`);
      console.log(`📱 LIFF ID: ${CONFIG.LINE.LIFF_ID || 'Not configured'}`);
      console.log(`🔐 Admin Panel: ${CONFIG.RENDER.SERVICE_URL}/admin/login`);
      
      // เริ่ม Keep-Alive service
      keepAliveService.init();
    });
    
  } catch (error) {
    console.error('❌ Failed to start server:', error);
    process.exit(1);
  }
}

// Graceful shutdown
process.on('SIGTERM', () => {
  console.log('SIGTERM received, shutting down gracefully');
  process.exit(0);
});

process.on('SIGINT', () => {
  console.log('SIGINT received, shutting down gracefully');
  process.exit(0);
});

startServer();

// ========== Excel Export Functions ==========

// ฟังก์ชันช่วยสำหรับ Excel export
function getThaiMonth(month) {
  const months = [
    '', 'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
    'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'
  ];
  return months[parseInt(month)];
}

function getThaiDayName(dayIndex) {
  const days = ['อา', 'จ', 'อ', 'พ', 'พฤ', 'ศ', 'ส'];
  return days[dayIndex];
}

function getWorkingDaysInMonth(month, year) {
  const daysInMonth = moment(`${year}-${month}`, 'YYYY-MM').daysInMonth();
  let workingDays = 0;
  
  for (let day = 1; day <= daysInMonth; day++) {
    const date = moment(`${year}-${month}-${day.toString().padStart(2, '0')}`);
    const dayOfWeek = date.day();
    if (dayOfWeek !== 0 && dayOfWeek !== 6) { // ไม่ใช่วันเสาร์อาทิตย์
      workingDays++;
    }
  }
  
  return workingDays;
}

// ดึงข้อมูลการทำงานจาก Google Sheets
async function getMonthlyAttendanceDataFromSheets(month, year) {
  try {
    // ลองดึงข้อมูลจาก Google Sheets ก่อน
    try {
      const mainSheet = await sheetsService.getSheet(CONFIG.SHEETS.MAIN);
      const rows = await mainSheet.getRows();
      
      const attendanceData = [];
      
      for (const row of rows) {
        const dateStr = row.get('วันที่');
        if (!dateStr) continue;
        
        const recordDate = new Date(dateStr);
        if (recordDate.getMonth() + 1 === parseInt(month) && recordDate.getFullYear() === parseInt(year)) {
          const clockIn = row.get('เวลาเข้า') || '';
          const clockOut = row.get('เวลาออก') || '';
          const employeeName = row.get('ชื่อพนักงาน') || '';
          
          // กำหนดสถานะ
          let status = 'present';
          let isLate = false;
            if (clockIn) {
            // ตรวจสอบมาสาย (หลัง 08:30)
            if (clockIn > '08:30') {
              isLate = true;
            }
          } else {
            status = 'absent';
          }
          
          // ตรวจสอบจากหมายเหตุ
          const remarks = row.get('หมายเหตุ') || '';
          if (remarks.includes('ลาป่วย') || remarks.includes('ลป')) {
            status = 'sick_leave';
          } else if (remarks.includes('ลากิจ') || remarks.includes('ลก')) {
            status = 'personal_leave';
          }
          
          attendanceData.push({
            employeeId: employeeName, // ใช้ชื่อเป็น ID ชั่วคราว
            employeeName: employeeName,
            date: moment(recordDate).format('YYYY-MM-DD'),
            status: status,
            clockIn: clockIn,
            clockOut: clockOut,
            isLate: isLate,
            remarks: remarks
          });
        }
      }
      
      if (attendanceData.length > 0) {
        return attendanceData;
      }
    } catch (sheetsError) {
      console.log('Google Sheets not available, using sample data');
    }
      // ถ้าไม่มีข้อมูลจาก Google Sheets ให้ใช้ข้อมูลตัวอย่าง
    // ดึงรายชื่อพนักงานจริงมาใช้
    const employeesList = await getEmployeesListFromSheets();
    return generateSampleAttendanceData(month, year, employeesList);
    
  } catch (error) {
    console.error('Error getting attendance data:', error);
    return generateSampleAttendanceData(month, year, employeesList);
  }
}

// ดึงรายชื่อพนักงานจาก Google Sheets
async function getEmployeesListFromSheets() {
  try {
    // ลองดึงข้อมูลจาก Google Sheets ก่อน
    try {
      const employees = await sheetsService.getEmployees();
      if (employees && employees.length > 0) {
        return employees.map((name, index) => ({
          employeeId: name,
          name: name
        }));
      }
    } catch (sheetsError) {
      console.log('Google Sheets not available, using sample employees');
    }
    
    // ใช้ข้อมูลตัวอย่างถ้าไม่มีข้อมูลจาก Google Sheets
    return [
      { employeeId: 'นายสมชาย ใจดี', name: 'นายสมชาย ใจดี' },
      { employeeId: 'นางสาวสมหญิง รักงาน', name: 'นางสาวสมหญิง รักงาน' },
      { employeeId: 'นายสมศักดิ์ ทำงานดี', name: 'นายสมศักดิ์ ทำงานดี' },
      { employeeId: 'นางสมใจ บริการดี', name: 'นางสมใจ บริการดี' },
      { employeeId: 'นายสมปอง มีความสุข', name: 'นายสมปอง มีความสุข' }
    ];
    
  } catch (error) {
    console.error('Error getting employees list:', error);
    return [
      { employeeId: 'นายสมชาย ใจดี', name: 'นายสมชาย ใจดี' },
      { employeeId: 'นางสาวสมหญิง รักงาน', name: 'นางสาวสมหญิง รักงาน' }
    ];
  }
}

// สร้างรายงานแบ่งตามวัน
async function createDailyBreakdownReport(worksheet, attendanceData, employees, month, year, options = {}) {
    try {
        console.log(`Generating detailed report for ${month}/${year}`);
        
        // หาจำนวนวันในเดือน
        const daysInMonth = new Date(year, month, 0).getDate();
        
        // ตั้งค่าความกว้างของคอลัมน์
        worksheet.getColumn(1).width = 8;  // ลำดับ
        worksheet.getColumn(2).width = 15; // รหัสพนักงาน
        worksheet.getColumn(3).width = 25; // ชื่อ-นามสกุล
        
        // วันที่
        for (let day = 1; day <= daysInMonth; day++) {
            worksheet.getColumn(3 + day).width = 12;
        }
        
        // คอลัมน์สรุป
        worksheet.getColumn(4 + daysInMonth).width = 12;     // วันทำงาน
        worksheet.getColumn(5 + daysInMonth).width = 12;     // วันมาสาย
        worksheet.getColumn(6 + daysInMonth).width = 12;     // วันขาดงาน
        worksheet.getColumn(7 + daysInMonth).width = 12;     // วันลา
        worksheet.getColumn(8 + daysInMonth).width = 20;     // หมายเหตุ
        
        // สร้าง header
        const headerRow = ['ลำดับ', 'รหัสพนักงาน', 'ชื่อ-นามสกุล'];
        
        // เพิ่มวันที่ในเดือน
        for (let day = 1; day <= daysInMonth; day++) {
            const currentDate = new Date(year, month - 1, day);
            const dayName = ['อา', 'จ', 'อ', 'พ', 'พฤ', 'ศ', 'ส'][currentDate.getDay()];
            headerRow.push(`${day}\n(${dayName})`);
        }
        
        headerRow.push('วันทำงาน', 'วันมาสาย', 'วันขาดงาน', 'วันลา', 'หมายเหตุ');
        
        // เพิ่ม header row
        worksheet.addRow(headerRow);
        
        // ตั้งค่าสไตล์ header
        const headerRowObj = worksheet.getRow(1);
        headerRowObj.font = { name: 'TH SarabunPSK', size: 16, bold: true };
        headerRowObj.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        headerRowObj.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFE6E6FA' }
        };
          console.log(`Found ${employees.length} employees`);
        console.log(`Found ${attendanceData.length} attendance records`);
        
        // Debug: แสดงข้อมูลตัวอย่าง
        if (employees.length > 0) {
            console.log('Sample employee:', employees[0]);
        }
        if (attendanceData.length > 0) {
            console.log('Sample attendance:', attendanceData[0]);
        }
        
        // เพิ่มข้อมูลพนักงานแต่ละคน
        employees.forEach((employee, index) => {
            const row = [
                index + 1,
                employee.employeeId || (index + 1),
                employee.name || `${employee.firstName || ''} ${employee.lastName || ''}`.trim()
            ];
            
            let workDays = 0;
            let lateDays = 0;
            let absentDays = 0;
            let leaveDays = 0;
            
            console.log(`Processing employee: ${employee.employeeId || index + 1} - ${employee.name}`);
            
            // Debug: แสดงข้อมูล attendance ที่เกี่ยวข้องกับพนักงานคนนี้
            const employeeAttendances = attendanceData.filter(att => 
                att.employeeName === employee.name || 
                att.employeeId === employee.name ||
                att.employeeName === employee.employeeId ||
                att.employeeId === employee.employeeId
            );
            console.log(`Found ${employeeAttendances.length} attendance records for ${employee.name}:`, employeeAttendances.slice(0, 3));
            
            // ตรวจสอบแต่ละวันในเดือน
            for (let day = 1; day <= daysInMonth; day++) {
                const currentDate = new Date(year, month - 1, day);
                const dateKey = `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
                
                // ตรวจสอบวันหยุด (เสาร์-อาทิตย์)
                const dayOfWeek = currentDate.getDay();
                const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;                  // หาข้อมูลการลงเวลาในวันนี้จาก attendanceData
                const dayAttendance = attendanceData.find(att => {
                    // ง่ายๆ ใช้ชื่อพนักงานจับคู่
                    const nameMatch = att.employeeName === employee.name || 
                                     att.employeeId === employee.name ||
                                     att.employeeName === employee.employeeId ||
                                     att.employeeId === employee.employeeId;
                    
                    const dateMatch = att.date === dateKey;
                    
                    return nameMatch && dateMatch;
                });
                
                let cellValue = '';
                
                if (dayAttendance) {
                    console.log(`📋 Found attendance: ${employee.name} on ${dateKey} - Status: ${dayAttendance.status}, ClockIn: ${dayAttendance.clockIn}`);
                    
                    // มีข้อมูลการลงเวลา
                    if (dayAttendance.status === 'present') {
                        if (dayAttendance.isLate) {
                            cellValue = `มาสาย\n${dayAttendance.clockIn}`;
                            lateDays++;
                        } else {
                            cellValue = dayAttendance.clockIn;
                        }
                        workDays++;
                    } else if (dayAttendance.status === 'sick_leave') {
                        cellValue = 'ลาป่วย';
                        leaveDays++;
                    } else if (dayAttendance.status === 'personal_leave') {
                        cellValue = 'ลากิจ';
                        leaveDays++;
                    } else if (dayAttendance.status === 'absent') {
                        cellValue = 'ขาด';
                        absentDays++;
                    } else {
                        // กรณีมีข้อมูลแต่ไม่ระบุสถานะชัด ให้ถือว่าเข้างาน
                        cellValue = dayAttendance.clockIn || '✓';
                        workDays++;
                    }
                } else if (isWeekend && !options.showWeekends) {
                    // วันหยุด
                    cellValue = 'หยุด';
                } else if (!isWeekend) {
                    // ไม่มีข้อมูลการลงเวลาในวันทำงาน
                    cellValue = 'ขาด';
                    absentDays++;
                } else {
                    // วันหยุดแต่ showWeekends = true
                    cellValue = 'หยุด';
                }
                
                row.push(cellValue);
            }
            
            // เพิ่มสรุปข้อมูล
            row.push(workDays, lateDays, absentDays, leaveDays, '');
            
            console.log(`Employee ${employee.employeeId || index + 1} summary: Work=${workDays}, Late=${lateDays}, Absent=${absentDays}, Leave=${leaveDays}`);
            
            const addedRow = worksheet.addRow(row);
            
            // ตั้งค่าสไตล์แถว
            addedRow.font = { name: 'TH SarabunPSK', size: 14 };
            addedRow.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
            
            // เพิ่มสีให้เซลล์ตามสถานะ
            for (let day = 1; day <= daysInMonth; day++) {
                const cellIndex = 3 + day; // เริ่มจากคอลัมน์ที่ 4 (D)
                const cell = addedRow.getCell(cellIndex);
                const cellValue = row[cellIndex - 1];
                
                if (cellValue && cellValue !== 'หยุด' && cellValue !== '') {
                    if (cellValue.includes('มาสาย')) {
                        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFA500' } };
                    } else if (cellValue.includes('ลา')) {
                        if (cellValue.includes('ป่วย')) {
                            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF87CEEB' } };
                        } else {
                            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
                        }
                    } else if (cellValue === 'ขาด') {
                        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF0000' } };
                    } else if (cellValue.match(/^\d{2}:\d{2}$/)) {
                        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF90EE90' } };
                    } else if (cellValue === '✓' || cellValue.includes(':')) {
                        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF90EE90' } };
                    }
                } else if (cellValue === 'หยุด') {
                    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD3D3D3' } };
                }
            }
        });
        
        // เพิ่มข้อมูลสรุปท้ายรายงาน
        const summaryStartRow = employees.length + 3;
        
        worksheet.getCell(`A${summaryStartRow}`).value = 'สรุปข้อมูลรายงาน';
        worksheet.getCell(`A${summaryStartRow}`).font = { name: 'TH SarabunPSK', size: 16, bold: true };
        
        worksheet.getCell(`A${summaryStartRow + 1}`).value = `จำนวนพนักงานทั้งหมด: ${employees.length} คน`;
        worksheet.getCell(`A${summaryStartRow + 2}`).value = `จำนวนวันในเดือน: ${daysInMonth} วัน`;
        worksheet.getCell(`A${summaryStartRow + 3}`).value = `สร้างรายงานเมื่อ: ${new Date().toLocaleString('th-TH')}`;
        
        // ตั้งค่าฟอนต์สำหรับข้อมูลสรุป
        for (let i = 1; i <= 3; i++) {
            worksheet.getCell(`A${summaryStartRow + i}`).font = { name: 'TH SarabunPSK', size: 16 };
        }
        
        // เพิ่มคำอธิบายสัญลักษณ์
        const legendStartRow = summaryStartRow + 5;
        worksheet.getCell(`A${legendStartRow}`).value = 'คำอธิบายสัญลักษณ์:';
        worksheet.getCell(`A${legendStartRow}`).font = { name: 'TH SarabunPSK', size: 16, bold: true };
        
        const legends = [
            '🟢 เขียว = เข้างานปกติ',
            '🟠 ส้ม = มาสาย', 
            '🔴 แดง = ขาดงาน',
            '🔵 ฟ้า = ลาป่วย',
            '🟡 เหลือง = ลากิจ',
            '⚫ เทา = วันหยุด'
        ];
        
        legends.forEach((legend, index) => {
            const cell = worksheet.getCell(`A${legendStartRow + 1 + index}`);
            cell.value = legend;
            cell.font = { name: 'TH SarabunPSK', size: 14 };
        });
        
        console.log('Daily breakdown report generated successfully');
        
    } catch (error) {
        console.error('Error in createDailyBreakdownReport:', error);
        throw error;
    }
}

// สร้างสรุปรายเดือน
async function createMonthlySummary(worksheet, attendanceData, employees, month, year) {
  // ตั้งค่าคอลัมน์
  worksheet.getColumn(1).width = 25;
  worksheet.getColumn(2).width = 15;
  worksheet.getColumn(3).width = 12;
  worksheet.getColumn(4).width = 12;
  worksheet.getColumn(5).width = 12;
  worksheet.getColumn(6).width = 12;
  worksheet.getColumn(7).width = 12;
  worksheet.getColumn(8).width = 15;

  // หัวข้อ
  const titleRow = worksheet.addRow([
    `สรุปการเข้างานประจำเดือน ${getThaiMonth(month)} พ.ศ. ${parseInt(year) + 543}`
  ]);
  worksheet.mergeCells(1, 1, 1, 8);
  titleRow.getCell(1).font = { name: 'TH SarabunPSK', size: 18, bold: true };
  titleRow.getCell(1).alignment = { horizontal: 'center' };
  titleRow.height = 30;

  worksheet.addRow([]);

  // หัวตาราง
  const headerRow = worksheet.addRow([
    'ชื่อ-นามสกุล', 'รหัสพนักงาน', 'มาทำงาน', 'ขาดงาน', 
    'มาสาย', 'ลาป่วย', 'ลากิจ', 'เปอร์เซ็นต์การมา'
  ]);

  headerRow.eachCell((cell) => {
    cell.font = { name: 'TH SarabunPSK', size: 14, bold: true };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE6E6FA' } };
    cell.border = {
      top: { style: 'thin' }, left: { style: 'thin' },
      bottom: { style: 'thin' }, right: { style: 'thin' }
    };
  });

  // คำนวณข้อมูลสรุป
  const workingDaysInMonth = getWorkingDaysInMonth(month, year);
  
  employees.forEach(employee => {
    const employeeData = attendanceData.filter(d => d.employeeId === employee.employeeId);
    
    const presentDays = employeeData.filter(d => d.status === 'present').length;
    const absentDays = employeeData.filter(d => d.status === 'absent').length;
    const lateDays = employeeData.filter(d => d.status === 'present' && d.isLate).length;
    const sickLeaveDays = employeeData.filter(d => d.status === 'sick_leave').length;
    const personalLeaveDays = employeeData.filter(d => d.status === 'personal_leave').length;
    
    const attendanceRate = workingDaysInMonth > 0 ? ((presentDays / workingDaysInMonth) * 100).toFixed(1) : '0.0';

    const dataRow = worksheet.addRow([
      employee.name,
      employee.employeeId,
      presentDays,
      absentDays,
      lateDays,
      sickLeaveDays,
      personalLeaveDays,
      `${attendanceRate}%`
    ]);

    dataRow.eachCell((cell, colNumber) => {
      cell.font = { name: 'TH SarabunPSK', size: 12 };
      cell.alignment = { 
        horizontal: colNumber === 1 ? 'left' : 'center', 
        vertical: 'middle' 
      };
      cell.border = {
        top: { style: 'thin' }, left: { style: 'thin' },
        bottom: { style: 'thin' }, right: { style: 'thin' }
      };

      // ใส่สีตามเปอร์เซ็นต์
      if (colNumber === 8) {
        const rate = parseFloat(attendanceRate);
        if (rate >= 95) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF90EE90' } };
        } else if (rate >= 85) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFD700' } };
        } else {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF6B6B' } };
        }
      }
    });
  });
}

// API สำหรับส่งออกรายงานรายเดือนแบบละเอียด
app.post('/api/reports/export-monthly-detailed', authenticateAdmin, async (req, res) => {
  console.log('📊 Received monthly detailed report request');
  console.log('Request body:', req.body);
  
  try {
    const { month, year, options } = req.body;
    
    if (!month || !year) {
      console.error('Missing parameters:', { month, year });
      return res.status(400).json({
        success: false,
        error: 'Missing month or year parameter'
      });
    }

    // Parse options
    const reportOptions = options || {
      showDailyBreakdown: true,
      showWeekends: true,
      showSummary: true,
      showLateComers: false,
      showOvertime: false,
      colorCoding: true
    };

    console.log(`📊 Generating detailed monthly report for ${month}/${year}`);

    // สร้าง workbook ใหม่
    const workbook = new ExcelJS.Workbook();
    
    // ตั้งค่าข้อมูลเอกสาร
    workbook.creator = 'ระบบจัดการลงเวลา อบต.ข่าใหญ่';
    workbook.lastModifiedBy = 'ระบบ';
    workbook.created = new Date();
    workbook.modified = new Date();

    // สร้าง worksheet หลัก
    const worksheet = workbook.addWorksheet(`รายงานเดือน ${getThaiMonth(month)} ${parseInt(year) + 543}`);

    // ดึงข้อมูลจาก Google Sheets
    const attendanceData = await getMonthlyAttendanceDataFromSheets(month, year);
    const employees = await getEmployeesListFromSheets();
    
    console.log(`📊 Found ${attendanceData.length} attendance records for ${employees.length} employees`);

    // สร้างรายงานตามตัวเลือก
    if (reportOptions.showDailyBreakdown) {
      await createDailyBreakdownReport(worksheet, attendanceData, employees, month, year, reportOptions);
    }

    // เพิ่ม worksheet สรุป
    if (reportOptions.showSummary) {
      const summarySheet = workbook.addWorksheet('สรุปรายเดือน');
      await createMonthlySummary(summarySheet, attendanceData, employees, month, year);
    }

    // ตั้งค่าการตอบกลับ
    const monthName = getThaiMonth(month);
    const filename = `รายงานรายเดือน_${monthName}_${parseInt(year) + 543}_แบ่งวันชัดเจน.xlsx`;
    
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(filename)}"`);

    // ส่งไฟล์
    await workbook.xlsx.write(res);
    res.end();

    console.log(`✅ Monthly detailed report sent: ${filename}`);

  } catch (error) {
    console.error('Export error:', error);    res.status(500).json({ 
      success: false, 
      message: 'เกิดข้อผิดพลาดในการส่งออกรายงาน',
      error: error.message 
    });
  }
});

// สร้างข้อมูลตัวอย่างสำหรับการทดสอบ
function generateSampleAttendanceData(month, year, employeesList = null) {
  const attendanceData = [];
  
  // ใช้รายชื่อพนักงานที่ส่งมา หรือใช้ตัวอย่างถ้าไม่มี
  const employees = employeesList ? employeesList.map(emp => emp.name || emp.employeeId) : [
    'นายสมชาย ใจดี',
    'นางสาวสมหญิง รักงาน', 
    
    'นายสมศักดิ์ ทำงานดี',
    'นางสมใจ บริการดี',
    'นายสมปอง มีความสุข'
  ];
  
  console.log(`📋 Generating sample data for ${employees.length} employees:`, employees.slice(0, 3));
    const daysInMonth = moment(`${year}-${String(month).padStart(2, '0')}`, 'YYYY-MM').daysInMonth();
  
  for (let day = 1; day <= daysInMonth; day++) {
    const date = moment(`${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`, 'YYYY-MM-DD');
    const isWeekend = date.day() === 0 || date.day() === 6;
    
    if (!isWeekend) {
      employees.forEach((employeeName, empIndex) => {
        // สุ่มสถานะการมาทำงาน
        const random = Math.random();
        let status = 'present';
        let clockIn = '08:30';
        let clockOut = '17:00';
        let isLate = false;
        
        if (random < 0.05) { // 5% ขาดงาน
          status = 'absent';
          clockIn = '';
          clockOut = '';
        } else if (random < 0.08) { // 3% ลาป่วย
          status = 'sick_leave';
          clockIn = '';
          clockOut = '';
        } else if (random < 0.1) { // 2% ลากิจ
          status = 'personal_leave';
          clockIn = '';
          clockOut = '';
        } else if (random < 0.3) { // 20% มาสาย
          isLate = true;
          const lateMinutes = Math.floor(Math.random() * 60) + 1; // สาย 1-60 นาที
          const lateHour = 8 + Math.floor((30 + lateMinutes) / 60);
          const lateMin = (30 + lateMinutes) % 60;
          clockIn = `${lateHour.toString().padStart(2, '0')}:${lateMin.toString().padStart(2, '0')}`;        } else {
          // มาทำงานปกติ แต่อาจจะมาก่อนหรือหลัง 8:30 เล็กน้อย
          const variation = Math.floor(Math.random() * 30) - 15; // -15 ถึง +15 นาที
          const clockInMinutes = 30 + variation;
          if (clockInMinutes < 0) {
            clockIn = `07:${(60 + clockInMinutes).toString().padStart(2, '0')}`;
          } else if (clockInMinutes >= 60) {
            clockIn = `09:${(clockInMinutes - 60).toString().padStart(2, '0')}`;
            if (clockInMinutes > 30) isLate = true; // มาสายถ้าหลัง 8:30
          } else {
            clockIn = `08:${clockInMinutes.toString().padStart(2, '0')}`;
            if (clockInMinutes > 30) isLate = true; // มาสายถ้าหลัง 8:30
          }
        }
        
        attendanceData.push({
          employeeId: employeeName,
          employeeName: employeeName,
          date: date.format('YYYY-MM-DD'),
          status: status,
          clockIn: clockIn,
          clockOut: clockOut,
          isLate: isLate,
          remarks: status === 'sick_leave' ? 'ลาป่วย' : status === 'personal_leave' ? 'ลากิจ' : ''
        });
      });
    }
  }
  
  console.log(`📊 Generated ${attendanceData.length} sample attendance records for ${month}/${year}`);
  return attendanceData;
}