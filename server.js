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
    // เพิ่มระบบ caching เพื่อลดการเรียก API
    this.cache = {
      employees: { data: null, timestamp: null, ttl: 300000 }, // 5 นาที
      onWork: { data: null, timestamp: null, ttl: 60000 }, // 1 นาที  
      main: { data: null, timestamp: null, ttl: 30000 }, // 30 วินาที
      stats: { data: null, timestamp: null, ttl: 120000 } // 2 นาที
    };
    this.emergencyMode = false; // เริ่มต้นปิดระบบ emergency mode
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
  }  // เพิ่มฟังก์ชัน cache helper
  isCacheValid(cacheKey) {
    const cache = this.cache[cacheKey];
    if (!cache || !cache.data || !cache.timestamp) return false;
    return (Date.now() - cache.timestamp) < cache.ttl;
  }

  setCache(cacheKey, data) {
    if (!this.cache[cacheKey]) {
      this.cache[cacheKey] = { data: null, timestamp: null, ttl: 300000 }; // default 5 min
    }
    this.cache[cacheKey] = {
      data: data,
      timestamp: Date.now(),
      ttl: this.cache[cacheKey].ttl
    };
  }

  getCache(cacheKey) {
    const cache = this.cache[cacheKey];
    return cache && cache.data ? cache.data : null;
  }

  clearCache(cacheKey = null) {
    if (cacheKey) {
      this.cache[cacheKey].data = null;
      this.cache[cacheKey].timestamp = null;
    } else {
      // Clear all cache
      Object.keys(this.cache).forEach(key => {
        this.cache[key].data = null;
        this.cache[key].timestamp = null;
      });
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
  }  // เพิ่มฟังก์ชันดึงข้อมูลพร้อม cache และ rate limiting
  async getCachedSheetData(sheetName) {
    const cacheKey = sheetName.toLowerCase().replace(/\s+/g, '');
    
    // ตรวจสอบ cache ก่อน
    if (this.isCacheValid(cacheKey)) {
      console.log(`📋 Using cached data for ${sheetName}`);
      return this.getCache(cacheKey);
    }

    // ตรวจสอบ rate limit ก่อนเรียก API
    if (!apiMonitor.canMakeAPICall()) {
      console.warn(`⚠️ API rate limit reached, using stale cache for ${sheetName}`);
      const staleData = this.getCache(cacheKey);
      if (staleData) {
        return staleData;
      }
      throw new Error('Rate limit exceeded and no cached data available');
    }

    console.log(`🔄 Fetching fresh data from ${sheetName}`);
    apiMonitor.logAPICall(`getCachedSheetData:${sheetName}`);
    
    try {
      const sheet = await this.getSheet(sheetName);
      
      let rows;
      if (sheetName === CONFIG.SHEETS.ON_WORK) {
        rows = await sheet.getRows({ offset: 1 }); // เริ่มจากแถว 3
      } else {
        rows = await sheet.getRows();
      }

      // บันทึกลง cache
      this.setCache(cacheKey, rows);
      return rows;
      
    } catch (error) {
      console.error(`❌ API Error for ${sheetName}:`, error.message);
      
      // ถ้าเป็น quota error, ใช้ stale cache
      if (error.message.includes('quota') || error.message.includes('limit') || 
          error.message.includes('429') || error.message.includes('RATE_LIMIT')) {
        console.warn(`⚠️ Quota exceeded for ${sheetName}, using stale cache`);
        const staleData = this.getCache(cacheKey);
        if (staleData) {
          console.log(`📋 Using stale cache for ${sheetName} (${staleData.length} items)`);
          return staleData;
        }
      }
      
      throw error;
    }
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
      // ใช้ cached data แทนการเรียก API ใหม่
      const rows = await this.getCachedSheetData(CONFIG.SHEETS.EMPLOYEES);
      
      const employees = rows.map(row => row.get('ชื่อ-นามสกุล')).filter(name => name);
      return employees;
      
    } catch (error) {
      console.error('Error getting employees:', error);
      return [];
    }
  }  async getEmployeeStatus(employeeName) {
    try {
      // ใช้ safe method แทน
      const rows = await this.safeGetCachedSheetData(CONFIG.SHEETS.ON_WORK);
      
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
      // ตรวจสอบ cache สำหรับ stats ก่อน
      if (this.isCacheValid('stats')) {
        console.log('📊 Using cached admin stats');
        return this.getCache('stats');
      }

      console.log('🔄 Fetching fresh admin stats data');      // ใช้ safe method แทน การเรียก API
      const [employees, onWorkRows, mainRows] = await Promise.all([
        this.safeGetCachedSheetData(CONFIG.SHEETS.EMPLOYEES),
        this.safeGetCachedSheetData(CONFIG.SHEETS.ON_WORK),
        this.safeGetCachedSheetData(CONFIG.SHEETS.MAIN)
      ]);

      const totalEmployees = employees.length;
      const workingNow = onWorkRows.length;// หาจำนวนคนที่มาทำงานวันนี้ (ใช้ข้อมูลจาก ON_WORK sheet ที่มีวันที่วันนี้)
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
      
      // บันทึกลง cache
      this.setCache('stats', stats);
      
      return stats;

    } catch (error) {
      console.error('Error getting admin stats:', error);
      throw error;
    }
  }
  async getReportData(type, params) {
    try {
      // ใช้ cached data แทนการเรียก API ใหม่
      const rows = await this.getCachedSheetData(CONFIG.SHEETS.MAIN);

      let filteredRows = [];switch (type) {
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
    }  }

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
          clockInTime: employeeStatus.workRecord?.clockIn
        };
      }

      const timestamp = moment().tz(CONFIG.TIMEZONE).format('YYYY-MM-DD HH:mm:ss');
      
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
        locationName,       
        '',                 
        '',                 
        ''                  
      ]);

      const mainRowIndex = newRow.rowNumber;
      console.log(`✅ Added to MAIN sheet at row: ${mainRowIndex}`);

      const onWorkSheet = await this.getSheet(CONFIG.SHEETS.ON_WORK);
      await onWorkSheet.addRow([
        timestamp,          
        employee,           
        timestamp,          
        'ทำงาน',           
        userinfo || '',     
        `${lat},${lon}`,    
        locationName,       
        mainRowIndex,       
        line_name,          
        line_picture,       
        mainRowIndex,       
        employee            
      ]);      // Clear cache เนื่องจากมีการเพิ่มข้อมูลใหม่
      this.clearCache('onwork');
      this.clearCache('main');
      this.clearCache('stats');

      console.log(`✅ Clock In successful: ${employee} at ${this.formatTime(timestamp)}, Main row: ${mainRowIndex}`);

      // ทำการ warm cache อัตโนมัติ
      setTimeout(async () => {
        try {
          await this.getCachedSheetData(CONFIG.SHEETS.ON_WORK);
          await this.getAdminStats();
        } catch (error) {
          console.error('⚠️ Auto cache warming error:', error);
        }
      }, 2000);

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
        
        // ใช้ cached data แทนการเรียก API ใหม่
        const rows = await this.getCachedSheetData(CONFIG.SHEETS.ON_WORK);
        
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
          suggestions: suggestions.length > 0 ? suggestions : undefined
        };
      }

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
      console.log(`📍 Clock out location: ${locationName}`);      console.log(`✅ Proceeding with clock out for "${employee}"`);
      
      // ใช้ cached data แทนการเรียก API ใหม่
      const mainSheet = await this.getSheet(CONFIG.SHEETS.MAIN);
      const rows = await this.getCachedSheetData(CONFIG.SHEETS.MAIN);
      
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
      }      try {
        await workRecord.row.delete();
        console.log('✅ Removed from ON_WORK sheet');
          // Clear cache เนื่องจากมีการเปลี่ยนแปลงข้อมูล
        this.clearCache('onwork');
        this.clearCache('main');
        this.clearCache('stats');
        
        // ทำการ warm cache อัตโนมัติ
        setTimeout(async () => {
          try {
            await this.getCachedSheetData(CONFIG.SHEETS.ON_WORK);
            await this.getAdminStats();
          } catch (error) {
            console.error('⚠️ Auto cache warming error:', error);
          }
        }, 2000);
        
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
      }    } catch (error) {
      console.warn(`⚠️ Location lookup failed for ${lat}, ${lon}:`, error.message);
      // ถ้าเกิดข้อผิดพลาด ใช้พิกัดแทน
      return `${lat}, ${lon}`;
    }
  }

  // เพิ่มระบบ Emergency Mode เมื่อ API quota หมด
  setEmergencyMode(enabled) {
    this.emergencyMode = enabled;
    if (enabled) {
      console.log('🚨 Emergency mode ENABLED - Using cached data only');
      // ขยาย TTL ของ cache เป็น 1 ชั่วโมง
      Object.keys(this.cache).forEach(key => {
        this.cache[key].ttl = 3600000; // 1 hour
      });
    } else {
      console.log('✅ Emergency mode DISABLED - Normal operation resumed');
      // คืนค่า TTL เดิม
      this.cache.employees.ttl = 300000; // 5 minutes
      this.cache.onwork.ttl = 60000;     // 1 minute
      this.cache.main.ttl = 30000;       // 30 seconds
      this.cache.stats.ttl = 120000;     // 2 minutes
    }
  }

  async safeGetCachedSheetData(sheetName) {
    try {
      return await this.getCachedSheetData(sheetName);
    } catch (error) {
      console.error(`❌ Failed to get data for ${sheetName}:`, error.message);
      
      // เข้าสู่ emergency mode
      if (!this.emergencyMode) {
        this.setEmergencyMode(true);
      }
      
      // คืนค่า cache เก่า (ถ้ามี)
      const staleData = this.getCache(sheetName.toLowerCase().replace(/\s+/g, ''));
      if (staleData) {
        console.log(`📋 Using emergency cache for ${sheetName}`);
        return staleData;
      }
      
      // ถ้าไม่มี cache เลย คืนค่า array ว่าง
      console.warn(`⚠️ No cache available for ${sheetName}, returning empty data`);
      return [];
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

// ========== API Rate Limiting และ Monitoring ==========
class APIMonitor {
  constructor() {
    this.apiCalls = [];
    this.maxCallsPerMinute = 30; // จำกัด API calls ไม่เกิน 30 ครั้งต่อนาที
    this.maxCallsPerHour = 300; // จำกัด API calls ไม่เกิน 300 ครั้งต่อชั่วโมง
  }

  logAPICall(operation) {
    const now = new Date();
    this.apiCalls.push({
      timestamp: now,
      operation: operation
    });

    // ลบ logs ที่เก่าเกิน 1 ชั่วโมง
    this.apiCalls = this.apiCalls.filter(call => 
      (now - call.timestamp) < 3600000 // 1 hour
    );

    console.log(`📊 API Call: ${operation} (Total in last hour: ${this.apiCalls.length})`);
  }

  canMakeAPICall() {
    const now = new Date();
    
    // นับจำนวน API calls ในนาทีที่แล้ว
    const callsInLastMinute = this.apiCalls.filter(call => 
      (now - call.timestamp) < 60000 // 1 minute
    ).length;

    // นับจำนวน API calls ในชั่วโมงที่แล้ว
    const callsInLastHour = this.apiCalls.length;

    if (callsInLastMinute >= this.maxCallsPerMinute) {
      console.warn(`⚠️ Rate limit exceeded: ${callsInLastMinute} calls in last minute`);
      return false;
    }

    if (callsInLastHour >= this.maxCallsPerHour) {
      console.warn(`⚠️ Rate limit exceeded: ${callsInLastHour} calls in last hour`);
      return false;
    }

    return true;
  }

  getStats() {
    const now = new Date();
    const callsInLastMinute = this.apiCalls.filter(call => 
      (now - call.timestamp) < 60000
    ).length;
    const callsInLastHour = this.apiCalls.length;

    return {
      callsInLastMinute,
      callsInLastHour,
      maxCallsPerMinute: this.maxCallsPerMinute,
      maxCallsPerHour: this.maxCallsPerHour,
      percentageUsedPerMinute: (callsInLastMinute / this.maxCallsPerMinute) * 100,
      percentageUsedPerHour: (callsInLastHour / this.maxCallsPerHour) * 100
    };
  }
}

const apiMonitor = new APIMonitor();

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

    // ตรวจสอบ rate limit
    if (!apiMonitor.canMakeAPICall()) {
      return res.status(429).json({
        success: false,
        error: 'Too many requests, please try again later'
      });
    }

    apiMonitor.logAPICall('clockIn');
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

    // ตรวจสอบ rate limit
    if (!apiMonitor.canMakeAPICall()) {
      return res.status(429).json({
        success: false,
        error: 'Too many requests, please try again later'
      });
    }

    apiMonitor.logAPICall('clockOut');
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
    }    const employeeStatus = await sheetsService.getEmployeeStatus(employee);

    // ใช้ cached data แทนการเรียก API ใหม่
    const rows = await sheetsService.getCachedSheetData(CONFIG.SHEETS.ON_WORK);
    
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

// API Monitoring endpoint
app.get('/api/admin/api-stats', authenticateAdmin, (req, res) => {
  const stats = apiMonitor.getStats();
  res.json({
    success: true,
    data: stats
  });
});

// API สำหรับ manual cache refresh (สำหรับ admin เท่านั้น)
app.post('/api/admin/refresh-cache', authenticateAdmin, async (req, res) => {
  try {
    console.log('🔄 Manual cache refresh initiated by admin');
    
    // Clear all cache
    sheetsService.clearCache();
    
    // Warm critical caches
    await sheetsService.getCachedSheetData(CONFIG.SHEETS.ON_WORK);
    await sheetsService.getCachedSheetData(CONFIG.SHEETS.EMPLOYEES);
    
    res.json({
      success: true,
      message: 'Cache refreshed successfully'
    });
  } catch (error) {
    console.error('Cache refresh error:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to refresh cache'
    });
  }
});

// API สำหรับตรวจสอบสถานะ API quota
app.get('/api/admin/quota-status', authenticateAdmin, async (req, res) => {
  try {
    const apiStats = apiMonitor.getStats();
    const isEmergencyMode = sheetsService.emergencyMode || false;
    
    // ทดสอบการเชื่อมต่อ API
    let apiHealthy = true;
    let lastError = null;
    
    try {
      await sheetsService.getCachedSheetData(CONFIG.SHEETS.EMPLOYEES);
    } catch (error) {
      apiHealthy = false;
      lastError = error.message;
    }
    
    res.json({
      success: true,
      data: {
        apiHealthy,
        emergencyMode: isEmergencyMode,
        lastError,
        apiStats,
        recommendations: apiHealthy ? 
          ['ระบบทำงานปกติ'] : 
          [
            'รอให้ quota reset (ภายใน 24 ชั่วโมง)',
            'ใช้ cached data ในระยะนี้',
            'ลดการใช้งานฟีเจอร์ที่ต้องใช้ API'
          ]
      }
    });
  } catch (error) {
    res.status(500).json({
      success: false,
      error: error.message
    });
  }
});

// API สำหรับเปิด/ปิด emergency mode
app.post('/api/admin/emergency-mode', authenticateAdmin, async (req, res) => {
  try {
    const { enabled } = req.body;
    sheetsService.setEmergencyMode(enabled);
    
    res.json({
      success: true,
      message: `Emergency mode ${enabled ? 'enabled' : 'disabled'}`,
      emergencyMode: enabled
    });
  } catch (error) {
    res.status(500).json({
      success: false,
      error: error.message
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
      
      // เริ่ม Cache Warming แบบ Background
      console.log('🔥 Starting initial cache warming...');
      setTimeout(async () => {
        try {
          await sheetsService.getCachedSheetData(CONFIG.SHEETS.EMPLOYEES);
          await sheetsService.getCachedSheetData(CONFIG.SHEETS.ON_WORK);
          console.log('✅ Initial cache warming completed');
          
          // เริ่ม periodic cache warming
          setInterval(async () => {
            try {
              console.log('🔄 Background cache refresh...');
              await sheetsService.getCachedSheetData(CONFIG.SHEETS.ON_WORK);
              await sheetsService.getAdminStats();
            } catch (error) {
              console.error('⚠️ Background cache refresh error:', error);
            }
          }, 60000); // ทุก 1 นาที
          
        } catch (error) {
          console.error('⚠️ Initial cache warming error:', error);
        }
      }, 3000); // รอ 3 วินาทีก่อนเริ่ม
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
          clockIn = `${lateHour.toString().padStart(2, '0')}:${lateMin.toString().padStart(2, '0')}`;
        } else {
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