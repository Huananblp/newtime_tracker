// server.js - Time Tracker with Keep-Alive และ Environment Variables ที่ปรับปรุง
const express = require('express');
const cors = require('cors');
const { GoogleSpreadsheet } = require('google-spreadsheet');
const { JWT } = require('google-auth-library');
const path = require('path');
const cron = require('node-cron');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 3000;

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
  SHEETS: {
    MAIN: 'MAIN',
    EMPLOYEES: 'EMPLOYEES',
    ON_WORK: 'ON WORK'
  },
  RENDER: {
    SERVICE_URL: process.env.RENDER_SERVICE_URL || `https://${process.env.RENDER_EXTERNAL_HOSTNAME}` || 'http://localhost:3000',
    KEEP_ALIVE_ENABLED: process.env.KEEP_ALIVE_ENABLED === 'true',
    GSA_WEBHOOK_SECRET: process.env.GSA_WEBHOOK_SECRET || 'your-secret-key'
  },
  TIMEZONE: 'Asia/Bangkok'
};

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

// Serve static files
app.use(express.static('public'));

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

  async isEmployeeOnWork(employeeName) {
    try {
      const sheet = await this.getSheet(CONFIG.SHEETS.ON_WORK);
      const rows = await sheet.getRows();
      
      return rows.some(row => row.get('ชื่อในระบบ') === employeeName);
      
    } catch (error) {
      console.error('Error checking employee work status:', error);
      return false;
    }
  }

  async getEmployeeWorkRecord(employeeName) {
    try {
      const sheet = await this.getSheet(CONFIG.SHEETS.ON_WORK);
      const rows = await sheet.getRows();
      
      const workRecord = rows.find(row => row.get('ชื่อในระบบ') === employeeName);
      return workRecord ? {
        row: workRecord,
        mainRowIndex: workRecord.get('แถวอ้างอิง')
      } : null;
      
    } catch (error) {
      console.error('Error getting employee work record:', error);
      return null;
    }
  }

  async clockIn(data) {
    try {
      const { employee, userinfo, lat, lon, line_name, line_picture } = data;
      
      // Check if already clocked in
      const isOnWork = await this.isEmployeeOnWork(employee);
      if (isOnWork) {
        return {
          success: false,
          message: 'คุณต้องลงเวลากลับก่อน',
          employee
        };
      }

      const timestamp = new Date();
      const location = `${lat},${lon}`;
      
      // Add to MAIN sheet
      const mainSheet = await this.getSheet(CONFIG.SHEETS.MAIN);
      const newRow = await mainSheet.addRow([
        employee,
        line_name,
        `=IMAGE("${line_picture}")`,
        timestamp,
        userinfo || '',
        '', // เวลาออก
        `${lat},${lon}`,
        location,
        '', // พิกัดออก
        '', // ที่อยู่ออก
        '' // ชั่วโมงทำงาน
      ]);

      // Add to ON_WORK sheet
      const onWorkSheet = await this.getSheet(CONFIG.SHEETS.ON_WORK);
      await onWorkSheet.addRow([
        timestamp,
        employee,
        timestamp,
        'ทำงาน',
        userinfo || '',
        `${lat},${lon}`,
        location,
        newRow.rowIndex,
        line_name,
        line_picture,
        newRow.rowIndex,
        employee
      ]);

      console.log(`✅ Clock In: ${employee} at ${this.formatTime(timestamp)}`);

      // เรียก GSA webhook สำหรับสร้างแผนที่และส่ง Telegram
      this.triggerMapGeneration('clockin', {
        employee, lat, lon, line_name, userinfo, timestamp
      });

      return {
        success: true,
        message: 'บันทึกเวลาเข้างานสำเร็จ',
        employee,
        time: this.formatTime(timestamp)
      };

    } catch (error) {
      console.error('Clock in error:', error);
      throw error;
    }
  }

  async clockOut(data) {
    try {
      const { employee, lat, lon, line_name } = data;
      
      // Check if clocked in
      const workRecord = await this.getEmployeeWorkRecord(employee);
      if (!workRecord) {
        return {
          success: false,
          message: 'คุณต้องลงเวลามาทำงานก่อน',
          employee
        };
      }

      const timestamp = new Date();
      const clockInTime = workRecord.row.get('เวลาเข้า');
      const hoursWorked = (timestamp - new Date(clockInTime)) / (1000 * 60 * 60);
      const location = `${lat},${lon}`;

      // Update MAIN sheet
      const mainSheet = await this.getSheet(CONFIG.SHEETS.MAIN);
      const rows = await mainSheet.getRows();
      const mainRow = rows[workRecord.mainRowIndex - 2];
      
      mainRow.set('เวลาออก', timestamp);
      mainRow.set('พิกัดออก', `${lat},${lon}`);
      mainRow.set('ที่อยู่ออก', location);
      mainRow.set('ชั่วโมงทำงาน', hoursWorked.toFixed(2));
      await mainRow.save();

      // Remove from ON_WORK sheet
      await workRecord.row.delete();

      console.log(`✅ Clock Out: ${employee} at ${this.formatTime(timestamp)} (${hoursWorked.toFixed(2)} hours)`);

      // เรียก GSA webhook สำหรับสร้างแผนที่และส่ง Telegram
      this.triggerMapGeneration('clockout', {
        employee, lat, lon, line_name, timestamp, hoursWorked
      });

      return {
        success: true,
        message: 'บันทึกเวลาออกงานสำเร็จ',
        employee,
        time: this.formatTime(timestamp),
        hours: hoursWorked.toFixed(2)
      };

    } catch (error) {
      console.error('Clock out error:', error);
      throw error;
    }
  }

  async triggerMapGeneration(action, data) {
    try {
      const gsaWebhookUrl = process.env.GSA_MAP_WEBHOOK_URL;
      if (!gsaWebhookUrl) {
        console.log('⚠️ GSA webhook URL not configured');
        return;
      }

      const payload = {
        action,
        data,
        timestamp: new Date().toISOString()
      };

      // ส่งข้อมูลไปยัง Google Apps Script เพื่อสร้างแผนที่
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
  }

  formatTime(date) {
    return date.toLocaleTimeString('th-TH', {
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit',
      timeZone: CONFIG.TIMEZONE
    }) + ' น.';
  }
}

// ========== Initialize Services ==========
const sheetsService = new GoogleSheetsService();
const keepAliveService = new KeepAliveService();

// ========== Routes ==========

// Home page
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Health check และ ping endpoint
app.get('/api/health', (req, res) => {
  res.json({
    status: 'healthy',
    timestamp: new Date().toISOString(),
    uptime: process.uptime(),
    keepAlive: keepAliveService.getStats(),
    environment: process.env.NODE_ENV || 'development'
  });
});

// Ping endpoint สำหรับ keep-alive
app.get('/api/ping', (req, res) => {
  res.json({
    status: 'pong',
    timestamp: new Date().toISOString(),
    uptime: process.uptime()
  });
});

// Webhook endpoint สำหรับรับ ping จาก GSA
app.post('/api/webhook/ping', (req, res) => {
  console.log('📨 Received ping from GSA');
  res.json({
    status: 'received',
    timestamp: new Date().toISOString()
  });
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
    console.log('🚀 Starting Time Tracker Server...');
    console.log(`🌍 Environment: ${process.env.NODE_ENV || 'development'}`);
    console.log('📊 Initializing Google Sheets service...');
    
    await sheetsService.initialize();
    
    app.listen(PORT, () => {
      console.log(`✅ Server running on port ${PORT}`);
      console.log(`📊 Google Sheets connected: ${CONFIG.GOOGLE_SHEETS.SPREADSHEET_ID}`);
      console.log(`🌐 Service URL: ${CONFIG.RENDER.SERVICE_URL}`);
      
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