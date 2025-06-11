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
    
    console.log(`🔍 Looking for employee: ${employeeName}`);
    console.log(`📊 Total rows in ON_WORK: ${rows.length}`);
    
    // ดึง headers จริงมาดู
    const headers = await sheet.getHeaderValues();
    console.log('📋 Actual headers:', headers);
    
    // Debug: แสดงข้อมูลใน ON_WORK sheet
    rows.forEach((row, index) => {
      // ลองหาจากหลายคอลัมน์ที่เป็นไปได้
      const systemName = row.get('ชื่อในระบบ');
      const employeeName2 = row.get('ชื่อพนักงาน');
      const mainRowRef1 = row.get('แถวอ้างอิง');
      const mainRowRef2 = row.get('แถวในMain');
      
      console.log(`Row ${index + 1}:`, {
        ชื่อในระบบ: systemName,
        ชื่อพนักงาน: employeeName2,
        แถวอ้างอิง: mainRowRef1,
        แถวในMain: mainRowRef2,
        เวลาเข้า: row.get('เวลาเข้า')
      });
    });
    
    // หาข้อมูลพนักงาน - ลองหลายวิธี
    const workRecord = rows.find(row => {
      const systemName = row.get('ชื่อในระบบ');
      const employeeName2 = row.get('ชื่อพนักงาน');
      
      // ลองเปรียบเทียบแบบ exact match และ contains
      return systemName === employeeName || 
             employeeName2 === employeeName ||
             (systemName && systemName.includes(employeeName)) ||
             (employeeName2 && employeeName2.includes(employeeName));
    });
    
    if (workRecord) {
      // ลองดึงหมายเลขแถวจากหลายคอลัมน์
      let mainRowIndex = null;
      
      // ลองจาก แถวอ้างอิง ก่อน
      const rowRef1 = workRecord.get('แถวอ้างอิง');
      if (rowRef1 && !isNaN(parseInt(rowRef1))) {
        mainRowIndex = parseInt(rowRef1);
      }
      
      // ถ้าไม่ได้ ลองจาก แถวในMain
      if (!mainRowIndex) {
        const rowRef2 = workRecord.get('แถวในMain');
        if (rowRef2 && !isNaN(parseInt(rowRef2))) {
          mainRowIndex = parseInt(rowRef2);
        }
      }
      
      console.log(`✅ Found work record for ${employeeName}:`, {
        mainRowIndex: mainRowIndex,
        clockIn: workRecord.get('เวลาเข้า'),
        systemName: workRecord.get('ชื่อในระบบ'),
        employeeName: workRecord.get('ชื่อพนักงาน')
      });
      
      return {
        row: workRecord,
        mainRowIndex: mainRowIndex
      };
    } else {
      console.log(`❌ No work record found for ${employeeName}`);
      
      // แสดงรายชื่อที่มีอยู่เพื่อ debug
      const availableNames = rows.map(row => ({
        systemName: row.get('ชื่อในระบบ'),
        employeeName: row.get('ชื่อพนักงาน')
      }));
      console.log('Available employees in ON_WORK:', availableNames);
      
      return null;
    }
    
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
        
        console.log(`⏰ Clock Out request for: "${employee}"`);
        console.log(`📍 Location: ${lat}, ${lon}`);
        
        // Check if clocked in
        const workRecord = await this.getEmployeeWorkRecord(employee);
        if (!workRecord) {
        console.log(`❌ Employee "${employee}" is not clocked in`);
        return {
            success: false,
            message: 'คุณต้องลงเวลามาทำงานก่อน หรือชื่อไม่ตรงกับที่ลงเวลาเข้า',
            employee
        };
        }

        const timestamp = new Date();
        
        // ดึงเวลาเข้าจาก work record
        const clockInTime = workRecord.row.get('เวลาเข้า');
        console.log(`⏰ Clock in time: ${clockInTime}`);
        
        // คำนวณชั่วโมงทำงาน
        let hoursWorked = 0;
        if (clockInTime) {
        const clockInDate = new Date(clockInTime);
        hoursWorked = (timestamp - clockInDate) / (1000 * 60 * 60);
        console.log(`⏱️ Hours worked: ${hoursWorked.toFixed(2)}`);
        }
        
        const location = `${lat},${lon}`;

        // Update MAIN sheet
        console.log(`📝 Updating MAIN sheet...`);
        const mainSheet = await this.getSheet(CONFIG.SHEETS.MAIN);
        const rows = await mainSheet.getRows();
        
        console.log(`📊 Total rows in MAIN: ${rows.length}`);
        console.log(`🎯 Target row index: ${workRecord.mainRowIndex}`);
        
        let mainRow = null;
        
        // วิธีที่ 1: ใช้ index โดยตรง (ลบ 2 เพื่อ offset header)
        if (workRecord.mainRowIndex && workRecord.mainRowIndex > 1) {
        const targetIndex = workRecord.mainRowIndex - 2; // ลบ 2 เพราะ array index เริ่มจาก 0 และ header row
        if (targetIndex >= 0 && targetIndex < rows.length) {
            mainRow = rows[targetIndex];
            console.log(`✅ Found main row by index: ${targetIndex}`);
        }
        }
        
        // วิธีที่ 2: หาจากชื่อพนักงานและเวลาเข้า (ถ้าวิธีแรกไม่ได้)
        if (!mainRow) {
        console.log('🔍 Searching by employee name and clock in time...');
        
        // หาแถวที่มีชื่อตรงกันและยังไม่ได้ clock out
        const candidateRows = rows.filter(row => {
            const rowEmployee = row.get('ชื่อพนักงาน');
            const rowClockOut = row.get('เวลาออก');
            
            return (rowEmployee === employee || 
                    (rowEmployee && rowEmployee.includes(employee)) ||
                    (employee && employee.includes(rowEmployee))) && 
                !rowClockOut; // ยังไม่ได้ clock out
        });
        
        console.log(`Found ${candidateRows.length} candidate rows`);
        
        if (candidateRows.length === 1) {
            mainRow = candidateRows[0];
            console.log(`✅ Found unique candidate row`);
        } else if (candidateRows.length > 1) {
            // หาแถวที่เวลาเข้าใกล้เคียงที่สุด
            let closestRow = null;
            let minTimeDiff = Infinity;
            
            candidateRows.forEach(row => {
            const rowClockIn = row.get('เวลาเข้า');
            if (rowClockIn && clockInTime) {
                const timeDiff = Math.abs(new Date(rowClockIn) - new Date(clockInTime));
                if (timeDiff < minTimeDiff) {
                minTimeDiff = timeDiff;
                closestRow = row;
                }
            }
            });
            
            if (closestRow && minTimeDiff < 300000) { // ห่างกันไม่เกิน 5 นาที
            mainRow = closestRow;
            console.log(`✅ Found closest matching row (time diff: ${minTimeDiff}ms)`);
            }
        }
        }
        
        if (!mainRow) {
        console.log('❌ Cannot find main row to update');
        return {
            success: false,
            message: 'ไม่พบข้อมูลการลงเวลาเข้างานที่ตรงกัน กรุณาตรวจสอบชื่อ',
            employee
        };
        }
        
        console.log('✅ Found main row, updating...');
        
        // Update main row with error handling
        try {
        mainRow.set('เวลาออก', timestamp);
        mainRow.set('พิกัดออก', `${lat},${lon}`);
        mainRow.set('ที่อยู่ออก', location);
        mainRow.set('ชั่วโมงทำงาน', hoursWorked.toFixed(2));
        await mainRow.save();
        console.log('✅ Main row updated successfully');
        } catch (updateError) {
        console.error('❌ Error updating main row:', updateError);
        throw new Error('ไม่สามารถอัปเดตข้อมูลได้: ' + updateError.message);
        }

        // Remove from ON_WORK sheet
        try {
        await workRecord.row.delete();
        console.log('✅ Removed from ON_WORK sheet');
        } catch (deleteError) {
        console.error('❌ Error deleting from ON_WORK:', deleteError);
        // ไม่ throw error เพราะข้อมูลหลักอัปเดตแล้ว
        }

        console.log(`✅ Clock Out: ${employee} at ${this.formatTime(timestamp)} (${hoursWorked.toFixed(2)} hours)`);

        // เรียก GSA webhook สำหรับสร้างแผนที่และส่ง Telegram
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
        hours: hoursWorked.toFixed(2)
        };

    } catch (error) {
        console.error('❌ Clock out error:', error);
        
        // Return more specific error message
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
