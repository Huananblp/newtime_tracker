// server.js - Time Tracker with Fixed Status Checking
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

  // ฟังก์ชันช่วยเหลือสำหรับการเปรียบเทียบชื่อ - ปรับปรุงใหม่
  normalizeEmployeeName(name) {
    if (!name) return '';
    
    return name.toString()
      .trim()
      .replace(/\s+/g, ' ') // แทนที่ whitespace หลายตัวด้วย space เดียว
      .toLowerCase(); // แปลงเป็นตัวเล็กเพื่อการเปรียบเทียบ
  }

  isNameMatch(inputName, compareName) {
    if (!inputName || !compareName) return false;
    
    const normalizedInput = this.normalizeEmployeeName(inputName);
    const normalizedCompare = this.normalizeEmployeeName(compareName);
    
    // เปรียบเทียบแบบต่างๆ
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

  // ฟังก์ชันตรวจสอบสถานะพนักงาน - ปรับปรุงใหม่
  async getEmployeeStatus(employeeName) {
    try {
      const sheet = await this.getSheet(CONFIG.SHEETS.ON_WORK);
      const rows = await sheet.getRows();
      
      console.log(`🔍 Checking status for: "${employeeName}"`);
      console.log(`📊 Total rows in ON_WORK: ${rows.length}`);
      
      if (rows.length === 0) {
        console.log('📋 ON_WORK sheet is empty');
        return { isOnWork: false, workRecord: null };
      }
      
      // หาข้อมูลพนักงาน
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
        // ลองดึงหมายเลขแถวจากหลายคอลัมน์
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
        
        // แสดงรายชื่อที่มีอยู่เพื่อ debug
        const availableNames = rows.map(row => ({
          systemName: row.get('ชื่อในระบบ'),
          employeeName: row.get('ชื่อพนักงาน')
        })).filter(emp => emp.systemName || emp.employeeName);
        
        console.log('📋 Currently working employees:', availableNames);
        
        return { isOnWork: false, workRecord: null };
      }
      
    } catch (error) {
      console.error('❌ Error checking employee status:', error);
      return { isOnWork: false, workRecord: null };
    }
  }

  // ใช้ฟังก์ชัน getEmployeeStatus แทน
  async isEmployeeOnWork(employeeName) {
    const status = await this.getEmployeeStatus(employeeName);
    return status.isOnWork;
  }

  async getEmployeeWorkRecord(employeeName) {
    const status = await this.getEmployeeStatus(employeeName);
    return status.workRecord;
  }

  async clockIn(data) {
    try {
      const { employee, userinfo, lat, lon, line_name, line_picture } = data;
      
      console.log(`⏰ Clock In request for: "${employee}"`);
      
      // ตรวจสอบสถานะพนักงาน
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

      const timestamp = new Date();
      const location = `${lat},${lon}`;
      
      console.log(`✅ Proceeding with clock in for "${employee}"`);
      
      // Add to MAIN sheet
      const mainSheet = await this.getSheet(CONFIG.SHEETS.MAIN);
      
      // เพิ่มข้อมูลแถวใหม่
      const newRow = await mainSheet.addRow([
        employee,           // ชื่อพนักงาน
        line_name,          // ชื่อไลน์
        `=IMAGE("${line_picture}")`, // รูปโปรไฟล์
        timestamp,          // เวลาเข้า
        userinfo || '',     // หมายเหตุ
        '',                 // เวลาออก
        `${lat},${lon}`,    // พิกัดเข้า
        location,           // ที่อยู่เข้า
        '',                 // พิกัดออก
        '',                 // ที่อยู่ออก
        ''                  // ชั่วโมงทำงาน
      ]);

      // ดึงหมายเลขแถวที่เพิ่งเพิ่ม
      const mainRowIndex = newRow.rowNumber;
      console.log(`✅ Added to MAIN sheet at row: ${mainRowIndex}`);

      // Add to ON_WORK sheet พร้อมหมายเลขแถวอ้างอิง
      const onWorkSheet = await this.getSheet(CONFIG.SHEETS.ON_WORK);
      await onWorkSheet.addRow([
        timestamp,          // วันที่
        employee,           // ชื่อพนักงาน
        timestamp,          // เวลาเข้า
        'ทำงาน',           // สถานะ
        userinfo || '',     // หมายเหตุ
        `${lat},${lon}`,    // พิกัดเข้า
        location,           // ที่อยู่เข้า
        mainRowIndex,       // แถวในMain
        line_name,          // ชื่อไลน์
        line_picture,       // รูปโปรไฟล์
        mainRowIndex,       // แถวอ้างอิง
        employee            // ชื่อในระบบ
      ]);

      console.log(`✅ Clock In successful: ${employee} at ${this.formatTime(timestamp)}, Main row: ${mainRowIndex}`);

      // เรียก GSA webhook สำหรับสร้างแผนที่และส่ง Telegram
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
      
      // ตรวจสอบสถานะพนักงาน
      const employeeStatus = await this.getEmployeeStatus(employee);
      
      if (!employeeStatus.isOnWork) {
        console.log(`❌ Employee "${employee}" is not clocked in`);
        
        // หาชื่อที่ใกล้เคียงที่สุด
        const onWorkSheet = await this.getSheet(CONFIG.SHEETS.ON_WORK);
        const rows = await onWorkSheet.getRows();
        
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

      const timestamp = new Date();
      const workRecord = employeeStatus.workRecord;
      
      // ดึงเวลาเข้าจาก work record
      const clockInTime = workRecord.clockIn;
      console.log(`⏰ Clock in time: ${clockInTime}`);
      
      // คำนวณชั่วโมงทำงาน
      let hoursWorked = 0;
      if (clockInTime) {
        const clockInDate = new Date(clockInTime);
        hoursWorked = (timestamp - clockInDate) / (1000 * 60 * 60);
        console.log(`⏱️ Hours worked: ${hoursWorked.toFixed(2)}`);
      }
      
      const location = `${lat},${lon}`;

      console.log(`✅ Proceeding with clock out for "${employee}"`);
      
      // Update MAIN sheet
      console.log(`📝 Updating MAIN sheet...`);
      const mainSheet = await this.getSheet(CONFIG.SHEETS.MAIN);
      const rows = await mainSheet.getRows();
      
      console.log(`📊 Total rows in MAIN: ${rows.length}`);
      console.log(`🎯 Target row index: ${workRecord.mainRowIndex}`);
      
      let mainRow = null;
      
      // วิธีที่ 1: ใช้ index ถ้ามี
      if (workRecord.mainRowIndex && workRecord.mainRowIndex > 1) {
        const targetIndex = workRecord.mainRowIndex - 2;
        
        if (targetIndex >= 0 && targetIndex < rows.length) {
          const candidateRow = rows[targetIndex];
          const candidateEmployee = candidateRow.get('ชื่อพนักงาน');
          
          // ตรวจสอบว่าชื่อตรงกันด้วย
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
      
      // วิธีที่ 2: หาจากชื่อพนักงานและเงื่อนไข (ถ้าวิธีแรกไม่ได้)
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
          // หาแถวที่เวลาเข้าใกล้เคียงที่สุด
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
          
          if (closestRow && minTimeDiff < 300000) { // ห่างกันไม่เกิน 5 นาที
            mainRow = closestRow;
            console.log(`✅ Found closest matching row (time diff: ${minTimeDiff}ms)`);
          } else {
            console.log(`❌ No close time match found (min diff: ${minTimeDiff}ms)`);
          }
        }
      }
      
      // วิธีที่ 3: หาแถวล่าสุดของพนักงานนี้
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
      
      // ตรวจสอบว่าหาเจอไหม
      if (!mainRow) {
        console.log('❌ Cannot find main row to update');
        
        return {
          success: false,
          message: 'ไม่พบข้อมูลการลงเวลาเข้างานที่ตรงกัน กรุณาตรวจสอบระบบ',
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

      console.log(`✅ Clock Out successful: ${employee} at ${this.formatTime(timestamp)} (${hoursWorked.toFixed(2)} hours)`);

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

// API สำหรับตรวจสอบสถานะพนักงาน - ปรับปรุงใหม่
app.post('/api/check-status', async (req, res) => {
  try {
    const { employee } = req.body;
    
    if (!employee) {
      return res.status(400).json({
        success: false,
        error: 'Missing employee name'
      });
    }

    // ใช้ฟังก์ชันใหม่ที่ปรับปรุงแล้ว
    const employeeStatus = await sheetsService.getEmployeeStatus(employee);

    // ดึงข้อมูล ON_WORK ทั้งหมด
    const onWorkSheet = await sheetsService.getSheet(CONFIG.SHEETS.ON_WORK);
    const rows = await onWorkSheet.getRows();
    
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

// API สำหรับดู debug info - ปรับปรุงใหม่
app.post('/api/debug-info', async (req, res) => {
  try {
    const { employee } = req.body;

    // ข้อมูลจาก MAIN sheet
    const mainSheet = await sheetsService.getSheet(CONFIG.SHEETS.MAIN);
    const mainRows = await mainSheet.getRows();
    
    // ข้อมูลจาก ON_WORK sheet
    const onWorkSheet = await sheetsService.getSheet(CONFIG.SHEETS.ON_WORK);
    const onWorkRows = await onWorkSheet.getRows();

    // หาข้อมูลที่เกี่ยวข้องกับพนักงานนี้
    const relatedMainRows = mainRows
      .filter(row => {
        const rowEmployee = row.get('ชื่อพนักงาน');
        return sheetsService.isNameMatch(employee, rowEmployee);
      })
      .slice(-5) // เอาแค่ 5 แถวล่าสุด
      .map(row => ({
        employee: row.get('ชื่อพนักงาน'),
        clockIn: row.get('เวลาเข้า'),
        clockOut: row.get('เวลาออก'),
        rowNumber: row.rowNumber
      }));

    const relatedOnWorkRows = onWorkRows
      .filter(row => {
        const systemName = row.get('ชื่อในระบบ');
        const employeeName2 = row.get('ชื่อพนักงาน');
        return sheetsService.isNameMatch(employee, systemName) ||
               sheetsService.isNameMatch(employee, employeeName2);
      })
      .map(row => ({
        systemName: row.get('ชื่อในระบบ'),
        employeeName: row.get('ชื่อพนักงาน'),
        clockIn: row.get('เวลาเข้า'),
        mainRowIndex: row.get('แถวในMain') || row.get('แถวอ้างอิง'),
        rowNumber: row.rowNumber
      }));

    res.json({
      success: true,
      data: {
        searchTerm: employee,
        mainSheetData: relatedMainRows,
        onWorkSheetData: relatedOnWorkRows,
        totalMainRows: mainRows.length,
        totalOnWorkRows: onWorkRows.length,
        headers: {
          main: await mainSheet.getHeaderValues(),
          onWork: await onWorkSheet.getHeaderValues()
        }
      }
    });

  } catch (error) {
    console.error('API Error - debug-info:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to get debug info'
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