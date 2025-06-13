// Google Apps Script - Fixed Map Service & Keep-Alive สำหรับ Time Tracker
// Deploy เป็น Web App และตั้งค่า Execute as: Me, Access: Anyone

// ========== Configuration ==========
const CONFIG = {
  TELEGRAM: {
    BOT_TOKEN: "7741909675:AAEs2v-1moVtbHdna2hmxmMj0DioQNy0CGg", //"7610983723:AAEFXDbDlq5uTHeyID8Fc5XEmIUx-LT6rJM",
    CHAT_ID: "-4587553843" //"7809169283"
  },
  RENDER_SERVICE: {
    BASE_URL: "https://newtime-tracker.onrender.com", // เปลี่ยนเป็น URL จริงของ Render
    WEBHOOK_SECRET: "https://newtime-tracker.onrender.com" // ใช้ secret เดียวกับ server
  },
  MAPS: {
    SIZE: 600,
    MAP_TYPE: Maps.StaticMap.Type.HYBRID,
    LANGUAGE: 'TH'
  }
};

// ========== Main Web App Handler ==========
function doPost(e) {
  try {
    console.log('📨 Received POST request');
    console.log('Content-Type:', e.postData.type);
    console.log('Raw data:', e.postData.contents);
    
    let requestData;
    try {
      requestData = JSON.parse(e.postData.contents);
    } catch (parseError) {
      console.error('❌ JSON Parse Error:', parseError);
      throw new Error('Invalid JSON format');
    }
    
    console.log('📋 Parsed data:', requestData);
    
    const { action, data } = requestData;
    
    if (!action || !data) {
      throw new Error('Missing action or data in request');
    }
    
    console.log(`🎯 Processing action: ${action} for ${data.employee || 'unknown'}`);
    
    let result;
    switch (action) {
      case 'clockin':
        result = handleClockIn(data);
        break;
      case 'clockout':
        result = handleClockOut(data);
        break;
      default:
        throw new Error(`Unknown action: ${action}`);
    }
    
    console.log('✅ Request processed successfully');
    return result;
    
  } catch (error) {
    console.error('❌ Error in doPost:', error);
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.message,
      timestamp: new Date().toISOString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  const action = e.parameter.action;
  
  console.log(`📨 GET request with action: ${action}`);
  
  switch (action) {
    case 'ping':
      return handlePing();
    case 'test':
      return handleTest();
    case 'debug':
      return handleDebug();
    default:
      return ContentService.createTextOutput(JSON.stringify({
        service: "Time Tracker Map Service",
        status: "running",
        timestamp: new Date().toISOString(),
        availableActions: ["ping", "test", "debug"]
      })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ========== Clock In Handler ==========
function handleClockIn(data) {
  try {
    const { employee, lat, lon, line_name, userinfo, timestamp } = data;
    
    console.log(`⏰ Processing Clock In for: ${employee}`);
    console.log(`📍 Location: ${lat}, ${lon}`);
    console.log(`💬 Line user: ${line_name}`);
    console.log(`📝 Note: ${userinfo || 'none'}`);
    
    if (!lat || !lon) {
      throw new Error("Latitude or Longitude is missing.");
    }

    if (!employee) {
      throw new Error("Employee name is missing.");
    }

    // ตรวจสอบว่าค่าพิกัดถูกต้อง
    const latitude = parseFloat(lat);
    const longitude = parseFloat(lon);
    
    if (isNaN(latitude) || isNaN(longitude)) {
      throw new Error("Invalid coordinates format.");
    }
    
    if (latitude < -90 || latitude > 90 || longitude < -180 || longitude > 180) {
      throw new Error("Coordinates out of valid range.");
    }

    // สร้างข้อความสำหรับ Telegram
    const formattedDate = Utilities.formatDate(new Date(timestamp), Session.getScriptTimeZone(), "dd/MM/yyyy");
    const formattedTime = Utilities.formatDate(new Date(timestamp), Session.getScriptTimeZone(), "HH:mm:ss") + " น.";
    
    console.log('🗓️ Formatted date/time:', formattedDate, formattedTime);
    
    // ดึงที่อยู่จาก Google Maps Geocoding
    console.log('📍 Getting location address...');
    const location = getLocationAddress(latitude, longitude);
    console.log('🏠 Address:', location);
    
    const message =
      `⏱ ลงเวลาเข้างาน\n` +
      `👤 ชื่อ-นามสกุล: *${employee}*\n` +
      `📅 วันที่: *${formattedDate}*\n` +
      `🕒 เวลา: *${formattedTime}*\n` +
      `💬 ชื่อไลน์: *${line_name || 'ไม่ระบุ'}*\n` +
      (userinfo ? `📝 หมายเหตุ: *${userinfo}*\n` : "") +
      `📍 พิกัด: *${location}*\n` +
      `🗺 [📍 ดูตำแหน่งบนแผนที่](https://www.google.com/maps/place/${latitude},${longitude})`;

    console.log('💬 Message prepared:', message.substring(0, 100) + '...');

    // สร้างแผนที่
    console.log('🗺️ Creating map...');
    const mapBlob = createMapImage(latitude, longitude);
    console.log('✅ Map created successfully');
    
    // ส่งไปยัง Telegram
    console.log('📤 Sending to Telegram...');
    const telegramResult = sendMapToTelegram(mapBlob, message);
    
    if (telegramResult.success) {
      console.log('✅ Clock In notification sent successfully');
    } else {
      console.error('❌ Failed to send Telegram notification:', telegramResult.error);
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      success: telegramResult.success,
      message: telegramResult.success ? "Clock In map sent successfully" : "Failed to send notification",
      employee: employee,
      location: location,
      timestamp: new Date().toISOString()
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    console.error('❌ Error in handleClockIn:', error);
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.message,
      employee: data.employee || 'unknown',
      timestamp: new Date().toISOString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ========== Clock Out Handler ==========
function handleClockOut(data) {
  try {
    const { employee, lat, lon, line_name, timestamp, hoursWorked } = data;
    
    console.log(`⏰ Processing Clock Out for: ${employee}`);
    console.log(`📍 Location: ${lat}, ${lon}`);
    console.log(`⏱️ Hours worked: ${hoursWorked}`);
    
    if (!lat || !lon) {
      throw new Error("Latitude or Longitude is missing.");
    }

    if (!employee) {
      throw new Error("Employee name is missing.");
    }

    // ตรวจสอบว่าค่าพิกัดถูกต้อง
    const latitude = parseFloat(lat);
    const longitude = parseFloat(lon);
    
    if (isNaN(latitude) || isNaN(longitude)) {
      throw new Error("Invalid coordinates format.");
    }

    // สร้างข้อความสำหรับ Telegram
    const formattedDate = Utilities.formatDate(new Date(timestamp), Session.getScriptTimeZone(), "dd/MM/yyyy");
    const formattedTime = Utilities.formatDate(new Date(timestamp), Session.getScriptTimeZone(), "HH:mm:ss") + " น.";
    
    // ดึงที่อยู่จาก Google Maps Geocoding
    const location = getLocationAddress(latitude, longitude);
    
    const message =
      `⏱ ลงเวลาออกงาน\n` +
      `👤 ชื่อ-นามสกุล: *${employee}*\n` +
      `📅 วันที่: *${formattedDate}*\n` +
      `🕒 เวลา: *${formattedTime}*\n` +
      `💬 ชื่อไลน์: *${line_name || 'ไม่ระบุ'}*\n` +
      `🕑 จำนวนชั่วโมงทำงาน: *${parseFloat(hoursWorked).toFixed(2)} ชั่วโมง*\n` +
      `📍 พิกัด: *${location}*\n` +
      `🗺 [📍 ดูตำแหน่งบนแผนที่](https://www.google.com/maps/place/${latitude},${longitude})`;

    // สร้างแผนที่
    const mapBlob = createMapImage(latitude, longitude);
    
    // ส่งไปยัง Telegram
    const telegramResult = sendMapToTelegram(mapBlob, message);
    
    if (telegramResult.success) {
      console.log('✅ Clock Out notification sent successfully');
    } else {
      console.error('❌ Failed to send Telegram notification:', telegramResult.error);
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      success: telegramResult.success,
      message: telegramResult.success ? "Clock Out map sent successfully" : "Failed to send notification",
      employee: employee,
      location: location,
      hoursWorked: hoursWorked,
      timestamp: new Date().toISOString()
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    console.error('❌ Error in handleClockOut:', error);
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.message,
      employee: data.employee || 'unknown',
      timestamp: new Date().toISOString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ========== Map Generation Functions ==========
function createMapImage(lat, lon) {
  try {
    console.log(`🗺️ Creating map for coordinates: ${lat}, ${lon}`);
    
    const map = Maps.newStaticMap()
      .setSize(CONFIG.MAPS.SIZE, CONFIG.MAPS.SIZE)
      .setLanguage(CONFIG.MAPS.LANGUAGE)
      .setMobile(true)
      .setMapType(CONFIG.MAPS.MAP_TYPE)
      .addMarker(lat, lon);

    const mapBlob = map.getBlob();
    
    console.log(`✅ Map created successfully, size: ${mapBlob.getBytes().length} bytes`);
    return mapBlob;
    
  } catch (error) {
    console.error('❌ Error creating map:', error);
    throw new Error(`Failed to create map image: ${error.message}`);
  }
}

function getLocationAddress(lat, lon) {
  try {
    console.log(`📍 Reverse geocoding: ${lat}, ${lon}`);
    
    const response = Maps.newGeocoder()
      .setRegion('th')
      .setLanguage('th-TH')
      .reverseGeocode(lat, lon);
      
    if (response.results && response.results.length > 0) {
      const address = response.results[0].formatted_address;
      console.log(`🏠 Address found: ${address}`);
      return address;
    } else {
      console.log('⚠️ No address found, using coordinates');
      return `${lat}, ${lon}`;
    }
    
  } catch (error) {
    console.error('❌ Error getting location address:', error);
    return `${lat}, ${lon}`;
  }
}

// ========== Telegram Functions ==========
function sendMapToTelegram(mapBlob, caption) {
  try {
    console.log('📤 Sending map to Telegram...');
    console.log(`📝 Caption length: ${caption.length} characters`);
    
    const payload = {
      "chat_id": CONFIG.TELEGRAM.CHAT_ID,
      "photo": mapBlob,
      "caption": caption,
      "parse_mode": "Markdown"
    };

    const options = {
      "method": "post",
      "payload": payload,
    };

    const url = `https://api.telegram.org/bot${CONFIG.TELEGRAM.BOT_TOKEN}/sendPhoto`;
    
    console.log(`🌐 Telegram API URL: ${url.substring(0, 50)}...`);

    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());
    
    console.log(`📊 Telegram response status: ${response.getResponseCode()}`);
    console.log('📋 Telegram response:', result);
    
    if (!result.ok) {
      throw new Error(`Telegram API error: ${result.description || 'Unknown error'}`);
    }
    
    console.log("✅ Message sent to Telegram successfully");
    return { success: true, result: result };
    
  } catch (error) {
    console.error("❌ Error sending to Telegram:", error);
    return { success: false, error: error.message };
  }
}

function sendTextToTelegram(message) {
  try {
    const payload = {
      "chat_id": CONFIG.TELEGRAM.CHAT_ID,
      "text": message,
      "parse_mode": "Markdown"
    };

    const options = {
      "method": "post",
      "contentType": "application/json",
      "payload": JSON.stringify(payload)
    };

    const url = `https://api.telegram.org/bot${CONFIG.TELEGRAM.BOT_TOKEN}/sendMessage`;
    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());
    
    if (!result.ok) {
      throw new Error(`Telegram API error: ${result.description}`);
    }
    
    return { success: true, result: result };
    
  } catch (error) {
    console.error("❌ Error sending text to Telegram:", error);
    return { success: false, error: error.message };
  }
}

// ========== Keep-Alive Functions ==========
function handlePing() {
  console.log('📨 Ping received from Render service');
  
  return ContentService.createTextOutput(JSON.stringify({
    status: "pong",
    timestamp: new Date().toISOString(),
    service: "GSA Map Service",
    config: {
      telegramConfigured: !!CONFIG.TELEGRAM.BOT_TOKEN,
      renderUrl: CONFIG.RENDER_SERVICE.BASE_URL
    }
  })).setMimeType(ContentService.MimeType.JSON);
}

function handleTest() {
  try {
    console.log('🧪 Running test function...');
    
    // ทดสอบการส่งข้อความไปยัง Telegram
    const testMessage = `🧪 ทดสอบระบบ\n📅 เวลา: ${new Date().toLocaleString('th-TH')}\n✅ Google Apps Script ทำงานปกติ`;
    
    const result = sendTextToTelegram(testMessage);
    
    console.log('🧪 Test completed:', result);
    
    return ContentService.createTextOutput(JSON.stringify({
      success: result.success,
      message: "Test completed",
      telegram: result,
      timestamp: new Date().toISOString()
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    console.error('❌ Test failed:', error);
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.message,
      timestamp: new Date().toISOString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function handleDebug() {
  try {
    const debugInfo = {
      timestamp: new Date().toISOString(),
      timezone: Session.getScriptTimeZone(),
      config: {
        telegramBotConfigured: !!CONFIG.TELEGRAM.BOT_TOKEN,
        telegramChatId: CONFIG.TELEGRAM.CHAT_ID,
        renderServiceUrl: CONFIG.RENDER_SERVICE.BASE_URL,
        mapSize: CONFIG.MAPS.SIZE,
        mapType: CONFIG.MAPS.MAP_TYPE
      },
      quotas: {
        mapsQuota: "Check manually in console",
        telegramQuota: "Check manually in console"
      }
    };
    
    console.log('🔧 Debug info:', debugInfo);
    
    return ContentService.createTextOutput(JSON.stringify(debugInfo)).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.message,
      timestamp: new Date().toISOString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ========== Scheduled Functions สำหรับ Keep-Alive ==========
function pingRenderService() {
  const now = new Date();
  const hour = now.getHours();
  
  // เช็คว่าอยู่ในเวลาทำงานไหม (04:00-23:00) - 19 ชั่วโมง/วัน
  const isWorkingHour = (hour >= 4 && hour < 23);
  
  if (!isWorkingHour) {
    console.log(`😴 Outside working hours (${hour}:00), skipping ping`);
    return;
  }
  
  try {
    const url = `${CONFIG.RENDER_SERVICE.BASE_URL}/api/webhook/ping`;
    
    const options = {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'X-Webhook-Secret': CONFIG.RENDER_SERVICE.WEBHOOK_SECRET
      },
      payload: JSON.stringify({
        timestamp: new Date().toISOString(),
        source: 'GSA Keep-Alive'
      })
    };
    
    const response = UrlFetchApp.fetch(url, options);
    
    if (response.getResponseCode() === 200) {
      console.log(`✅ Pinged Render service successfully (${hour}:00)`);
    } else {
      console.log(`⚠️ Ping response: ${response.getResponseCode()}`);
    }
    
  } catch (error) {
    console.error('❌ Error pinging Render service:', error);
  }
}

// ========== Setup Functions ==========
function setupTriggers() {
  // ลบ triggers เก่า
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'pingRenderService') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // สร้าง trigger ใหม่ทุก 10 นาที
  ScriptApp.newTrigger('pingRenderService')
    .timeBased()
    .everyMinutes(10)
    .create();
  
  console.log('✅ Keep-alive triggers created successfully');
}

function deleteTriggers() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'pingRenderService') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  console.log('🗑️ All keep-alive triggers deleted');
}

// ========== Test Functions ==========
function testMapGeneration() {
  console.log('🧪 Testing map generation...');
  
  const testData = {
    employee: 'นายทดสอบ ระบบ',
    lat: 14.0583,
    lon: 100.6014,
    line_name: 'ผู้ทดสอบ',
    userinfo: 'ทดสอบระบบ',
    timestamp: new Date().toISOString()
  };
  
  return handleClockIn(testData);
}

function checkConfiguration() {
  console.log('🔧 Checking configuration...');
  
  const issues = [];
  
  if (!CONFIG.TELEGRAM.BOT_TOKEN) {
    issues.push('❌ Telegram Bot Token not configured');
  }
  
  if (!CONFIG.TELEGRAM.CHAT_ID) {
    issues.push('❌ Telegram Chat ID not configured');
  }
  
  if (!CONFIG.RENDER_SERVICE.BASE_URL) {
    issues.push('❌ Render Service URL not configured');
  }
  
  if (issues.length === 0) {
    console.log('✅ All configuration looks good');
    return true;
  } else {
    console.log('⚠️ Configuration issues found:');
    issues.forEach(issue => console.log(issue));
    return false;
  }
}