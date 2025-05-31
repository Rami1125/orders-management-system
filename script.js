const CONFIG = {
  SPREADSHEET_ID: '1Xsvc-d6vvmBB9iy_aCS4T_6KTW8Urhp6L6yZwkbCGDI',
  SHEET_NAME: 'CRM1'
};

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index');
}
function validateUser(credentials) {
  try {
    const { username, password } = credentials;
    const sheet = SpreadsheetApp.getActive().getSheetByName('Users');
    const data = sheet.getDataRange().getValues();
    
    // מצא משתמש תואם
    const user = data.find(row => row[0] === username && row[1] === password);
    
    if (user) {
      return {
        success: true,
        user: {
          name: user[0],
          role: user[2] || 'user'
        }
      };
    }
    
    return {
      success: false,
      message: 'שם משתמש או סיסמה לא נכונים'
    };
    
  } catch (error) {
    return {
      success: false,
      message: 'שגיאה במערכת: ' + error.message
    };
  }
}
function getOrders() {
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)
    .getSheetByName(CONFIG.SHEET_NAME);
  return sheet.getDataRange().getValues();
}

function processFormData(data) {
  // קוד לעיבוד נתוני הטופס
  return { success: true, data };
}
function testDirectFetch() {
  const orders = loadAllOrders();
  Logger.log(orders.length);
}
function initSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // וודא שאתה משתמש בשמות הגיליונות המדויקים כפי שהם מופיעים בקובץ הגיליונות שלך
  const usersSheet = ss.getSheetByName('Users') || ss.insertSheet('Users');
  const crmSheet = ss.getSheetByName('CRM') || ss.insertSheet('CRM');
  
  // אתחול גיליון Users
  if (usersSheet.getLastRow() < 2) {
    usersSheet.clearContents();
    usersSheet.appendRow(['Username', 'Password', 'Role', 'Last Login']);
    usersSheet.appendRow(['סבן', '1234', 'admin', '']);
    usersSheet.getRange('A1:D1').setFontWeight('bold');
  }
  
  // אתחול גיליון CRM
  if (crmSheet.getLastRow() < 1) {
    crmSheet.clearContents();
    crmSheet.appendRow([
      'מספר הזמנה', 'תאריך', 'שעה', 'לקוח', 'כתובת', 
      'מחסן', 'סטטוס', 'נהג', 'סוג פעולה', 
      'זמן אספקה משוער (דקות)', 'הערות', 'תאריך יצירה'
    ]);
    crmSheet.getRange('A1:L1').setFontWeight('bold');
  }
  
  console.log('המערכת אותחלה בהצלחה!');
  return 'אתחול הושלם. בדוק את הגיליונות.';
}
function testSheetsAccess() {
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  console.log('גיליונות זמינים:');
  sheets.forEach(sheet => {
    console.log('- ' + sheet.getName());
  });
  return sheets.map(s => s.getName());
}
function addSampleOrders() {
  const crmSheet = SpreadsheetApp.getActive().getSheetByName('CRM');
  const today = new Date();
  const tomorrow = new Date();
  tomorrow.setDate(today.getDate() + 1);
  
  const sampleData = [
    [
      'ORD-' + Math.floor(1000 + Math.random() * 9000),
      Utilities.formatDate(today, 'Israel', 'yyyy-MM-dd'),
      '08:30',
      'דוד לוי',
      'הברזל 12 תל אביב',
      'החרש',
      'בהמתנה',
      'חכמת',
      'משלוח',
      45,
      'יש להתקשר לפני המשלוח',
      Utilities.formatDate(today, 'Israel', 'yyyy-MM-dd')
    ],
    [
      'ORD-' + Math.floor(1000 + Math.random() * 9000),
      Utilities.formatDate(tomorrow, 'Israel', 'yyyy-MM-dd'),
      '10:00',
      'חברת בניה בע"מ',
      'החרושת 5 רמת גן',
      'התלמיד',
      'בביצוע',
      'עלי',
      'איסוף',
      30,
      '',
      Utilities.formatDate(today, 'Israel', 'yyyy-MM-dd')
    ]
  ];
  
  sampleData.forEach(row => {
    crmSheet.appendRow(row);
  });
  
  console.log('נוספו הזמנות לדוגמה');
}
function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  const sheet = SpreadsheetApp.openById("1Xsvc-d6vvmBB9iy_aCS4T_6KTW8Urhp6L6yZwkbCGDI").getSheetByName("CRM");

  if (data.action === "getData") {
    const values = sheet.getDataRange().getValues();
    return ContentService.createTextOutput(JSON.stringify(values)).setMimeType(ContentService.MimeType.JSON);
  }

  if (data.action === "addRow") {
    sheet.appendRow([data.name, data.email]);
    return ContentService.createTextOutput(JSON.stringify({ success: true })).setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput(JSON.stringify({ error: "Invalid action" })).setMimeType(ContentService.MimeType.JSON);
}


function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}





/**
 * חילוץ כתובת, ETA וקישור מתוך טקסט תא עמודת Waze
 * @param {string} link - טקסט מתוך תא גיליון עמודת Waze
 * @return {Object} { address, eta, link }
 */
function extractWazeData(link) {
  if (!link || typeof link !== 'string') {
    return { address: '', eta: '', link: '' };
  }

  // חילוץ כתובת מתוך הטקסט: "אני בדרך אל {כתובת} עם Waze"
  const addressMatch = link.match(/אני בדרך אל (.+?) עם Waze/);
  const address = addressMatch ? addressMatch[1].trim() : '';

  // חילוץ ETA מתוך הטקסט: "אגיע בשעה 08:45"
  const etaMatch = link.match(/אגיע בשעה (\\d{1,2}:\\d{2})/);
  const eta = etaMatch ? etaMatch[1].trim() : '';

  // חילוץ קישור URL תקני
const urlMatch = link.match(/https?:\/\/[\w.-]+\/ul\?[^ \n"]+/);
  const extractedLink = urlMatch ? urlMatch[0] : '';

  return {
    address: address,
    eta: eta,
    link: extractedLink
  };
}

function updateCharts(orders) {
  const durations = Array(11).fill(0);  // שעות 7–17
  const counts = Array(11).fill(0);

  orders.forEach(order => {
    if (order.time && order.driveTime) {
      const parsed = new Date(order.time);
      const hour = parsed.getHours();
      const index = hour - 7;

      let minutes = 0;
      // אם זה בפורמט 00:25
      if (typeof order.driveTime === 'string' && order.driveTime.includes(':')) {
        const [h, m] = order.driveTime.split(':');
        minutes = parseInt(h) * 60 + parseInt(m);
      } else {
        minutes = parseInt(order.driveTime);
      }

      if (!isNaN(minutes) && index >= 0 && index < 11) {
        durations[index] += minutes;
        counts[index]++;
      }
    }
  });

  const avgDurations = durations.map((sum, i) =>
    counts[i] > 0 ? Math.round(sum / counts[i]) : 0
  );

  // עדכון גרף זמן נסיעה
  lineChart.setOption({
    xAxis: { type: 'category', data: ['07:00','08:00','09:00','10:00','11:00','12:00','13:00','14:00','15:00','16:00','17:00'] },
    yAxis: { type: 'value', name: 'דקות' },
    series: [{
      name: 'זמן נסיעה ממוצע',
      type: 'line',
      smooth: true,
      data: avgDurations,
      itemStyle: { color: '#28C76F' },
      areaStyle: {
        color: {
          type: 'linear',
          x: 0, y: 0, x2: 0, y2: 1,
          colorStops: [
            { offset: 0, color: 'rgba(40, 199, 111, 0.5)' },
            { offset: 1, color: 'rgba(40, 199, 111, 0.1)' }
          ]
        }
      }
    }]
  });
}

function loadAllOrders() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error("גיליון לא נמצא");

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

    const col = name => headers.indexOf(name);

    return data.map(row => {
      const wazeRaw = row[col('Waze')] || '';
      const wazeData = extractWazeData(wazeRaw);

      return {
        date: row[col('תאריך')],
        time: row[col('שעה')],
        warehouse: row[col('מחסן')],
        status: row[col('סטטוס')],
        driver: row[col('נהג')],
        actionType: row[col('סוג פעולה')],
        client: row[col('לקוח')],
        address: wazeData.address || row[col('כתובת')],
        wazeLink: wazeData.link,
        eta: wazeData.eta,
        orderNumber: row[col('מספר הזמנה')],
        driveTime: row[col('זמן נסיעה')] || ''
      };
    });
  } catch (e) {
    console.error("שגיאה ב-loadAllOrders:", e);
    return [];
  }
}


function getOrdersWithCache() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('orders');
  
  if (cached) return JSON.parse(cached);
  
  const freshData = loadAllOrders();
  cache.put('orders', JSON.stringify(freshData), CONFIG.CACHE_EXPIRATION);
  return freshData;
}
// פונקציות דוחות
function generateWhatsAppReport() {
  const orders = loadAllOrders();
  const summary = getOrdersSummary(orders);
  const drivers = getActiveDrivers(orders);
  
  return `דוח הזמנות ונהגים:\n\n${summary}\n\n${drivers}`;
}
function getPreviewData() {
  return {
    message: "הקוד טוען בהצלחה!",
    time: new Date().toLocaleString()
  };
}

function generateMorningReport() {
  const orders = loadAllOrders();
  const waitingOrders = orders.filter(o => o.status === 'בהמתנה');
  
  const details = waitingOrders.map(o => 
    `${o.time} - ${o.client} (${o.address})`
  ).join('\n');
  
  return `דוח בוקר - ${new Date().toLocaleDateString('he-IL')}\n\nהזמנות ממתינות (${waitingOrders.length}):\n${details}`;
}

function getOrdersSummary() {
  const orders = loadAllOrders();
  
  // הגנה מפני undefined/null
  if (!orders || !Array.isArray(orders)) {
    return "שגיאה: לא ניתן לטעון הזמנות.";
  }
  console.log("orders loaded:", orders);
console.log("orders type:", typeof orders);

  const waiting = orders.filter(o => o.status === "בהמתנה").length;
  const inProgress = orders.filter(o => o.status === "בביצוע").length;
  const completed = orders.filter(o => o.status === "הושלם").length;
  const cancelled = orders.filter(o => o.status === "מבוטל").length;

  return `סיכום הזמנות:\nממתינות: ${waiting}\nבביצוע: ${inProgress}\nהושלמו: ${completed}\nבוטלו: ${cancelled}\nסה״כ: ${orders.length}`;
}


function getActiveDrivers() {
  const orders = loadAllOrders();
  
  if (!orders || !Array.isArray(orders)) {
    return "שגיאה: לא ניתן לטעון הזמנות מהגיליון.";
  }

  const drivers = orders
    .filter(o => o.driver && o.driver.trim() !== '')
    .map(o => o.driver.trim());

  const uniqueDrivers = [...new Set(drivers)];

  return `נהגים פעילים (${uniqueDrivers.length}):\n${uniqueDrivers.join('\n')}`;
}
