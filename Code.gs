function doGet(e) {
  // 1. รองรับการดึงข้อมูลสรุปแดชบอร์ดผ่าน API ตรง: ?api=data
  if (e && e.parameter && e.parameter.api === 'data') {
    return ContentService.createTextOutput(getDashboardData())
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  // 2. รองรับการดึงข้อมูลรายการละเอียดในพื้นหลังผ่าน API: ?api=records
  if (e && e.parameter && e.parameter.api === 'records') {
    return ContentService.createTextOutput(getAllRecordsCompact())
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  // 3. แสดงผลหน้าเว็บแดชบอร์ด พร้อมฝังข้อมูลจริงลงไปในหน้า HTML ทันที
  let html = HtmlService.createHtmlOutputFromFile('index').getContent();
  try {
    const initialData = getDashboardData();
    html = html.replace('window.INITIAL_DATA = null;', 'window.INITIAL_DATA = ' + initialData + ';');
  } catch (err) {
    Logger.log('getDashboardData error in doGet: ' + err);
  }

  return HtmlService.createHtmlOutput(html)
    .setTitle('สถานะงานสำรวจโคมไฟสาธารณะ')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ============================================================
// 🔍 ตรวจสอบโครงสร้างคอลัมน์ใน Sheet
// เลือก "checkColumns" แล้วกด ▶️ เรียกใช้
// ============================================================
function checkColumns() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    ss = SpreadsheetApp.openById('1ItTEV7wSB5M-99TUgYzl8v2YL0NZoZXYREzwE-a9u40');
  }
  const sheet = ss.getSheetByName('ชีต1') || ss.getSheetByName('ชีท1') || ss.getSheetByName('Sheet1') || ss.getSheets()[0];
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const firstRow = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  Logger.log('📋 จำนวนคอลัมน์ทั้งหมด: ' + headers.length);
  Logger.log('📋 จำนวนแถวข้อมูลทั้งหมด: ' + (sheet.getLastRow() - 1));
  Logger.log('');
  Logger.log('=== หัวคอลัมน์และตัวอย่างข้อมูลแถวแรก ===');
  
  for (let i = 0; i < headers.length; i++) {
    const val = firstRow[i];
    const valStr = val instanceof Date ? val.toString() : String(val);
    Logger.log('คอลัมน์ ' + i + ' [' + String.fromCharCode(65+i) + ']: "' + headers[i] + '" → ตัวอย่าง: ' + valStr.substring(0, 60));
  }
}

// ============================================================
// ⭐ ขั้นตอนที่ 1: ขอสิทธิ์การเข้าถึง Sheet
// เลือก "authorizeSheets" แล้วกด ▶️ เรียกใช้
// ============================================================
function authorizeSheets() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    ss = SpreadsheetApp.openById('1ItTEV7wSB5M-99TUgYzl8v2YL0NZoZXYREzwE-a9u40');
  }
  const sheet = ss.getSheetByName('ชีต1') || ss.getSheetByName('ชีท1') || ss.getSheetByName('Sheet1') || ss.getSheets()[0];
  Logger.log('✅ ได้รับสิทธิ์แล้ว! ชื่อ Sheet: ' + sheet.getName() + ' | แถวทั้งหมด: ' + sheet.getLastRow());
}

// ============================================================
// ✅ ฟังก์ชันทดสอบ
// วิธีใช้: เลือก "testDashboard" ในช่องเลือกฟังก์ชัน แล้วกด "เรียกใช้"
// ============================================================
function testDashboard() {
  try {
    const start = new Date().getTime();
    const result = getDashboardData();
    const elapsed = (new Date().getTime() - start) / 1000;
    const parsed = JSON.parse(result);
    if (parsed.status === 'success') {
      Logger.log('✅ SUCCESS! ดึงข้อมูลสำเร็จภายใน ' + elapsed + ' วินาที');
      Logger.log('📊 จำนวนจุดทั้งหมด: ' + parsed.data.totalPoints);
      Logger.log('📍 todayCount: ' + parsed.data.todayCount);
      Logger.log('📍 yesterdayCount: ' + parsed.data.yesterdayCount);
      Logger.log('🏘️ จำนวนชุมชน: ' + parsed.data.totalCommunities);
      Logger.log('👥 จำนวนผู้สำรวจ: ' + parsed.data.totalSurveyors);
      Logger.log('📅 อัปเดตล่าสุด: ' + parsed.data.lastUpdateStr);
      Logger.log('💡 ตู้ควบคุม: ' + parsed.data.cabinetTotal);
      Logger.log('💡 ชนิดโคมไฟ (' + parsed.data.lampTypes.length + ' ชนิด): ' + JSON.stringify(parsed.data.lampTypes));
      Logger.log('🏘️ Top 3 ชุมชน: ' + JSON.stringify(parsed.data.densityByCommunity.slice(0,3)));
    } else {
      Logger.log('❌ ERROR ใน JSON: ' + parsed.message);
    }
  } catch(e) {
    Logger.log('❌ EXCEPTION: ' + e.toString());
  }
}

function getDashboardData() {
  try {
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      ss = SpreadsheetApp.openById('1ItTEV7wSB5M-99TUgYzl8v2YL0NZoZXYREzwE-a9u40');
    }
    // ดึงเฉพาะชีต "ชีต1" (งานสำรวจ) สะกดด้วย ต-เต่า เพื่อความถูกต้องและแม่นยำ
    const sheet = ss.getSheetByName('ชีต1') || ss.getSheetByName('ชีท1') || ss.getSheetByName('Sheet1') || ss.getSheets()[0];
    
    // ⚡ ดึงเฉพาะ getValues() โดยไม่เรียก getDisplayValues() เพื่อลดเวลาจาก 35 วินาที เหลือต่ำกว่า 1 วินาที!
    const dataRange = sheet.getDataRange();
    const data = dataRange.getValues();
    if (data.length <= 1) {
      return JSON.stringify({ status: 'success', data: { empty: true } });
    }
    
    const records = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;
      
      const val = row[1];
      if (!val) continue;
      
      let dateObj;
      if (val instanceof Date) {
        dateObj = val;
        let y = dateObj.getFullYear();
        if (y > 2500) {
          dateObj = new Date(y - 543, dateObj.getMonth(), dateObj.getDate(), dateObj.getHours(), dateObj.getMinutes(), dateObj.getSeconds());
        }
      } else {
        try {
          let dateStr = String(val).trim();
          let dateParts = dateStr.split(' ');
          let dateOnly = dateParts[0];
          let timeOnly = dateParts[1] || "00:00:00";
          let parts = dateOnly.split(/[-/]/);
          if (parts.length < 3) continue;
          
          let p1 = parseInt(parts[0], 10);
          let p2 = parseInt(parts[1], 10);
          let p3 = parseInt(parts[2], 10);
          
          let year, month, day;
          if (p3 > 1000) {
            year = p3;
            if (p1 > 12) { day = p1; month = p2; }
            else if (p2 > 12) { month = p1; day = p2; }
            else { day = p1; month = p2; }
          } else if (p1 > 1000) {
            year = p1; month = p2; day = p3;
          } else {
            continue;
          }
          
          if (year > 2500) year -= 543;
          
          let timeParts = timeOnly.split(':');
          let h = parseInt(timeParts[0], 10) || 0;
          let min = parseInt(timeParts[1], 10) || 0;
          let sec = parseInt(timeParts[2], 10) || 0;
          
          dateObj = new Date(year, month - 1, day, h, min, sec);
        } catch (e) {
          continue;
        }
      }
      
      if (!dateObj || isNaN(dateObj.getTime())) continue;
      records.push({
        id: row[0],
        timestamp: dateObj,
        surveyor: row[2] ? row[2].toString().trim().split('@')[0] : '',
        community: row[3] ? row[3].toString().trim() : 'ไม่ระบุ',
        soi: row[4] ? row[4].toString().trim() : 'ไม่ระบุ',
        lampType: row[5] ? row[5].toString().trim() : 'ไม่ระบุ',
        picture: row[6] || '',
        gps: row[7] || '',
        remark: row[8] || ''
      });
    }
    
    if (records.length === 0) {
      return JSON.stringify({ status: 'success', data: { empty: true } });
    }
    
    records.sort((a, b) => a.timestamp.getTime() - b.timestamp.getTime());
    const totalPoints = records.length;
    const communitiesSet = new Set();
    records.forEach(r => communitiesSet.add(r.community));
    const totalCommunities = communitiesSet.size;
    
    const TZ = 'Asia/Bangkok';
    const now = new Date();
    let todayStr = Utilities.formatDate(now, TZ, 'yyyy-MM-dd');
    let yd = new Date(now.getTime() - 86400000);
    let yesterdayStr = Utilities.formatDate(yd, TZ, 'yyyy-MM-dd');
    
    const fixBuddhistYear = (dateStr) => {
      let parts = dateStr.split('-');
      let y = parseInt(parts[0]);
      if (y > 2500) parts[0] = (y - 543).toString();
      return parts.join('-');
    };
    todayStr = fixBuddhistYear(todayStr);
    yesterdayStr = fixBuddhistYear(yesterdayStr);
    
    let todayCount = 0;
    let yesterdayCount = 0;
    let firstDate = null;
    let latestDate = null;
    const lampTypeCounts = {};
    const communityCounts = {};
    const soiCounts = {};
    const dailyCountMap = {};
    const communityLampTypes = {};
    
    records.forEach(r => {
      let d = r.timestamp;
      let y = d.getFullYear();
      if (y > 2500) y -= 543;
      let m = (d.getMonth() + 1).toString().padStart(2, '0');
      let day = d.getDate().toString().padStart(2, '0');
      let rDateStr = `${y}-${m}-${day}`;
      
      if (!firstDate || r.timestamp < firstDate) firstDate = r.timestamp;
      if (!latestDate || r.timestamp > latestDate) latestDate = r.timestamp;
      if (rDateStr === todayStr) todayCount++;
      if (rDateStr === yesterdayStr) yesterdayCount++;
      if (dailyCountMap.hasOwnProperty(rDateStr)) {
        dailyCountMap[rDateStr]++;
      } else {
        dailyCountMap[rDateStr] = 1;
      }
      lampTypeCounts[r.lampType] = (lampTypeCounts[r.lampType] || 0) + 1;
      
      if (r.lampType !== 'ตู้ควบคุม') {
        communityCounts[r.community] = (communityCounts[r.community] || 0) + 1;
        const soiKey = r.soi + ' (' + r.community + ')';
        soiCounts[soiKey] = (soiCounts[soiKey] || 0) + 1;
      }
      if (!communityLampTypes[r.community]) communityLampTypes[r.community] = {};
      communityLampTypes[r.community][r.lampType] = (communityLampTypes[r.community][r.lampType] || 0) + 1;
    });
    
    const months = ["ม.ค.","ก.พ.","มี.ค.","เม.ย.","พ.ค.","มิ.ย.","ก.ค.","ส.ค.","ก.ย.","ต.ค.","พ.ย.","ธ.ค."];
    const dailyStats = Object.keys(dailyCountMap)
      .sort()
      .slice(-10)
      .map(key => {
        const parts = key.split('-');
        return {
          dateLabel: parseInt(parts[2]) + ' ' + months[parseInt(parts[1]) - 1],
          fullDate: key,
          count: dailyCountMap[key]
        };
      });
      
    let cabinetTotal = 0;
    const filteredLampTypes = Object.keys(lampTypeCounts).filter(k => {
      if (k === 'ตู้ควบคุม') {
        cabinetTotal = lampTypeCounts[k];
        return false;
      }
      return true;
    });
    const lampTotalPoints = filteredLampTypes.reduce((sum, k) => sum + lampTypeCounts[k], 0);
    const lampTypes = filteredLampTypes.map(k => ({
      type: k,
      count: lampTypeCounts[k],
      percentage: ((lampTypeCounts[k] / lampTotalPoints) * 100).toFixed(1)
    })).sort((a, b) => b.count - a.count);
    
    const densityByCommunity = Object.keys(communityCounts).map(k => ({
      name: k, count: communityCounts[k]
    })).sort((a, b) => b.count - a.count);
    
    const communityLampStats = Object.keys(communityLampTypes).map(comm => {
      let totalLampInCommunity = 0;
      for (let t in communityLampTypes[comm]) {
         if (t !== 'ตู้ควบคุม') totalLampInCommunity += communityLampTypes[comm][t];
      }
      return {
        community: comm,
        total: totalLampInCommunity,
        types: communityLampTypes[comm]
      };
    }).sort((a, b) => b.total - a.total);
    
    const densityBySoi = Object.keys(soiCounts).map(k => {
      const parts = k.match(/(.*)\s\((.*)\)$/);
      return {
        soi: parts ? parts[1] : k,
        community: parts ? parts[2] : '',
        count: soiCounts[k]
      };
    }).sort((a, b) => b.count - a.count);
    
    const recentActivities = records.slice(-10).reverse().map(r => {
      let d = r.timestamp;
      let timeStr = d.getDate() + ' ' + months[d.getMonth()] + ' ' + d.getHours().toString().padStart(2, '0') + ':' + d.getMinutes().toString().padStart(2, '0');
      return {
        timeStr: timeStr,
        surveyor: r.surveyor ? r.surveyor.split('@')[0] : 'ไม่ระบุ',
        community: r.community,
        soi: r.soi,
        lampType: r.lampType
      };
    });
    
    const allSurveyorsSet = new Set();
    records.forEach(r => { if (r.surveyor) allSurveyorsSet.add(r.surveyor); });
    let totalSurveyDays = 1;
    if (firstDate && latestDate) {
      totalSurveyDays = Math.ceil(Math.abs(latestDate - firstDate) / 86400000) + 1;
    }
    const averagePerDay = Math.round(totalPoints / totalSurveyDays);
    
    let lastUpdateStr = '';
    const updateTime = new Date();
    let yStr = Utilities.formatDate(updateTime, TZ, 'yyyy');
    let year = parseInt(yStr);
    if (year < 2500) year += 543;
    
    lastUpdateStr = Utilities.formatDate(updateTime, TZ, 'd') + ' '
      + months[parseInt(Utilities.formatDate(updateTime, TZ, 'M')) - 1] + ' '
      + year + ' • '
      + Utilities.formatDate(updateTime, TZ, 'HH:mm');

    const allSurveyorsList = Array.from(allSurveyorsSet).sort();

    return JSON.stringify({
      status: 'success',
      data: {
        totalPoints,
        totalCommunities,
        todayCount,
        yesterdayCount,
        totalSurveyors: allSurveyorsSet.size,
        allSurveyors: allSurveyorsList,
        averagePerDay,
        lastUpdateStr,
        dailyStats,
        lampTypes,
        cabinetTotal,
        lampTotalPoints,
        communityLampStats,
        densityByCommunity,
        densityBySoi,
        recentActivities
      }
    });
  } catch (error) {
    throw new Error('getDashboardData failed: ' + error.toString());
  }
}

// ============================================================
// 📦 ดึงรายการข้อมูลแบบละเอียดในพื้นหลัง (Background Fetch)
// สำหรับระบบค้นหาและตารางข้อมูล โดยไม่ทำให้หน้าแรกกระตุกหรือหมุนค้าง
// ============================================================
function getAllRecordsCompact() {
  try {
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      ss = SpreadsheetApp.openById('1ItTEV7wSB5M-99TUgYzl8v2YL0NZoZXYREzwE-a9u40');
    }
    const sheet = ss.getSheetByName('ชีต1') || ss.getSheetByName('ชีท1') || ss.getSheetByName('Sheet1') || ss.getSheets()[0];
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return JSON.stringify({ status: 'success', records: [] });
    }
    
    const records = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0] || !row[1]) continue;
      
      let dateStr = '';
      let timeStr = '';
      const val = row[1];
      if (val instanceof Date) {
        let y = val.getFullYear();
        if (y > 2500) y -= 543;
        let m = (val.getMonth() + 1).toString().padStart(2, '0');
        let d = val.getDate().toString().padStart(2, '0');
        dateStr = `${y}-${m}-${d}`;
        timeStr = val.getHours().toString().padStart(2, '0') + ':' + val.getMinutes().toString().padStart(2, '0');
      } else {
        let s = String(val).trim();
        let p = s.split(' ');
        dateStr = p[0] || '';
        timeStr = p[1] || '';
      }
      
      records.push([
        row[0],
        dateStr,
        timeStr,
        row[2] ? String(row[2]).trim().split('@')[0] : '',
        row[3] ? String(row[3]).trim() : '',
        row[4] ? String(row[4]).trim() : '',
        row[5] ? String(row[5]).trim() : '',
        row[7] || '', // GPS
        row[6] || '', // Picture
        row[8] || ''  // Remark
      ]);
    }
    return JSON.stringify({ status: 'success', records: records });
  } catch (err) {
    return JSON.stringify({ status: 'error', message: err.toString() });
  }
}

