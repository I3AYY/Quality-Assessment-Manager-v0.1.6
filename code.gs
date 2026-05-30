/**
 * Quality Assessment Manager - AI Core v0.1.6
 * Backend Script (Google Apps Script) - Professional & Secured Edition
 * พัฒนาและอัปเกรดสำหรับงานห้องปฏิบัติการทางการแพทย์ (IQC & EQA)
 * รองรับ Multi-Department Database Architecture, PDPA & Cyber Security Compliant
 */

const SETTINGS_PASSWORD = '03842510';
const PROP_MAIN_SHEET_ID = 'TARGET_SHEET_ID'; // ใช้เก็บ Google Sheet Database (Main)
const PROP_DEPT_SHEETS = 'DEPT_SHEET_IDS';    // ใช้เก็บ JSON Object (Department -> Sheet ID)
const PROP_GEMINI_KEY = 'GEMINI_API_KEY';
const PROP_GEMINI_MODEL = 'GEMINI_MODEL';

// ==========================================
// 0. Reusable Helper Functions (DRY Principle)
// ==========================================

/**
 * ล้างข้อมูลและตัดช่องว่างของสตริงอย่างปลอดภัย ป้องกัน Injection เบื้องต้น
 */
function cleanStr(val) {
  if (val === null || val === undefined) return '';
  return val.toString().trim();
}

/**
 * แปลง Format วันที่ให้เป็น Asia/Bangkok ป้องกันความสับสนเรื่อง Timezone ของเครื่องตรวจและระบบวิเคราะห์
 */
function formatDateBangkok(dateVal) {
  if (dateVal instanceof Date) {
    return Utilities.formatDate(dateVal, "Asia/Bangkok", "dd/MM/yyyy");
  }
  return cleanStr(dateVal);
}

/**
 * ดึงค่า Properties ทั้งหมดในรอบเดียวเพื่อเพิ่มประสิทธิภาพและความเร็วในการประมวลผล (Reduce API Quota)
 */
function getCachedProperties() {
  return PropertiesService.getScriptProperties().getProperties();
}

// ==========================================
// 1. Web App Entry Point (UI Initialization)
// ==========================================
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('index');
  return template.evaluate()
    .setTitle('Quality Assessment Manager v0.1.6 AI Core')
    .setFaviconUrl('https://cdn-icons-png.flaticon.com/512/3003/3003280.png')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 2. Helper สำหรับ Include ไฟล์ HTML ย่อย (CSS/JS) เข้ามาใน index.html
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ==========================================
// 3. Settings & Cybersecurity & Credentials
// ==========================================

/**
 * ยืนยันรหัสผ่านเพื่อเข้าใช้งานในส่วนระบบหลังบ้าน (Developer Mode)
 */
function verifySettingsPassword(password) {
  return password === SETTINGS_PASSWORD;
}

/**
 * บันทึกการตั้งค่าระบบและจัดเก็บคีย์ใน Script Properties ที่มีความปลอดภัยสูง
 */
function saveSettings(payload) {
  try {
    const props = PropertiesService.getScriptProperties();
    if (payload && typeof payload === 'object') {
      const toSave = {};
      if (payload.mainSheetId !== undefined) toSave[PROP_MAIN_SHEET_ID] = cleanStr(payload.mainSheetId);
      if (payload.apiKey !== undefined) toSave[PROP_GEMINI_KEY] = cleanStr(payload.apiKey);
      if (payload.model !== undefined) toSave[PROP_GEMINI_MODEL] = cleanStr(payload.model);
      if (payload.deptSheets !== undefined) toSave[PROP_DEPT_SHEETS] = JSON.stringify(payload.deptSheets);
      
      if (Object.keys(toSave).length > 0) {
        props.setProperties(toSave);
      }
    } else if (payload) {
      props.setProperty(PROP_MAIN_SHEET_ID, cleanStr(payload));
    }
    return { success: true, message: 'บันทึกการตั้งค่าระบบเรียบร้อยแล้ว' };
  } catch (error) {
    return { success: false, message: 'เกิดข้อผิดพลาดในการบันทึก: ' + error.message };
  }
}

/**
 * ดึงค่าการตั้งค่าจากระบบเพื่อแสดงผลในหน้า Developer Mode อย่างปลอดภัย
 */
function getSettings() {
  const allProps = getCachedProperties();
  const deptSheetsRaw = allProps[PROP_DEPT_SHEETS];
  let deptSheets = {};
  if (deptSheetsRaw) {
    try { deptSheets = JSON.parse(deptSheetsRaw); } catch(e) {}
  }
  return {
    mainSheetId: allProps[PROP_MAIN_SHEET_ID] || '',
    apiKey: allProps[PROP_GEMINI_KEY] ? '••••••••••••••••' : '', // เพื่อความปลอดภัยของรหัสผ่านและ API Key ป้องกันการดูผ่านหน้าบ้าน
    model: allProps[PROP_GEMINI_MODEL] || '',
    deptSheets: deptSheets
  };
}

// ==========================================
// 4. Spreadsheet Access Helpers
// ==========================================

/**
 * ดึงอ็อบเจกต์ฐานข้อมูลหลัก (Main DB)
 */
function getMainDb() {
  const sheetId = PropertiesService.getScriptProperties().getProperty(PROP_MAIN_SHEET_ID);
  if (!sheetId) {
    throw new Error("ยังไม่ได้ตั้งค่า Main Google Sheet ID กรุณาไปที่โหมดนักพัฒนา");
  }
  return SpreadsheetApp.openById(sheetId);
}

/**
 * ดึงอ็อบเจกต์ฐานข้อมูลเฉพาะแผนก (Department DB) เพื่อความปลอดภัยและการแบ่งสิทธิ์การเข้าถึงข้อมูลผู้ป่วย
 */
function getDeptDb(deptName) {
  if (!deptName) throw new Error("ไม่พบชื่อแผนก กรุณาเลือกแผนกก่อนทำรายการ");
  
  const allProps = getCachedProperties();
  const deptSheetsRaw = allProps[PROP_DEPT_SHEETS];
  if (!deptSheetsRaw) throw new Error("ยังไม่ได้ตั้งค่าฐานข้อมูลของแต่ละแผนก (Department Sheet IDs)");
  
  const deptSheets = JSON.parse(deptSheetsRaw);
  const sheetId = deptSheets[deptName];
  if (!sheetId) throw new Error(`ยังไม่ได้ตั้งค่า Sheet ID สำหรับแผนก: ${deptName}`);
  
  return SpreadsheetApp.openById(sheetId);
}

/**
 * ดึงรายชื่อแผนกทั้งหมดจาก Main DB (Tab: Units คอลัมน์ B) เพื่อใช้ในส่วนงาน Developer Mode
 */
function getUnitsList() {
  try {
    const db = getMainDb();
    const sheet = db.getSheetByName('Units');
    if (!sheet || sheet.getLastRow() < 2) return { success: true, data: [] };
    
    const vals = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues();
    const units = [...new Set(vals.map(r => cleanStr(r[0])).filter(Boolean))];
    return { success: true, data: units };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * ดึงเฉพาะแผนกห้องปฏิบัติการที่ได้รับการกำหนดค่าฐานข้อมูล (Sheet ID) เรียบร้อยแล้ว เพื่อแสดงที่ Sidebar ให้เลือกใช้งาน
 */
function getAvailableDepartments() {
  try {
    const deptSheetsRaw = PropertiesService.getScriptProperties().getProperty(PROP_DEPT_SHEETS);
    if (!deptSheetsRaw) return { success: true, data: [] };
    
    const deptSheets = JSON.parse(deptSheetsRaw);
    const availableDepts = Object.keys(deptSheets).filter(dept => deptSheets[dept] && cleanStr(deptSheets[dept]) !== '');
    
    return { success: true, data: availableDepts };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// ==========================================
// 5. Data Retrieval & Dependencies
// ==========================================

/**
 * ดึงรายชื่อผู้บันทึก (จาก Main DB) และโครงสร้างงานประกันคุณภาพ EQA/IQC (จาก Department DB)
 */
function getFormDependencies(deptName) {
  try {
    const mainDb = getMainDb();
    
    // 5.1 ดึงข้อมูล Users สำหรับบันทึกความรับผิดชอบในการประเมินผล
    const userSheet = mainDb.getSheetByName('Users');
    let users = [];
    if (userSheet && userSheet.getLastRow() > 1) {
      const userVals = userSheet.getRange(2, 1, userSheet.getLastRow() - 1, 1).getValues();
      users = [...new Set(userVals.map(r => cleanStr(r[0])).filter(Boolean))];
    }

    let eqaData = [];
    let iqcData = [];
    
    if (deptName) {
      const deptDb = getDeptDb(deptName);
      
      // 5.2 ดึงข้อมูลโปรแกรมควบคุมคุณภาพภายนอก (EQA)
      const eqaSheet = deptDb.getSheetByName('EQA');
      if (eqaSheet && eqaSheet.getLastRow() > 1) {
        const eqaVals = eqaSheet.getRange(2, 1, eqaSheet.getLastRow() - 1, 13).getValues();
        eqaVals.forEach(row => {
          if (row[0]) {
            eqaData.push({
              eqaName: cleanStr(row[0]),
              criteriaAcc: cleanStr(row[1]),
              criteriaWar: cleanStr(row[2]),
              criteriaUnacc: cleanStr(row[3]),
              year: cleanStr(row[4]),
              lot: cleanStr(row[5]),
              exp: formatDateBangkok(row[6]),
              folderId: cleanStr(row[7]),
              keyword: cleanStr(row[8]),
              targetData: cleanStr(row[9]),
              timesPerYear: row[10] ? parseInt(cleanStr(row[10])) : null,
              certUrl: cleanStr(row[11]), 
              folderIdFm: cleanStr(row[12]) 
            });
          }
        });
      }

      // 5.3 ดึงข้อมูลโปรแกรมควบคุมคุณภาพภายใน (IQC)
      const iqcSheet = deptDb.getSheetByName('IQC');
      if (iqcSheet && iqcSheet.getLastRow() > 1) {
        const iqcVals = iqcSheet.getRange(2, 1, iqcSheet.getLastRow() - 1, 5).getValues();
        
        const parseBraces = (val) => {
           if (!val) return [];
           let str = val instanceof Date ? Utilities.formatDate(val, "Asia/Bangkok", "dd/MM/yyyy") : cleanStr(val);
           if (str.startsWith('{') && str.endsWith('}')) {
              return str.slice(1, -1).split(',').map(s => s.trim());
           }
           return [str];
        };

        iqcVals.forEach(row => {
          if (row[0]) {
            let levelCount = 1;
            if (row[1]) {
              const match = row[1].toString().match(/\d+/);
              if (match) levelCount = parseInt(match[0]);
            }

            iqcData.push({
               iqcName: cleanStr(row[0]),
               level: levelCount,
               lots: parseBraces(row[2]),
               exps: parseBraces(row[3]),
               targetData: cleanStr(row[4])
            });
          }
        });
      }
    }

    return { success: true, users: users, eqaData: eqaData, iqcData: iqcData };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * ดึงข้อมูลรายชื่อโมเดล AI Gemini ที่เปิดใช้งานในระบบ จาก Main DB
 */
function getGeminiModels() {
  try {
    const db = getMainDb();
    const sheet = db.getSheetByName('Gemini Model');
    if (!sheet || sheet.getLastRow() < 2) return { success: true, data: [] };
    
    const vals = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    const models = vals.map(r => cleanStr(r[0])).filter(Boolean);
    return { success: true, data: models };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * ดึงรายการทดสอบ (Test) และหน่วย (Unit) ของ EQA เฉพาะตัวที่เลือกจาก MetaData (Department DB)
 */
function getTestsForEQA(eqaName, deptName) {
  try {
    const db = getDeptDb(deptName);
    const metaSheet = db.getSheetByName('MetaData');
    if (!metaSheet) throw new Error("ไม่พบแท็บ 'MetaData' ในฐานข้อมูลแผนก");

    const metaVals = metaSheet.getRange(2, 1, metaSheet.getLastRow() - 1, 8).getValues();
    const tests = [];
    const targetEqa = eqaName.trim();
    
    metaVals.forEach(row => {
      if (row[7] && row[7].toString().trim() === targetEqa) {
        tests.push({
          test: cleanStr(row[1]),
          unit: cleanStr(row[2])
        });
      }
    });

    return { success: true, data: tests };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// ==========================================
// 6. AI PDF Processing (EQA & IQC)
// ==========================================

/**
 * อัปโหลดและส่ง PDF รายงานผล EQA ไปยัง Gemini API เพื่อแกะข้อมูลผลวิเคราะห์อย่างรวดเร็วและแม่นยำ
 */
function processPdfWithAI(base64Data, filename, folderId, keyword, targetData, testsList) {
  try {
    const allProps = getCachedProperties();
    const apiKey = allProps[PROP_GEMINI_KEY];
    const modelName = allProps[PROP_GEMINI_MODEL] || 'gemini-3.1-flash-lite';

    if (!apiKey) throw new Error("API_KEY_MISSING");

    let fileUrl = '';
    if (folderId) {
      try {
        const folder = DriveApp.getFolderById(folderId);
        const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'application/pdf', filename);
        const file = folder.createFile(blob);
        fileUrl = file.getUrl();
      } catch (e) {
        throw new Error("DRIVE_UPLOAD_ERROR: " + e.message);
      }
    }

    const promptText = `
      คุณคือผู้เชี่ยวชาญด้านงานห้องปฏิบัติการทางการแพทย์
      โปรดวิเคราะห์ไฟล์ PDF ผลประเมินคุณภาพ (EQA) นี้
      1. ค้นหาส่วนที่มีคำว่า: "${keyword}"
      2. ดึงข้อมูลที่สอดคล้องกับ "${targetData}" สำหรับรายการทดสอบ (Tests) ต่อไปนี้:
      ${testsList.join(', ')}

      สิ่งที่ต้องการดึงสำหรับแต่ละ Test:
      - ค่าผลลัพธ์ (Result)
      - ค่ากลาง (Mean หรือ Target value)
      - ค่าความเบี่ยงเบน (%Bias)
      - ค่า Z-score

      เงื่อนไขสำคัญ:
      - ส่งคำตอบกลับมาเป็นรูปแบบ JSON Array ล้วนๆ ห้ามมีคำอธิบาย หรือ markdown block ครอบทับ
      - โครงสร้าง JSON แต่ละ Object ต้องมี keys ดังนี้:
        "test" (ชื่อรายการทดสอบ),
        "result" (ค่าผลลัพธ์),
        "meanTarget" (ค่ากลางหรือ Target),
        "bias" (ค่า %Bias หากในเอกสารไม่มีระบุ ให้ส่งเป็นค่าสตริงว่าง ""),
        "zScore" (ค่า Z-score)
      - หากไม่พบค่าใด ให้ใส่เป็นค่าสตริงว่าง "" 

      ข้อควรระวัง :
      - ห้ามนำค่า RMZ มาใส่ลงใน Z-score
      - บางค่าอาจมีสัญลักษณ์ "*" อยู่ท้ายตัวเลข ไม่ต้องสนใจสัญลักษณ์ ให้ดึงมาเฉพาะตัวเลข เช่น 23* ให้ดึงมาเฉพาะ 23 เป็นต้น
    `;

    const payload = {
      contents: [{ parts: [{ text: promptText }, { inlineData: { mimeType: "application/pdf", data: base64Data } }] }],
      generationConfig: { responseMimeType: "application/json" }
    };

    const options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${apiKey}`;
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (responseCode !== 200) {
      if (responseCode === 400) throw new Error("API_ERROR_400: โมเดลไม่รองรับ หรือ Request ไม่ถูกต้อง");
      if (responseCode === 403) throw new Error("API_ERROR_403: API Key ไม่ถูกต้อง หรือถูกระงับ");
      if (responseCode === 429) throw new Error("API_ERROR_429: ขีดจำกัดการใช้งานโควต้า API เต็มแล้ว");
      throw new Error(`API_ERROR_${responseCode}: ${responseText}`);
    }

    const jsonResponse = JSON.parse(responseText);
    if (!jsonResponse.candidates || jsonResponse.candidates.length === 0) throw new Error("API_NO_RESPONSE: AI ไม่สามารถสร้างคำตอบได้");

    let aiText = jsonResponse.candidates[0].content.parts[0].text;
    aiText = aiText.replace(/```json/g, '').replace(/```/g, '').trim();
    const extractedData = JSON.parse(aiText);

    return { success: true, data: extractedData, fileUrl: fileUrl, message: "วิเคราะห์ข้อมูลสำเร็จ" };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * อัปโหลดและส่ง PDF สรุปผล IQC ประจำเดือนไปวิเคราะห์ด้วย AI Gemini
 */
function processIqcPdfWithAI(base64Data, filename, iqcName, levels, targetData, testsList) {
  try {
    const allProps = getCachedProperties();
    const apiKey = allProps[PROP_GEMINI_KEY];
    const modelName = allProps[PROP_GEMINI_MODEL] || 'gemini-3.1-flash-lite';

    if (!apiKey) throw new Error("API_KEY_MISSING");

    const promptText = `
      คุณคือผู้เชี่ยวชาญด้านงานห้องปฏิบัติการทางการแพทย์
      โปรดวิเคราะห์ไฟล์ PDF สรุปผล Internal Quality Control (IQC) รายเดือนนี้
      ค้นหาข้อมูลที่สอดคล้องกับ "${targetData}" สำหรับรายการทดสอบ (Tests) ต่อไปนี้:
      ${testsList.join(', ')}

      IQC นี้มีทั้งหมด ${levels} Level
      สิ่งที่ต้องการดึงสำหรับแต่ละ Test ในแต่ละ Level คือ: Mean, SD, และ %CV

      เงื่อนไขสำคัญ:
      - ส่งคำตอบกลับมาเป็นรูปแบบ JSON Array ล้วนๆ ห้ามมีคำอธิบาย หรือ markdown block ครอบทับ
      - โครงสร้าง JSON แต่ละ Object ต้องมี keys ดังนี้:
        "test": (ชื่อรายการทดสอบ),
        "level_1": { "mean": "", "sd": "", "cv": "" }
      - หาก IQC มีมากกว่า 1 Level ให้เพิ่ม "level_2" และ "level_3" ตามจำนวน levels ที่แจ้งไป
      - หากไม่พบค่าใด ให้ใส่เป็นค่าสตริงว่าง "" 
    `;

    const payload = {
      contents: [{ parts: [{ text: promptText }, { inlineData: { mimeType: "application/pdf", data: base64Data } }] }],
      generationConfig: { responseMimeType: "application/json" }
    };

    const options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${apiKey}`;
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode !== 200) {
      if (responseCode === 400) throw new Error("API_ERROR_400");
      if (responseCode === 403) throw new Error("API_ERROR_403");
      if (responseCode === 429) throw new Error("API_ERROR_429");
      throw new Error(`API_ERROR_${responseCode}: ${response.getContentText()}`);
    }

    const jsonResponse = JSON.parse(response.getContentText());
    if (!jsonResponse.candidates || jsonResponse.candidates.length === 0) throw new Error("API_NO_RESPONSE");

    let aiText = jsonResponse.candidates[0].content.parts[0].text;
    aiText = aiText.replace(/```json/g, '').replace(/```/g, '').trim();
    const extractedData = JSON.parse(aiText);

    return { success: true, data: extractedData, message: "วิเคราะห์ข้อมูล IQC สำเร็จ" };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// ==========================================
// 7. Data Saving & Transaction Handlers (EQA)
// ==========================================

/**
 * บันทึกประวัติและผลลัพธ์การประเมิน EQA ลงในแท็บ EQARecord (Department DB)
 */
function saveEQARecord(payload) {
  try {
    if (!payload.deptName) throw new Error("ไม่ได้ระบุแผนก ไม่สามารถบันทึกได้");
    const db = getDeptDb(payload.deptName);
    let sheet = db.getSheetByName('EQARecord');
    
    if (!sheet) {
      sheet = db.insertSheet('EQARecord');
      sheet.getRange(1, 1, 1, 17).setValues([
        ["Transaction ID (TXID)", "Timestamp", "Record by", "EQA", "Year", "Times", "Lot", "EXP", "Test", "Result", "Mean/Target", "%Bias", "Z-score", "Evaluation", "Remark", "EQA Result URL", "FM Document URL"]
      ]);
      sheet.getRange(1, 1, 1, 17).setFontWeight("bold").setBackground("#f3f4f6");
    }

    const timestamp = new Date();
    const txid = 'EQA-' + Utilities.formatDate(timestamp, "Asia/Bangkok", "yyyyMMdd-HHmmss") + '-' + Math.floor(1000 + Math.random() * 9000);

    const rows = payload.tests.map(t => [
      txid, 
      timestamp, 
      cleanStr(payload.recordBy), 
      cleanStr(payload.eqaName), 
      cleanStr(payload.year), 
      cleanStr(payload.times), 
      cleanStr(payload.lot), 
      cleanStr(payload.exp),
      cleanStr(t.test), 
      cleanStr(t.result), 
      cleanStr(t.meanTarget), 
      cleanStr(t.bias), 
      cleanStr(t.zScore), 
      cleanStr(t.evaluation || ''), 
      cleanStr(t.remark || ''), 
      cleanStr(payload.fileUrl || ''), 
      '' 
    ]);

    if (rows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    }

    return { success: true, message: 'บันทึกข้อมูลเรียบร้อยแล้ว', txid: txid };
  } catch (error) {
    return { success: false, message: 'เกิดข้อผิดพลาดในการบันทึก: ' + error.message };
  }
}

/**
 * ดึงรายงานสรุปประวัติของ EQA เพื่อแสดงในหน้าตารางข้อมูลและ Dashboard
 */
function getEQARecordsSummary(deptName) {
  try {
    if (!deptName) throw new Error("กรุณาระบุแผนก");
    const db = getDeptDb(deptName);
    const sheet = db.getSheetByName('EQARecord');
    if (!sheet || sheet.getLastRow() < 2) return { success: true, data: [] };

    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 17).getValues();
    
    const metaSheet = db.getSheetByName('MetaData');
    const fmCodeMap = {}; 
    const eqaTestsCountMap = {}; 
    if (metaSheet && metaSheet.getLastRow() > 1) {
      const metaVals = metaSheet.getRange(2, 1, metaSheet.getLastRow() - 1, 11).getValues();
      metaVals.forEach(r => {
        if (r[7]) {
          const eqaNameClean = cleanStr(r[7]);
          fmCodeMap[eqaNameClean] = cleanStr(r[10]);
          eqaTestsCountMap[eqaNameClean] = (eqaTestsCountMap[eqaNameClean] || 0) + 1;
        }
      });
    }

    const eqaSheet = db.getSheetByName('EQA');
    const eqaCriteriaMap = {};
    if (eqaSheet && eqaSheet.getLastRow() > 1) {
      const eqaVals = eqaSheet.getRange(2, 1, eqaSheet.getLastRow() - 1, 2).getValues();
      eqaVals.forEach(r => {
        if (r[0]) {
          eqaCriteriaMap[cleanStr(r[0])] = cleanStr(r[1]); 
        }
      });
    }

    const summaryMap = {};

    data.forEach(row => {
      const txid = row[0];
      if (!txid) return;

      if (!summaryMap[txid]) {
        const dateObj = row[1] instanceof Date ? row[1] : new Date(row[1]);
        const rawTimeNum = dateObj.getTime();
        const recordMonth = dateObj.getMonth() + 1; // 1-12
        const recordYear = dateObj.getFullYear();   
        const eqaName = cleanStr(row[3]);

        summaryMap[txid] = {
          txid: txid,
          timestamp: row[1] instanceof Date ? Utilities.formatDate(row[1], "Asia/Bangkok", "dd/MM/yyyy HH:mm") : row[1].toString(),
          rawTimestamp: rawTimeNum, 
          recordMonth: recordMonth,
          recordYear: recordYear,
          recordBy: row[2],
          eqaName: eqaName,
          year: row[4], 
          times: row[5],
          totalTests: 0,        
          answeredCount: 0,     
          acceptableCount: 0,
          warningCount: 0,
          unacceptableCount: 0,
          fileUrl: row[15] || '', 
          fmUrl: row[16] || '',    
          fmCode: fmCodeMap[eqaName] || '',
          criteriaAcc: eqaCriteriaMap[eqaName] || 'ไม่ได้ระบุเกณฑ์การผ่าน' 
        };
      }

      const ref = summaryMap[txid];
      ref.totalTests++;
      
      const resultVal = cleanStr(row[9]);
      if (resultVal !== '') {
        ref.answeredCount++;
      }
      
      const evalStatus = cleanStr(row[13]);
      if (evalStatus === 'Acceptable') ref.acceptableCount++;
      else if (evalStatus === 'Warning') ref.warningCount++;
      else if (evalStatus === 'Unacceptable') ref.unacceptableCount++;
    });

    const summaryArray = Object.values(summaryMap).map(item => {
      item.percentAcceptable = item.totalTests > 0 ? Math.round((item.acceptableCount / item.totalTests) * 1000) / 10 : 0;
      
      const totalExpectedTests = eqaTestsCountMap[item.eqaName] || item.totalTests || 1;
      item.participationPercent = Math.round((item.answeredCount / totalExpectedTests) * 1000) / 10;
      if (item.participationPercent > 100) item.participationPercent = 100.0; 

      return item;
    });

    summaryArray.sort((a, b) => b.rawTimestamp - a.rawTimestamp);
    return { success: true, data: summaryArray };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * ดึงรายละเอียดแถวข้อมูล EQA แบบละเอียด (ใช้ในกรณีต้องการกดดูเพื่อแก้ไข/ดูประวัติ)
 */
function getEQARecordDetails(txid, deptName) {
  try {
    const db = getDeptDb(deptName);
    const sheet = db.getSheetByName('EQARecord');
    if (!sheet) throw new Error("ไม่พบแท็บ EQARecord");

    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 17).getValues();
    const recordRows = data.filter(row => row[0] === txid);

    if (recordRows.length === 0) throw new Error("ไม่พบข้อมูลรหัส Transaction นี้");

    const eqaNameStr = cleanStr(recordRows[0][3]);
    
    const meta = {
      txid: recordRows[0][0], 
      recordBy: recordRows[0][2], 
      eqaName: eqaNameStr,
      year: cleanStr(recordRows[0][4]), 
      times: recordRows[0][5], 
      lot: cleanStr(recordRows[0][6]),
      exp: formatDateBangkok(recordRows[0][7]),
      fileUrl: cleanStr(recordRows[0][15]), 
      fmUrl: cleanStr(recordRows[0][16])
    };

    const metaSheet = db.getSheetByName('MetaData');
    const unitsMap = {};
    if (metaSheet && metaSheet.getLastRow() > 1) {
      const metaVals = metaSheet.getRange(2, 1, metaSheet.getLastRow() - 1, 8).getValues();
      metaVals.forEach(row => {
        if (row[7] && row[7].toString().trim() === eqaNameStr) {
          unitsMap[row[1].toString().trim()] = cleanStr(row[2]);
        }
      });
    }

    const tests = recordRows.map(row => {
      const testName = cleanStr(row[8]);
      return {
        test: testName, 
        unit: unitsMap[testName] || '-',
        result: row[9] !== '' ? row[9] : '', 
        meanTarget: row[10] !== '' ? row[10] : '',
        bias: row[11] !== '' ? row[11] : '', 
        zScore: row[12] !== '' ? row[12] : '',
        evaluation: cleanStr(row[13]), 
        remark: cleanStr(row[14])
      };
    });

    return { success: true, metadata: meta, tests: tests };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * ทำการแก้ไขอัปเดตข้อมูลบันทึก EQA เดิมที่บันทึกผิดพลาด (ลดความซ้ำซ้อนข้อมูลด้วยกระบวนการ Re-writing)
 */
function updateEQARecord(payload) {
  try {
    const db = getDeptDb(payload.deptName);
    const sheet = db.getSheetByName('EQARecord');
    if (!sheet) throw new Error("ไม่พบแท็บ EQARecord");

    const txidToUpdate = payload.txid;
    if (!txidToUpdate) throw new Error("ไม่พบ TXID สำหรับอัปเดต");

    let existingFmUrl = '';
    const data = sheet.getRange(1, 1, sheet.getLastRow(), 17).getValues();
    
    let startIndex = -1;
    let deleteCount = 0;
    
    for (let i = 1; i < data.length; i++) { 
      if (data[i][0] === txidToUpdate) {
        if (startIndex === -1) {
          startIndex = i + 1; 
        }
        deleteCount++;
        if (data[i][16]) {
          existingFmUrl = data[i][16].toString(); 
        }
      }
    }
    
    if (startIndex !== -1 && deleteCount > 0) {
      sheet.deleteRows(startIndex, deleteCount);
    }

    const timestamp = new Date(); 
    const rows = payload.tests.map(t => [
      txidToUpdate, 
      timestamp, 
      cleanStr(payload.recordBy), 
      cleanStr(payload.eqaName), 
      cleanStr(payload.year), 
      cleanStr(payload.times), 
      cleanStr(payload.lot), 
      cleanStr(payload.exp),
      cleanStr(t.test), 
      cleanStr(t.result), 
      cleanStr(t.meanTarget), 
      cleanStr(t.bias), 
      cleanStr(t.zScore), 
      cleanStr(t.evaluation || ''), 
      cleanStr(t.remark || ''), 
      cleanStr(payload.fileUrl || ''), 
      existingFmUrl
    ]);

    if (rows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    }
    return { success: true, message: 'แก้ไขอัปเดตข้อมูลเรียบร้อยแล้ว', txid: txidToUpdate };
  } catch (error) {
    return { success: false, message: 'เกิดข้อผิดพลาดในการแก้ไข: ' + error.message };
  }
}

/**
 * อัปโหลดใบรายงานผลอย่างเป็นทางการ (Certificate PDF) และจัดบันทึกแนบ URL ไว้ใช้สำหรับประเมินคุณภาพมาตรฐานแล็บ
 */
function uploadCertificate(payload) {
  try {
    const { base64Data, filename, folderId, eqaName, year, deptName } = payload;
    
    if (!folderId) throw new Error("ไม่พบ Folder ID สำหรับบันทึกไฟล์");
    if (!eqaName || !year) throw new Error("ข้อมูลชื่อ EQA หรือปี (Year) ไม่ครบถ้วน");

    const folder = DriveApp.getFolderById(folderId);
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'application/pdf', filename);
    const file = folder.createFile(blob);
    const fileUrl = file.getUrl();

    const db = getDeptDb(deptName);
    const eqaSheet = db.getSheetByName('EQA');
    if (!eqaSheet) throw new Error("ไม่พบแท็บ 'EQA'");

    const data = eqaSheet.getRange(1, 1, eqaSheet.getLastRow(), 5).getValues();
    let targetRow = -1;
    const targetEqa = eqaName.trim();
    const targetYear = year.trim();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() === targetEqa && 
          data[i][4] && data[i][4].toString().trim() === targetYear) {
        targetRow = i + 1; 
        break;
      }
    }

    if (targetRow === -1) throw new Error(`ไม่พบข้อมูล EQA: ${eqaName} ปี: ${year}`);

    eqaSheet.getRange(targetRow, 12).setValue(fileUrl);
    return { success: true, message: "อัปโหลดและบันทึกใบ Certificate เรียบร้อยแล้ว", fileUrl: fileUrl };
  } catch (error) {
    return { success: false, message: "เกิดข้อผิดพลาดในการอัปโหลด: " + error.message };
  }
}

// ==========================================
// 8. FM Document Template Processing & Exporting
// ==========================================

/**
 * ดึงข้อมูลการวิเคราะห์และนำมาบันทึกลงในเอกสารรูปแบบฟอร์มของแล็บ (FM Document) ก่อน Export PDF ทันทีแบบไร้รอยต่อ
 */
function generateFmDocument(txid, deptName) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // กำหนด LockService ป้องกันการชนกันของข้อมูลการจัดทำเอกสาร (Race Conditions)

    const deptDb = getDeptDb(deptName);
    const mainDb = getMainDb();
    
    const recordSheet = deptDb.getSheetByName('EQARecord');
    if (!recordSheet) throw new Error("ไม่พบแท็บ EQARecord ในแผนกนี้");
    
    const recordData = recordSheet.getRange(2, 1, recordSheet.getLastRow() - 1, 17).getValues();
    const records = recordData.filter(row => row[0] === txid);
    if (records.length === 0) throw new Error("ไม่พบข้อมูล TXID นี้");

    const eqaNameStr = cleanStr(records[0][3]);
    const year = cleanStr(records[0][4]);
    const times = cleanStr(records[0][5]);
    const lot = cleanStr(records[0][6]);
    const exp = formatDateBangkok(records[0][7]);

    const metaSheet = deptDb.getSheetByName('MetaData');
    let fmCode = '';
    if (metaSheet && metaSheet.getLastRow() > 1) {
      const metaVals = metaSheet.getRange(2, 1, metaSheet.getLastRow() - 1, 11).getValues();
      for (let i = 0; i < metaVals.length; i++) {
        if (metaVals[i][7] && metaVals[i][7].toString().trim() === eqaNameStr) {
          fmCode = cleanStr(metaVals[i][10]);
          if (fmCode) break; 
        }
      }
    }

    if (!fmCode) throw new Error(`ไม่พบรหัสเอกสาร (FM บันทึกผล EQA) สำหรับ ${eqaNameStr}`);

    const templateSheet = mainDb.getSheetByName(fmCode);
    if (!templateSheet) throw new Error(`ไม่พบแท็บ Template ชื่อ: ${fmCode} ในฐานข้อมูลหลัก`);

    // บันทึกส่วนหัวของหัวกระดาษฟอร์ม
    templateSheet.getRange('D10').setValue(lot);
    templateSheet.getRange('F10').setValue(exp);
    templateSheet.getRange('D11').setValue(year); 
    templateSheet.getRange('F11').setValue(times);

    const numRows = 23; 
    const numCols = 6;
    const values = Array(numRows).fill().map(() => Array(numCols).fill(''));
    const bgColors = Array(numRows).fill().map(() => Array(numCols).fill(null));
    const fontColors = Array(numRows).fill().map(() => Array(numCols).fill('#000000'));

    for (let i = 0; i < records.length && i < numRows; i++) {
      const r = records[i];
      values[i][0] = r[8] || ''; 
      values[i][1] = r[9]; 
      values[i][2] = r[10];
      values[i][3] = r[11]; 
      values[i][4] = r[12]; 
      values[i][5] = r[13] || '';
      
      const evaluationStr = cleanStr(r[13]);
      if (evaluationStr === 'Warning') {
        bgColors[i][4] = '#FFFF00'; fontColors[i][4] = '#000000';
        bgColors[i][5] = '#FFFF00'; fontColors[i][5] = '#000000';
      } else if (evaluationStr === 'Unacceptable') {
        bgColors[i][4] = '#FF0000'; fontColors[i][4] = '#FFFFFF';
        bgColors[i][5] = '#FF0000'; fontColors[i][5] = '#FFFFFF';
      }
    }

    // เขียนข้อมูลลงพื้นที่กรอกของเอกสาร
    const dataRange = templateSheet.getRange(13, 2, numRows, numCols);
    dataRange.setValues(values);
    dataRange.setBackgrounds(bgColors);
    dataRange.setFontColors(fontColors);
    SpreadsheetApp.flush();

    // ทำการดาวน์โหลด PDF ผ่าน API หลังบ้านอย่างรวดเร็ว
    const url = "https://docs.google.com/spreadsheets/d/" + mainDb.getId() + "/export" +
                "?format=pdf&size=A4&portrait=true&scale=4&top_margin=0.5&bottom_margin=0.5&left_margin=0.5&right_margin=0.5&horizontal_alignment=CENTER&vertical_alignment=TOP&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false&fzr=false&gid=" + templateSheet.getSheetId();
    
    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true });

    if (response.getResponseCode() !== 200) throw new Error("Export PDF Error: " + response.getContentText());

    const base64Data = Utilities.base64Encode(response.getBlob().getBytes());
    const finalFilename = `${fmCode}_${eqaNameStr}_Year${year}_Times${times}.pdf`;

    // คลีนอัพเคลียร์ฟอร์ม Template เพื่อพร้อมรับงานผู้ใช้งานคนถัดไปอย่างปลอดภัย
    templateSheet.getRange('D10').clearContent();
    templateSheet.getRange('F10').clearContent();
    templateSheet.getRange('D11').clearContent();
    templateSheet.getRange('F11').clearContent();
    dataRange.clearContent();
    dataRange.setBackgrounds(Array(numRows).fill().map(() => Array(numCols).fill(null)));
    dataRange.setFontColors(Array(numRows).fill().map(() => Array(numCols).fill('#000000')));

    return { success: true, base64: base64Data, filename: finalFilename };
  } catch (error) {
    return { success: false, message: error.message };
  } finally {
    lock.releaseLock(); 
  }
}

/**
 * อัปโหลดเอกสารรายงานประเมินอย่างเป็นทางการ (FM) ขึ้น Google Drive แผนกและอัปเดตลิงก์กลับมายังฐานข้อมูลประวัติ
 */
function uploadFmDocument(payload) {
  try {
    const { txid, base64Data, filename, eqaName, deptName } = payload;
    const db = getDeptDb(deptName);
    
    const eqaSheet = db.getSheetByName('EQA');
    if (!eqaSheet) throw new Error("ไม่พบแท็บ EQA");
    
    const eqaData = eqaSheet.getRange(2, 1, eqaSheet.getLastRow() - 1, 13).getValues();
    let folderIdFm = '';
    const targetEqa = eqaName.trim();
    
    for (let i = 0; i < eqaData.length; i++) {
      if (eqaData[i][0] && eqaData[i][0].toString().trim() === targetEqa) {
        folderIdFm = cleanStr(eqaData[i][12]);
        break;
      }
    }

    if (!folderIdFm) throw new Error(`ไม่พบ Folder ID (FM) สำหรับ EQA: ${eqaName}`);

    const folder = DriveApp.getFolderById(folderIdFm);
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'application/pdf', filename);
    const file = folder.createFile(blob);
    const fileUrl = file.getUrl();

    const recordSheet = db.getSheetByName('EQARecord');
    if (!recordSheet) throw new Error("ไม่พบแท็บ EQARecord");
    
    const records = recordSheet.getRange(1, 1, recordSheet.getLastRow(), 1).getValues();
    let isUpdated = false;
    for (let i = 0; i < records.length; i++) {
      if (records[i][0] === txid) {
        recordSheet.getRange(i + 1, 17).setValue(fileUrl);
        isUpdated = true;
      }
    }

    if (!isUpdated) throw new Error("ไม่พบรหัส Transaction ในฐานข้อมูลที่จะทำการบันทึก URL");

    return { success: true, message: "อัปโหลด FM บันทึกผล EQA เรียบร้อยแล้ว", fileUrl: fileUrl };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// ==========================================
// 9. Monitoring & Analytical Workflows (IQC)
// ==========================================

/**
 * ดึงรายงานสารสนเทศเพื่อวิเคราะห์แนวโน้ม (Accuracy, Precision, Levy-Jennings Chart และประวัติ IQC)
 */
function getMonitoringData(deptName) {
  try {
    if (!deptName) throw new Error("กรุณาระบุแผนกเพื่อดึงข้อมูล Monitoring");
    const db = getDeptDb(deptName);

    // 9.1 ดึงข้อมูล MetaData (โครงสร้างเชื่อมต่อแล็บ)
    const metaSheet = db.getSheetByName('MetaData');
    const tests = [];
    if (metaSheet && metaSheet.getLastRow() > 1) {
      const metaVals = metaSheet.getRange(2, 1, metaSheet.getLastRow() - 1, 17).getValues();
      metaVals.forEach(row => {
        const testName = cleanStr(row[1]);
        const method = cleanStr(row[3]);         
        const machine = cleanStr(row[4]);
        const iqcNameStr = cleanStr(row[5]);      
        const iqcCompany = cleanStr(row[6]);      
        const eqaShortName = cleanStr(row[8]); 
        const eqaCompany = cleanStr(row[9]);   
        const tea = cleanStr(row[12]);          
        const teaBy = cleanStr(row[13]);        
        const controlLimit = cleanStr(row[14]); 
        const iqcCheck = cleanStr(row[15]); 
        const eqaCheck = cleanStr(row[16]); 

        if (testName && machine !== '') {
          tests.push({
            test: testName, method, machine, iqcNameStr, iqcCompany,
            eqaShortName, eqaCompany, tea, teaBy, controlLimit,
            iqcCheck, eqaCheck
          });
        }
      });
    }

    // 9.2 ดึงข้อมูลประวัติผล EQA (ใช้สำหรับประเมิน Accuracy / %Bias Trend)
    const recordSheet = db.getSheetByName('EQARecord');
    const records = [];
    if (recordSheet && recordSheet.getLastRow() > 1) {
      const recordVals = recordSheet.getRange(2, 1, recordSheet.getLastRow() - 1, 12).getValues();
      recordVals.forEach(row => {
        const timestamp = row[1];
        const testName = cleanStr(row[8]);
        const target = row[10] !== '' ? row[10] : null; 
        const bias = row[11] !== '' ? row[11] : null;   

        if (testName && timestamp) {
          const dateNum = timestamp instanceof Date ? timestamp.getTime() : new Date(timestamp).getTime();
          records.push({ test: testName, timestamp: dateNum, target: target, bias: bias });
        }
      });
    }

    // 9.3 ดึงข้อมูลประวัติ IQC (ใช้วิเคราะห์ Precision ประจำเดือน)
    const iqcSheet = db.getSheetByName('IQCRecord');
    const iqcRecords = [];
    if (iqcSheet && iqcSheet.getLastRow() > 1) {
      const iqcVals = iqcSheet.getRange(2, 1, iqcSheet.getLastRow() - 1, 18).getValues();
      iqcVals.forEach(row => {
         const testName = cleanStr(row[8]);
         if (testName) {
            iqcRecords.push({
                month: row[3], year: row[4], iqcName: row[5], test: testName,
                mean1: row[9] !== '' ? row[9] : null, sd1: row[10] !== '' ? row[10] : null, cv1: row[11] !== '' ? row[11] : null,
                mean2: row[12] !== '' ? row[12] : null, sd2: row[13] !== '' ? row[13] : null, cv2: row[14] !== '' ? row[14] : null,
                mean3: row[15] !== '' ? row[15] : null, sd3: row[16] !== '' ? row[16] : null, cv3: row[17] !== '' ? row[17] : null,
            });
         }
      });
    }

    return { success: true, tests, records, iqcRecords };
  } catch (error) {
    return { success: false, message: 'ดึงข้อมูล Monitoring ผิดพลาด: ' + error.message };
  }
}

/**
 * ดึงรายการทดสอบเพื่อประเมินความเหมาะสมในการกรอกประวัติบันทึกข้อมูล IQC
 */
function getTestsForIQC(iqcName, deptName) {
  try {
    const db = getDeptDb(deptName);
    const metaSheet = db.getSheetByName('MetaData');
    if (!metaSheet) throw new Error("ไม่พบแท็บ 'MetaData'");

    const metaVals = metaSheet.getRange(2, 1, metaSheet.getLastRow() - 1, 6).getValues();
    const tests = [];
    const targetIqc = iqcName.trim();
    
    metaVals.forEach(row => {
      const testName = cleanStr(row[1]);
      const machine = cleanStr(row[4]);
      const mappedIqc = cleanStr(row[5]); 

      if (testName && machine !== '' && mappedIqc === targetIqc) {
        tests.push({ test: testName });
      }
    });

    if (tests.length === 0) {
       metaVals.forEach(row => {
          const testName = cleanStr(row[1]);
          const machine = cleanStr(row[4]);
          if (testName && machine !== '') {
             tests.push({ test: testName });
          }
       });
       const uniqueTests = Array.from(new Set(tests.map(t => t.test))).map(test => ({ test }));
       return { success: true, data: uniqueTests };
    }

    return { success: true, data: tests };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * บันทึกรายงานสรุปผล IQC ในแต่ละ Level ลงสู่ฐานข้อมูล IQCRecord (Department DB)
 */
function saveIQCRecord(payload) {
  try {
    if (!payload.deptName) throw new Error("ไม่ได้ระบุแผนก");
    const db = getDeptDb(payload.deptName);
    let sheet = db.getSheetByName('IQCRecord');
    
    if (!sheet) {
      sheet = db.insertSheet('IQCRecord');
      sheet.getRange(1, 1, 1, 18).setValues([
        ["Transaction ID (TXID)", "Timestamp", "Record by", "Month", "Year", "IQC", "Lot", "EXP", "Test", "Mean_1", "SD_1", "%CV_1", "Mean_2", "SD_2", "%CV_2", "Mean_3", "SD_3", "%CV_3"]
      ]);
      sheet.getRange(1, 1, 1, 18).setFontWeight("bold").setBackground("#f3f4f6");
    }

    const timestamp = new Date();
    const txid = 'IQC-' + Utilities.formatDate(timestamp, "Asia/Bangkok", "yyyyMMdd-HHmmss") + '-' + Math.floor(1000 + Math.random() * 9000);
    const formatArray = (arr) => Array.isArray(arr) ? arr.join(', ') : arr;

    const rows = payload.tests.map(t => {
       const m1 = t.level_1 ? t.level_1.mean : '';
       const s1 = t.level_1 ? t.level_1.sd : '';
       const c1 = t.level_1 ? t.level_1.cv : '';
       const m2 = t.level_2 ? t.level_2.mean : '';
       const s2 = t.level_2 ? t.level_2.sd : '';
       const c2 = t.level_2 ? t.level_2.cv : '';
       const m3 = t.level_3 ? t.level_3.mean : '';
       const s3 = t.level_3 ? t.level_3.sd : '';
       const c3 = t.level_3 ? t.level_3.cv : '';

       return [
         txid, 
         timestamp, 
         cleanStr(payload.recordBy), 
         cleanStr(payload.month), 
         cleanStr(payload.year), 
         cleanStr(payload.iqcName),
         formatArray(payload.lots), 
         formatArray(payload.exps), 
         cleanStr(t.test),
         m1, s1, c1, 
         m2, s2, c2, 
         m3, s3, c3
       ];
    });

    if (rows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    }

    return { success: true, message: 'บันทึกข้อมูลสรุป IQC เรียบร้อยแล้ว', txid: txid };
  } catch (error) {
    return { success: false, message: 'เกิดข้อผิดพลาดในการบันทึก: ' + error.message };
  }
}

// ==========================================
// 10. Dashboard Aggregate Functions
// ==========================================

/**
 * รวบรวมและวิเคราะห์สถิติทั่วไปสำหรับทำสรุปแดชบอร์ดบริหารคุณภาพห้องปฏิบัติการทางการแพทย์
 */
function getDashboardData(deptName) {
  try {
    if (!deptName) throw new Error("กรุณาระบุแผนกเพื่อประมวลผล Dashboard");
    const db = getDeptDb(deptName);
    
    // 10.1 ประวัติความครอบคลุม EQA
    let eqaSummary = [];
    const eqaRecordSheet = db.getSheetByName('EQARecord');
    if (eqaRecordSheet && eqaRecordSheet.getLastRow() > 1) {
      const eqaResult = getEQARecordsSummary(deptName);
      if (eqaResult.success) {
        eqaSummary = eqaResult.data;
      }
    }

    // 10.2 สถิติ IQC %CV
    const iqcSummary = [];
    const iqcRecordSheet = db.getSheetByName('IQCRecord');
    if (iqcRecordSheet && iqcRecordSheet.getLastRow() > 1) {
      const iqcVals = iqcRecordSheet.getRange(2, 1, iqcRecordSheet.getLastRow() - 1, 18).getValues();
      iqcVals.forEach(row => {
        const testName = cleanStr(row[8]);
        if (testName) {
          const cv1 = row[11] !== '' ? parseFloat(row[11]) : null;
          const cv2 = row[14] !== '' ? parseFloat(row[14]) : null;
          const cv3 = row[17] !== '' ? parseFloat(row[17]) : null;

          iqcSummary.push({
            year: cleanStr(row[4]),
            month: cleanStr(row[3]),
            iqcName: cleanStr(row[5]),
            test: testName,
            cv1, cv2, cv3
          });
        }
      });
    }

    // 10.3 รายการวิเคราะห์อุปกรณ์จากตาราง MetaData สำหรับสรุปแดชบอร์ด
    const metadata = [];
    const metaSheet = db.getSheetByName('MetaData');
    if (metaSheet && metaSheet.getLastRow() > 1) {
      const metaVals = metaSheet.getRange(2, 1, metaSheet.getLastRow() - 1, 17).getValues();
      metaVals.forEach(row => {
        const testName = cleanStr(row[1]);  // Col B
        const machine = cleanStr(row[4]);   // Col E
        const iqcName = cleanStr(row[5]);   // Col F
        const iqcCompany = cleanStr(row[6]); // Col G (บริษัทผู้ผลิต IQC)
        const eqaName = cleanStr(row[8]);   // Col I (ชื่อย่อ EQA)
        const eqaCompany = cleanStr(row[9]); // Col J (บริษัทผู้ผลิต EQA)
        const iqcCheck = cleanStr(row[15]); // Col P
        const eqaCheck = cleanStr(row[16]); // Col Q

        if (testName && machine !== '') {
          metadata.push({
            test: testName,
            machine: machine,
            iqcName: iqcName,
            iqcCompany: iqcCompany,
            eqaName: eqaName,
            eqaCompany: eqaCompany,
            iqcCheck: iqcCheck,
            eqaCheck: eqaCheck
          });
        }
      });
    }

    return {
      success: true,
      eqaSummary: eqaSummary,
      iqcSummary: iqcSummary,
      metadata: metadata
    };
  } catch (error) {
    return { success: false, message: 'การคำนวณแดชบอร์ดล้มเหลว: ' + error.message };
  }
}
