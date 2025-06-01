// รหัสสคริปต์ของ Google Apps สำหรับระบบการจัดหาวัสดุ // ปรับใช้สิ่งนี้เป็นแอปเว็บด้วย URL: // https://script.google.com/macros/s/AKfycbzkXVLXc-DV_HzSHHxpihNNz-Jb5ISeoGMRGn6WxV3zMs-9nUGuVUCz33rZEt3jqDH92w/exec // ID ชีต Google: 1ZjLX7Rwbsn9MC6PkJLvx0qE-3DylsxHAfXOPRYjcFqU // ชื่อชีต: ฟังก์ชัน StockData doGet(e) { // ตั้งค่าส่วนหัว CORS สำหรับคำขอการตรวจสอบล่วงหน้า var headers = { 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Methods': 'GET, POST, OPTIONS', 'Access-Control-Allow-Headers': 'Content-Type' }; const action = e.parameter.action; if (action === 'getMaterials') { return getMaterials(headers); } else if (action === 'getHistory') { return getHistory(headers); } else { return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'Invalid action' })) .setMimeType(ContentService.MimeType.JSON) .setHeaders(headers); } } function doPost(e) { // ตั้งค่าส่วนหัว CORS สำหรับคำขอการตรวจสอบก่อนการตรวจสอบ var headers = { 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Methods': 'GET, POST, OPTIONS', 'Access-Control-Allow-Headers': 'Content-Type' }; ลอง { // วิเคราะห์ข้อมูลขาเข้า const data = JSON.parse(e.postData.contents); const action = data.action; if (action === 'requisition') { return processRequisition(data, headers); } else { return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'Invalid action' })) .setMimeType(ContentService.MimeType.JSON) .setHeaders(headers); } } catch (error) { return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.toString() })) .setMimeType(ContentService.MimeType.JSON) .setHeaders(headers); } } ฟังก์ชัน getMaterials(ส่วนหัว) { ลอง { const ss = SpreadsheetApp.openById('1ZjLX7Rwbsn9MC6PkJLvx0qE-3DylsxHAfXOPRYjcFqU'); const sheet = ss.getSheetByName('StockData'); // รับข้อมูลทั้งหมดจากแผ่นงาน const data = sheet.getDataRange().getValues(); const headers = data[0]; // สร้างแผนที่ของวัสดุ const materialsMap = {}; // ข้ามแถวส่วนหัวสำหรับ (ให้ i = 1; i < data.length; i++) { const row = data[i]; const code = row[1]; // คอลัมน์ materialCode ถ้า (รหัส) { // ถ้ารหัสมีอยู่แล้ว ให้ใช้รายการล่าสุด (ซึ่งควรเป็นสต๊อกล่าสุด) ถ้า (!materialsMap[code] || new Date(row[6]) > new Date(materialsMap[code].timestamp)) { materialsMap[code] = { code: code, name: row[2], // คอลัมน์ materialName สต๊อก: row[5], // คอลัมน์ remainQuantity timestamp:แถว[6] // คอลัมน์ timestamp }; } } } // แปลงแผนที่เป็นอาร์เรย์และลบคุณสมบัติ timestamp const materials = Object.values(materialsMap).map(item => { const { timestamp, ...rest } = item; return rest; }); return ContentService.createTextOutput(JSON.stringify({ status: 'success', materials: materials })) .setMimeType(ContentService.MimeType.JSON) .setHeaders(headers); } catch (error) { return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.toString() })) .setMimeType(ContentService.MimeType.JSON) .setHeaders(headers); } } ฟังก์ชัน getHistory(ส่วนหัว) { ลอง { const ss = SpreadsheetApp.openById('1ZjLX7Rwbsn9MC6PkJLvx0qE-3DylsxHAfXOPRYjcFqU'); const sheet = ss.getSheetByName('StockData'); // รับข้อมูลทั้งหมดจากแผ่นงาน const data = sheet.getDataRange().getValues(); // สร้างอาร์เรย์ประวัติ const history = []; // ข้ามแถวส่วนหัวสำหรับ (let i = 1; i < data.length; i++) { const row = data[i]; if (row[4] > 0) { // รวมเฉพาะแถวที่มี requestQuantity เชิงบวกเท่านั้น history.push({ date: formatDate(row[6]), // timestamp column code: row[1], // materialCode column name: row[2], // materialName column quantity: row[4] // requestQuantity column }); } } // เรียงลำดับตามวันที่ (ใหม่สุดก่อน) history.sort((a, b) => new Date(b.date) - new Date(a.date)); return ContentService.createTextOutput(JSON.stringify({ status: 'success', history: history })) .setMimeType(ContentService.MimeType.JSON) .setHeaders(headers); } catch (error) { return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.toString() })) .setMimeType(ContentService.MimeType.JSON) .setHeaders(headers); } } ฟังก์ชัน processRequisition(data, headers) { ลอง { const ss = SpreadsheetApp.openById('1ZjLX7Rwbsn9MC6PkJLvx0qE-3DylsxHAfXOPRYjcFqU'); const sheet = ss.getSheetByName('StockData'); // ผนวกข้อมูลไปที่แผ่นงาน sheet.appendRow([ data.qrcode, data.materialCode, data.materialName, data.stockQuantity, data.requestQuantity, data.remainingQuantity, new Date() // Timestamp ]); return ContentService.createTextOutput(JSON.stringify({ status: 'success', message: 'Data saved successful' })) .setMimeType(ContentService.MimeType.JSON) .setHeaders(headers); } catch (error) { return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.toString() })) .setMimeType(ContentService.MimeType.JSON) .setHeaders(headers); } } function formatDate(date) { if (!date) return '';
 // return Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), 'yyyy-MM-dd');
//}

const SHEET_ID = '14XwCMLNYYllVSRdl1jqzU48H2ilFbAFgFuhbtuKHLbc';
const SHEET_NAME = 'StockData';
const HISTORY_SHEET = 'RequisitionLog';

// ช่วยตั้งค่า Header สำหรับ CORS
function setCorsHeaders(output) {
  return output
    .setMimeType(ContentService.MimeType.JSON)
    .setHeader("Access-Control-Allow-Origin", "*")
    .setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
    .setHeader("Access-Control-Allow-Headers", "Content-Type");
}

// GET → ใช้ดึงข้อมูลวัสดุทั้งหมด
function getMaterials() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  const rows = sheet.getDataRange().getValues();
  const materials = [];

  for (let i = 1; i < rows.length; i++) {
    materials.push({
      code: rows[i][0],
      name: rows[i][1],
      type: rows[i][2],
      stock: Number(rows[i][3]),
      unit: rows[i][4],
      reorder_point: Number(rows[i][5])
    });
  }

  return setCorsHeaders(ContentService.createTextOutput(JSON.stringify({
    status: "success",
    materials: materials
  })));
}

// GET → ใช้ดึงประวัติการเบิก
function getHistory() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(HISTORY_SHEET);
  const rows = sheet.getDataRange().getValues();
  const history = [];

  for (let i = 1; i < rows.length; i++) {
    history.push({
      code: rows[i][0],
      name: rows[i][1],
      qty: Number(rows[i][2]),
      date: rows[i][3],
      user: rows[i][4] || ""
    });
  }

  return setCorsHeaders(ContentService.createTextOutput(JSON.stringify({
    status: "success",
    history: history
  })));
}

// POST → ใช้เบิกของ และบันทึกใน History
function processRequisition(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const stockSheet = ss.getSheetByName(SHEET_NAME);
  const historySheet = ss.getSheetByName(HISTORY_SHEET);

  const code = data.code;
  const qty = Number(data.qty);
  const user = data.user || "ไม่ระบุ";

  const stockData = stockSheet.getDataRange().getValues();
  for (let i = 1; i < stockData.length; i++) {
    if (stockData[i][0] === code) {
      let currentQty = Number(stockData[i][3]);
      if (currentQty >= qty) {
        const newQty = currentQty - qty;
        stockSheet.getRange(i + 1, 4).setValue(newQty);
        stockSheet.getRange(i + 1, 7).setValue(new Date()); // คอลัมน์วันที่ล่าสุด
        const name = stockData[i][1];

        // บันทึกประวัติ
        historySheet.appendRow([code, name, qty, new Date(), user]);

        return setCorsHeaders(ContentService.createTextOutput(JSON.stringify({
          status: "success",
          message: `เบิก ${qty} หน่วยสำเร็จ`
        })));
      } else {
        return setCorsHeaders(ContentService.createTextOutput(JSON.stringify({
          status: "error",
          message: "จำนวนในคลังไม่พอ"
        })));
      }
    }
  }

  return setCorsHeaders(ContentService.createTextOutput(JSON.stringify({
    status: "error",
    message: "ไม่พบรหัสวัสดุ"
  })));
}

// รองรับ GET
function doGet(e) {
  const action = e.parameter.action;

  if (action === "getMaterials") return getMaterials();
  if (action === "getHistory") return getHistory();

  return setCorsHeaders(ContentService.createTextOutput(JSON.stringify({
    status: "error",
    message: "Invalid GET action"
  })));
}

// รองรับ POST
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.action === "requisition") {
      return processRequisition(data);
    } else {
      return setCorsHeaders(ContentService.createTextOutput(JSON.stringify({
        status: "error",
        message: "Invalid POST action"
      })));
    }
  } catch (err) {
    return setCorsHeaders(ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: err.toString()
    })));
  }
}


function doPost(e) {
  if (!e || !e.postData || !e.postData.contents) {
    return ContentService.createTextOutput("Error: No data received").setMimeType(ContentService.MimeType.TEXT);
  }

  var data = JSON.parse(e.postData.contents);
  var sheetStock = SpreadsheetApp.openById("14XwCMLNYYllVSRdl1jqzU48H2ilFbAFgFuhbtuKHLbc").getSheetByName("StockData");
  var sheetLog = SpreadsheetApp.openById("14XwCMLNYYllVSRdl1jqzU48H2ilFbAFgFuhbtuKHLbc").getSheetByName("RequisitionLog");

  // อัปเดต StockData
  sheetStock.appendRow([data.qrCode, data.materialCode, data.materialName, data.stockQty, data.remainingQty]);

  // บันทึกข้อมูลใบงานไปยัง RequisitionLog
  sheetLog.appendRow([data.workOrderNo, data.qrCode, data.materialCode, data.materialName, data.requisitionQty, new Date()]);

  return ContentService.createTextOutput("Data Saved Successfully").setMimeType(ContentService.MimeType.TEXT);
}

function doGet(e) {
  var sheet = SpreadsheetApp.openById("14XwCMLNYYllVSRdl1jqzU48H2ilFbAFgFuhbtuKHLbc").getSheetByName("StockData");
  var data = sheet.getDataRange().getValues();
  
  var jsonData = [];
  for (var i = 1; i < data.length; i++) {
    jsonData.push({
      qrCode: data[i][0],
      materialCode: data[i][1],
      materialName: data[i][2],
      stockQty: data[i][3],
      remainingQty: data[i][4]
    });
  }
  
  return ContentService.createTextOutput(JSON.stringify(jsonData))
    .setMimeType(ContentService.MimeType.JSON);
}


// Function to initialize the spreadsheet with sample data
function initializeSpreadsheet() {
  try {
    const ss = SpreadsheetApp.openById('14XwCMLNYYllVSRdl1jqzU48H2ilFbAFgFuhbtuKHLbc');
    let sheet = ss.getSheetByName('StockData');
    
    // Create the sheet if it doesn't exist
    if (!sheet) {
      sheet = ss.insertSheet('StockData');
    }
    
    // Clear existing data
    sheet.clear();
    
    // Add headers
    sheet.appendRow(['QRcode', 'รหัสวัสดุ', 'ชื่อวัสดุ', 'จำนวนในสต๊อก', 'จำนวนที่เบิก', 'จำนวนคงเหลือ', 'วันที่']);
    
    // Sample data
    const sampleData = [
      // Initial stock entries
      ['M001', 'M001', 'กระดาษ A4', 50, 0, 50, new Date(2023, 5, 1)],
      ['M002', 'M002', 'ปากกาลูกลื่น', 100, 0, 100, new Date(2023, 5, 1)],
      ['M003', 'M003', 'แฟ้มเอกสาร', 30, 0, 30, new Date(2023, 5, 1)],
      ['M004', 'M004', 'ลวดเย็บกระดาษ', 40, 0, 40, new Date(2023, 5, 1)],
      ['M005', 'M005', 'คลิปหนีบกระดาษ', 60, 0, 60, new Date(2023, 5, 1)],
      ['M006', 'M006', 'กรรไกร', 15, 0, 15, new Date(2023, 5, 1)],
      ['M007', 'M007', 'เทปกาว', 25, 0, 25, new Date(2023, 5, 1)],
      ['M008', 'M008', 'ยางลบ', 45, 0, 45, new Date(2023, 5, 1)],
      ['M009', 'M009', 'ดินสอ', 80, 0, 80, new Date(2023, 5, 1)],
      ['M010', 'M010', 'ไม้บรรทัด', 20, 0, 20, new Date(2023, 5, 1)],
      
      // Requisition entries
      ['M001', 'M001', 'กระดาษ A4', 50, 5, 45, new Date(2023, 5, 5)],
      ['M002', 'M002', 'ปากกาลูกลื่น', 100, 10, 90, new Date(2023, 5, 6)],
      ['M003', 'M003', 'แฟ้มเอกสาร', 30, 3, 27, new Date(2023, 5, 7)],
      ['M001', 'M001', 'กระดาษ A4', 45, 10, 35, new Date(2023, 5, 10)],
      ['M004', 'M004', 'ลวดเย็บกระดาษ', 40, 5, 35, new Date(2023, 5, 12)],
      ['M005', 'M005', 'คลิปหนีบกระดาษ', 60, 15, 45, new Date(2023, 5, 15)],
      ['M002', 'M002', 'ปากกาลูกลื่น', 90, 20, 70, new Date(2023, 5, 18)],
      ['M006', 'M006', 'กรรไกร', 15, 2, 13, new Date(2023, 5, 20)],
      ['M007', 'M007', 'เทปกาว', 25, 5, 20, new Date(2023, 5, 22)],
      ['M001', 'M001', 'กระดาษ A4', 35, 5, 30, new Date(2023, 5, 25)]
    ];
    
    // Add sample data
    sampleData.forEach(row => {
      sheet.appendRow(row);
    });
    
    return "Spreadsheet initialized with sample data";
  } catch (error) {
    return "Error:" + ข้อผิดพลาด.toString(); } }
