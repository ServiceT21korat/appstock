// รหัสสคริปต์ของ Google Apps สำหรับระบบการจัดหาวัสดุ // ปรับใช้สิ่งนี้เป็นแอปเว็บด้วย URL: // https://script.google.com/macros/s/AKfycbzkXVLXc-DV_HzSHHxpihNNz-Jb5ISeoGMRGn6WxV3zMs-9nUGuVUCz33rZEt3jqDH92w/exec // ID ชีต Google: 1ZjLX7Rwbsn9MC6PkJLvx0qE-3DylsxHAfXOPRYjcFqU // ชื่อชีต: ฟังก์ชัน StockData doGet(e) { // ตั้งค่าส่วนหัว CORS สำหรับคำขอการตรวจสอบล่วงหน้า var headers = { 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Methods': 'GET, POST, OPTIONS', 'Access-Control-Allow-Headers': 'Content-Type' }; const action = e.parameter.action; if (action === 'getMaterials') { return getMaterials(headers); } else if (action === 'getHistory') { return getHistory(headers); } else { return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'Invalid action' })) .setMimeType(ContentService.MimeType.JSON) .setHeaders(headers); } } function doPost(e) { // ตั้งค่าส่วนหัว CORS สำหรับคำขอการตรวจสอบก่อนการตรวจสอบ var headers = { 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Methods': 'GET, POST, OPTIONS', 'Access-Control-Allow-Headers': 'Content-Type' }; ลอง { // วิเคราะห์ข้อมูลขาเข้า const data = JSON.parse(e.postData.contents); const action = data.action; if (action === 'requisition') { return processRequisition(data, headers); } else { return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'Invalid action' })) .setMimeType(ContentService.MimeType.JSON) .setHeaders(headers); } } catch (error) { return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.toString() })) .setMimeType(ContentService.MimeType.JSON) .setHeaders(headers); } } ฟังก์ชัน getMaterials(ส่วนหัว) { ลอง { const ss = SpreadsheetApp.openById('1ZjLX7Rwbsn9MC6PkJLvx0qE-3DylsxHAfXOPRYjcFqU'); const sheet = ss.getSheetByName('StockData'); // รับข้อมูลทั้งหมดจากแผ่นงาน const data = sheet.getDataRange().getValues(); const headers = data[0]; // สร้างแผนที่ของวัสดุ const materialsMap = {}; // ข้ามแถวส่วนหัวสำหรับ (ให้ i = 1; i < data.length; i++) { const row = data[i]; const code = row[1]; // คอลัมน์ materialCode ถ้า (รหัส) { // ถ้ารหัสมีอยู่แล้ว ให้ใช้รายการล่าสุด (ซึ่งควรเป็นสต๊อกล่าสุด) ถ้า (!materialsMap[code] || new Date(row[6]) > new Date(materialsMap[code].timestamp)) { materialsMap[code] = { code: code, name: row[2], // คอลัมน์ materialName สต๊อก: row[5], // คอลัมน์ remainQuantity timestamp:แถว[6] // คอลัมน์ timestamp }; } } } // แปลงแผนที่เป็นอาร์เรย์และลบคุณสมบัติ timestamp const materials = Object.values(materialsMap).map(item => { const { timestamp, ...rest } = item; return rest; }); return ContentService.createTextOutput(JSON.stringify({ status: 'success', materials: materials })) .setMimeType(ContentService.MimeType.JSON) .setHeaders(headers); } catch (error) { return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.toString() })) .setMimeType(ContentService.MimeType.JSON) .setHeaders(headers); } } ฟังก์ชัน getHistory(ส่วนหัว) { ลอง { const ss = SpreadsheetApp.openById('1ZjLX7Rwbsn9MC6PkJLvx0qE-3DylsxHAfXOPRYjcFqU'); const sheet = ss.getSheetByName('StockData'); // รับข้อมูลทั้งหมดจากแผ่นงาน const data = sheet.getDataRange().getValues(); // สร้างอาร์เรย์ประวัติ const history = []; // ข้ามแถวส่วนหัวสำหรับ (let i = 1; i < data.length; i++) { const row = data[i]; if (row[4] > 0) { // รวมเฉพาะแถวที่มี requestQuantity เชิงบวกเท่านั้น history.push({ date: formatDate(row[6]), // timestamp column code: row[1], // materialCode column name: row[2], // materialName column quantity: row[4] // requestQuantity column }); } } // เรียงลำดับตามวันที่ (ใหม่สุดก่อน) history.sort((a, b) => new Date(b.date) - new Date(a.date)); return ContentService.createTextOutput(JSON.stringify({ status: 'success', history: history })) .setMimeType(ContentService.MimeType.JSON) .setHeaders(headers); } catch (error) { return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.toString() })) .setMimeType(ContentService.MimeType.JSON) .setHeaders(headers); } } ฟังก์ชัน processRequisition(data, headers) { ลอง { const ss = SpreadsheetApp.openById('1ZjLX7Rwbsn9MC6PkJLvx0qE-3DylsxHAfXOPRYjcFqU'); const sheet = ss.getSheetByName('StockData'); // ผนวกข้อมูลไปที่แผ่นงาน sheet.appendRow([ data.qrcode, data.materialCode, data.materialName, data.stockQuantity, data.requestQuantity, data.remainingQuantity, new Date() // Timestamp ]); return ContentService.createTextOutput(JSON.stringify({ status: 'success', message: 'Data saved successful' })) .setMimeType(ContentService.MimeType.JSON) .setHeaders(headers); } catch (error) { return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.toString() })) .setMimeType(ContentService.MimeType.JSON) .setHeaders(headers); } } function formatDate(date) { if (!date) return '';
 // return Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), 'yyyy-MM-dd');
//}

// Function to initialize the spreadsheet with sample data
function initializeSpreadsheet() {
  try {
    const ss = SpreadsheetApp.openById('1ZjLX7Rwbsn9MC6PkJLvx0qE-3DylsxHAfXOPRYjcFqU');
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