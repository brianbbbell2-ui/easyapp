/**
 * Google Apps Script Backend for Repair Request System V5
 * Final Version with 100% Accurate PDF Layout (Updated based on FM-ASM-01 Google Doc)
 */

const SPREADSHEET_ID = '1soSnnMGnZwnqou3SdeOkxULWFBpd_ygLuKWAE5nbyXo'; 
const PDF_FOLDER_ID = '1vm8eALNA5cghIPcxZ62ll79D9_8oXSuM';      
const SERVER_TEMP_FOLDER_ID = '1hz8T6kocbQvq7xX_oKI-jrOcP88y5rFU'; 
const LINE_CHANNEL_ACCESS_TOKEN = '/LG/Ffu2KMZHUC0nzQbNLkZXw4DU+jvdRXqkFJ9EYhk8NsGFMAWoip3Hp3FwgWzGWYDWxE5i1JTOun6XhN+XjqT7h/55PxZ5RcZYv410iQkbVNmPPl7usoT58fm05RFERytNGH88g3ffZnO7aI5nMQdB04t89/1O/w1cDnyilFU='; 
const LINE_GROUP_ID = 'Ufa4a2300682877df44751bf53cce0e7b'; 
const DELIVERY_PIC_FOLDER_ID = '1yljdw045Cn8JY--ZqRdqnsgZnByLRojc';

function doGet(e) {
  try {
    return HtmlService.createTemplateFromFile('Index')
        .evaluate()
        .setTitle('EasyAPP EIO-DMK')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    return HtmlService.createHtmlOutput("Error: " + err.toString());
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getSS() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function login(username, password) {
  try {
    const ss = getSS();
    const sheet = ss.getSheetByName('Users');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][4]).trim() === String(username).trim() && String(data[i][5]).trim() === String(password).trim()) {
        return { success: true, user: { id: data[i][0], name: data[i][1], role: data[i][6], phone: data[i][2], email: data[i][3], username: data[i][4] } };
      }
    }
    return { success: false, message: 'Username หรือ Password ไม่ถูกต้อง' };
  } catch (e) { return { success: false, message: e.toString() }; }
}

function getAssetDataById(id) {
  try {
    const ss = getSS();
    const sheet = ss.getSheetByName('Data_RepairRequests');
    if (!sheet) return { success: false, message: 'ไม่พบชีต Data_RepairRequests' };
    const data = sheet.getDataRange().getValues();
    const searchId = String(id).trim().toLowerCase();
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (String(row[0]).trim().toLowerCase() === searchId) {
        let expireDate = "-";
        if (row[9] instanceof Date) {
          expireDate = Utilities.formatDate(row[9], "GMT+7", "dd/MM/yyyy");
        } else if (row[9]) { 
          expireDate = row[9]; 
        }
        const assetNameFull = [row[5], row[6], row[7]].filter(Boolean).join(' / ');
        return {
          success: true,
          data: { 
            id: row[0],
            company: row[1],
            costCenter: row[10],
            department: row[2],
            location: row[3],
            codeAsset: row[4],
            assetName: assetNameFull,
            serialNo: row[8] || "",
            warrantyExp: expireDate
          }
        };
      }
    }
    return { success: false, message: 'ไม่พบข้อมูลรหัส ID นี้ในระบบ' };
  } catch (e) { return { success: false, message: 'เกิดข้อผิดพลาด: ' + e.toString() }; }
}

function submitRepairRequestV3(formData) {
  try {
    const ss = getSS();
    if (formData.id && formData.costCenter) {
      const dataSheet = ss.getSheetByName('Data_RepairRequests');
      if (dataSheet) {
        const dataValues = dataSheet.getDataRange().getValues();
        const searchId = String(formData.id).trim().toLowerCase();
        for (let i = 1; i < dataValues.length; i++) {
          if (String(dataValues[i][0]).trim().toLowerCase() === searchId) {
            if (String(dataValues[i][10]).trim() !== String(formData.costCenter).trim()) {
              dataSheet.getRange(i + 1, 11).setValue(formData.costCenter);
            }
            break;
          }
        }
      }
    }
    const sheet = ss.getSheetByName('RepairRequests');
    const now = new Date();
    const dateReceived = Utilities.formatDate(now, "GMT+7", "dd/MM/yyyy HH:mm");
    let jobNo = formData.jobNo;
    if (!jobNo || jobNo.includes('XXXX')) {
      const lastRow = sheet.getLastRow();
      const count = lastRow;
      const jobPrefix = "REQ" + now.getFullYear() + String(now.getMonth() + 1).padStart(2, '0');
      jobNo = jobPrefix + "-" + String(count).padStart(4, '0');
    }
    const pdfUrl = generateRepairPDF_V3(jobNo, formData);
    sheet.appendRow([
      jobNo, dateReceived, formData.id, formData.company, formData.costCenter, 
      formData.department, formData.location, formData.codeAsset, formData.assetName,
      formData.serialNo, formData.warrantyExp, formData.issueDetail, formData.requester,
      formData.contactInfo, formData.acknowledgedBy, formData.status,
      (formData.status === 'ไม่ส่งซ่อมเนื่องจาก' ? formData.fix2Detail : formData.fix1Detail),
      formData.inspectedBy, formData.fix1Detail, formData.fix2Detail, pdfUrl
    ]);
    return { success: true, jobNo: jobNo, pdfUrl: pdfUrl };
  } catch (e) { return { success: false, message: e.toString() }; }
}

function getRepairHistoryV3() {
  try {
    const ss = getSS();
    const sheet = ss.getSheetByName('RepairRequests');
    if (!sheet) return { success: false, message: 'ไม่พบชีต RepairRequests' };
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: true, data: [] };
    const numRows = Math.min(lastRow - 1, 50);
    const startRow = lastRow - numRows + 1;
    const data = sheet.getRange(startRow, 1, numRows, 21).getValues();
    const history = [];
    for (let i = data.length - 1; i >= 0; i--) {
      let dateVal = data[i][1];
      let dateStr = dateVal instanceof Date ? Utilities.formatDate(dateVal, "GMT+7", "dd/MM/yyyy HH:mm") : String(dateVal);
      history.push({ 
        jobNo: data[i][0], date: dateStr, assetId: data[i][7] || data[i][2],
        assetName: data[i][8], issueDetail: data[i][11], status: data[i][15], pdfUrl: data[i][20] 
      });
    }
    return { success: true, data: history };
  } catch (e) { return { success: false, message: e.toString() }; }
}

/**
 * 100% Accurate PDF Generation based on FM-ASM-01 Google Doc
 */
function generateRepairPDF_V3(jobNo, formData) {
  try {
    const folder = DriveApp.getFolderById(PDF_FOLDER_ID);
    const dateNow = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy");
    const html = `
    <!DOCTYPE html>
    <html>
      <head>
        <link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@400;700&display=swap" rel="stylesheet">
        <style>
          @page { size: A4; margin: 10mm; }
          body { font-family: 'Sarabun', sans-serif; font-size: 10pt; color: #000; line-height: 1.2; margin: 0; padding: 0; }
          .container { width: 100%; box-sizing: border-box; }
          
          /* Header Style */
          .header-table { width: 100%; border-collapse: collapse; margin-bottom: 5px; }
          .header-table td { border: none; padding: 0; vertical-align: top; }
          .logo-container { width: 35%; }
          .logo-text { font-size: 18pt; font-weight: bold; letter-spacing: 1px; display: block; }
          .crown-icon { font-size: 12pt; display: block; margin-bottom: -3px; }
          .title-container { width: 35%; text-align: center; border-left: 1px solid #000; padding: 0 10px; }
          .title-text { font-size: 12pt; font-weight: bold; }
          .subtitle-text { font-size: 9pt; }
          .job-container { width: 30%; border-left: 1px solid #000; padding-left: 15px; }
          .job-row { font-size: 10pt; margin-bottom: 3px; }
          .underline { border-bottom: 1px solid #000; display: inline-block; min-width: 90px; text-align: center; font-weight: bold; }
          
          /* Section Title */
          .section-title { background-color: #f2f2f2; font-weight: bold; text-align: center; border: 1px solid #000; padding: 4px; font-size: 10pt; margin-top: 5px; }
          
          /* Tables */
          table.main-table { width: 100%; border-collapse: collapse; table-layout: fixed; }
          table.main-table td, table.main-table th { border: 1px solid #000; padding: 5px 8px; vertical-align: top; }
          .bg-gray { background-color: #f2f2f2; font-weight: bold; text-align: center; font-size: 9pt; }
          .center { text-align: center; }
          .bold { font-weight: bold; }
          
          /* Signature */
          .sig-table { width: 100%; border-collapse: collapse; table-layout: fixed; margin-top: -1px; }
          .sig-table td { border: 1px solid #000; height: 100px; vertical-align: top; padding: 10px; }
          .sig-label { font-weight: bold; text-align: center; display: block; margin-bottom: 40px; font-size: 10pt; }
          .sig-line { border-bottom: 1px dotted #000; width: 80%; margin: 0 auto 5px; height: 15px; }
          .sig-name { text-align: center; font-size: 10pt; }
          
          /* Checkbox */
          .checkbox { display: inline-block; width: 13px; height: 13px; border: 1px solid #000; margin-right: 8px; text-align: center; line-height: 12px; font-size: 10pt; font-weight: bold; vertical-align: middle; }
          .check-row { margin-bottom: 8px; display: flex; align-items: center; font-size: 10pt; }
          .dot-line { border-bottom: 1px dotted #000; flex-grow: 1; margin-left: 5px; font-weight: bold; min-height: 18px; }
          
          .footer-notes { font-size: 8pt; margin: 5px 0; line-height: 1.2; }
        </style>
      </head>
      <body>
        <div class="container">
          <!-- HEADER -->
          <table class="header-table">
            <tr>
              <td class="logo-container">
                <span class="crown-icon">♔</span>
                <span class="logo-text">KING POWER</span>
              </td>
              <td class="title-container">
                <div class="title-text">แบบฟอร์มแจ้งซ่อมทรัพย์สิน</div>
                <div class="subtitle-text">FM-ASM-01 Rev.0</div>
              </td>
              <td class="job-container">
                <div class="job-row">Job No. <span class="underline">${jobNo}</span></div>
                <div class="job-row">Date Received <span class="underline">${dateNow}</span></div>
              </td>
            </tr>
          </table>

          <!-- ASSET DETAILS -->
          <div class="section-title">รายละเอียดทรัพย์สิน</div>
          <table class="main-table">
            <tr>
              <td width="15%">บริษัท</td>
              <td width="35%" class="bold">${formData.company || 'King Power'}</td>
              <td width="15%">ฝ่าย / แผนก</td>
              <td width="35%" class="bold">${formData.department}</td>
            </tr>
            <tr>
              <td>Cost Center</td>
              <td class="bold">${formData.costCenter}</td>
              <td>สถานที่</td>
              <td class="bold">${formData.location}</td>
            </tr>
          </table>

          <!-- ITEM TABLE -->
          <table class="main-table" style="margin-top: -1px;">
            <tr class="bg-gray">
              <th width="5%">ที่</th>
              <th width="15%">รหัสทรัพย์สิน</th>
              <th width="30%">รายการทรัพย์สิน (ชื่อ/ยี่ห้อ/รุ่น/ขนาด)</th>
              <th width="15%">Serial No.</th>
              <th width="25%">รายละเอียดที่ขอซ่อม อาการเสีย / ชำรุด</th>
              <th width="10%">การรับประกัน สิ้นสุดวันที่</th>
            </tr>
            <tr>
              <td class="center">1</td>
              <td class="center">${formData.codeAsset}</td>
              <td><span class="bold">${formData.assetName}</span><br><span style="font-size: 8pt;">ID: ${formData.id}</span></td>
              <td class="center">${formData.serialNo || '-'}</td>
              <td>${formData.issueDetail}</td>
              <td class="center">${formData.warrantyExp || '-'}</td>
            </tr>
            ${[2,3,4,5].map(i => `<tr><td class="center" style="height: 25px;">${i}</td><td></td><td></td><td></td><td></td><td></td></tr>`).join('')}
          </table>

          <div class="footer-notes">
            สามารถแนบเอกสารเพิ่มเติมได้ เช่น มีรายการขอซ่อมมากกว่า 5 รายการ, ภาพทรัพย์สิน, เอกสารแจ้งซ่อมของหน่วยงานอื่น<br>
            สอบถามข้อมูลทรัพย์สินพื้นที่ รางน้ำ, เชียงใหม่ โทร 1416, 1440, 1439 / ภูเก็ต โทร 1457<br>
            สุวรรณภูมิ, บางบ่อ, บ้านวารีฯ, พัทยา, อู่ตะเภา, ดอนเมือง โทร 1478, 1477, 1488, 1474
          </div>

          <!-- SIGNATURES -->
          <table class="sig-table">
            <tr>
              <td>
                <span class="sig-label">ผู้แจ้งซ่อมทรัพย์สิน</span>
                <div class="sig-name">
                  <div class="sig-line"></div>
                  ( <span class="bold">${formData.requester}</span> )<br>
                  <div style="font-size: 9pt; margin-top: 10px;">เบอร์ติดต่อ / อีเมล์ <span style="border-bottom: 1px dotted #000; font-weight: bold;">${formData.contactInfo || '........................................'}</span></div>
                </div>
              </td>
              <td>
                <span class="sig-label">รับทราบโดยผู้จัดการฝ่ายขึ้นไป</span>
                <div class="sig-name">
                  <div class="sig-line"></div>
                  ( ............................................................ )<br>
                  <div style="font-size: 9pt; margin-top: 10px;">รับทราบโดย</div>
                </div>
              </td>
            </tr>
          </table>

          <!-- UNIT CHECK -->
          <div class="section-title">สำหรับหน่วยงานตรวจสอบทรัพย์สิน เช่น ส่วนงาน ITO, ส่วนงาน ITI, ฝ่าย AMF, ฝ่าย ASM</div>
          <table class="main-table" style="margin-top: -1px;">
            <tr class="bg-gray">
              <th width="60%">ผลการตรวจสอบ</th>
              <th width="40%">ตรวจสอบโดย</th>
            </tr>
            <tr>
              <td style="height: 100px;">
                <div class="check-row"><div class="checkbox">${formData.status === 'แก้ไขได้' ? '/' : ''}</div> แก้ไขได้</div>
                <div class="check-row">
                  <div class="checkbox">${formData.status === 'ส่งซ่อมรายละเอียด' ? '/' : ''}</div> 
                  ส่งซ่อม รายละเอียด <span class="dot-line">${formData.status === 'ส่งซ่อมรายละเอียด' ? formData.fix1Detail : ''}</span>
                </div>
                <div class="check-row">
                  <div class="checkbox">${formData.status === 'ไม่ส่งซ่อมเนื่องจาก' ? '/' : ''}</div> 
                  ไม่ส่งซ่อม เนื่องจาก <span class="dot-line">${formData.status === 'ไม่ส่งซ่อมเนื่องจาก' ? formData.fix2Detail : ''}</span>
                </div>
              </td>
              <td class="sig-name">
                <div style="height: 30px;"></div>
                <div class="sig-line"></div>
                ( <span class="bold">${formData.inspectedBy || '................................'}</span> )<br>
                <div style="margin-top: 10px;">ฝ่าย <span style="border-bottom: 1px dotted #000; min-width: 80px; display: inline-block; font-weight: bold;">EIO</span></div>
              </td>
            </tr>
          </table>

          <!-- ASSET MANAGEMENT -->
          <div class="section-title">สำหรับฝ่ายบริหารทรัพย์สิน</div>
          <table class="main-table" style="margin-top: -1px;">
            <tr class="bg-gray">
              <th width="50%">กรณีส่งซ่อมกับ Supplier</th>
              <th width="50%">ข้อมูล Supplier</th>
            </tr>
            <tr>
              <td style="height: 110px;">
                <div style="font-size: 9pt; margin-bottom: 5px; font-weight: bold;">ค่าใช้จ่าย</div>
                <div class="check-row"><div class="checkbox"></div> ไม่มีค่าใช้จ่าย</div>
                <div class="check-row"><div class="checkbox"></div> มีค่าใช้จ่าย .......................................... บาท</div>
                <div style="font-size: 9pt; margin-bottom: 5px;">ใบเสนอราคาเลขที่ ...........................................................</div>
                <div style="font-size: 9pt;">เลขที่ P/O ........................................................................</div>
              </td>
              <td>
                <div style="font-size: 9pt; margin-bottom: 5px;">ชื่อบริษัท : ....................................................................</div>
                <div style="font-size: 9pt; margin-bottom: 5px;">ชื่อและเบอร์ผู้ติดต่อ : .....................................................</div>
                <div style="font-size: 9pt; margin-bottom: 5px;">วันที่ส่งซ่อม : ...............................................................</div>
                <div style="font-size: 9pt; margin-bottom: 5px;">วันที่รับคืน : .................................................................</div>
                <div style="font-size: 9pt;">หมายเหตุ : ...................................................................</div>
              </td>
            </tr>
            <tr class="bg-gray">
              <th>ผู้ดำเนินการ</th>
              <th>รับทราบโดยผู้จัดการฝ่ายขึ้นไป</th>
            </tr>
            <tr>
              <td class="sig-name" style="height: 80px;">
                <div style="height: 20px;"></div>
                <div class="sig-line"></div>
                ( ............................................................ )
              </td>
              <td class="sig-name" style="height: 80px;">
                <div style="height: 20px;"></div>
                <div class="sig-line"></div>
                ( ............................................................ )
              </td>
            </tr>
          </table>
        </div>
      </body>
    </html>
    `;
    const pdfFile = folder.createFile(Utilities.newBlob(html, 'text/html').getAs('application/pdf')).setName(jobNo + ".pdf");
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return pdfFile.getUrl();
  } catch (e) { return ""; }
}

/**
 * ===== DELIVERY FUNCTIONS (แก้ไข: field names ตรงกับ DeliveryUI.html + เพิ่ม submitDeliveryRequest) =====
 */

/**
 * บันทึกรายการส่งเอกสาร/พัสดุใหม่
 * Column layout: A(1)=dateTime, B(2)=type, C(3)=contactName, D(4)=department,
 *                E(5)=phone, F(6)=address, G(7)=sender, H(8)=details,
 *                I(9)=completeTime, J(10)=imageUrl, K(11)=signatureUrl, L(12)=status, M(13)=jobNo
 */
function submitDeliveryRequest(formData) {
  try {
    const ss = getSS();
    const sheet = ss.getSheetByName('DeliveryRecords');
    if (!sheet) return { success: false, message: 'ไม่พบชีต DeliveryRecords' };
    const now = new Date();
    const dateTime = formData.dateTime || Utilities.formatDate(now, 'GMT+7', 'dd/MM/yyyy HH:mm');
    const lastRow = sheet.getLastRow();
    const jobNo = 'DLV' + Utilities.formatDate(now, 'GMT+7', 'yyyyMM') + '-' + String(lastRow).padStart(4, '0');
    let imageUrl = '';
    if (formData.imageData && formData.imageData.length > 100) {
      const imgFileName = 'DLV_' + Utilities.formatDate(now, 'GMT+7', 'yyyyMMdd_HHmmss') + '.jpg';
      imageUrl = uploadImageToFolder(formData.imageData, imgFileName, DELIVERY_PIC_FOLDER_ID);
    }
    sheet.appendRow([
      dateTime,
      formData.type || '',
      formData.contactName || '',
      formData.department || '',
      formData.phone || '',
      formData.address || '',
      formData.sender || '',
      formData.details || '',
      '',          // completeTime (col I)
      imageUrl,    // imageUrl (col J)
      '',          // signatureUrl (col K)
      'รอดำเนินการ', // status (col L)
      jobNo        // jobNo (col M)
    ]);
    return { success: true, jobNo: jobNo };
  } catch (e) { return { success: false, message: e.toString() }; }
}

/**
 * ดึงรายการที่ยังไม่ได้ส่ง (status != 'ส่งเรียบร้อย')
 * ส่งกลับ field ชื่อที่ตรงกับ DeliveryUI.html
 */
function getPendingDeliveries() {
  try {
    const ss = getSS();
    const sheet = ss.getSheetByName('DeliveryRecords');
    if (!sheet) return { success: false, message: 'ไม่พบชีต DeliveryRecords' };
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: true, data: [] };
    const data = sheet.getDataRange().getValues();
    const pending = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const status = String(row[11] || '').trim();
      if (status !== 'ส่งเรียบร้อย') {
        pending.push({
          rowIndex: i + 1,
          jobNo: String(row[12] || ('ROW' + (i + 1))),
          type: String(row[1] || ''),
          contactName: String(row[2] || ''),
          department: String(row[3] || ''),
          phone: String(row[4] || ''),
          address: String(row[5] || ''),
          sender: String(row[6] || ''),
          details: String(row[7] || ''),
          dateTime: row[0] instanceof Date ? Utilities.formatDate(row[0], 'GMT+7', 'dd/MM/yyyy HH:mm') : String(row[0] || ''),
          imageUrl: String(row[9] || ''),
          status: status || 'รอดำเนินการ'
        });
      }
    }
    return { success: true, data: pending };
  } catch (e) { return { success: false, message: e.toString() }; }
}

/**
 * ดึงประวัติการส่งที่เสร็จแล้ว (status = 'ส่งเรียบร้อย')
 * ส่งกลับ field ชื่อที่ตรงกับ DeliveryUI.html
 */
function getDeliveryHistory() {
  try {
    const ss = getSS();
    const sheet = ss.getSheetByName('DeliveryRecords');
    if (!sheet) return { success: false, message: 'ไม่พบชีต DeliveryRecords' };
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: true, data: [] };
    const data = sheet.getDataRange().getValues();
    const history = [];
    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      const status = String(row[11] || '').trim();
      if (status === 'ส่งเรียบร้อย') {
        history.push({
          rowIndex: i + 1,
          jobNo: String(row[12] || ('ROW' + (i + 1))),
          type: String(row[1] || ''),
          contactName: String(row[2] || ''),
          department: String(row[3] || ''),
          phone: String(row[4] || ''),
          address: String(row[5] || ''),
          sender: String(row[6] || ''),
          details: String(row[7] || ''),
          dateTime: row[0] instanceof Date ? Utilities.formatDate(row[0], 'GMT+7', 'dd/MM/yyyy HH:mm') : String(row[0] || ''),
          deliveryTime: row[8] instanceof Date ? Utilities.formatDate(row[8], 'GMT+7', 'dd/MM/yyyy HH:mm') : String(row[8] || ''),
          imageUrl: String(row[9] || ''),
          signatureUrl: String(row[10] || ''),
          status: status
        });
      }
    }
    return { success: true, data: history };
  } catch (e) { return { success: false, message: e.toString() }; }
}

/**
 * บันทึกการส่งเสร็จสิ้น พร้อมลายเซ็นผู้รับ
 * รับ jobNo (string) แทน rowIndex เพื่อความปลอดภัย
 */
function completeDelivery(jobNo, signatureData) {
  try {
    const ss = getSS();
    const sheet = ss.getSheetByName('DeliveryRecords');
    if (!sheet) return { success: false, message: 'ไม่พบชีต DeliveryRecords' };
    const data = sheet.getDataRange().getValues();
    let targetRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][12] || '').trim() === String(jobNo).trim()) {
        targetRow = i + 1;
        break;
      }
    }
    if (targetRow === -1) return { success: false, message: 'ไม่พบรายการ: ' + jobNo };
    const now = new Date();
    const timeStr = Utilities.formatDate(now, 'GMT+7', 'dd/MM/yyyy HH:mm');
    const sigFileName = 'SIG_' + Utilities.formatDate(now, 'GMT+7', 'yyyyMMdd_HHmmss') + '.png';
    const sigUrl = uploadImageToFolder(signatureData, sigFileName, DELIVERY_PIC_FOLDER_ID);
    sheet.getRange(targetRow, 9).setValue(timeStr);   // completeTime
    sheet.getRange(targetRow, 11).setValue(sigUrl);   // signatureUrl
    sheet.getRange(targetRow, 12).setValue('ส่งเรียบร้อย'); // status
    return { success: true, message: 'บันทึกการส่งเรียบร้อย' };
  } catch (e) { return { success: false, message: e.toString() }; }
}

function submitServerTemp(formData) {
  try {
    const ss = getSS();
    const sheet = ss.getSheetByName('ServerTempDT1');
    const now = new Date();
    const recordTime = Utilities.formatDate(now, "GMT+7", "HH:mm:ss");
    let imageTempUrl = "";
    if (formData.imageDataTemp && formData.imageDataTemp.length > 100) {
      const tempFileName = "Temp_" + Utilities.formatDate(now, "GMT+7", "yyyyMMdd_HHmmss") + ".jpg";
      imageTempUrl = uploadImageToFolder(formData.imageDataTemp, tempFileName, SERVER_TEMP_FOLDER_ID);
    }
    let imageWifiUrl = "";
    if (formData.imageDataWifi && formData.imageDataWifi.length > 100) {
      const wifiFileName = "WIFI_" + Utilities.formatDate(now, "GMT+7", "yyyyMMdd_HHmmss") + ".jpg";
      imageWifiUrl = uploadImageToFolder(formData.imageDataWifi, wifiFileName, SERVER_TEMP_FOLDER_ID);
    }
    sheet.appendRow([
      formData.date, formData.shift, formData.temperature, formData.status, recordTime,
      formData.recorder, imageTempUrl, formData.wifiStatus, imageWifiUrl,
      formData.note || "", formData.tempNote || ""
    ]);
    try {
      const lineMsg = `ตรวจบันทึกอุณหภูมิ Server DT1\n` +
        `วันที่/เวลา: ${formData.date}\nShift: ${formData.shift}\nอุณหภูมิ: ${formData.temperature} °C\nสถานะ: ${formData.status}\n` +
        `ถ่ายรูปหน้าจอนิเตอร์อุณหภูมิ\nตรวจสอบสัญญาณ WIFI\nWIFI: ${formData.wifiStatus}\n` +
        `หมายเหตุ: ${formData.note || formData.tempNote || '-'}\nถ่ายรูปหน้าจอการเชื่อมต่อ WIFI`;
      sendLineNotification(lineMsg);
    } catch (lineErr) {}
    return { success: true, message: 'บันทึกข้อมูลสำเร็จ' };
  } catch (e) { return { success: false, message: e.toString() }; }
}

function getServerTempHistory() {
  try {
    const ss = getSS();
    const sheet = ss.getSheetByName('ServerTempDT1');
    if (!sheet) return { success: false, message: 'ไม่พบชีต ServerTempDT1' };
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: true, data: [] };
    const numRows = Math.min(lastRow - 1, 50);
    const startRow = lastRow - numRows + 1;
    const data = sheet.getRange(startRow, 1, numRows, 11).getValues();
    const history = [];
    for (let i = data.length - 1; i >= 0; i--) {
      const row = data[i];
      history.push({
        date: String(row[0] || ''), shift: String(row[1] || ''), temperature: String(row[2] || ''),
        status: String(row[3] || ''), recordTime: String(row[4] || ''), recorder: String(row[5] || ''),
        imageTemp: String(row[6] || ''), wifiStatus: String(row[7] || ''), imageWifi: String(row[8] || ''),
        note: String(row[9] || ''), tempNote: String(row[10] || '')
      });
    }
    return { success: true, data: history };
  } catch (e) { return { success: false, message: e.toString() }; }
}

function sendLineNotification(message) {
  const url = 'https://api.line.me/v2/bot/message/push';
  const payload = JSON.stringify({ to: LINE_GROUP_ID, messages: [{ type: 'text', text: message }] });
  const options = {
    method: 'post', contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + LINE_CHANNEL_ACCESS_TOKEN },
    payload: payload, muteHttpExceptions: true
  };
  return UrlFetchApp.fetch(url, options);
}

function getUsers() {
  try {
    const ss = getSS();
    const sheet = ss.getSheetByName('Users');
    const data = sheet.getDataRange().getValues();
    const users = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;
      users.push({ id: row[0], name: row[1], phone: row[2], email: row[3], username: row[4], password: row[5], role: row[6] });
    }
    return { success: true, data: users };
  } catch (e) { return { success: false, message: e.toString() }; }
}

function addUser(formData) {
  try {
    if (formData.requestorRole !== 'Admin') return { success: false, message: 'ไม่มีสิทธิ์เพิ่มผู้ใช้งาน' };
    const ss = getSS();
    const sheet = ss.getSheetByName('Users');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][4]).trim().toLowerCase() === String(formData.username).trim().toLowerCase()) {
        return { success: false, message: 'Username นี้ถูกใช้งานแล้ว' };
      }
    }
    const newId = 'USR' + Utilities.formatDate(new Date(), "GMT+7", "yyyyMMddHHmmss");
    sheet.appendRow([newId, formData.name, formData.phone || '', formData.email || '', formData.username, formData.password, formData.role || 'User']);
    return { success: true, message: 'เพิ่มผู้ใช้งานสำเร็จ' };
  } catch (e) { return { success: false, message: e.toString() }; }
}

function updateUser(formData) {
  try {
    const ss = getSS();
    const sheet = ss.getSheetByName('Users');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(formData.id).trim()) {
        const rowNum = i + 1;
        sheet.getRange(rowNum, 2).setValue(formData.name);
        sheet.getRange(rowNum, 3).setValue(formData.phone || '');
        sheet.getRange(rowNum, 4).setValue(formData.email || '');
        sheet.getRange(rowNum, 5).setValue(formData.username);
        if (formData.password) sheet.getRange(rowNum, 6).setValue(formData.password);
        if (formData.requestorRole === 'Admin') sheet.getRange(rowNum, 7).setValue(formData.role || 'User');
        return { success: true, message: 'อัปเดตข้อมูลสำเร็จ' };
      }
    }
    return { success: false, message: 'ไม่พบผู้ใช้งานที่ต้องการแก้ไข' };
  } catch (e) { return { success: false, message: e.toString() }; }
}

function deleteUser(userId) {
  try {
    const ss = getSS();
    const sheet = ss.getSheetByName('Users');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(userId).trim()) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'ลบผู้ใช้งานสำเร็จ' };
      }
    }
    return { success: false, message: 'ไม่พบผู้ใช้งานที่ต้องการลบ' };
  } catch (e) { return { success: false, message: e.toString() }; }
}

function uploadImageToFolder(base64Data, fileName, folderId) {
  try {
    const parts = base64Data.split(',');
    const base64String = parts[1];
    const mimeType = parts[0].split(':')[1].split(';')[0];
    const blob = Utilities.newBlob(Utilities.base64Decode(base64String), mimeType, fileName);
    const file = DriveApp.getFolderById(folderId).createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (e) { return ""; }
}
