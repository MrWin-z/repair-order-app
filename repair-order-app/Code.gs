/** Repair Order Web App — Google Apps Script */
const APP = {
  name: 'Repair Order',
  sheetName: 'Repair Orders',
  imageFolderName: 'Repair Order Attachments',
  headers: ['เลขที่ใบงาน', 'วันที่แจ้ง', 'สถานะ', 'ประเภทพื้นที่', 'ชื่อผู้แจ้ง', 'บ้านเลขที่/พื้นที่', 'โทรศัพท์', 'อาคาร', 'ชั้น', 'ประเภทงานซ่อม', 'หัวข้อ', 'รายละเอียด', 'ผู้รับแจ้ง', 'วันที่รับแจ้ง', 'ผู้รับผิดชอบ', 'วันนัดหมาย', 'รายละเอียดดำเนินการ', 'ผลการปฏิบัติงาน', 'วันที่ดำเนินการ', 'ผู้ตรวจรับงาน', 'วันที่ตรวจรับ', 'ความพึงพอใจ', 'จุดเด่นบริการ', 'ข้อเสนอแนะ', 'ลิงก์รูปภาพ', 'Folder ID']
};

function doGet(e) {
  const view = e && e.parameter.view;
  const template = HtmlService.createTemplateFromFile(view === 'print' ? 'Print' : 'Index');
  template.ticketId = (e && e.parameter.id) || '';
  return template.evaluate()
    .setTitle(view === 'print' ? 'ใบงานแจ้งซ่อม' : 'ระบบรับแจ้งซ่อม')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) { return HtmlService.createHtmlOutputFromFile(filename).getContent(); }

/** Run this once before the first use. */
function setup() {
  const props = PropertiesService.getScriptProperties();
  let spreadsheetId = props.getProperty('SPREADSHEET_ID');
  let sheet;
  if (!spreadsheetId) {
    const ss = SpreadsheetApp.create('Repair Order Database');
    spreadsheetId = ss.getId();
    props.setProperty('SPREADSHEET_ID', spreadsheetId);
    sheet = ss.getActiveSheet().setName(APP.sheetName);
  } else {
    sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(APP.sheetName);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(APP.headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, APP.headers.length).setFontWeight('bold').setBackground('#f6d365');
    sheet.autoResizeColumns(1, APP.headers.length);
  }
  if (!props.getProperty('IMAGE_FOLDER_ID')) {
    const folder = DriveApp.createFolder(APP.imageFolderName);
    props.setProperty('IMAGE_FOLDER_ID', folder.getId());
  }
  return { spreadsheetUrl: SpreadsheetApp.openById(spreadsheetId).getUrl(), folderUrl: DriveApp.getFolderById(props.getProperty('IMAGE_FOLDER_ID')).getUrl() };
}

function getConfig() {
  setup();
  const props = PropertiesService.getScriptProperties();
  return { categories: ['ไฟฟ้า', 'ประปา', 'แอร์', 'โครงสร้าง', 'ประตู/หน้าต่าง', 'สุขภัณฑ์', 'ส่วนกลาง', 'อื่นๆ'], spreadsheetUrl: SpreadsheetApp.openById(props.getProperty('SPREADSHEET_ID')).getUrl() };
}

function submitTicket(form) {
  setup();
  validate_(form);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const now = new Date();
    const ticketNo = nextTicketNo_(now);
    const folder = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('IMAGE_FOLDER_ID')).createFolder(ticketNo);
    const imageUrls = (form.attachments || []).map((item, index) => saveAttachment_(item, folder, ticketNo, index + 1));
    const row = [ticketNo, now, 'รับแจ้งแล้ว', form.areaType, form.requesterName, form.serviceArea, form.phone, form.building || '', form.floor || '', form.category, form.subject, form.detail || '', '', '', '', '', '', '', '', '', '', '', '', '', imageUrls.join('\n'), folder.getId()];
    const sheet = getSheet_();
    sheet.appendRow(row);
    const data = rowToObject_(row);
    return { ok: true, ticket: data, printUrl: ScriptApp.getService().getUrl() + '?view=print&id=' + encodeURIComponent(ticketNo) };
  } finally { lock.releaseLock(); }
}

function getTicket(ticketNo) {
  setup();
  const rows = getSheet_().getDataRange().getValues();
  const row = rows.slice(1).find(r => String(r[0]) === String(ticketNo));
  if (!row) throw new Error('ไม่พบเลขที่ใบงาน ' + ticketNo);
  return rowToObject_(row);
}

/** Recent repair orders for the in-app print archive. */
function listTickets() {
  setup();
  const rows = getSheet_().getDataRange().getValues().slice(1).reverse();
  const baseUrl = ScriptApp.getService().getUrl();
  return rows.slice(0, 500).map(row => {
    const ticket = rowToObject_(row);
    ticket.printUrl = baseUrl + '?view=print&id=' + encodeURIComponent(ticket.ticketNo);
    return ticket;
  });
}

/** Keeps customer photos private while still rendering them in the print view. */
function getPrintableTicket(ticketNo) {
  const ticket = getTicket(ticketNo);
  ticket.printImages = ticket.imageUrls.map(url => {
    const id = (/[?&]id=([^&]+)/.exec(url) || [])[1];
    if (!id) return '';
    const blob = DriveApp.getFileById(id).getBlob();
    return 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
  }).filter(Boolean);
  return ticket;
}

function getSheet_() { return SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID')).getSheetByName(APP.sheetName); }
function nextTicketNo_(date) {
  const props = PropertiesService.getScriptProperties();
  const yymm = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyMM');
  const key = 'COUNTER_' + yymm;
  const next = Number(props.getProperty(key) || '0') + 1;
  props.setProperty(key, String(next));
  return yymm + '-' + String(next).padStart(4, '0');
}
function saveAttachment_(item, folder, ticketNo, index) {
  const match = /^data:(.+);base64,(.+)$/.exec(item.dataUrl || '');
  if (!match) throw new Error('ไฟล์แนบไม่ถูกต้อง');
  const ext = (item.name || '').split('.').pop() || 'jpg';
  const blob = Utilities.newBlob(Utilities.base64Decode(match[2]), match[1], ticketNo + '-' + index + '.' + ext);
  const file = folder.createFile(blob);
  // Direct-view URL can be displayed by the printable repair form.
  return 'https://drive.google.com/uc?export=view&id=' + file.getId();
}
function validate_(f) {
  ['areaType', 'requesterName', 'serviceArea', 'phone', 'category', 'subject'].forEach(k => { if (!f[k] || !String(f[k]).trim()) throw new Error('กรุณากรอกข้อมูลให้ครบ'); });
  if ((f.attachments || []).length > 3) throw new Error('แนบรูปได้สูงสุด 3 รูป');
}
function rowToObject_(r) {
  const o = {}; APP.headers.forEach((h, i) => o[h] = r[i]);
  return { ticketNo: o['เลขที่ใบงาน'], requestedAt: formatDate_(o['วันที่แจ้ง']), status: o['สถานะ'], areaType: o['ประเภทพื้นที่'], requesterName: o['ชื่อผู้แจ้ง'], serviceArea: o['บ้านเลขที่/พื้นที่'], phone: o['โทรศัพท์'], building: o['อาคาร'], floor: o['ชั้น'], category: o['ประเภทงานซ่อม'], subject: o['หัวข้อ'], detail: o['รายละเอียด'], receivedBy: o['ผู้รับแจ้ง'], receivedAt: formatDate_(o['วันที่รับแจ้ง']), assignee: o['ผู้รับผิดชอบ'], appointment: formatDate_(o['วันนัดหมาย']), fixingDetail: o['รายละเอียดดำเนินการ'], jobResult: o['ผลการปฏิบัติงาน'], operatedAt: formatDate_(o['วันที่ดำเนินการ']), verifiedBy: o['ผู้ตรวจรับงาน'], verifiedAt: formatDate_(o['วันที่ตรวจรับ']), satisfaction: o['ความพึงพอใจ'], serviceLikes: o['จุดเด่นบริการ'], comment: o['ข้อเสนอแนะ'], imageUrls: String(o['ลิงก์รูปภาพ'] || '').split('\n').filter(Boolean) };
}
function formatDate_(v) { return v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'd MMM yyyy HH:mm') : (v || ''); }
