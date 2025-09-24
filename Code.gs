/** ============ Entry ============ **/
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ระบบยืม-คืนหนังสือ')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** ============ Sheets Helpers ============ **/
const _TZ = 'Asia/Bangkok';
const _DATE_FMT = 'dd/MM/yyyy';

const _ss        = () => SpreadsheetApp.getActiveSpreadsheet();
const _bookSheet = () => _ss().getSheetByName('Book');
const _brSheet   = () => _ss().getSheetByName('Borrow-Return');

function _readBooks() {
  const sh = _bookSheet();
  if (!sh || sh.getLastRow() < 2) return [];
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
  return rows.map(r => ({ id: r[0], title: r[1], status: r[2] }));
}

function _nextId(sh) {
  const last = sh.getLastRow();
  if (last < 2) return 1;
  const v = sh.getRange(last, 1).getValue();
  return Number(v || 0) + 1;
}

/** ============ Date Utils ============ **/
function _coerceDate(v) {
  if (!v) return null;
  if (v instanceof Date) return v;
  if (typeof v === 'number') {
    // Google Sheets serial → Date
    const base = new Date(Date.UTC(1899, 11, 30));
    return new Date(base.getTime() + v * 86400000);
  }
  const d = new Date(v);
  return isNaN(d.getTime()) ? null : d;
}
function _fmt(d) {
  return d instanceof Date ? Utilities.formatDate(d, _TZ, _DATE_FMT) : '-';
}

/** ============ Header Map (Borrow-Return) ============ **/
function _headerMap(sh) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
    .map(h => String(h || '').trim().toLowerCase().replace(/\s+/g, '').replace(/[–—-]/g, '-'));

  const map = {};
  headers.forEach((k, i) => {
    const col = i + 1;
    if (['id','รหัส'].includes(k)) map.id = col;
    if (['bookid','รหัสหนังสือ','รหัสหนังสือid'].includes(k)) map.bookId = col;
    if (['booktitle','ชื่อหนังสือ','ชื่อเรื่อง'].includes(k)) map.bookTitle = col;
    if (['name','borrower','ผู้ยืม','ชื่อผู้ยืม'].includes(k)) map.borrower = col;
    if (['borrowdate','วันที่ยืม','วันยืม'].includes(k)) map.borrowDate = col;
    if (['duedate','กำหนดคืน','วันกำหนดคืน'].includes(k)) map.dueDate = col;
    if (['returndate','วันที่คืน','วันคืน'].includes(k)) map.returnDate = col;
    if (['status','สถานะ'].includes(k)) map.status = col;
  });
  return map;
}

/** ============ Read APIs ============ **/
function getBooks() {
  return _readBooks();
}
function getDashboardStats() {
  const books = _readBooks();
  const total = books.length;
  const borrowed = books.filter(b => b.status === 'ถูกยืม').length;
  return { total, borrowed, available: total - borrowed };
}
function getBorrowableBooks() {
  // เฉพาะเล่มที่ยังไม่ถูกยืม
  return _readBooks().filter(b => b.status !== 'ถูกยืม');
}
function getReturnableBooks() {
  // เล่มที่มีเรคคอร์ดสถานะ "ยังไม่คืน"
  const books = _readBooks();
  const history = _readHistoryRaw();
  const active = new Set(history.filter(h => h.status === 'ยังไม่คืน').map(h => h.bookTitle));
  return books.filter(b => active.has(b.title));
}

/** ส่งประวัติแบบพร้อมใช้บน client (string ล้วน ปลอดภัยต่อ serialization) */
function getBorrowHistoryForClient() {
  const rows = _readHistoryRaw();
  return rows.map(r => ({
    id: String(r.id ?? ''),
    bookId: String(r.bookId ?? ''),
    bookTitle: String(r.bookTitle ?? ''),
    borrower: String(r.borrower ?? ''),
    borrowDateTxt: _fmt(r.borrowDate),
    dueDateTxt: _fmt(r.dueDate),
    returnDateTxt: _fmt(r.returnDate),
    status: String(r.status ?? '')
  }));
}

/** อ่าน Borrow-Return เป็น object ภายใน (ใช้หัวตารางจริง) */
function _readHistoryRaw() {
  const sh = _brSheet();
  if (!sh || sh.getLastRow() < 2) return [];
  const map = _headerMap(sh);
  const get = (r, key, defIdx) => r[(map[key] || defIdx) - 1];

  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  const grid = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();

  return grid
    .filter(r => r.some(c => c !== '' && c !== null))
    .map(r => ({
      id:        get(r, 'id', 1),
      bookId:    get(r, 'bookId', 2),
      bookTitle: get(r, 'bookTitle', 3),
      borrower:  get(r, 'borrower', 4),
      borrowDate: _coerceDate(get(r, 'borrowDate', 5)),
      dueDate:    _coerceDate(get(r, 'dueDate', 6)),
      returnDate: _coerceDate(get(r, 'returnDate', 7)),
      status:     get(r, 'status', 8)
    }));
}

/** ============ Mutations (Borrow / Return) ============ **/
function borrowBook(form) {
  // form = { bookTitle, borrowerName, borrowDateISO, dueDateISO }
  const bkSh = _bookSheet(), brSh = _brSheet();
  if (!bkSh || !brSh) return { success:false, message:'ไม่พบชีตข้อมูล' };

  const books = _readBooks();
  const book = books.find(b => String(b.title) === String(form.bookTitle));
  if (!book) return { success:false, message:'ไม่พบหนังสือที่เลือก' };
  if (book.status === 'ถูกยืม') return { success:false, message:'หนังสือเล่มนี้ถูกยืมอยู่แล้ว' };

  const borrowDate = form.borrowDateISO ? new Date(form.borrowDateISO) : new Date();
  const dueDate = form.dueDateISO ? new Date(form.dueDateISO) : null;
  if (dueDate && borrowDate > dueDate) {
    return { success:false, message:'วันที่กำหนดคืนต้องไม่น้อยกว่าวันที่ยืม' };
  }

  // อัปเดตสถานะหนังสือ
  const rng = bkSh.getRange(2, 1, bkSh.getLastRow() - 1, 3).getValues();
  for (let i = 0; i < rng.length; i++) {
    if (rng[i][1] === book.title) { bkSh.getRange(i + 2, 3).setValue('ถูกยืม'); break; }
  }

  // เพิ่มบันทึกประวัติ
  const newId = _nextId(brSh);
  brSh.appendRow([newId, book.id, book.title, form.borrowerName || '', borrowDate, dueDate, '', 'ยังไม่คืน']);

  return { success:true, message:'ยืมหนังสือสำเร็จ', data:{ id:newId } };
}

function returnBook(payload) {
  // payload = { bookTitle }
  const bkSh = _bookSheet(), brSh = _brSheet();
  if (!bkSh || !brSh) return { success:false, message:'ไม่พบชีตข้อมูล' };

  const title = String(payload.bookTitle);
  // หา row ใน Book
  const bGrid = bkSh.getRange(2, 1, bkSh.getLastRow() - 1, 3).getValues();
  let bookRow = -1, bookId = null;
  for (let i = 0; i < bGrid.length; i++) {
    if (bGrid[i][1] === title) { bookRow = i + 2; bookId = bGrid[i][0]; break; }
  }
  if (bookRow === -1) return { success:false, message:'ไม่พบบันทึกหนังสือที่เลือก' };

  // หาเรคคอร์ดยืมที่ยังไม่คืนล่าสุด
  const rGrid = brSh.getRange(2, 1, brSh.getLastRow() - 1, 8).getValues();
  let hRow = -1;
  for (let i = rGrid.length - 1; i >= 0; i--) {
    if (rGrid[i][2] === title && rGrid[i][7] === 'ยังไม่คืน') { hRow = i + 2; break; }
  }
  if (hRow === -1) return { success:false, message:'ไม่พบเรคคอร์ดที่ยังไม่คืนของเล่มนี้' };

  // อัปเดตคืน
  brSh.getRange(hRow, 7).setValue(new Date()); // Return Date
  brSh.getRange(hRow, 8).setValue('คืนแล้ว');  // Status
  bkSh.getRange(bookRow, 3).setValue('พร้อมใช้งาน');

  return { success:true, message:'คืนหนังสือสำเร็จ', data:{ bookId } };
}
