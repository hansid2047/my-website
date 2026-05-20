// ═══════════════════════════════════════════════════════════════════
//  全日型身障機構｜探視 & 返家預約系統 — Google Apps Script 後端
//  使用方式：
//    1. 在 Google 試算表中開啟 擴充功能 > Apps Script
//    2. 貼上此程式碼（取代全部內容）
//    3. 執行 initSheets() 一次初始化工作表
//    4. 部署 > 新增部署 > 網路應用程式
//       - 以身分執行：我（你的帳號）
//       - 誰可以存取：所有人
//    5. 複製部署 URL 貼到 index.html 的 GAS_URL 變數
// ═══════════════════════════════════════════════════════════════════

const SHEET_BOOKINGS = "預約紀錄";
const SHEET_NOTICES  = "機構公告";
const SHEET_ADMINS   = "管理員帳號";

// ─── 初始化工作表 ────────────────────────────────────────────────
function initSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 預約紀錄
  if (!ss.getSheetByName(SHEET_BOOKINGS)) {
    const s = ss.insertSheet(SHEET_BOOKINGS);
    s.appendRow(["編號","類型","院生姓名","家長姓名","日期","開始時間","結束時間","預計返院日期","備註","申請時間"]);
    styleHeader(s, 10, "#1B5E75");
    s.setFrozenRows(1);
    [60,70,90,90,100,80,80,100,200,160].forEach((w,i)=>s.setColumnWidth(i+1,w));
  }

  // 機構公告
  if (!ss.getSheetByName(SHEET_NOTICES)) {
    const s = ss.insertSheet(SHEET_NOTICES);
    s.appendRow(["編號","標題","內容","發布日期","置頂"]);
    styleHeader(s, 5, "#2A7A5C");
    s.setFrozenRows(1);
    [60,200,420,100,60].forEach((w,i)=>s.setColumnWidth(i+1,w));
    s.appendRow([1,"探視注意事項","1. 探視時間：每日 08:30–11:30、13:30–17:00。\n2. 中午 12:00–13:30 為院生休息時間，禁止探視。\n3. 探視前請於 48 小時前完成線上預約。\n4. 每次探視人數不超過 4 人。\n5. 探視時請勿攜帶來路不明食品。","2025-05-01","是"]);
    s.appendRow([2,"返家申請規定","1. 返家申請請至少 72 小時前提出。\n2. 返家期間家長須負完整監護責任。\n3. 返家前請確認院生身體狀況良好。\n4. 預計返院時間請務必準時，如有異動請立即電話通知。","2025-05-01","是"]);
  }

  // 管理員帳號
  if (!ss.getSheetByName(SHEET_ADMINS)) {
    const s = ss.insertSheet(SHEET_ADMINS);
    s.appendRow(["帳號","密碼","姓名","職稱","角色"]);
    styleHeader(s, 5, "#7C3AED");
    s.setFrozenRows(1);
    s.appendRow(["admin","fufu","系統管理員","最高管理員","superadmin"]);
  }
}

function styleHeader(sheet, cols, color) {
  sheet.getRange(1,1,1,cols)
    .setFontWeight("bold")
    .setBackground(color)
    .setFontColor("#ffffff");
}

// ─── CORS helper ─────────────────────────────────────────────────
function corsResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── GET 路由（支援 JSONP callback 參數） ────────────────────────
function doGet(e) {
  const p      = e.parameter || {};
  const action = p.action || "";
  const cb     = p.callback || ""; // JSONP callback

  let result;
  try {
    if (action === "getBookings") result = getBookings();
    else if (action === "getNotices")  result = getNotices();
    else if (action === "login")       result = loginAdmin(p.username, p.password);
    else if (action === "deleteNotice")result = deleteNotice(p.id);
    else if (action === "pinNotice")   result = pinNotice(p.id, p.pinned === "true");
    else result = { ok: false, error: "unknown action" };
  } catch(err) {
    result = { ok: false, error: err.message };
  }

  const json = JSON.stringify(result);

  // JSONP 模式（跨域使用）
  if (cb) {
    return ContentService
      .createTextOutput(`${cb}(${json})`)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return corsResponse(result);
}

// ─── POST 路由 ────────────────────────────────────────────────────
function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); }
  catch(err) { return corsResponse({ ok:false, error:"invalid JSON" }); }

  const action = body.action || "";
  let result;
  try {
    if      (action === "addBooking")   result = addBooking(body);
    else if (action === "addNotice")    result = addNotice(body);
    else if (action === "deleteNotice") result = deleteNotice(body.id);
    else if (action === "pinNotice")    result = pinNotice(body.id, body.pinned);
    else result = { ok:false, error:"unknown action" };
  } catch(err) {
    result = { ok:false, error:err.message };
  }

  return corsResponse(result);
}

// ─── 預約 ────────────────────────────────────────────────────────
function getBookings() {
  const s = getSheet(SHEET_BOOKINGS);
  const rows = s.getDataRange().getValues();
  if (rows.length <= 1) return { ok:true, data:[] };
  return {
    ok: true,
    data: rows.slice(1).map(r => ({
      id: r[0], type: r[1], child: r[2], parentName: r[3],
      date: fmtDate(r[4]), startTime: r[5], endTime: r[6],
      returnDate: fmtDate(r[7]), note: r[8], appliedAt: r[9],
    }))
  };
}

function addBooking(body) {
  const s   = getSheet(SHEET_BOOKINGS);
  const id  = s.getLastRow(); // row count (header = 1, so id starts at 1)
  const now = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy-MM-dd HH:mm:ss");

  s.appendRow([
    id,
    body.type === "visit" ? "探視" : "返家",
    body.child      || "",
    body.parentName || "",
    body.date       || "",
    body.startTime  || "",
    body.endTime    || "",
    body.returnDate || "",
    body.note       || "",
    now,
  ]);

  // 交替底色
  const row   = s.getLastRow();
  const color = row % 2 === 0 ? "#EAF4F0" : "#FFFFFF";
  s.getRange(row, 1, 1, 10).setBackground(color);

  return { ok:true, id };
}

// ─── 公告 ────────────────────────────────────────────────────────
function getNotices() {
  const s = getSheet(SHEET_NOTICES);
  const rows = s.getDataRange().getValues();
  if (rows.length <= 1) return { ok:true, data:[] };
  return {
    ok: true,
    data: rows.slice(1).map(r => ({
      id: r[0], title: r[1], content: r[2], date: fmtDate(r[3]), pinned: r[4] === "是",
    }))
  };
}

function addNotice(body) {
  const s     = getSheet(SHEET_NOTICES);
  const id    = s.getLastRow();
  const today = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy-MM-dd");
  s.appendRow([id, body.title, body.content, today, body.pinned ? "是" : "否"]);
  return { ok:true, id };
}

function deleteNotice(id) {
  const s    = getSheet(SHEET_NOTICES);
  const rows = s.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) { s.deleteRow(i+1); return { ok:true }; }
  }
  return { ok:false, error:"not found" };
}

function pinNotice(id, pinned) {
  const s    = getSheet(SHEET_NOTICES);
  const rows = s.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      s.getRange(i+1, 5).setValue(pinned ? "是" : "否");
      return { ok:true };
    }
  }
  return { ok:false };
}

// ─── 登入 ────────────────────────────────────────────────────────
function loginAdmin(username, password) {
  const s    = getSheet(SHEET_ADMINS);
  const rows = s.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === username && rows[i][1] === password) {
      return { ok:true, user:{ username:rows[i][0], name:rows[i][2], jobTitle:rows[i][3], role:rows[i][4] } };
    }
  }
  return { ok:false, error:"invalid credentials" };
}

// ─── 工具 ────────────────────────────────────────────────────────
function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s    = ss.getSheetByName(name);
  if (!s) { initSheets(); s = ss.getSheetByName(name); }
  return s;
}

function fmtDate(val) {
  if (!val) return "";
  if (val instanceof Date) return Utilities.formatDate(val, "Asia/Taipei", "yyyy-MM-dd");
  return String(val);
}
