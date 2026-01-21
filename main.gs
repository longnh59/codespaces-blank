/*********** CẤU HÌNH FILE ************/
const SPREADSHEET_ID = '1ltiAK_lyktLcFFfjcwExW7XUY_aoqWsBHSWMlCdo-vk';

  const SHEET_DATA    = "Thông tin đặt hàng"; // chi tiết món (A:K, K = Mã đơn)
  const SHEET_DEBT    = "congno";             // sheet công nợ (A→M)
  const SHEET_KH      = "Khách hàng";
  const SHEET_DM      = "Danh mục hàng";
  const SHEET_ACCOUNT = "account";
  const SHEET_LOG     = "log";
  const SHEET_PENDING = "doncho";

/*********** TỐI ƯU GOOGLE SHEETS API (ADVANCED SERVICE) ************/
// Bật trong Apps Script: Services -> Google Sheets API
// Bật trong Google Cloud: APIs & Services -> Library -> Google Sheets API
const USE_SHEETS_API = true;


function jsonSafe_(obj) {
  try {
    return JSON.parse(JSON.stringify(obj));
  } catch (e) {
    return obj;
  }
}

function hasSheetsApi_() {
  try { return !!(USE_SHEETS_API && typeof Sheets !== "undefined" && Sheets.Spreadsheets && Sheets.Spreadsheets.Values); }
  catch(e){ return false; }
}

// Chuyển username (có thể là email) sang tên hiển thị
function formatDisplayName_(username) {
  if (!username) return "";
  const str = String(username).trim();
  // Nếu là email, lấy phần trước @
  if (str.includes('@')) {
    return str.split('@')[0];
  }
  return str;
}

function sh_valuesGet_(rangeA1) {
  if (!hasSheetsApi_()) return null;
  const res = Sheets.Spreadsheets.Values.get(SPREADSHEET_ID, rangeA1, {
    valueRenderOption: "FORMATTED_VALUE",
    dateTimeRenderOption: "FORMATTED_STRING"
  });
  return (res && res.values) ? res.values : [];
}

function sh_valuesBatchGet_(ranges) {
  if (!hasSheetsApi_()) return null;
  const res = Sheets.Spreadsheets.Values.batchGet(SPREADSHEET_ID, {
    ranges: ranges,
    valueRenderOption: "FORMATTED_VALUE",
    dateTimeRenderOption: "FORMATTED_STRING"
  });
  const vrs = (res && res.valueRanges) ? res.valueRanges : [];
  return vrs.map(v => (v && v.values) ? v.values : []);
}

function sh_valuesBatchUpdate_(data) {
  if (!hasSheetsApi_()) return null;
  // data: [{range:"Sheet!A1:B1", values:[[...]]}, ...]
  return Sheets.Spreadsheets.Values.batchUpdate(SPREADSHEET_ID, {
    valueInputOption: "USER_ENTERED",
    data: data
  });
}

function sh_valuesAppend_(rangeA1, values2d) {
  if (!hasSheetsApi_()) return null;
  return Sheets.Spreadsheets.Values.append(SPREADSHEET_ID, rangeA1, {
    valueInputOption: "USER_ENTERED",
    insertDataOption: "INSERT_ROWS"
  }, { values: values2d });
}

// Dùng cho thao tác cấu trúc (delete row, insert row...). sheetId PHẢI là số.
function sh_batchUpdate_(requests) {
  if (!hasSheetsApi_()) return null;
  return Sheets.Spreadsheets.batchUpdate(SPREADSHEET_ID, { requests: requests });
}

function sh_getSheetIdByName_(sheetName) {
  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error("Không tìm thấy sheet '" + sheetName + "'.");
  return sh.getSheetId(); // number
}

function sh_deleteRowFast_(sheetName, row1based) {
  row1based = Number(row1based || 0);
  if (!row1based || row1based < 1) throw new Error("Row không hợp lệ.");

  // Prefer reliable SpreadsheetApp deletion; Sheets API deletion is optional + fallback
  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error("Không tìm thấy sheet '" + sheetName + "'.");

  // If Sheets API not enabled, just delete via SpreadsheetApp
  if (!hasSheetsApi_()) {
    sh.deleteRow(row1based);
    return;
  }

  // Try Sheets API deleteDimension (faster for large sheets), but ALWAYS fallback if it errors
  try {
    const sheetId = Number(sh.getSheetId());
    if (!isFinite(sheetId)) throw new Error("sheetId không hợp lệ.");
    Sheets.Spreadsheets.batchUpdate(SPREADSHEET_ID, {
      requests: [{
        deleteDimension: {
          range: {
            sheetId: sheetId,
            dimension: "ROWS",
            startIndex: row1based - 1,
            endIndex: row1based
          }
        }
      }]
    });
  } catch (e) {
    // Fallback: SpreadsheetApp (reliable)
    sh.deleteRow(row1based);
  }
}


function toMoneyNumber_(v) {
  if (v === null || v === undefined) return 0;
  if (typeof v === "number") return isFinite(v) ? v : 0;
  var s = String(v || "").trim();
  if (!s) return 0;

  // keep digits, dot, comma, minus
  s = s.replace(/[^0-9,\.\-]/g, "");
  if (!s) return 0;

  var hasDot = s.indexOf(".") >= 0;
  var hasComma = s.indexOf(",") >= 0;

  if (hasDot && hasComma) {
    // assume '.' thousands, ',' decimal
    s = s.replace(/\./g, "").replace(/,/g, ".");
  } else if (hasComma) {
    var parts = s.split(",");
    if (parts.length === 2 && parts[1].length === 3) {
      // comma as thousands separator
      s = s.replace(/,/g, "");
    } else {
      // comma as decimal
      s = s.replace(/,/g, ".");
    }
  } else if (hasDot) {
    var dotParts = s.split(".");
    if (dotParts.length > 2) {
      // multiple dots => thousands separators
      s = s.replace(/\./g, "");
    } else if (dotParts.length === 2 && dotParts[1].length === 3) {
      // single dot with 3 digits tail => thousands
      s = s.replace(/\./g, "");
    }
  }

  var n = Number(s);
  return isFinite(n) ? n : 0;
}

function sh_readRowA1_(sheetName, row1based, lastColIndex) {
  lastColIndex = Number(lastColIndex || 1);
  const lastColA1 = colToA1_(lastColIndex);
  const rg = sheetName + "!A" + row1based + ":" + lastColA1 + row1based;
  const vals = sh_valuesGet_(rg);
  return (vals && vals[0]) ? vals[0] : [];
}



  /*********** CẤU HÌNH IN ẤN ************/
  const CFG_ENABLE_TAX = "CFG_ENABLE_TAX"; // "1" = có tính thuế, "0" = không tính thuế

  function getAppProps_() {
    // Dùng ScriptProperties để chạy ổn cho cả project standalone & container-bound
    return PropertiesService.getScriptProperties();
  }

  function getEnableTax_() {
    const v = getAppProps_().getProperty(CFG_ENABLE_TAX);
    if (v === null || v === "") return true; // mặc định: có tính thuế
    const s = String(v).toLowerCase();
    return (s === "1" || s === "true");
  }

  function ui_getTaxConfig() {
    return { ok: true, enabled: getEnableTax_() };
  }

  function ui_setTaxConfig(enabled, username) {
    const ctx = _resolveUserCtx_({ meta: { username: username } });
    if (String(ctx.role || "").toLowerCase() !== "admin") throw new Error("Bạn không có quyền cấu hình.");

    const on = !!enabled;
    getAppProps_().setProperty(CFG_ENABLE_TAX, on ? "1" : "0");
    return { ok: true, enabled: on };
  }


  /*********** DANH MỤC HÀNG - CỘT ************/
  const DM_COL_NAME   = 1; // A
  const DM_COL_UNIT   = 2; // B
  const DM_COL_PRICE  = 3; // C
  const DM_COL_TYPE   = 4; // D
  const DM_COL_STATUS = 5;
  const DM_COL_CODE   = 6; // F: Mã món

  /*********** CÔNG NỢ - CỘT ************/
  const DEBT_COL_DATE       = 1;   // A: Ngày tháng
  const DEBT_COL_INFO       = 2;   // B: Thông tin tiệc (Tên - SĐT - Địa chỉ)
  const DEBT_COL_SOMAM      = 3;   // C: Số mâm
  const DEBT_COL_DONGIA_MAM = 4;   // D: Đơn giá 1 mâm
  const DEBT_COL_TONG_DON   = 5;   // E: Thành tiền (tổng đơn)
  const DEBT_COL_KM_NOTE    = 6;   // F: Nội dung KM
  const DEBT_COL_KM_AMOUNT  = 7;   // G: Số tiền KM
  const DEBT_COL_DOANHSO    = 8;   // H: Doanh số
  const DEBT_COL_CONGNO     = 9;   // I: Công nợ
  const DEBT_COL_STATUS     = 10;  // J: Tình trạng
  const DEBT_COL_NGAYTT     = 11;  // K: Ngày thanh toán
  const DEBT_COL_THUNGAN    = 12;  // L: Thu ngân
  const DEBT_COL_ORDER_ID   = 13;  
const DEBT_COL_ITEM_COUNT = 14;
const DEBT_COL_ITEM_START = 15;
const DEBT_COL_ITEM_END   = 16;
// M: Mã đơn

  
  // === Pending (Đơn chờ) schema (giống congno) ===
  const PENDING_COL_DATE       = 1;   // A: Ngày
  const PENDING_COL_INFO       = 2;   // B: Thông tin tiệc (Tên - SĐT - Địa chỉ, có thể có dấu '-')
  const PENDING_COL_SOMAM      = 3;   // C: Số mâm
  const PENDING_COL_DONGIA_MAM = 4;   // D: Đơn giá 1 mâm
  const PENDING_COL_TONG_DON   = 5;   // E: Thành tiền (tổng đơn)
  const PENDING_COL_KM_NOIDUNG = 6;   // F: Nội dung KM
  const PENDING_COL_KM_SOTIEN  = 7;   // G: Số tiền KM
  const PENDING_COL_CASHIER    = 12;  // L: Thu ngân
  const PENDING_COL_ORDER_ID   = 13;  // M: Mã đơn
  const PENDING_COL_DEPOSIT    = 17;  // Q: Đặt cọc
  var PENDING_DATA_START = 2; // dữ liệu đơn chờ bắt đầu từ hàng 2 (sau header)


const PENDING_COL_ITEM_COUNT = 14;  // N: Số món (cache)
  const PENDING_COL_ITEM_START = 15;  // O: item_start_row (sheet data)
  const PENDING_COL_ITEM_END   = 16;  // P: item_end_row (sheet data)
  const DEBT_HEADER_ROW = 1;
  const DEBT_DATA_START = 2;
// Cập nhật lại các hằng số cột
const DEBT_COL_DEPOSIT = 17; // Cột Q: Đặt cọc 
const DEBT_LAST_COL = 17;
const PENDING_LAST_COL = 17;
  const STATUS_PAID   = "Đã thanh toán";
  const STATUS_DEBT   = "Ghi nợ";
  const STATUS_UNPAID = "Chưa thanh toán";


  function _getRole_(username) {
    username = String(username || "").trim();
    if (!username) return "";
    var acc = getAccountByUsername_(username);
    if (!acc && username.indexOf("@") > 0) {
      // nếu người dùng gửi email, thử map về username trước dấu @
      acc = getAccountByUsername_(String(username.split("@")[0] || "").trim());
    }
    if (!acc) return "";
    return String(acc.role || "").trim().toLowerCase();
  }

  /**
   * Resolve username/role từ:
   *  - string username
   *  - object { username | user | actor | token | meta:{...} }
   *  - token (validateToken_)
   *
   * Lưu ý: Nếu không resolve được → role mặc định 'cashier' để tránh lỗi lưu đơn,
   * nhưng các thao tác xoá vẫn bị chặn (chỉ admin).
   */
  function _resolveUserCtx_(input) {
    var u = "";
    var token = "";

    if (input && typeof input === "object") {
      var meta = input.meta || {};
      u = String(meta.username || meta.user || meta.actor || input.username || input.user || input.actor || "").trim();
      token = String(meta.token || input.token || "").trim();
    } else {
      u = String(input || "").trim();
    }

    if (!u && token && typeof validateToken_ === "function") {
      var vu = validateToken_(token);
      if (vu && vu.username) u = String(vu.username || "").trim();
    }

    // fallback: ActiveUser email (trong nhiều trường hợp consumer sẽ rỗng)
    if (!u) {
      try {
        var email = Session.getActiveUser().getEmail();
        if (email) u = String(email || "").trim();
      } catch (e) {}
    }

    var role = _getRole_(u);

    // nếu u là email mà account dùng username (không @), thử lại
    if (!role && u && u.indexOf("@") > 0) {
      var u2 = String(u.split("@")[0] || "").trim();
      if (u2) {
        var role2 = _getRole_(u2);
        if (role2) {
          u = u2;
          role = role2;
        }
      }
    }

    return { username: u, role: role };
  }

  function assertCanEdit_(actorOrPayload) {
    var ctx = _resolveUserCtx_(actorOrPayload);
    var role = ctx.role || "cashier";

    // manager chỉ được xem
    if (role === "manager") throw new Error("Tài khoản manager chỉ được xem, không được sửa/xoá.");

    // trả về ctx để caller log/ghi vào sheet
    if (!ctx.username) ctx.username = "unknown";
    ctx.role = role;
    return ctx;
  }

  function assertCanDelete_(actorOrPayload) {
    var ctx = _resolveUserCtx_(actorOrPayload);
    var role = ctx.role || "cashier";

    // chỉ admin được xoá
    if (role === "manager") throw new Error("Bạn không có quyền xoá.");

    if (!ctx.username) ctx.username = "unknown";
    ctx.role = role;
    return ctx;
  }

  /* =====================================================================================
    1) SPREADSHEET + MENU (GIỮ NGUYÊN)
    ===================================================================================== */
  // ===== CACHE Spreadsheet/Sheet (tăng tốc trong 1 lần execution) =====
  var __SS_CACHE__ = null;
  var __SHEET_CACHE__ = null;

  /** Mỗi execution chỉ openById 1 lần */
  function getSpreadsheet_() {
    if (__SS_CACHE__) return __SS_CACHE__;
    __SS_CACHE__ = SpreadsheetApp.openById(SPREADSHEET_ID);
    return __SS_CACHE__;
  }

  /** Cache sheet theo tên trong 1 execution */
  function getSheet_(sheetName) {
    sheetName = String(sheetName || "").trim();
    if (!sheetName) return null;
    if (!__SHEET_CACHE__) __SHEET_CACHE__ = {};
    if (__SHEET_CACHE__[sheetName]) return __SHEET_CACHE__[sheetName];
    var sh = getSpreadsheet_().getSheetByName(sheetName);
    __SHEET_CACHE__[sheetName] = sh || null;
    return sh || null;
  }

  function rememberSheetName_(propKey, sh) {
    try {
      if (!propKey || !sh) return;
      PropertiesService.getScriptProperties().setProperty(String(propKey), String(sh.getName()));
    } catch(e) {}
  }
  function normalizeDataSheet_(shDataMaybe) {
    // Accept either Sheet or Spreadsheet; fallback to DATA sheet by name.
    if (shDataMaybe && typeof shDataMaybe.getLastRow === 'function' && typeof shDataMaybe.getRange === 'function') {
      return shDataMaybe;
    }
    if (shDataMaybe && typeof shDataMaybe.getSheetByName === 'function') {
      var sh = shDataMaybe.getSheetByName(SHEET_DATA);
      if (sh) return sh;
    }
    var ss = getSpreadsheet_();
    var sh2 = ss.getSheetByName(SHEET_DATA);
    if (!sh2) throw new Error("Không tìm thấy sheet '" + SHEET_DATA + "'.");
    return sh2;
  }

  function onOpen(e) {
    SpreadsheetApp.getUi()
      .createMenu("🍽 Phiếu nhập (Web)")
      .addItem("📋 Mở form phiếu nhập", "openWebForm_")
      .addItem("📊 Mở Dashboard", "openDashboard_")
      .addSeparator()
      .addItem("✅ Áp dụng dropdown trạng thái Công nợ", "applyDebtRules_")
      .addToUi();
  }

  function openWebForm_() {
    const html = HtmlService.createHtmlOutputFromFile("index")
      .setWidth(1200)
      .setHeight(700);
    SpreadsheetApp.getUi().showModalDialog(html, "Phiếu nhập (Web)");
  }

  function openDashboard_() {
    const html = HtmlService.createHtmlOutputFromFile("index")
      .setWidth(1200)
      .setHeight(700);
    SpreadsheetApp.getUi().showModalDialog(html, "Phiếu nhập (Web)");
  }

  function applyDebtRules_() {
    const ss = getSpreadsheet_();
    const sh = ss.getSheetByName(SHEET_DEBT);
    if (!sh) throw new Error("Không tìm thấy sheet '" + SHEET_DEBT + "'.");

    const lastRow = sh.getLastRow();
    if (lastRow < DEBT_DATA_START) return;

    const numRows = lastRow - DEBT_DATA_START + 1;
    const ruleStatus = SpreadsheetApp.newDataValidation()
      .requireValueInList([STATUS_PAID, STATUS_DEBT, STATUS_UNPAID], true)
      .setAllowInvalid(false)
      .build();

    sh.getRange(DEBT_DATA_START, DEBT_COL_STATUS, numRows, 1).setDataValidation(ruleStatus);
  }


  /* =====================================================================================
    2) DONCHO SHEET (GIỮ NGUYÊN)
    ===================================================================================== */
function getPendingSheet_() {
  const ss = getSpreadsheet_();

  // Ưu tiên tuyệt đối sheet doncho nếu tồn tại (kể cả trống)
  const sh = ss.getSheetByName(SHEET_PENDING);
  if (sh) {
    try { ensurePendingHeader_(sh); } catch (e) {}
    try { PropertiesService.getScriptProperties().setProperty("PENDING_SHEET_NAME", SHEET_PENDING); } catch (e) {}
    return sh;
  }

  // Nếu không có sheet doncho thì tạo mới + header (nhanh, tránh scan nhầm sang congno)
  const shNew = ss.insertSheet(SHEET_PENDING);
  try { ensurePendingHeader_(shNew); } catch (e) {}
  try { PropertiesService.getScriptProperties().setProperty("PENDING_SHEET_NAME", SHEET_PENDING); } catch (e) {}
  return shNew;
}

  /* =====================================================================================
    3) UTIL CHUNG (GIỮ NGUYÊN)
    ===================================================================================== */
  function normalizePhone_(s) {
    let digits = String(s || "").replace(/\D/g, "");
    if (digits.startsWith("84")) digits = digits.slice(2);
    if (digits.startsWith("0"))  digits = digits.slice(1);
    return digits;
  }

function extractPhoneFromInfo_(info) {
  const s = String(info || "");
  const m = s.match(/(\+?84|0)?\d{6,12}/);
  if (!m) return "";
  return normalizePhone_(m[0]);
}

function getNameFromInfo_(info) {
  var s = String(info || "").trim();
  if (!s) return "";
  var phone = extractPhoneFromInfo_(s);
  if (!phone) return s;
  var idx = s.indexOf(phone);
  if (idx <= 0) return s;
  var namePart = s.slice(0, idx).replace(/[-|]+$/g, "").trim();
  // remove trailing separators like " - "
  namePart = namePart.replace(/\s*-\s*$/g, "").trim();
  return namePart;
}

function getAddressFromInfo_(info) {
  var s = String(info || "").trim();
  if (!s) return "";
  var phone = extractPhoneFromInfo_(s);
  if (!phone) return "";
  var idx = s.indexOf(phone);
  if (idx < 0) return "";
  var rest = s.slice(idx + phone.length).trim();
  rest = rest.replace(/^[-|]+/g, "").trim();
  rest = rest.replace(/^\s*-\s*/g, "").trim();
  return rest;
}



function getLeadingIndexOffset_(sh) {
  try {
    if (!sh) return 0;
    var lastCol = sh.getLastColumn() || 1;
    var h = sh.getRange(1, 1, 1, Math.min(6, lastCol)).getValues()[0] || [];
    var a = normalizeText_(h[0] || "");
    if (!a || a === "stt" || a.indexOf("unnamed") >= 0) {
      // confirm by looking at next header
      var b = normalizeText_(h[1] || "");
      if (!b) return 1;
      return 1;
    }
  } catch (e) {}
  return 0;
}

function extractPhoneFromInfo_(info) {
  var s = String(info || "");
  if (!s) return "";
  // Prefer VN-style numbers: 0xxxxxxxxx or +84xxxxxxxxx
  var m = s.match(/(?:\+?84|0)\d{8,11}/g);
  if (m && m.length) {
    // choose the longest
    var best = m[0];
    for (var i = 1; i < m.length; i++) if (String(m[i]).length > String(best).length) best = m[i];
    return String(best || "").trim();
  }
  // fallback: any 6-12 digit group
  var m2 = s.match(/\b\d{6,12}\b/g);
  if (m2 && m2.length) {
    var best2 = m2[0];
    for (var j = 1; j < m2.length; j++) if (String(m2[j]).length > String(best2).length) best2 = m2[j];
    return String(best2 || "").trim();
  }
  return "";
}



function buildCustomerAggFromDebt_() {
  const ss = getSpreadsheet_();
  const shDebt = ss.getSheetByName(SHEET_DEBT);
  const agg = {};
  if (!shDebt) return agg;

  const last = shDebt.getLastRow();
  if (last < DEBT_DATA_START) return agg;

  const vals = shDebt.getRange(DEBT_DATA_START, 1, last - DEBT_DATA_START + 1, DEBT_LAST_COL).getValues();

  vals.forEach(row => {
    const phoneKey = extractPhoneFromInfo_(row[DEBT_COL_INFO - 1]);
    if (!phoneKey) return;

    const tongDon = toMoneyNumber_(row[DEBT_COL_TONG_DON - 1]);
    let congNo = toMoneyNumber_(row[DEBT_COL_CONGNO - 1]);
    const doanhSo = toMoneyNumber_(row[DEBT_COL_DOANHSO - 1]);
    const st = String(row[DEBT_COL_STATUS - 1] || "").trim();

    if (!congNo && !doanhSo) {
      if (st === STATUS_PAID) congNo = 0;
      else congNo = tongDon;
    }

    if (!agg[phoneKey]) agg[phoneKey] = { totalOrdered: 0, totalDebt: 0 };
    agg[phoneKey].totalOrdered += tongDon;
    agg[phoneKey].totalDebt += congNo;
  });

  Object.keys(agg).forEach(k => {
    agg[k].totalOrdered = roundVnd_(agg[k].totalOrdered);
    agg[k].totalDebt = roundVnd_(agg[k].totalDebt);
  });

  return agg;
}


  function isValidPhone_(s) {
    const core = normalizePhone_(s);
    return /^\d{6,12}$/.test(core);
  }

  function normalizeText_(s) {
    return String(s || "")
      .replace(/[\u200B-\u200D\uFEFF]/g, "")
      .replace(/\u00A0/g, " ")
      .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
      .replace(/\s+/g, " ")
      .trim()
      .toLowerCase();
  }


 function parseDateCell_(v) {
  if (v === null || v === undefined || v === "") return null;

  // Date object từ sheet
  if (Object.prototype.toString.call(v) === "[object Date]") {
    const t = v.getTime();
    if (isNaN(t)) return null;
    const d = new Date(v);
    d.setHours(0, 0, 0, 0);
    return d;
  }

  // Timestamp (ms) - chỉ dùng nếu bạn chắc input là ms
  if (typeof v === "number") {
    const d = new Date(v);
    if (isNaN(d.getTime())) return null;
    d.setHours(0, 0, 0, 0);
    return d;
  }

  const s = String(v).trim();
  if (!s) return null;

  let m;

  // yyyy-MM-dd
  m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) {
    const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    d.setHours(0, 0, 0, 0);
    return d;
  }

  // dd/MM/yyyy
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) {
    const d = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
    d.setHours(0, 0, 0, 0);
    return d;
  }

  // dd-MM-yyyy
  m = s.match(/^(\d{1,2})-(\d{1,2})-(\d{4})/);
  if (m) {
    const d = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
    d.setHours(0, 0, 0, 0);
    return d;
  }

  // fallback
  const d2 = new Date(s);
  if (isNaN(d2.getTime())) return null;
  d2.setHours(0, 0, 0, 0);
  return d2;
}


  function fmtDateYmd_(d, tz) {
    const dd = parseDateCell_(d) || new Date();
    const yyyy = Utilities.formatDate(dd, tz, "yyyy");
    const mm   = Utilities.formatDate(dd, tz, "MM");
    const ddd  = Utilities.formatDate(dd, tz, "dd");
    return yyyy + "-" + mm + "-" + ddd;
  }

  function parseYmdToDate_(ymd) {
    const s = String(ymd || "").trim();
    if (!s) return null;
    const p = s.split("-").map(Number);
    if (p.length !== 3 || isNaN(p[0]) || isNaN(p[1]) || isNaN(p[2])) return null;
    const d = new Date(p[0], p[1] - 1, p[2]);
    d.setHours(0, 0, 0, 0);
    return d;
  }

  function roundVnd_(n) {
    n = Number(n) || 0;
    return Math.round(n);
  }

  function makeOrderKey_(ngayDate, sdt) {
    const tz = Session.getScriptTimeZone();
    return fmtDateYmd_(ngayDate, tz) + "|" + normalizePhone_(sdt);
  }

  function makeOrderId_(ngayDate, sdt) {
    const tz   = Session.getScriptTimeZone();
    const base = fmtDateYmd_(ngayDate, tz) + "|" + normalizePhone_(sdt);
    const now  = new Date();
    const ts   = Utilities.formatDate(now, tz, "HHmmss");
    const rnd  = Math.floor(Math.random() * 1000);
    return base + "|" + ts + "|" + rnd;
  }

  function normalizeName_(s) {
    return (s || "")
      .toString()
      .trim()
      .toLowerCase()
      .replace(/\s+/g, " ");
  }


  /* =====================================================================================
    4) DANH MỤC HÀNG (GIỮ NGUYÊN)
    ===================================================================================== */
  function getProductMap_() {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh = ss.getSheetByName(SHEET_DM);
    const lastRow = sh ? sh.getLastRow() : 0;
    if (!sh || lastRow < 2) return {};

    // Đọc đủ đến cột Mã món (F)
    const values = sh.getRange(2, 1, lastRow - 1, DM_COL_CODE).getValues();
    const map = {};
    values.forEach(r => {
      const name = String(r[DM_COL_NAME - 1] || "").trim();
      if (!name) return;

      map[name.toLowerCase()] = {
        name,
        unit: String(r[DM_COL_UNIT - 1] || "").trim(),
        price: Number(r[DM_COL_PRICE - 1] || 0),
        type: String(r[DM_COL_TYPE - 1] || "").trim(),
        status: String(r[DM_COL_STATUS - 1] || "").trim(),
        code: String(r[DM_COL_CODE - 1] || "").trim()
      };
    });
    return map;
  }

  function getProductInfo(productName) {
    const key = String(productName || "").trim().toLowerCase();
    const map = getProductMap_();
    return map[key] || null;
  }

  function getDanhMucHangFull_() {
    const ss = getSpreadsheet_();
    const sh = ss.getSheetByName(SHEET_DM);
    if (!sh) return [];
    const last = sh.getLastRow();
    if (last < 2) return [];

    // A..F: tên, dvt, giá, loại, status, mã món
    const values = sh.getRange(2, 1, last - 1, DM_COL_CODE).getValues();
    return values
      .filter(r => r[0])
      .map(r => ({
        ten: r[DM_COL_NAME - 1],
        dvt: (r[DM_COL_UNIT - 1] || "Đĩa"),
        gia: Number(r[DM_COL_PRICE - 1]) || 0,
        loaiMon: String(r[DM_COL_TYPE - 1] || "").trim(),
        status: String(r[DM_COL_STATUS - 1] || "").trim(),
        maMon: String(r[DM_COL_CODE - 1] || "").trim()
      }));
  }

  function clearMenuCache_() {
    try {
      const cache = CacheService.getScriptCache();
      cache.remove("DM_FULL_V1");

      const tz = Session.getScriptTimeZone();
      const now = new Date();
      const todayStr = Utilities.formatDate(now, tz, "yyyy-MM-dd");
      cache.remove("INIT_DATA_V1:" + todayStr);

      const y = new Date(now.getTime() - 24 * 3600 * 1000);
      const yStr = Utilities.formatDate(y, tz, "yyyy-MM-dd");
      cache.remove("INIT_DATA_V1:" + yStr);
    } catch (e) {}
  }

  function getDanhMucHangFullCached_() {
    const cache = CacheService.getScriptCache();
    const key = "DM_FULL_V1";
    const cached = cache.get(key);
    if (cached) {
      try { return JSON.parse(cached) || []; } catch(e) {}
    }
    const data = getDanhMucHangFull_();
    try { cache.put(key, JSON.stringify(data), 300); } catch(e) {}
    return data;
  }


  function getInitData() {
    const tz = Session.getScriptTimeZone();
    const now = new Date();
    const todayStr = Utilities.formatDate(now, tz, "yyyy-MM-dd");

    const cache = CacheService.getScriptCache();
    const key = "INIT_DATA_V1:" + todayStr;
    const cached = cache.get(key);
    if (cached) {
      try { return JSON.parse(cached); } catch(e) {}
    }

    const out = { today: todayStr, timeZone: tz, menu: getDanhMucHangFullCached_() };
    try { cache.put(key, JSON.stringify(out), 300); } catch(e) {}
    return out;
  }

  function ensureDanhMucColumns_(shDm) {
    // Đảm bảo sheet Danh mục hàng có đủ cột A..F
    try {
      if (!shDm) return;
      const lastCol = shDm.getLastColumn();
      if (lastCol >= DM_COL_CODE) return;

      const missing = DM_COL_CODE - lastCol;
      shDm.insertColumnsAfter(lastCol, missing);

      // Nếu header đang trống, set header chuẩn
      const header = shDm.getRange(1, 1, 1, DM_COL_CODE).getValues()[0] || [];
      const h = header.map(x => String(x || "").trim());
      if (!h[0]) {
        shDm.getRange(1, 1, 1, DM_COL_CODE).setValues([[
          "Tên hàng", "ĐVT", "Đơn giá", "Loại món", "Status", "Mã món"
        ]]);
      } else {
        // bổ sung tên cột thiếu (nếu có)
        const need = ["Tên hàng","ĐVT","Đơn giá","Loại món","Status","Mã món"];
        for (let i = 0; i < need.length; i++) if (!h[i]) h[i] = need[i];
        shDm.getRange(1, 1, 1, DM_COL_CODE).setValues([h.slice(0, DM_COL_CODE)]);
      }
    } catch (e) {}
  }

  function getNextMonCode_(shDm) {
    // Sinh mã món dạng M0001, M0002... dựa trên cột F (Mã món)
    try {
      if (!shDm) return "M0001";
      const last = shDm.getLastRow();
      if (last < 2) return "M0001";

      const vals = shDm.getRange(2, DM_COL_CODE, last - 1, 1).getValues();
      let maxN = 0;
      for (let i = 0; i < vals.length; i++) {
        const s = String(vals[i][0] || "").trim().toUpperCase();
        const m = s.match(/^M(\d{1,})$/);
        if (m) {
          const n = parseInt(m[1], 10);
          if (n > maxN) maxN = n;
        }
      }
      const next = maxN + 1;
      return "M" + String(next).padStart(4, "0");
    } catch (e) {
      return "M0001";
    }
  }

  function upsertMenuFromItems_(items) {
    if (!items || !items.length) return;

    var ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
    var shDm = ss.getSheetByName(SHEET_DM);
    if (!shDm) return;

    ensureDanhMucColumns_(shDm);

    var lastRow = shDm.getLastRow();
    var existing = {};
    var usedCodes = {};
    var maxCodeN = 0;

    if (lastRow >= 2) {
      var rng = shDm.getRange(2, 1, lastRow - 1, DM_COL_CODE).getValues();
      rng.forEach(function(r) {
        var name = normalizeName_(r[DM_COL_NAME - 1]);
        if (name) existing[name] = true;

        var code = String(r[DM_COL_CODE - 1] || "").trim().toUpperCase();
        if (code) usedCodes[code] = true;
        var m = code.match(/^M(\d{1,})$/);
        if (m) {
          var n = parseInt(m[1], 10);
          if (n > maxCodeN) maxCodeN = n;
        }
      });
    }

    var toAppend = [];
    var nextN = maxCodeN;

    items.forEach(function(it) {
      var rawName = (it.tenMon || it.ten || "").toString().trim();
      var norm = normalizeName_(rawName);
      if (!norm) return;

      if (!existing[norm]) {
        existing[norm] = true;

        var dvt  = (it.dvt || "Đĩa").toString().trim();
        var gia  = Number(it.donGia || it.dg || it.gia || 0);
        if (!isFinite(gia)) gia = 0;

        var loaiMon = (it.loaiMon || it.loai || it.category || it.type || "").toString().trim() || "Món chính";
        var status  = normalizeMenuStatus01_(it.status || it.st || it.trangThai || it.isActive || "");

        // Sinh mã món
        nextN++;
        var code = "M" + String(nextN).padStart(4, "0");
        while (usedCodes[code]) {
          nextN++;
          code = "M" + String(nextN).padStart(4, "0");
        }
        usedCodes[code] = true;

        toAppend.push([rawName, dvt, gia, loaiMon, status, code]);
      }
    });

    if (toAppend.length) {
      shDm.getRange(lastRow + 1, 1, toAppend.length, DM_COL_CODE).setValues(toAppend);
    }

    clearMenuCache_();
  }


  /* =====================================================================================
    5) KHÁCH HÀNG (GIỮ NGUYÊN)
    ===================================================================================== */
  function getAllCustomersData() {
    const ss = getSpreadsheet_();
    const sh = ss.getSheetByName(SHEET_KH);
    if (!sh) return { headers: [], rows: [], agg: {} };

    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow < 1) return { headers: [], rows: [], agg: {} };

    const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    const dataRows = lastRow > 1 ? sh.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];

    const aggMap = buildCustomerAggFromDebt_();

    return { headers: headers, rows: dataRows, agg: aggMap };
  }
  function getCustomerByPhone_(phone) {
    const target = normalizePhone_(phone);
    if (!target) return null;

    // cache theo SĐT (giảm scan sheet KH khi user gõ)
    try {
      const cache = CacheService.getScriptCache();
      const key = "KH_BY_PHONE_V1:" + target;
      const cached = cache.get(key);
      if (cached) {
        if (cached === "NULL") return null;
        try { return JSON.parse(cached); } catch(e) {}
      }
    } catch(e) {}

    const sh = getSheet_(SHEET_KH);
    if (!sh) return null;

    const vals = sh.getDataRange().getValues();
    for (let i = 1; i < vals.length; i++) {
      const sdtRaw  = (vals[i][0] || "").toString().trim();
      const sdtNorm = normalizePhone_(sdtRaw);
      if (sdtNorm && sdtNorm === target) {
        const display = sdtRaw || ("0" + sdtNorm);
        const out = { phone: display, ten: vals[i][1] || "", diaChi: vals[i][2] || "" };
        try {
          CacheService.getScriptCache().put("KH_BY_PHONE_V1:" + target, JSON.stringify(out), 600);
        } catch(e) {}
        return out;
      }
    }

    try {
      CacheService.getScriptCache().put("KH_BY_PHONE_V1:" + target, "NULL", 300);
    } catch(e) {}
    return null;
  }


  function findCustomerByPhone(phone) {
    if (!phone) return null;
    return getCustomerByPhone_(phone) || null;
  }

  function upsertCustomer_(phone, name, address, allowOverwrite) {
    const ss = getSpreadsheet_();
    const sh = ss.getSheetByName(SHEET_KH);
    if (!sh) return;

    const p = String(phone || "").trim();
    const n = String(name || "").trim();
    const a = String(address || "").trim();
    if (!p) return;

    const targetNorm   = normalizePhone_(p);
    const displayPhone = targetNorm ? ("0" + targetNorm) : p;

    try { CacheService.getScriptCache().remove("KH_INDEX_V1"); } catch(e) {}

    try { if (targetNorm) CacheService.getScriptCache().remove("KH_BY_PHONE_V1:" + targetNorm); } catch(e) {}

    const last = sh.getLastRow();
    if (last >= 2) {
      const vals = sh.getRange(2, 1, last - 1, 3).getValues();
      for (let i = 0; i < vals.length; i++) {
        const rowIndex = 2 + i;
        const sdtRaw   = (vals[i][0] || "").toString().trim();
        const sdtNorm  = normalizePhone_(sdtRaw);

        if (sdtNorm && sdtNorm === targetNorm) {
          const curName = vals[i][1] || "";
          const curAddr = vals[i][2] || "";

          if (displayPhone && displayPhone !== sdtRaw) sh.getRange(rowIndex, 1).setValue(displayPhone);

          if (allowOverwrite) {
            sh.getRange(rowIndex, 2).setValue(n || curName);
            sh.getRange(rowIndex, 3).setValue(a || curAddr);
          } else {
            if (!curName && n) sh.getRange(rowIndex, 2).setValue(n);
            if (!curAddr && a) sh.getRange(rowIndex, 3).setValue(a);
          }
          return;
        }
      }
    }

    sh.getRange(sh.getLastRow() + 1, 1, 1, 3).setValues([[displayPhone, n, a]]);
  }

// ===== SEARCH KHÁCH HÀNG (autocomplete tên/sđt/địa chỉ) =====
function ui_searchCustomers(q, limit) {
  q = String(q || "").trim();
  limit = Math.min(Math.max(Number(limit || 12) || 12, 1), 30);
  if (!q) return [];

  const idx = getCustomerIndexCached_();
  if (!idx.length) return [];

  const qPhone = normalizePhone_(q);
  const qText  = normalizeText_(q);

  const out = [];
  for (let i = 0; i < idx.length; i++) {
    const c = idx[i];

    let score = 0;
    if (qPhone) {
      const pos = c.p.indexOf(qPhone);
      if (pos >= 0) score = 1000 - pos;
    } else if (qText) {
      const posName = c.t.indexOf(qText);
      if (posName >= 0) score = 900 - posName;
      else {
        const posAddr = c.a.indexOf(qText);
        if (posAddr >= 0) score = 600 - posAddr;
      }
    }

    if (score > 0) out.push({ score: score, ten: c.ten, phone: c.phone, diaChi: c.diaChi });
  }

  out.sort((x, y) => y.score - x.score);
  return out.slice(0, limit).map(x => ({ ten: x.ten, phone: x.phone, diaChi: x.diaChi }));
}

function getCustomerIndexCached_() {
  const cache = CacheService.getScriptCache();
  const key = "KH_INDEX_V1";
  const hit = cache.get(key);
  if (hit) {
    try { return JSON.parse(hit) || []; } catch(e) {}
  }
  const idx = buildCustomerIndex_();
  cache.put(key, JSON.stringify(idx), 600); // 10 phút
  return idx;
}

function buildCustomerIndex_() {
  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName(SHEET_KH);
  if (!sh) return [];

  const last = sh.getLastRow();
  if (last < 2) return [];

  // Cột: A=SĐT, B=Tên, C=Địa chỉ
  const vals = sh.getRange(2, 1, last - 1, 3).getValues();

  const idx = [];
  for (let i = 0; i < vals.length; i++) {
    const r = vals[i];
    const phoneRaw = String(r[0] || "").trim();
    const ten      = String(r[1] || "").trim();
    const diaChi   = String(r[2] || "").trim();

    const p = normalizePhone_(phoneRaw);
    if (!p && !ten && !diaChi) continue;

    idx.push({
      phone: phoneRaw,
      ten: ten,
      diaChi: diaChi,
      p: p,
      t: normalizeText_(ten),
      a: normalizeText_(diaChi)
    });
  }
  return idx;
}
// ===== END SEARCH KHÁCH HÀNG =====



  /* =====================================================================================
    6) ĐẾM ĐƠN TRÙNG (Ngày + SĐT) TRONG CÔNG NỢ (GIỮ NGUYÊN)
    ===================================================================================== */
  function countOrdersByDateAndPhone(ngayStr, phone) {
    try {
      if (!ngayStr || !phone) return { ok: true, count: 0 };

      const ss = getSpreadsheet_();
      const sh = ss.getSheetByName(SHEET_DEBT);
      if (!sh) return { ok: true, count: 0 };

      const lastRow = sh.getLastRow();
      if (lastRow < DEBT_DATA_START) return { ok: true, count: 0 };

      const tz = Session.getScriptTimeZone();
      const d = parseYmdToDate_(ngayStr);
      if (!d) return { ok: false, error: "Ngày không hợp lệ." };

      const targetYmd = fmtDateYmd_(d, tz);
      const targetPhoneNorm = normalizePhone_(phone);

      const vals = sh.getRange(DEBT_DATA_START, 1, lastRow - DEBT_DATA_START + 1, DEBT_LAST_COL).getValues();
      let count = 0;

      for (let i = 0; i < vals.length; i++) {
        const row = vals[i];
        const dateVal = row[DEBT_COL_DATE - 1];
        const info    = String(row[DEBT_COL_INFO - 1] || "");
        if (!dateVal || !info) continue;

        const ymd = fmtDateYmd_(dateVal, tz);
        if (ymd !== targetYmd) continue;

        const phoneRaw = extractPhoneFromInfo_(info) || "";
        if (!phoneRaw) continue;

        const phoneNorm = normalizePhone_(phoneRaw);
        if (phoneNorm === targetPhoneNorm) count++;
      }

      return { ok: true, count: count };
    } catch (e) {
      return { ok: false, error: e && e.message ? e.message : String(e) };
    }
  }


  function getDataSheetMap_(shData) {
  shData = normalizeDataSheet_(shData);

  var lastCol = Math.max(1, shData.getLastColumn());
  var headers = shData.getRange(1, 1, 1, lastCol).getValues()[0] || [];

  function normHeader(x) {
    // normalizeText_ removes accents + lowercases; also collapse whitespace/newlines
    return normalizeText_(String(x || "").replace(/\n/g, " ").replace(/\s+/g, " "));
  }
  var norm = headers.map(normHeader);

  function findIdx(pred) {
    for (var i = 0; i < norm.length; i++) if (pred(norm[i] || "")) return i + 1;
    return -1;
  }

  // detect leading index/STT/blank column
  var base = 1;
  if (norm.length && (!norm[0] || norm[0] === "stt" || norm[0].indexOf("unnamed") >= 0)) {
    base = 2;
  }

  var idxNgay  = findIdx(h => h.indexOf("ngay") >= 0);
  var idxTenKH = findIdx(h => h.indexOf("khach hang") >= 0);
  var idxSdt   = findIdx(h => h.indexOf("so dien thoai") >= 0 || h === "sdt");
  var idxDc    = findIdx(h => h.indexOf("dia chi") >= 0);

  var idxDgmam = findIdx(h => h.indexOf("don gia/mam") >= 0 || h.indexOf("don gia mam") >= 0 || h.indexOf("don gia mam") >= 0);
  var idxTen   = findIdx(h => h.indexOf("ten hang") >= 0 || h.indexOf("ten mon") >= 0);
  var idxDvt   = findIdx(h => h === "dvt" || h === "dvt" || h.indexOf("don vi") >= 0);
  var idxSl    = findIdx(h => h === "sl" || h.indexOf("so luong") >= 0);
  var idxDg    = findIdx(h => h.indexOf("don gia") >= 0 && h.indexOf("mam") < 0);
  var idxTt    = findIdx(h => h.indexOf("thanh tien") >= 0);

  var idxOid   = findIdx(h => h === "id" || h.indexOf("order") >= 0 || h.indexOf("ma don") >= 0);
  var idxLoai  = findIdx(h => h.indexOf("loai mon") >= 0 || h.indexOf("loai") === 0);
  var idxSt    = findIdx(h => h === "st" || h.indexOf("status") >= 0 || h.indexOf("trang thai") >= 0);
  if (idxSt < 1) idxSt = 0;

  // fallback theo schema 12 cột (không tính cột STT nếu có)
  if (idxNgay  < 1) idxNgay  = base + 0;
  if (idxTenKH < 1) idxTenKH = base + 1;
  if (idxSdt   < 1) idxSdt   = base + 2;
  if (idxDc    < 1) idxDc    = base + 3;
  if (idxDgmam < 1) idxDgmam = base + 4;
  if (idxTen   < 1) idxTen   = base + 5;
  if (idxDvt   < 1) idxDvt   = base + 6;
  if (idxSl    < 1) idxSl    = base + 7;
  if (idxDg    < 1) idxDg    = base + 8;
  if (idxTt    < 1) idxTt    = base + 9;
  if (idxOid   < 1) idxOid   = base + 10;
  if (idxLoai  < 1) idxLoai  = base + 11;

  return { idxNgay, idxTenKH, idxSdt, idxDc, idxDgmam, idxTen, idxDvt, idxSl, idxDg, idxTt, idxOid, idxLoai, idxSt };
}


  /* =====================================================================================
    7) DATA SHEET ITEMS (GIỮ NGUYÊN)
    ===================================================================================== */
  function appendItemsToData_(shData, ngay, tenKH, sdt, diaChi, items, orderId, donGiaMam) {
    shData = normalizeDataSheet_(shData);
    if (!Array.isArray(items) || !items.length) return { startRow: 0, endRow: 0, count: 0 };

    var m = getDataSheetMap_(shData);
    var width = Math.max(
      m.idxSt, m.idxLoai, m.idxOid, m.idxTt, m.idxDg, m.idxSl, m.idxDvt, m.idxTen,
      m.idxDgmam, m.idxDc, m.idxSdt, m.idxTenKH, m.idxNgay
    );

    var rows = [];
    for (var i = 0; i < items.length; i++) {
      var it = items[i] || {};
      var tenMon = String(it.tenMon || it.ten || "").trim();
      if (!tenMon) continue;

      var dvt = String(it.dvt || "").trim();
      var sl  = toMoneyNumber_(it.sl || 0) || 0;
      var dg  = toMoneyNumber_(it.donGia || it.dg || 0) || 0;
      var amount = toMoneyNumber_(it.thanhTien || it.amount || it.tt || 0) || 0;

      // ✅ auto-calc line total if client did not send it
      if (!amount && sl > 0 && dg >= 0) amount = sl * dg;

      if (sl <= 0) continue;

      var loaiMon = String(it.loaiMon || it.loai || "").trim();
      var stItem  = String(it.status || it.st || "").trim();

      var r = new Array(width);
      for (var j = 0; j < width; j++) r[j] = "";

      // ngày/kh/sdt/đc
      if (m.idxNgay)  r[m.idxNgay - 1]  = ngay instanceof Date ? new Date(ngay) : ngay;
      if (m.idxTenKH) r[m.idxTenKH - 1] = tenKH;
      if (m.idxSdt)   r[m.idxSdt - 1]   = sdt;
      if (m.idxDc)    r[m.idxDc - 1]    = diaChi;

      // item
      if (m.idxTen)   r[m.idxTen - 1]   = tenMon;
      if (m.idxDvt)   r[m.idxDvt - 1]   = dvt;
      if (m.idxSl)    r[m.idxSl - 1]    = sl;
      if (m.idxDg)    r[m.idxDg - 1]    = dg;
      if (m.idxTt)    r[m.idxTt - 1]    = amount;

      // đơn giá mâm (lưu trên từng dòng để trace)
      if (m.idxDgmam) r[m.idxDgmam - 1] = Number(donGiaMam || 0) || 0;

      if (m.idxLoai)  r[m.idxLoai - 1]  = loaiMon;
      if (m.idxSt)    r[m.idxSt - 1]    = stItem;
      if (m.idxOid)   r[m.idxOid - 1]   = orderId;

      rows.push(r);
    }

    if (!rows.length) return { startRow: 0, endRow: 0, count: 0 };

    // Append nhanh: ưu tiên Sheets API (Values.append) để lấy updatedRange => start/end row chính xác
    var startRow = shData.getLastRow() + 1;
    var endRow = startRow + rows.length - 1;

    var wrote = false;
    try {
      if (hasSheetsApi_()) {
        var res = sh_valuesAppend_(SHEET_DATA + "!A1", rows);
        if (res && res.updates && res.updates.updatedRange) {
          var rg = String(res.updates.updatedRange);
          // ví dụ: Sheet!A123:M130
          var mm = rg.match(/!A(\d+):[A-Z]+(\d+)$/);
          if (mm) {
            startRow = Number(mm[1] || startRow) || startRow;
            endRow = Number(mm[2] || endRow) || endRow;
          }
          wrote = true;
        }
      }
    } catch (e) {}

    if (!wrote) {
      shData.getRange(startRow, 1, rows.length, width).setValues(rows);
    }

    // Upsert menu (Danh mục hàng)
    upsertMenuFromItems_(items);

    return { startRow: startRow, endRow: endRow, count: rows.length };
  }




  
function getItemsBySpan_(shData, startRow, endRow) {
  startRow = Number(startRow || 0);
  endRow = Number(endRow || 0);
  if (!startRow || !endRow || endRow < startRow) return [];

  shData = normalizeDataSheet_(shData);
  var m = getDataSheetMap_(shData);
  var width = Math.max(m.idxSt, m.idxLoai, m.idxOid, m.idxTt, m.idxDg, m.idxSl, m.idxDvt, m.idxTen, m.idxDgmam, m.idxDc, m.idxSdt, m.idxTenKH, m.idxNgay);

  var nRows = endRow - startRow + 1;
  var values = shData.getRange(startRow, 1, nRows, width).getValues();

  var out = [];
  for (var i = 0; i < values.length; i++) {
    var r = values[i] || [];
    // bỏ qua dòng đã DELETED nếu có status
    if (m.idxSt) {
      var st = String(r[m.idxSt - 1] || "").trim().toUpperCase();
      if (st === "DELETED") continue;
    }
    out.push({
      tenMon: String(r[m.idxTen - 1] || ""),
      dvt: String(r[m.idxDvt - 1] || ""),
      sl: Number(r[m.idxSl - 1] || 0) || 0,
      donGia: toMoneyNumber_(r[m.idxDg - 1] || 0),
      thanhTien: toMoneyNumber_(r[m.idxTt - 1] || 0),
      loaiMon: String(r[m.idxLoai - 1] || "")
    });
  }
  return out;
}
function getItemsByOrderId_(shDataMaybe, orderIdMaybe) {
  var shData = shDataMaybe;
  var orderId = orderIdMaybe;

  if (orderId === undefined && (typeof shDataMaybe === "string" || typeof shDataMaybe === "number")) {
    orderId = shDataMaybe;
    shData = null;
  }

  shData = normalizeDataSheet_(shData);
  var m = getDataSheetMap_(shData);

  var lastRow = shData.getLastRow();
  if (lastRow < 2) return [];

  var oid = String(orderId || "").trim();
  if (!oid) return [];

  // width đủ để đọc các cột cần thiết
  var width = Math.max(
    m.idxLoai || 1, m.idxOid || 1, m.idxTt || 1, m.idxDg || 1,
    m.idxSl || 1, m.idxDvt || 1, m.idxTen || 1, m.idxSt || 1
  );

  // ✅ Tìm span các dòng có orderId bằng TextFinder -> chỉ đọc đúng block đó, không đọc full sheet
  var span = findOrderIdRowSpan_(shData, m.idxOid, oid);

  // fallback nếu TextFinder không ra (ít gặp)
  if (!span) {
    var oidVals = shData.getRange(2, m.idxOid, lastRow - 1, 1).getValues();
    var minR = 0, maxR = 0, cnt = 0;
    for (var i = 0; i < oidVals.length; i++) {
      if (String(oidVals[i][0] || "").trim() === oid) {
        var r = i + 2;
        if (!minR || r < minR) minR = r;
        if (!maxR || r > maxR) maxR = r;
        cnt++;
      }
    }
    if (!cnt) return [];
    span = { startRow: minR, endRow: maxR, count: cnt };
  }

  var nRows = span.endRow - span.startRow + 1;
  var values = shData.getRange(span.startRow, 1, nRows, width).getValues();

  var out = [];
  for (var j = 0; j < values.length; j++) {
    var r = values[j];

    // lọc theo orderId đúng
    if (String(r[m.idxOid - 1] || "").trim() !== oid) continue;

    // bỏ qua dòng đã DELETED nếu có status
    if (m.idxSt) {
      var st = String(r[m.idxSt - 1] || "").trim().toUpperCase();
      if (st === "DELETED") continue;
    }

    out.push({
      tenMon: String(r[m.idxTen - 1] || ""),
      dvt: String(r[m.idxDvt - 1] || ""),
      sl: Number(r[m.idxSl - 1] || 0) || 0,
      donGia: toMoneyNumber_(r[m.idxDg - 1] || 0),
      thanhTien: toMoneyNumber_(r[m.idxTt - 1] || 0),
      loaiMon: String(r[m.idxLoai - 1] || "")
    });
  }
  return out;
}


function getItemsByRowSpan_(shData, startRow, endRow, orderIdOpt) {
  startRow = Number(startRow || 0);
  endRow = Number(endRow || 0);
  if (!shData || !startRow || !endRow || endRow < startRow) return [];

  shData = normalizeDataSheet_(shData);
  const m = getDataSheetMap_(shData);
  if (!m) return [];

  const width = Math.max(
    m.idxTt || 1, m.idxDg || 1, m.idxSl || 1, m.idxDvt || 1, m.idxTen || 1, m.idxLoai || 1, m.idxOid || 1, m.idxSt || 1
  );

  const nRows = endRow - startRow + 1;
  const values = shData.getRange(startRow, 1, nRows, width).getValues();

  const oid = String(orderIdOpt || "").trim();
  const out = [];
  for (let j = 0; j < values.length; j++) {
    const r = values[j];

    if (oid && m.idxOid) {
      if (String(r[m.idxOid - 1] || "").trim() !== oid) continue;
    }

    if (m.idxSt) {
      const st = String(r[m.idxSt - 1] || "").trim().toUpperCase();
      if (st === "DELETED") continue;
    }

    out.push({
      tenMon: String(r[m.idxTen - 1] || ""),
      dvt: String(r[m.idxDvt - 1] || ""),
      sl: Number(r[m.idxSl - 1] || 0) || 0,
      donGia: toMoneyNumber_(r[m.idxDg - 1] || 0),
      thanhTien: toMoneyNumber_(r[m.idxTt - 1] || 0),
      loaiMon: String(r[m.idxLoai - 1] || "")
    });
  }
  return out;
}





function findOrderIdRowSpan_(shData, idxOid, orderId) {
  try {
    if (!shData || !idxOid || !orderId) return null;
    const lastRow = shData.getLastRow();
    if (lastRow < 2) return null;

    const finder = shData.getRange(2, idxOid, lastRow - 1, 1)
      .createTextFinder(String(orderId))
      .matchEntireCell(true);

    const found = finder.findAll();
    if (!found || !found.length) return null;

    let minR = found[0].getRow();
    let maxR = minR;
    for (let i = 1; i < found.length; i++) {
      const r = found[i].getRow();
      if (r < minR) minR = r;
      if (r > maxR) maxR = r;
    }
    return { startRow: minR, endRow: maxR, count: found.length };
  } catch (e) {
    return null;
  }
}

function deleteItemsByOrderId_(shData, orderId) {
  if (!shData || !orderId) return;

  shData = normalizeDataSheet_(shData);
  const m = getDataSheetMap_(shData);
  if (!m || !m.idxOid) return;

  orderId = String(orderId).trim();
  if (!orderId) return;

  // Ưu tiên xoá theo span (contiguous block) để nhanh
  const span = findOrderIdRowSpan_(shData, m.idxOid, orderId);
  if (span && span.startRow && span.endRow && span.endRow >= span.startRow) {
    const n = span.endRow - span.startRow + 1;
    try {
      shData.deleteRows(span.startRow, n);
      return;
    } catch (e) {
      // fallback tiếp bên dưới
    }
  }

  // Fallback: quét cột order_id và xoá tất cả dòng match (xoá từ dưới lên để không lệch index)
  const lastRow = shData.getLastRow();
  if (lastRow < 2) return;

  const oidVals = shData.getRange(2, m.idxOid, lastRow - 1, 1).getValues();
  const rowsToDelete = [];
  for (let i = 0; i < oidVals.length; i++) {
    if (String(oidVals[i][0] || "").trim() === orderId) rowsToDelete.push(i + 2);
  }
  if (!rowsToDelete.length) return;

  rowsToDelete.sort((a,b)=>b-a);
  for (let k = 0; k < rowsToDelete.length; k++) {
    try { shData.deleteRow(rowsToDelete[k]); } catch(e) {}
  }
}

  function calcTotalsFromItems_(items, soMam, kmSoTien) {
    var sum1Mam = 0;
    (items || []).forEach(function(it){
      var sl = toMoneyNumber_(it.sl || 0) || 0;
      var dg = toMoneyNumber_(it.donGia || it.dg || it.gia || 0) || 0;
      sum1Mam += sl * dg;
    });
    var donGiaMam = sum1Mam;
    var tongDon = donGiaMam * (Number(soMam || 1) || 1);
    var doanhSo = Math.max(0, tongDon - (Number(kmSoTien || 0) || 0));
    return { donGiaMam: donGiaMam, tongDon: tongDon, doanhSo: doanhSo };
  }

function sanitizeTextOneLine_(s) {
  s = String(s || "");
  s = s.replace(/\s+/g, " ").trim();      // gom khoảng trắng
  s = s.replace(/[\r\n\t]+/g, " ").trim(); // không cho xuống dòng/tab
  return s;
}

function normalizePhoneStrict_(raw) {
  const s = String(raw || "").trim();
  if (!s) return "";

  // Cho phép người dùng nhập: +84 941 068 777, 0941-068-777, (0941)068777...
  // Nhưng cấm chữ cái
  if (/[A-Za-z]/.test(s)) {
    throw new Error("Số điện thoại chỉ được chứa chữ số (có thể bắt đầu bằng +84). Vui lòng bỏ ký tự chữ.");
  }

  // chỉ giữ + và số
  let cleaned = s.replace(/[^\d+]/g, "");

  // chỉ cho phép 1 dấu + ở đầu
  if (cleaned.includes("+") && cleaned[0] !== "+") {
    throw new Error("Số điện thoại không hợp lệ. Dấu + chỉ được ở đầu.");
  }
  if ((cleaned.match(/\+/g) || []).length > 1) {
    throw new Error("Số điện thoại không hợp lệ. Dấu + chỉ được xuất hiện 1 lần.");
  }

  // chuyển +84 -> 0
  if (cleaned.startsWith("+84")) cleaned = "0" + cleaned.slice(3);

  // giờ chỉ còn số
  cleaned = cleaned.replace(/[^\d]/g, "");

  // validate độ dài cơ bản
  if (cleaned.length < 9 || cleaned.length > 12) {
    throw new Error("Số điện thoại không hợp lệ (độ dài " + cleaned.length + ").");
  }

  return cleaned;
}

function validateHeaderFields_(tenKH, sdt, diaChi) {
  tenKH = sanitizeTextOneLine_(tenKH);
  diaChi = sanitizeTextOneLine_(diaChi);

  if (!tenKH) throw new Error("Thiếu tên khách hàng.");
  if (!diaChi) throw new Error("Thiếu địa chỉ.");

  // chặn ký tự gây lỗi parse/ID
  if (/[|]/.test(tenKH) || /[|]/.test(diaChi)) {
    throw new Error('Tên/địa chỉ không được chứa ký tự "|".');
  }

  sdt = normalizePhoneStrict_(sdt);

  return { tenKH, sdt, diaChi };
}


  /* =====================================================================================
    8) SAVE ORDER (GIỮ NGUYÊN LOGIC, CHỈ SỬA PENDING)
    ===================================================================================== */
 function saveOrder(payload) {
  payload = payload || {};
  payload.meta = payload.meta || {};
  var ctx = assertCanEdit_(payload);
  
  var h = payload.header || {};
  var status = String(h.trangThai || "").trim();

  // Kiểm tra nếu payload yêu cầu tạo dòng mới (không có rowIndex)
  var pendingRow = Number(payload.meta.rowIndex || payload.meta.row || 0);

  if (pendingRow >= 2) {
    return updatePendingOrder(pendingRow, payload);
  }

  if (!status) return savePendingOrder_(payload);
  return saveConfirmedOrder_(payload);
}

function saveConfirmedOrder_(orderPayload) {
  try {
    const h = orderPayload.header || orderPayload;
    const items = orderPayload.items || [];
    const datCoc = toMoneyNumber_(h.datCoc || 0);
    const ngay = parseDateCell_(h.ngay || new Date());
    const sdt = normalizePhone_(h.sdt || "");
    const info = (h.tenKH || "") + " - " + sdt + " - " + (h.diaChi || "");
    const orderId = h.orderId || makeOrderId_(ngay, sdt);
    
    const calc = calcTotalsFromItems_(items, h.soMam, h.kmSoTien);
    const ss = getSpreadsheet_();
    const shDebt = ss.getSheetByName(SHEET_DEBT);
    const shData = ss.getSheetByName(SHEET_DATA);

    // Lưu chi tiết món (nếu có)
    let span = {startRow:0, endRow:0, count:0};
    if (items.length > 0) {
      span = appendItemsToData_(shData, ngay, h.tenKH, sdt, h.diaChi, items, orderId, calc.donGiaMam);
    }

    const doanhSo = Math.max(0, calc.doanhSo || 0);
    let trangThai = String(h.trangThai || STATUS_DEBT).trim();

    // Nếu khuyến mãi 100% (doanhSo = 0), mặc định "Đã thanh toán"
    if (doanhSo === 0 && calc.tongDon > 0) {
      trangThai = STATUS_PAID;
    }

    // Công nợ = Doanh số (sau KM) - Đặt cọc
    const congNo = (trangThai === STATUS_PAID) ? 0 : Math.max(0, doanhSo - datCoc);

    const debtRow = [
      ngay, info, h.soMam, calc.donGiaMam, calc.tongDon,
      h.kmNoiDung, h.kmSoTien, doanhSo, congNo,
      trangThai, (trangThai === STATUS_PAID ? new Date() : ""),
      orderPayload.meta.username || "", orderId,
      span.count, span.startRow, span.endRow, datCoc
    ];

    shDebt.appendRow(debtRow);
    upsertCustomer_(sdt, h.tenKH, h.diaChi);
    return { ok: true, orderId: orderId };
  } catch (e) { return { ok: false, error: e.message }; }
}

  /* ==========================
    PENDING SAVE (SỬA CẦN THIẾT)
    - Giữ chống duplicate
    - Không đổi nghiệp vụ
    ========================== */
 function savePendingOrder_(orderPayload) {
  try {
    const h = orderPayload.header || orderPayload;
    const items = orderPayload.items || [];
    const datCoc = toMoneyNumber_(h.datCoc || 0);
    const ngay = parseDateCell_(h.ngay || new Date());
    const sdt = normalizePhone_(h.sdt || "");
    const info = (h.tenKH || "") + " - " + sdt + " - " + (h.diaChi || "");
    const orderId = h.orderId || makeOrderId_(ngay, sdt);
    
    let calc = calcTotalsFromItems_(items, h.soMam, h.kmSoTien);
    const shP = getPendingSheet_();
    const shData = getSheet_(SHEET_DATA);

    let span = {startRow:0, endRow:0, count:0};
    let isOnlyDeposit = (items.length === 0 && datCoc > 0);

    if (items.length > 0) {
      span = appendItemsToData_(shData, ngay, h.tenKH, sdt, h.diaChi, items, orderId, calc.donGiaMam);
    }

    // Nếu chỉ có đặt cọc, set tổng đơn = đặt cọc
    if (isOnlyDeposit) {
      calc = { donGiaMam: datCoc, tongDon: datCoc, doanhSo: datCoc };
    }

    // Đơn chờ: Công nợ = Tổng đơn - KM - Đặt cọc
    const pendingCongNo = Math.max(0, calc.tongDon - h.kmSoTien - datCoc);

    const statusText = isOnlyDeposit ? "Đặt cọc" : "Chưa thanh toán";

    const pendingRow = [
      ngay, info, h.soMam, calc.donGiaMam, calc.tongDon,
      h.kmNoiDung, h.kmSoTien, 0, pendingCongNo,
      statusText, "", orderPayload.meta.username || "",
      orderId, span.count, span.startRow, span.endRow, datCoc
    ];

    shP.appendRow(pendingRow);
    upsertCustomer_(sdt, h.tenKH, h.diaChi);
    return { ok: true, orderId: orderId };
  } catch (e) { return { ok: false, error: e.message }; }
}


  /* =====================================================================================
    9) PENDING LIST + COUNT ITEM (SỬA CHẠY ỔN, KHÔNG VỚ VẨN)
    - Chuẩn hoá: chỉ dùng 1 map đếm món theo orderId
    - Chuẩn hoá: listPendingByPhone / listPendingByPhoneFull cùng output
    ===================================================================================== */

  /*********** ĐẾM SỐ MÓN THEO ORDERID (scan 1 lần) ************/
  function buildOrderItemCountByOrderId_() {
    var ss = getSpreadsheet_();
    var sh = ss.getSheetByName(SHEET_DATA);
    if (!sh || sh.getLastRow() < 2) return {};

    var m = getDataSheetMap_(sh);
    if (!m || !m.idxOid) return {};

    var mp = {};
    try {
      if (hasSheetsApi_()) {
        const colOid = colToA1_(m.idxOid);
        const ranges = [SHEET_DATA + "!" + colOid + "2:" + colOid];
        let colSt = null;
        if (m.idxSt && m.idxSt >= 1) {
          colSt = colToA1_(m.idxSt);
          ranges.push(SHEET_DATA + "!" + colSt + "2:" + colSt);
        }
        const got = sh_valuesBatchGet_(ranges);
        const oidVals = (got && got[0]) ? got[0] : [];
        const stVals = (colSt && got && got[1]) ? got[1] : null;

        for (var i = 0; i < oidVals.length; i++) {
          var oid = String((oidVals[i] && oidVals[i][0]) || "").trim();
          if (!oid) continue;

          if (stVals) {
            var st = (stVals[i] && stVals[i][0]) || "";
            var stStr = String(st).trim().toLowerCase();
            if (stStr === "0" || stStr === "false") continue;
          }
          mp[oid] = (mp[oid] || 0) + 1;
        }
        return mp;
      }
    } catch(e) {}

    // Fallback SpreadsheetApp
    var lastRow = sh.getLastRow();
    var valsOid = sh.getRange(2, m.idxOid, lastRow - 1, 1).getValues();
    var valsSt = (m.idxSt && m.idxSt >= 1) ? sh.getRange(2, m.idxSt, lastRow - 1, 1).getValues() : null;

    for (var j = 0; j < valsOid.length; j++) {
      var oid2 = String(valsOid[j][0] || "").trim();
      if (!oid2) continue;
      if (valsSt) {
        var st2 = valsSt[j][0];
        var st2s = String(st2).trim().toLowerCase();
        if (st2s === "0" || st2s === "false") continue;
      }
      mp[oid2] = (mp[oid2] || 0) + 1;
    }
    return mp;
  }



  /* Core list pending: 1 nơi, các wrapper gọi lại để UI khỏi lệch */
  function listPendingByPhoneCore_(phoneNorm) {
  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName(SHEET_PENDING);
  if (!sh) return [];

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const off = getLeadingIndexOffset_(sh);
  const baseCol = 1 + off;

  let vals = sh.getRange(2, baseCol, lastRow - 1, PENDING_LAST_COL).getValues();

  // fallback: nếu dữ liệu đang nằm từ cột A (schema cũ)
  if (off && (!vals || !vals.length)) {
    vals = sh.getRange(2, 1, lastRow - 1, PENDING_LAST_COL).getValues();
  }

  const out = [];
  for (let i = 0; i < vals.length; i++) {
    const r = vals[i] || [];
    const info = String(r[PENDING_COL_INFO - 1] || "");
    const sdt = normalizePhone_(extractPhoneFromInfo_(info));
    if (!sdt) continue;
    if (phoneNorm && sdt.indexOf(phoneNorm) === -1) continue;

    out.push({
      rowIndex: i + 2,
      date: r[PENDING_COL_DATE - 1],
      info: info,
      soMam: r[PENDING_COL_SOMAM - 1],
      tongDon: r[PENDING_COL_TONG_DON - 1],
      trangThai: STATUS_DEBT,
      orderId: r[PENDING_COL_ORDER_ID - 1]
    });
  }
  return out;
}

  /* API mà UI hay gọi */
  function listPendingByPhone(phone) {
    return listPendingByPhoneCore_(phone);
  }

  /* API “full” giữ lại cho tương thích */
  function listPendingByPhoneFull(phone) {
    return listPendingByPhoneCore_(phone);
  }

  /* Pending recent giữ nguyên, nhưng dùng cùng map + sort chuẩn */
  function listPendingRecent(maxRows) {
    maxRows = Number(maxRows || 50) || 50;

    const sh = getPendingSheet_();
    const last = sh.getLastRow();
    if (last < 2) return [];

    const tz = Session.getScriptTimeZone();
    const start = Math.max(2, last - maxRows + 1);
    const n = last - start + 1;

    const vals = sh.getRange(start, 1, n, 13).getValues(); // A:N
    const out = [];
    for (let i = 0; i < vals.length; i++) {
      const r = vals[i];
      const rowIndex = start + i;

      const dateRaw = r[0] ? new Date(r[0]) : null;
      const dateStr = dateRaw ? Utilities.formatDate(dateRaw, tz, "dd/MM/yyyy") : "";

      const orderId = String(r[12] || "").trim();
      const soMon = orderId ? (itemCountMap[orderId] || 0) : 0;

      out.push({
        rowIndex: rowIndex,
        row: rowIndex,
        dateRaw: dateRaw ? dateRaw.getTime() : 0,
        dateStr: dateStr,
        date: dateStr,
        info: String(r[1] || ""),
        soMam: Number(r[2] || 0) || 0,
        soMon: soMon,
        donGiaMam: Number(r[3] || 0) || 0,
        tongDon: Number(r[4] || 0) || 0,
        trangThai: "Đơn chờ",
        orderId: orderId
      });
    }

    out.sort((a, b) => (b.dateRaw || 0) - (a.dateRaw || 0));
    out.forEach(x => { delete x.dateRaw; });
    return out;
  }


  /* =====================================================================================
    10) PENDING DETAIL + UPDATE (GIỮ NGUYÊN, CHỈ DỌN CHO ĐỒNG BỘ)
    ===================================================================================== */
  function getOrderDetailByPendingRow(rowIndex) {
  const ss = getSpreadsheet_();
  const shP = ss.getSheetByName(SHEET_PENDING);
  if (!shP) return { ok: false, error: "Không tìm thấy sheet '" + SHEET_PENDING + "'." };

  const lastRow = shP.getLastRow();
  if (!rowIndex || rowIndex < 2 || rowIndex > lastRow) return { ok: false, error: "Dòng đơn chờ không hợp lệ." };

  // ưu tiên schema có cột index (A trống/STT), data bắt đầu từ cột B
  let off = getLeadingIndexOffset_(shP);
  let baseCol = 1 + off;

  let r = shP.getRange(rowIndex, baseCol, 1, PENDING_LAST_COL).getValues()[0] || [];
  let orderId = String(r[PENDING_COL_ORDER_ID - 1] || "").trim();

  // fallback schema cũ: data bắt đầu từ cột A
  if (!orderId && off) {
    off = 0;
    baseCol = 1;
    r = shP.getRange(rowIndex, baseCol, 1, PENDING_LAST_COL).getValues()[0] || [];
    orderId = String(r[PENDING_COL_ORDER_ID - 1] || "").trim();
  }

  if (!orderId) return { ok: false, error: "Đơn chờ thiếu Mã đơn (orderId)." };

  const ngay = r[PENDING_COL_DATE - 1];
  const info = String(r[PENDING_COL_INFO - 1] || "").trim();
  const soMam = toMoneyNumber_(r[PENDING_COL_SOMAM - 1] || 0);
  const donGiaMam = toMoneyNumber_(r[PENDING_COL_DONGIA_MAM - 1] || 0);
  const tongDon = toMoneyNumber_(r[PENDING_COL_TONG_DON - 1] || 0);
  const kmNoiDung = String(r[PENDING_COL_KM_NOIDUNG - 1] || "").trim();
  const kmSoTien = toMoneyNumber_(r[PENDING_COL_KM_SOTIEN - 1] || 0);
  const thungan = String(r[PENDING_COL_CASHIER - 1] || "").trim();
  const datCoc = toMoneyNumber_(r[PENDING_COL_DEPOSIT - 1] || 0);

  // Debug log để kiểm tra dữ liệu từ sheet
  Logger.log('SHEET DATA DEBUG: ' + JSON.stringify({
    rowIndex, soMam, donGiaMam, tongDon, datCoc,
    rawData: r.slice(0, 17), // Log 17 cột đầu
    datCocColumn: r[16], // Cột 17 (index 16)
    PENDING_COL_DEPOSIT: PENDING_COL_DEPOSIT
  }));

  // Also try console.log for browser
  console.log('SHEET DATA DEBUG:', {
    rowIndex, soMam, donGiaMam, tongDon, datCoc,
    rawData: r.slice(0, 17),
    datCocColumn: r[16],
    PENDING_COL_DEPOSIT: PENDING_COL_DEPOSIT
  });

  const shData = ss.getSheetByName(SHEET_DATA);
  if (!shData) return { ok: false, error: "Không tìm thấy sheet '" + SHEET_DATA + "'." };

  const items = getItemsByOrderId_(shData, orderId);

  const sdt = normalizePhone_(extractPhoneFromInfo_(info));
  const tenKH = getNameFromInfo_(info) || "";
  const diaChi = getAddressFromInfo_(info) || "";

  const header = {
    dateRaw: ngay,
    date: ngay instanceof Date ? Utilities.formatDate(ngay, Session.getScriptTimeZone(), "dd/MM/yyyy") : String(ngay || ""),
    tenKH: tenKH,
    sdt: sdt,
    diaChi: diaChi,
    soMam: soMam,
    donGiaMam: donGiaMam,
    tongDon: tongDon,
    kmNoiDung: kmNoiDung,
    kmSoTien: kmSoTien,
    datCoc: datCoc,
    thungan: thungan,
    trangThai: STATUS_UNPAID,
    status: STATUS_UNPAID,
    orderId: orderId
  };

  // ✅ Cho phép in đơn ngay cả khi không có món (đơn đặt cọc)
  return { ok: true, header: header, items: items || [] };
}

  function updatePendingOrder(rowIndex, payload) {
    try {
      rowIndex = Number(rowIndex || 0);
      payload = payload || {};
      if (!rowIndex || rowIndex < 2) return { ok: false, error: "rowIndex không hợp lệ" };

      const ss = getSpreadsheet_();
      const shP = getPendingSheet_();
      const shDebt = ss.getSheetByName(SHEET_DEBT);
      const shData = ss.getSheetByName(SHEET_DATA);
      if (!shDebt) return { ok: false, error: "Không tìm thấy sheet '" + SHEET_DEBT + "'." };
      if (!shData) return { ok: false, error: "Không tìm thấy sheet '" + SHEET_DATA + "'." };

      const off = getLeadingIndexOffset_(shP);
      const baseCol = 1 + off;
      const old = shP.getRange(rowIndex, baseCol, 1, PENDING_LAST_COL).getValues()[0] || [];
      const oldInfo = String(old[PENDING_COL_INFO - 1] || "");
      const orderId = String(old[PENDING_COL_ORDER_ID - 1] || "").trim();
      if (!orderId) return { ok: false, error: "Không có orderId" };

      const h = payload.header || {};
      const items = payload.items || [];

      let status = String(h.trangThai || "").trim(); // "" = vẫn pending, khác "" = chốt

      // Nếu khuyến mãi 100% (doanh thu = 0), mặc định "Đã thanh toán"
      if (calcFinal.doanhSo === 0 && calcFinal.tongDon > 0) {
        status = STATUS_PAID;
      }
      const ngayStr = String(h.ngay || "").trim();
      const ngay = ngayStr ? (parseYmdToDate_(ngayStr) || new Date(old[0] || new Date())) : new Date(old[0] || new Date());

      const tenKH = String(h.tenKH || "").trim();
      const sdt = String(h.sdt || "").trim();
      const diaChi = String(h.diaChi || "").trim();
      const soMam = Number(h.soMam || 0) || Number(old[2] || 1) || 1;

      const kmNoiDung = String(h.kmNoiDung || "").trim();
      const kmSoTien = Number(h.kmSoTien || 0) || 0;
      const nguoiLap = String((payload && payload.meta && payload.meta.username) || (payload && payload.username) || h.nguoiLap || old[11] || "").trim();
      const datCoc = toMoneyNumber_(h.datCoc || 0);

      let calc = calcTotalsFromItems_(items, soMam, kmSoTien);
      if (!tenKH || !sdt || !diaChi) {
        const tenOld = getNameFromInfo_(oldInfo) || "";
        const sdtOld = extractPhoneFromInfo_(oldInfo) || "";
        const dcOld  = getAddressFromInfo_(oldInfo) || "";

        if (!h.tenKH && tenOld) h.tenKH = tenOld;
        if (!h.sdt && sdtOld) h.sdt = sdtOld;
        if (!h.diaChi && dcOld) h.diaChi = dcOld;
      }

      const finalTen = String(h.tenKH || tenKH || "").trim();
      const finalSdt = String(h.sdt || sdt || "").trim();
      const finalDc  = String(h.diaChi || diaChi || "").trim();

      if (!finalTen) return { ok: false, error: "Thiếu tên khách hàng." };
      if (!finalSdt) return { ok: false, error: "Thiếu số điện thoại." };
      if (!isValidPhone_(finalSdt)) return { ok: false, error: "SĐT không hợp lệ (6–12 số)." };
      if (!finalDc)  return { ok: false, error: "Thiếu địa chỉ." };

      // Cho phép đơn chỉ đặt cọc (không có món)
      if (!Array.isArray(items) || !items.length) {
        if (datCoc <= 0) {
          return { ok: false, error: "Cần ít nhất 1 món ăn hoặc có đặt cọc." };
        }
        // Đơn chỉ đặt cọc - không cần validate thêm
      }

      const isOnlyDeposit = (!Array.isArray(items) || !items.length) && datCoc > 0;

      // Chỉ upsert menu nếu có món ăn
      if (Array.isArray(items) && items.length > 0) {
        upsertMenuFromItems_(items);
      }
      upsertCustomer_(finalSdt, finalTen, finalDc, false);

      const calcFinal = isOnlyDeposit ?
        { donGiaMam: datCoc, tongDon: datCoc, doanhSo: Math.max(0, datCoc - kmSoTien) } :
        calcTotalsFromItems_(items, soMam, kmSoTien);

      let span = { startRow: 0, endRow: 0, count: 0 };
      if (!isOnlyDeposit) {
        deleteItemsByOrderId_(shData, orderId);
        span = appendItemsToData_(shData, ngay, finalTen, finalSdt, finalDc, items, orderId, calcFinal.donGiaMam);
      }

      const itemCount = span && span.count ? span.count : (Array.isArray(items) && items.length > 0 ? items.length : 0);
      const itemStart = span && span.startRow ? span.startRow : 0;
      const itemEnd   = span && span.endRow ? span.endRow : 0;

      // Span debug removed
const info = finalTen + " - " + finalSdt + " - " + finalDc;

      if (!status) {
        const pendingCongNo = Math.max(0, calc.tongDon - kmSoTien - datCoc);
        const rowValuesPending = [
          new Date(ngay),           // A
          info,                    // B
          soMam,                   // C
          calc.donGiaMam,          // D
          calc.tongDon,            // E
          kmNoiDung,               // F
          kmSoTien,                // G
          0,                       // H: Doanh số (đơn chờ)
          pendingCongNo,           // I: Công nợ
          "Chưa thanh toán",       // J: Trạng thái
          "",                      // K: Ngày thanh toán
          nguoiLap || getCurrentUserName_(), // L: Thu ngân
          orderId,                 // M: Mã đơn
          itemCount,               // N
          itemStart,               // O
          itemEnd,
          datCoc                 // P
        ];

        const off2 = getLeadingIndexOffset_(shP);
        const baseCol2 = 1 + off2;
        if (off2) {
          try { shP.getRange(rowIndex, 1).setValue(rowIndex - PENDING_DATA_START + 1); } catch(e) {}
        }
        shP.getRange(rowIndex, baseCol2, 1, rowValuesPending.length).setValues([rowValuesPending]);
        try { resetPendingSheetCache_(); } catch(e) {}

        return { ok: true, moved: false, orderId: orderId };
      }

      const isPaid = (status === STATUS_PAID);
      const congNo = isPaid ? 0 : calcFinal.doanhSo;
      const payDateVal = isPaid ? new Date() : "";

      // Tạo debtRow hoàn chỉnh như saveConfirmedOrder_
      const debtRow = [
        new Date(ngay),           // A: Ngày
        info,                     // B: Thông tin
        soMam,                    // C: Số mâm
        calcFinal.donGiaMam,      // D: Đơn giá mâm
        calcFinal.tongDon,        // E: Tổng đơn
        kmNoiDung,                // F: KM nội dung
        kmSoTien,                 // G: KM số tiền
        calcFinal.doanhSo,        // H: Doanh số
        congNo,                   // I: Công nợ
        status,                   // J: Trạng thái
        payDateVal,               // K: Ngày thanh toán
        nguoiLap,                 // L: Thu ngân
        orderId,                  // M: Mã đơn
        itemCount,                // N: Số món
        itemStart,                // O: Dòng bắt đầu (Data)
        itemEnd,                  // P: Dòng kết thúc (Data)
        datCoc                    // Q: Tiền đặt cọc
      ];
      console.log(debtRow)
      shDebt.appendRow(debtRow);
      shP.deleteRow(rowIndex);
      try { 
        resetPendingSheetCache_(); 
        // Nếu có hàm reset cho sheet công nợ, hãy gọi ở đây
      } catch(e) {}
      // Trả về debug info cho client
      return {
        ok: true,
        moved: true,
        orderId: orderId,
        message: "Đã chuyển đơn sang Công nợ thành công"
      };
    } catch (e) {
      return { ok: false, error: e.message || String(e) };
    }
  }


  /* =====================================================================================
    11) LOG IN HÓA ĐƠN (GIỮ NGUYÊN)
    ===================================================================================== */
  function getPrintLogSheet_() {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sh = ss.getSheetByName("log_print");
    if (!sh) {
      sh = ss.insertSheet("log_print");
      sh.appendRow(["Thời gian", "Người in", "Số hóa đơn", "Row công nợ"]);
    }
    return sh;
  }

  function getInvoiceNoByDebtRow(rowIndex) {
    rowIndex = String(rowIndex || "").trim();
    if (!rowIndex) return { ok: true, invoiceNo: "" };

    const sh = getPrintLogSheet_();
    const last = sh.getLastRow();
    if (last < 2) return { ok: true, invoiceNo: "" };

    const vals = sh.getRange(2, 1, last - 1, 4).getValues();
    for (let i = vals.length - 1; i >= 0; i--) {
      const rRow = String(vals[i][3] || "").trim();
      const inv  = String(vals[i][2] || "").trim();
      if (rRow === rowIndex && inv) return { ok: true, invoiceNo: inv };
    }
    return { ok: true, invoiceNo: "" };
  }

  function getNextInvoiceNo_() {
    const sh = getPrintLogSheet_();
    const tz = Session.getScriptTimeZone();
    const todayKey = Utilities.formatDate(new Date(), tz, "yyyyMMdd");
    const prefix = "HD" + todayKey + "-";

    const last = sh.getLastRow();
    let maxSeq = 0;

    if (last >= 2) {
      const invCol = sh.getRange(2, 3, last - 1, 1).getValues();
      invCol.forEach(r => {
        const s = String(r[0] || "").trim();
        if (!s.startsWith(prefix)) return;
        const tail = s.slice(prefix.length);
        const n = parseInt(tail, 10);
        if (!isNaN(n) && n > maxSeq) maxSeq = n;
      });
    }

    const nextSeq = maxSeq + 1;
    const seqStr = String(nextSeq).padStart(4, "0");
    return prefix + seqStr;
  }

  function getNextInvoiceNo() {
    return { ok: true, invoiceNo: getNextInvoiceNo_() };
  }

  function reserveInvoiceNo(invoiceNo, username, rowIndex) {
    invoiceNo = String(invoiceNo || "").trim();
    rowIndex  = String(rowIndex  || "").trim();
    if (!invoiceNo) return { ok: false, error: "Thiếu số hóa đơn." };
    if (!rowIndex)  return { ok: false, error: "Thiếu row công nợ." };

    const existed = getInvoiceNoByDebtRow(rowIndex);
    if (existed && existed.ok && existed.invoiceNo) {
      return { ok: true, invoiceNo: existed.invoiceNo, reused: true };
    }

    const sh = getPrintLogSheet_();
    const last = sh.getLastRow();

    if (last >= 2) {
      const invCol = sh.getRange(2, 3, last - 1, 1).getValues();
      for (let i = 0; i < invCol.length; i++) {
        if (String(invCol[i][0] || "").trim() === invoiceNo) {
          return { ok: false, error: "Số hóa đơn đã tồn tại trong log_print: " + invoiceNo };
        }
      }
    }

    const tz = Session.getScriptTimeZone();
    const ts = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");
    sh.appendRow([ts, username || "", invoiceNo, rowIndex]);

    return { ok: true, invoiceNo: invoiceNo, reused: false };
  }

  function logPrintedInvoice(invoiceNo, username, rowIndex) {
    const r = reserveInvoiceNo(invoiceNo, username, rowIndex);
    if (!r || r.ok === false) return r;
    return { ok: true, invoiceNo: r.invoiceNo, reused: !!r.reused };
  }


  /* =====================================================================================
    12) CÔNG NỢ (GIỮ NGUYÊN)
    ===================================================================================== */
  function getAllDebtData() {
    const ss = getSpreadsheet_();
    const sh = ss.getSheetByName(SHEET_DEBT);

    if (!sh) return { headers: [], rows: [], rowIndices: [] };

    const lastRow = sh.getLastRow();
    const lastCol = DEBT_LAST_COL || sh.getLastColumn();
    if (lastRow < 1 || lastCol < 1) return { headers: [], rows: [], rowIndices: [] };

    const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);

    const rows = [];
    const rowIndices = [];

    if (lastRow > 1) {
      const tz = Session.getScriptTimeZone();

      const shData = ss.getSheetByName(SHEET_DATA);
      const orderItemCount = {};
      if (shData && shData.getLastRow() > 1) {
        const dataVals = shData.getDataRange().getValues(); // A:K
        for (let i = 1; i < dataVals.length; i++) {
          const rData = dataVals[i];
          const dVal  = rData[0];
          const sdt   = rData[2];
          const orderIdCell = rData[10];

          if (!dVal || !sdt) continue;

          let key;
          if (orderIdCell) key = String(orderIdCell);
          else key = fmtDateYmd_(dVal, tz) + "|" + normalizePhone_(sdt);

          orderItemCount[key] = (orderItemCount[key] || 0) + 1;
        }
      }

      const raw = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
      for (let i = 0; i < raw.length; i++) {
        const r = raw[i];

        const dateVal  = r[DEBT_COL_DATE - 1];
        const infoVal  = String(r[DEBT_COL_INFO - 1] || "");
        const soMamRaw = r[DEBT_COL_SOMAM - 1];
        const orderId  = r[DEBT_COL_ORDER_ID - 1] || "";

        let soMon = 0;
        let keyForCount = "";
        if (orderId) keyForCount = String(orderId);
        else if (dateVal && infoVal) {
          const phone = extractPhoneFromInfo_(infoVal) || "";
          if (phone) keyForCount = fmtDateYmd_(dateVal, tz) + "|" + normalizePhone_(phone);
        }
        if (keyForCount && orderItemCount[keyForCount]) soMon = orderItemCount[keyForCount];

        const out = [];
        for (let c = 0; c < lastCol; c++) {
          let v = r[c];
          if (v instanceof Date) v = Utilities.formatDate(v, tz, "yyyy-MM-dd");
          if (c === DEBT_COL_SOMAM - 1) {
            const base = soMamRaw;
            v = (soMon && soMon > 0) ? base + " (" + soMon + " món)" : base;
          }
          out.push(v);
        }

        rows.push(out);
        rowIndices.push(DEBT_DATA_START + i);
      }
    }

    return { headers: headers, rows: rows, rowIndices: rowIndices };
  }

  function listDebtByPhone(phone) {
    const sdtQRaw = String(phone || "").trim();
    const sdtQNorm = normalizePhone_(sdtQRaw);
    if (!sdtQNorm) return [];

    const ss = getSpreadsheet_();
    const sh = ss.getSheetByName(SHEET_DEBT);
    const shDat = ss.getSheetByName(SHEET_DATA);
    if (!sh) return [];

    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow < 2 || lastCol < 2) return [];

    const tz   = Session.getScriptTimeZone();
    const vals = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
    const dataVals = (shDat && shDat.getLastRow() > 1) ? shDat.getDataRange().getValues() : [];

    const out = [];
    for (let i = 0; i < vals.length; i++) {
      const rowIndex = 2 + i;
      const r        = vals[i];

      const dateVal   = r[DEBT_COL_DATE - 1];
      const info      = String(r[DEBT_COL_INFO - 1] || "");
      const soMam     = Number(r[DEBT_COL_SOMAM - 1] || 0);
      const donGiaMam = Number(r[DEBT_COL_DONGIA_MAM - 1] || 0);
      const tongDon   = Number(r[DEBT_COL_TONG_DON - 1] || 0);
      const trangThai = String(r[DEBT_COL_STATUS - 1] || "");
      const orderId   = r[DEBT_COL_ORDER_ID - 1] || "";

      if (!info) continue;

      const realPhone = extractPhoneFromInfo_(info) || "";
      const normPhone = normalizePhone_(realPhone);
      if (!normPhone || normPhone !== sdtQNorm) continue;

      const dObj = parseDateCell_(dateVal);

      const dateStr = dObj ? Utilities.formatDate(dObj, tz, "dd/MM/yyyy") : "";

      let soMon = 0;
      if (dataVals.length > 1) {
        if (orderId) {
          for (let j = 1; j < dataVals.length; j++) {
            const rd = dataVals[j];
            const rowOrderId = rd[10] || "";
            if (String(rowOrderId) === String(orderId)) soMon++;
          }
        } else if (dateVal && normPhone) {
          const dateYmdKey = fmtDateYmd_(dateVal, tz);
          const keyBase = dateYmdKey + "|" + normPhone;
          for (let j = 1; j < dataVals.length; j++) {
            const rd = dataVals[j];
            if (!rd[0] || !rd[2]) continue;
            const kRow = fmtDateYmd_(rd[0], tz) + "|" + normalizePhone_(rd[2]);
            if (kRow === keyBase) soMon++;
          }
        }
      }

      out.push({
        row: rowIndex,
        date: dateStr,
        info: info,
        soMam: soMam,
        soMon: soMon,
        donGiaMam: donGiaMam,
        tongDon: tongDon,
        trangThai: trangThai,
        orderId: orderId
      });
    }

    return out;
  }
function coerceGsResponse_(res) {
  if (res === null || res === undefined) return res;

  // nếu lỡ là String object
  if (typeof res === "object" && typeof res.toString === "function") {
    // object chuẩn có ok/header/error thì giữ nguyên
    if (("ok" in res) || ("header" in res) || ("error" in res) || ("message" in res)) return res;
    const sObj = String(res).trim();
    if (sObj) return coerceGsResponse_(sObj);
    return res;
  }

  if (typeof res !== "string") return res;

  let s = String(res || "").trim();
  if (!s) return res;

  // JSON parse 1 lớp
  try {
    const j1 = JSON.parse(s);
    if (j1 && typeof j1 === "object") return j1;

    // JSON parse 2 lớp ( '"{...}"' )
    if (typeof j1 === "string") {
      const s2 = j1.trim();
      if (s2 && (s2[0] === "{" || s2[0] === "[")) {
        const j2 = JSON.parse(s2);
        if (j2 && typeof j2 === "object") return j2;
      }
    }
  } catch (e) {}

  return res;
}


  /* =====================================================================================
    13) ORDER DETAIL BY DEBT ROW (GIỮ NGUYÊN)
    ===================================================================================== */
    function extractNameFromInfo_(info) {
  const s = String(info || "").trim();
  if (!s) return "";

  const m = s.match(/(\+?84|0)\d{6,12}/); // match số điện thoại
  if (!m) {
    // fallback dạng cũ nếu không tìm thấy sđt
    const parts = s.split(" - ");
    return String(parts[0] || "").trim();
  }

  const idx = s.indexOf(m[0]);
  let left = s.slice(0, idx).trim();           // phần trước SĐT
  left = left.replace(/[-–—|\s]+$/g, "").trim(); // bỏ dấu phân cách ở cuối
  return left;
}

function extractAddressFromInfo_(info) {
  const s = String(info || "").trim();
  if (!s) return "";

  const m = s.match(/(\+?84|0)\d{6,12}/);
  if (!m) {
    // fallback dạng cũ nếu không tìm thấy sđt
    const parts = s.split(" - ");
    return String(parts.slice(2).join(" - ") || "").trim();
  }

  const idx = s.indexOf(m[0]);
  let right = s.slice(idx + m[0].length).trim();     // phần sau SĐT
  right = right.replace(/^[-–—|\s]+/g, "").trim();   // bỏ dấu phân cách ở đầu
  return right;
}

  function getOrderDetailByDebtRow(rowIndex) {
  try {
    const ss    = getSpreadsheet_();
    const shDeb = ss.getSheetByName(SHEET_DEBT);
    const shDat = ss.getSheetByName(SHEET_DATA);
    if (!shDeb || !shDat) throw new Error("Thiếu sheet Công nợ hoặc Thông tin đặt hàng.");

    rowIndex = Number(rowIndex || 0);
    if (!rowIndex || rowIndex < DEBT_DATA_START || rowIndex > shDeb.getLastRow()) {
      throw new Error("rowIndex không hợp lệ.");
    }

    const tz      = Session.getScriptTimeZone();
    const dateVal = shDeb.getRange(rowIndex, DEBT_COL_DATE).getValue();
    const info    = String(shDeb.getRange(rowIndex, DEBT_COL_INFO).getValue() || "");

    const soMam     = toMoneyNumber_(shDeb.getRange(rowIndex, DEBT_COL_SOMAM).getValue() || 0);
    const donGiaMam = toMoneyNumber_(shDeb.getRange(rowIndex, DEBT_COL_DONGIA_MAM).getValue() || 0);
    const tongDon   = toMoneyNumber_(shDeb.getRange(rowIndex, DEBT_COL_TONG_DON).getValue() || 0);
    const trangThai = String(shDeb.getRange(rowIndex, DEBT_COL_STATUS).getValue() || "").trim();
    const kmNote    = String(shDeb.getRange(rowIndex, DEBT_COL_KM_NOTE).getValue() || "").trim();
    const kmAmount  = toMoneyNumber_(shDeb.getRange(rowIndex, DEBT_COL_KM_AMOUNT).getValue() || 0);
    const thungan   = String(shDeb.getRange(rowIndex, DEBT_COL_THUNGAN).getValue() || "").trim();
    const orderIdRaw = String(shDeb.getRange(rowIndex, DEBT_COL_ORDER_ID).getValue() || "").trim();

    const dateStr = (dateVal instanceof Date) ? Utilities.formatDate(dateVal, tz, "dd/MM/yyyy") : String(dateVal || "");
    const tenKH = extractNameFromInfo_(info) || "";
    const sdt   = extractPhoneFromInfo_(info) || "";
    const diaChi = extractAddressFromInfo_(info) || "";

    const orderId = orderIdRaw || makeOrderId_(dateVal instanceof Date ? dateVal : new Date(), sdt);

    // Lấy items theo orderId bằng map header (đúng schema)
    const items = getItemsByOrderId_(shDat, orderId) || [];
    if (!items.length) {
      // vẫn trả ok=true để UI mở được, nhưng cảnh báo nhẹ nếu cần
      // (bạn có thể đổi thành ok=false nếu muốn chặn in)
    }

    const payload = {
      ok: true,
      header: {
        dateRaw: (dateVal instanceof Date) ? dateVal.toISOString() : String(dateVal || ""),
        date: dateStr,
        tenKH: tenKH,
        sdt: sdt,
        diaChi: diaChi,
        soMam: soMam,
        donGiaMam: donGiaMam,
        tongDon: tongDon,
        trangThai: trangThai,
        status: trangThai,
        kmNoiDung: kmNote,
        kmSoTien: kmAmount,
        thungan: thungan,
        orderId: orderId
      },
      items: items
    };

    // ✅ QUAN TRỌNG: trả JSON string để client parse chắc chắn
    return JSON.stringify(payload);

  } catch (e) {
    return JSON.stringify({ ok: false, error: e && e.message ? e.message : String(e) });
  }
}



  /* =====================================================================================
    14) LOG XOÁ / THU NỢ + DELETE / PAYMENT / TIMELINE (GIỮ NGUYÊN)
    ===================================================================================== */
  function getLogSheet_() {
    const ss = getSpreadsheet_();
    let sh = ss.getSheetByName(SHEET_LOG);
    if (!sh) {
      sh = ss.insertSheet(SHEET_LOG);
      sh.getRange(1, 1, 1, 11).setValues([[
        "Thời gian",
        "Người thao tác",
        "Loại",
        "Row Công nợ cũ",
        "Ngày đơn",
        "Thông tin tiệc",
        "Trạng thái",
        "Mã đơn",
        "Số tiền thanh toán",
        "Khuyến mãi thêm",
        "Công nợ mới"
      ]]);
    }
    return sh;
  }

  function deleteOrderByDebtRow(rowIndex, deletedByOrPayload) {
    try {
      const delCtx = assertCanDelete_(deletedByOrPayload);
      const deletedBy = String(delCtx.username || "").trim();

      const ss    = getSpreadsheet_();
      const shDeb = ss.getSheetByName(SHEET_DEBT);
      const shDat = ss.getSheetByName(SHEET_DATA);
      if (!shDeb || !shDat) throw new Error("Thiếu sheet Công nợ hoặc Thông tin đặt hàng.");

      rowIndex = Number(rowIndex || 0);
      if (!rowIndex || rowIndex < DEBT_DATA_START || rowIndex > shDeb.getLastRow()) throw new Error("rowIndex không hợp lệ.");

      const tz       = Session.getScriptTimeZone();
      const dateVal  = shDeb.getRange(rowIndex, DEBT_COL_DATE).getValue();
      const info     = String(shDeb.getRange(rowIndex, DEBT_COL_INFO).getValue() || "");
      const trangThai= String(shDeb.getRange(rowIndex, DEBT_COL_STATUS).getValue() || "");
      const orderId  = String(shDeb.getRange(rowIndex, DEBT_COL_ORDER_ID).getValue() || "");
      const dateStr  = dateVal ? Utilities.formatDate(new Date(dateVal), tz, "yyyy-MM-dd") : "";

      const sdt = extractPhoneFromInfo_(info) || "";
      const baseKey   = dateVal ? fmtDateYmd_(dateVal, tz) + "|" + normalizePhone_(sdt) : "";

      // ✅ Xoá items theo orderId: dùng TextFinder để chỉ đụng đúng block, không scan toàn sheet
if (orderId) {
  deleteItemsByOrderId_(shDat, orderId);
} else if (baseKey) {
  // Fallback legacy theo (ngày + sđt) nếu thiếu orderId
  const m = getDataSheetMap_(shDat);
  const lastRowDat = shDat.getLastRow();
  if (m && lastRowDat >= 2) {
    const width = Math.max(m.idxOid || 1, m.idxSt || 1, m.idxNgay || 1, m.idxSdt || 1);
    const vals = shDat.getRange(2, 1, lastRowDat - 1, width).getValues();
    for (let i = 0; i < vals.length; i++) {
      const r = vals[i];
      if (!r) continue;
      if (!r[0] || !r[2]) continue;
      const kRow = fmtDateYmd_(r[0], tz) + "|" + normalizePhone_(r[2]);
      if (kRow === baseKey) {
        const rr = i + 2;
        try { shDat.getRange(rr, m.idxOid, 1, 1).clearContent(); } catch(e) {}
        if (m.idxSt) {
          try { shDat.getRange(rr, m.idxSt, 1, 1).setValue("DELETED"); } catch(e) {}
        }
      }
    }
  }
}

// Xoá dòng congno (nhanh)
sh_deleteRowFast_(SHEET_DEBT, rowIndex);

      try {
        const shLog = getLogSheet_();
        const ts    = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");
        shLog.appendRow([ts, deletedBy || "", "DELETE", rowIndex, dateStr, info, trangThai, orderId, "", "", ""]);
      } catch (logErr) {}

      const msg = "Đã xoá đơn.Ngày: " + dateStr + " " + info + "Trạng thái: " + trangThai;
      return { ok: true, message: msg };
    } catch (e) {
      return { ok: false, error: e && e.message ? e.message : String(e) };
    }
  }


  function processDebtPayment(rowIndex, payAmount, promoAmount, newStatus, payDateStr, username) {
    try {
      assertCanEdit_(username);
      const ss    = getSpreadsheet_();
      const shDeb = ss.getSheetByName(SHEET_DEBT);
      if (!shDeb) throw new Error("Không tìm thấy sheet Công nợ.");

      rowIndex = Number(rowIndex || 0);
      if (!rowIndex || rowIndex < DEBT_DATA_START || rowIndex > shDeb.getLastRow()) throw new Error("rowIndex không hợp lệ.");

      payAmount   = Number(payAmount)   || 0;
      promoAmount = Number(promoAmount) || 0;
      if (payAmount < 0 || promoAmount < 0) throw new Error("Số tiền thanh toán / khuyến mãi không được âm.");

      const tz = Session.getScriptTimeZone();

      const dateVal   = shDeb.getRange(rowIndex, DEBT_COL_DATE).getValue();
      const info      = String(shDeb.getRange(rowIndex, DEBT_COL_INFO).getValue() || "");
      const statusOld = String(shDeb.getRange(rowIndex, DEBT_COL_STATUS).getValue() || "");
      const kmOld     = Number(shDeb.getRange(rowIndex, DEBT_COL_KM_AMOUNT).getValue() || 0);
      const doanhOld  = Number(shDeb.getRange(rowIndex, DEBT_COL_DOANHSO).getValue() || 0);
      const congOld   = Number(shDeb.getRange(rowIndex, DEBT_COL_CONGNO).getValue() || 0);
      const orderId   = String(shDeb.getRange(rowIndex, DEBT_COL_ORDER_ID).getValue() || "");
      const dateStr   = dateVal ? Utilities.formatDate(new Date(dateVal), tz, "yyyy-MM-dd") : "";

      if (congOld <= 0 && statusOld === STATUS_PAID) return { ok: false, error: "Đơn này đã thanh toán đủ, không còn công nợ." };

      let newKm    = kmOld + promoAmount;
      if (newKm < 0) newKm = 0;

      let newDoanh = doanhOld - promoAmount;
      if (newDoanh < 0) newDoanh = 0;

      let newCong  = congOld - payAmount - promoAmount;
      if (newCong < 0) newCong = 0;

      let finalStatus = newStatus || statusOld;
      let payDateVal  = shDeb.getRange(rowIndex, DEBT_COL_NGAYTT).getValue();

      const payDateFromClient = payDateStr ? (parseYmdToDate_(payDateStr) || new Date()) : new Date();

      if (finalStatus === STATUS_PAID || newCong === 0) {
        finalStatus = STATUS_PAID;
        newCong     = 0;
        payDateVal  = payDateFromClient;
      } else {
        payDateVal = "";
      }

      shDeb.getRange(rowIndex, DEBT_COL_KM_AMOUNT).setValue(newKm);
      shDeb.getRange(rowIndex, DEBT_COL_DOANHSO).setValue(newDoanh);
      shDeb.getRange(rowIndex, DEBT_COL_CONGNO).setValue(newCong);
      shDeb.getRange(rowIndex, DEBT_COL_STATUS).setValue(finalStatus);
      shDeb.getRange(rowIndex, DEBT_COL_NGAYTT).setValue(payDateVal);

      try {
        const shLog = getLogSheet_();
        const ts    = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");
        shLog.appendRow([ts, username || "", "PAY", rowIndex, dateStr, info, finalStatus, orderId, payAmount, promoAmount, newCong]);
      } catch (logErr) {}

      const msg =
        "Đã cập nhật thu nợ.\n" +
        "Công nợ cũ: " + congOld.toLocaleString("vi-VN") + "\n" +
        "Thanh toán: " + payAmount.toLocaleString("vi-VN") + "\n" +
        "Khuyến mãi thêm: " + promoAmount.toLocaleString("vi-VN") + "\n" +
        "Công nợ mới: " + newCong.toLocaleString("vi-VN") + "\n" +
        "Trạng thái mới: " + finalStatus;

      return { ok: true, message: msg, congNoMoi: newCong, statusMoi: finalStatus };
    } catch (e) {
      return { ok: false, error: e && e.message ? e.message : String(e) };
    }
  }

function processAllDebtPaymentForCustomer(phone, username) {
  try {
    assertCanEdit_(username);

    const pNorm = normalizePhone_(phone);
    if (!pNorm) return { ok: false, error: "Thiếu SĐT." };

    const list = listDebtByPhone(pNorm) || [];
    if (!list.length) return { ok: false, error: "Không có đơn nào cho SĐT này." };

    const ss = getSpreadsheet_();
    const shDebt = ss.getSheetByName(SHEET_DEBT);
    if (!shDebt) return { ok: false, error: "Không tìm thấy sheet Công nợ." };

    const tz = Session.getScriptTimeZone();
    const todayStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");

    let totalPaid = 0;
    let paidOrders = 0;
    let skipped = 0;
    const errors = [];

    list.forEach(function(r) {
      const rowIndex = Number(r.row || r.rowIndex || 0);
      if (!rowIndex || rowIndex < 2) return;

      // đọc công nợ hiện tại của từng đơn
      const congOld = toMoneyNumber_(shDebt.getRange(rowIndex, DEBT_COL_CONGNO).getValue() || 0);

      if (!congOld || congOld <= 0) { skipped++; return; }

      // ✅ thu đúng số còn nợ của đơn, KHÔNG dùng số sentinel
      const rs = processDebtPayment(rowIndex, congOld, 0, STATUS_PAID, todayStr, username);

      if (!rs || rs.ok === false) {
        errors.push("Row " + rowIndex + ": " + ((rs && (rs.error || rs.message)) ? (rs.error || rs.message) : "Lỗi không xác định"));
        return;
      }

      paidOrders += 1;
      totalPaid += congOld;
    });

    if (!paidOrders && errors.length) {
      return { ok: false, error: "Không thu được đơn nào.\n" + errors.join("\n") };
    }

    let msg =
      "Đã thu nợ toàn bộ cho SĐT: " + pNorm +
      "\n- Số đơn đã thu: " + paidOrders +
      "\n- Tổng đã thu: " + totalPaid.toLocaleString("vi-VN") + " đ" +
      (skipped ? ("\n- Bỏ qua (đã hết nợ): " + skipped) : "") +
      (errors.length ? ("\n\nLỗi một số đơn:\n- " + errors.join("\n- ")) : "");

    return { ok: true, message: msg, totalPaid: totalPaid, paidOrders: paidOrders, skipped: skipped, errors: errors };
  } catch (e) {
    return { ok: false, error: e && e.message ? e.message : String(e) };
  }
}


  /* Timeline giữ nguyên như bạn gửi */
  function getDebtTimelineByPhone(phone, fromDateStr, toDateStr) {
    try {
      phone = String(phone || '').trim();
      const phoneNorm = normalizePhone_(phone);
      if (!phoneNorm) return { ok: false, error: 'Thiếu hoặc sai số điện thoại.' };

      const ss = getSpreadsheet_();
      const shDeb = ss.getSheetByName(SHEET_DEBT);
      if (!shDeb) return { ok: false, error: "Không tìm thấy sheet Công nợ." };

      const tz = Session.getScriptTimeZone();
      let fromDate = parseYmdToDate_(fromDateStr);
      let toDate   = parseYmdToDate_(toDateStr);

      const events = [];
      let custName = '';
      let displayPhone = '';

      const lastDebtRow = shDeb.getLastRow();
      if (lastDebtRow >= DEBT_DATA_START) {
        const valsDebt = shDeb.getRange(DEBT_DATA_START, 1, lastDebtRow - DEBT_DATA_START + 1, DEBT_LAST_COL).getValues();
        valsDebt.forEach(function (row) {
          const dateVal = row[DEBT_COL_DATE - 1];
          const info    = String(row[DEBT_COL_INFO - 1] || "");
          const doanhSo = Number(row[DEBT_COL_DOANHSO - 1] || 0);
          if (!info || !dateVal || doanhSo === 0) return;

          const rPhone = extractPhoneFromInfo_(info) || "";
          const rNorm  = normalizePhone_(rPhone);
          if (!rNorm || rNorm !== phoneNorm) return;

          if (!custName) custName = getNameFromInfo_(info) || "";
          if (!displayPhone) displayPhone = String(rPhone || "").trim() || String(phone || "").trim();


          const d = new Date(dateVal);
          d.setHours(0, 0, 0, 0);

          events.push({ date: d, type: 'SALE', desc: 'Bán hàng', increase: doanhSo, decrease: 0 });
        });
      }

      const shLog = ss.getSheetByName(SHEET_LOG);
      if (shLog && shLog.getLastRow() > 1) {
        const valsLog = shLog.getRange(2, 1, shLog.getLastRow() - 1, 11).getValues();
        valsLog.forEach(function (row) {
          const type = String(row[2] || '').toUpperCase();
          if (type !== 'PAY') return;

          const info = String(row[5] || "");
          if (!info) return;
          const rPhone = extractPhoneFromInfo_(info) || "";
          const rNorm  = normalizePhone_(rPhone);
          if (!rNorm || rNorm !== phoneNorm) return;

          if (!custName) custName = (parts[0] || "").trim();
          if (!displayPhone) displayPhone = rPhone;

          const tsStr = String(row[0] || "");
          let payDate = null;
          if (tsStr) {
            const dPart = tsStr.split(' ')[0];
            payDate = parseYmdToDate_(dPart);
          }
          if (!payDate) {
            payDate = new Date();
            payDate.setHours(0, 0, 0, 0);
          }

          const payAmount   = Number(row[8] || 0) || 0;
          const promoAmount = Number(row[9] || 0) || 0;
          const dec = payAmount + promoAmount;
          if (dec <= 0) return;

          events.push({ date: payDate, type: 'PAY', desc: 'Thu nợ khách hàng', increase: 0, decrease: dec });
        });
      }

      if (!events.length) return { ok: true, customer: { tenKhach: custName || '', phone: displayPhone || phone }, opening: 0, rows: [] };

      events.sort(function (a, b) { return a.date.getTime() - b.date.getTime(); });

      if (!fromDate) { fromDate = new Date(events[0].date); fromDate.setHours(0,0,0,0); }
      if (!toDate)   { toDate   = new Date(events[events.length - 1].date); toDate.setHours(0,0,0,0); }
      if (fromDate.getTime() > toDate.getTime()) { const tmp = fromDate; fromDate = toDate; toDate = tmp; }

      let opening = 0;
      events.forEach(function (ev) {
        if (ev.date.getTime() < fromDate.getTime()) opening += ev.increase - ev.decrease;
      });

      const rows = [];
      let running = opening;

      rows.push({
        date: Utilities.formatDate(fromDate, tz, 'dd/MM/yyyy'),
        soChungTu: '',
        dienGiai: 'Số nợ đầu kỳ',
        tang: 0,
        giam: 0,
        conNo: running
      });

      events.forEach(function (ev) {
        const t = ev.date.getTime();
        if (t < fromDate.getTime() || t > toDate.getTime()) return;

        running += ev.increase - ev.decrease;

        rows.push({
          date: Utilities.formatDate(ev.date, tz, 'dd/MM/yyyy'),
          soChungTu: '',
          dienGiai: ev.desc,
          tang: ev.increase,
          giam: ev.decrease,
          conNo: running
        });
      });

      return {
        ok: true,
        customer: { tenKhach: custName || '', phone: displayPhone || phone },
        fromDate: Utilities.formatDate(fromDate, tz, 'dd/MM/yyyy'),
        toDate: Utilities.formatDate(toDate, tz, 'dd/MM/yyyy'),
        opening: opening,
        rows: rows
      };
    } catch (err) {
      return { ok: false, error: 'Lỗi getDebtTimelineByPhone: ' + err };
    }
  }


  /* =====================================================================================
    15) DANH MỤC HÀNG CRUD (GIỮ NGUYÊN)
    ===================================================================================== */
  function dm_list() {
    const ss = getSpreadsheet_();
    const sh = ss.getSheetByName(SHEET_DM);
    if (!sh) return [];

    const last = sh.getLastRow();
    if (last < 2) return [];

    // đọc A:E (có status)
    const values = sh.getRange(2, 1, last - 1, 5).getValues();
    const out = [];

    for (let i = 0; i < values.length; i++) {
      const rowIndex = i + 2;
      const ten = String(values[i][0] || "").trim();
      const dvt = String(values[i][1] || "Đĩa").trim();
      const gia = Number(values[i][2] || 0) || 0;
      const loaiMon = String(values[i][3] || "").trim();
      const status = String(values[i][4] === undefined ? "" : values[i][4]).trim();

      if (!ten) continue;
      if (status === "0") continue; // ✅ soft delete

      out.push({ rowIndex, ten, dvt, gia, loaiMon });
    }
    return out;
  }


  function dm_upsert(payload) {
    payload = payload || {};
    const ten = String(payload.ten || "").trim();
    const dvt = String(payload.dvt || "Đĩa").trim();
    const gia = Number(payload.gia || 0) || 0;
    const loaiMon = String(payload.loaiMon || "").trim();

    if (!ten) return { ok: false, error: "Thiếu tên món." };
    if (gia < 0) return { ok: false, error: "Giá không được âm." };

    const ss = getSpreadsheet_();
    const sh = ss.getSheetByName(SHEET_DM);
    if (!sh) return { ok: false, error: "Không tìm thấy sheet Danh mục hàng." };

    let rowIndex = Number(payload.rowIndex || 0);

    // ✅ upsert sẽ set status = 1 (món đang active)
    const statusActive = 1;

    if (rowIndex >= 2 && rowIndex <= sh.getLastRow()) {
      sh.getRange(rowIndex, 1, 1, 5).setValues([[ten, dvt, gia, loaiMon, statusActive]]);
      clearMenuCache_();
      return { ok: true, rowIndex: rowIndex };
    }

    sh.appendRow([ten, dvt, gia, loaiMon, statusActive]);
    rowIndex = sh.getLastRow();
    clearMenuCache_();
    return { ok: true, rowIndex: rowIndex };
  }


  function dm_delete(rowIndex) {
    const ss = getSpreadsheet_();
    const sh = ss.getSheetByName(SHEET_DM);
    if (!sh) return { ok: false, error: "Không tìm thấy sheet Danh mục hàng." };

    rowIndex = Number(rowIndex || 0);
    if (rowIndex < 2 || rowIndex > sh.getLastRow()) return { ok: false, error: "rowIndex không hợp lệ." };

    // ✅ soft delete: set status = 0
    sh.getRange(rowIndex, DM_COL_STATUS).setValue(0);
    clearMenuCache_();
    return { ok: true };
  }



  /* =====================================================================================
    16) IN HÓA ĐƠN (GIỮ NGUYÊN STYLE/HTML), CHỈ SỬA: KHÔNG TRÙNG HÀM PENDING
    - QUAN TRỌNG: File bạn gửi có 2 hàm getInvoiceHtmlByPendingRow -> bị ghi đè
    - Ở đây: GIỮ 1 HÀM, giống cơ chế cũ: cấp số HĐ + log_print
    ===================================================================================== */
  function isAlcoholItem_(loaiMon) {
    const t = normalizeText_(loaiMon);
    return t === "do uong co con" || t.includes("do uong co con");
  }

  function buildInvoiceHtml(h, items, invoiceNo, includeTax) {
    h = h || {};
    items = items || [];
    invoiceNo = String(invoiceNo || "").trim();

    const tz = Session.getScriptTimeZone();

    includeTax = (typeof includeTax === "undefined" || includeTax === null) ? true : !!includeTax;
    const dateStr = h.date || (h.dateRaw ? Utilities.formatDate(parseDateCell_(h.dateRaw), tz, "dd/MM/yyyy") : "");

    const soMam = Number(h.soMam || 0) || 0;
    const tongDon = Number(h.tongDon || 0) || 0;
    const kmSoTien = Number(h.kmSoTien || 0) || 0;
    const datCoc = Number(h.datCoc || 0) || 0;

    // FORCE DEBUG - Add to HTML output
    const debugInfo = `DEBUG: soMam=${soMam}, datCoc=${datCoc}, tongDon=${tongDon}, items=${items.length}`;

    // FORCE ALERT FOR DEBUG
    // alert('DEBUG VALUES: ' + debugInfo);

    // Debug logs (tạm tắt)
    // console.log('buildInvoiceHtml DEBUG:', { soMam, tongDon, datCoc, itemsLength: items.length });
    const esc = s => String(s === undefined || s === null ? "" : s).replace(/[&<>"']/g, m => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));
    const money = n => (Number(n || 0) || 0).toLocaleString("vi-VN") + " đ";
    let vat10Base = 0;
    let totalPerMam = 0;
    items.forEach(it => {
      const tt = Number(it.thanhTien || (Number(it.sl || 0) * Number(it.donGia || 0))) || 0;
      totalPerMam += tt;
      if (isAlcoholItem_(it.loaiMon || "")) vat10Base += tt;
    });

    // Tổng đơn = tổng tiền hàng mỗi mâm × số mâm
    const calculatedTotal = totalPerMam * Math.max(1, soMam);
    const netAfterDiscount = Math.max(0, calculatedTotal - kmSoTien);
    const alcoholTotal = vat10Base * Math.max(0, soMam);
    const vat10 = includeTax ? Math.round(alcoholTotal * 0.10) : 0;

    // Nếu đơn chỉ có đặt cọc (không có món), tổng thanh toán = 0 (đã đặt cọc)
    // Ngược lại, tổng thanh toán = netAfterDiscount + vat10 - datCoc
    const isOnlyDeposit = items.length === 0 && datCoc > 0;
    const grandTotal = isOnlyDeposit ? 0 : Math.max(0, netAfterDiscount + vat10 - datCoc);

    // Debug chi tiết
    const calcDebug = `CALC: soMam=${soMam}, totalPerMam=${totalPerMam}, calculated=${calculatedTotal}, grand=${grandTotal}`;

    console.log('INVOICE CALC DEBUG: ' + JSON.stringify({
      soMam, totalPerMam, calculatedTotal, tongDon, datCoc,
      itemsCount: items.length,
      hasDeposit: datCoc > 0,
      isOnlyDeposit: items.length === 0 && datCoc > 0
    }));

    // console.log('INVOICE CALC DEBUG:', { soMam, totalPerMam, calculatedTotal, grandTotal });

    // Debug log
    console.log('buildInvoiceHtml CALC:', {
      totalPerMam, soMam, calculatedTotal, netAfterDiscount,
      vat10, datCoc, isOnlyDeposit, grandTotal,
      items: items.map(it => ({ ten: it.tenMon, sl: it.sl, gia: it.donGia, tt: it.thanhTien }))
    });

    const paid = String(h.trangThai || h.status || "").trim() === STATUS_PAID;
    const stampCls = paid ? "paid" : "unpaid";
    const stampText = paid ? "ĐÃ THANH TOÁN" : "CHƯA THANH TOÁN";

    let rowsHtml = "";
    (items || []).forEach((it, i) => {
      const ten = esc(it.tenMon || "");
      const dvt = esc(it.dvt || "");
      const sl  = Number(it.sl || 0) || 0;
      const gia = Number(it.donGia || 0) || 0;
      const tt  = Number(it.thanhTien || (sl * gia)) || 0;
      rowsHtml += `
        <tr>
          <td style="text-align:center;">${i + 1}</td>
          <td>${ten}</td>
          <td style="text-align:center;">${dvt}</td>
          <td style="text-align:right;">${sl}</td>
          <td style="text-align:right;">${money(gia)}</td>
          <td style="text-align:right;">${money(tt)}</td>
        </tr>
      `;
    });

    const css = `
      @page { size: A5; margin: 8mm; }
      * { box-sizing: border-box; }
      body { margin:0; font-family: Arial, Helvetica, sans-serif; color:#111; font-size: 12px; line-height: 1.25; }
      .invoice { width: 100%; }
      .inv-head { text-align:center; margin-bottom: 6px; position: relative; }
      .inv-brand { font-weight: 700; font-size: 13px; letter-spacing: 0.3px; }
      .inv-sub { font-weight: 600; font-size: 12px; margin-top: 2px; }
      .inv-title { font-weight: 800; font-size: 14px; margin-top: 4px; }
      .inv-no { font-size: 12px; margin-top: 2px; }
      .inv-stamp { position:absolute; right:0; top:0; font-size: 10px; padding: 3px 6px; border-radius: 10px; border: 1px solid #999; }
      .inv-stamp.paid { border-color:#16a34a; color:#16a34a; }
      .inv-stamp.unpaid { border-color:#dc2626; color:#dc2626; }
      .inv-meta { margin: 6px 0 8px; }
      .inv-meta div { margin: 2px 0; }
      .inv-table { width:100%; border-collapse: collapse; font-size: 11px; }
      .inv-table th, .inv-table td { border: 1px solid #cbd5e1; padding: 3px 4px; vertical-align: top; }
      .inv-table th { background: #f1f5f9; font-weight: 700; }
      .inv-total { margin-top: 8px; border-top: 1px dashed #94a3b8; padding-top: 6px; }
      .inv-total .line { display:flex; justify-content: space-between; padding: 2px 0; }
      .inv-total .grand { font-size: 12px; font-weight: 800; padding-top: 4px; }
      .inv-footer { margin-top: 8px; text-align:center; font-size: 11px; line-height: 1.15; }
      @media print { .no-print { display:none !important; } }
    `;

    const html = `
      <html>
        <head>
          <meta charset="utf-8">
          <title>Hóa đơn ${esc(invoiceNo)}</title>
          <style>${css}</style>
        </head>
        <body>
          <div class="invoice">
            <div class="inv-head">
              <div class="inv-brand">BÁO CÁO KHÁCH ĐOÀN 2</div>
              <div class="inv-sub">HÓA ĐƠN BÁN HÀNG</div>
              <div class="inv-stamp ${stampCls}">${stampText}</div>

              <div class="inv-title">HÓA ĐƠN</div>
              <div class="inv-no">Số: <b>${esc(invoiceNo || "")}</b></div>
            </div>

            <div class="inv-meta">
              <div><b>Ngày:</b> ${esc(dateStr)}</div>
              <div><b>Khách:</b> ${esc(h.tenKH || "")} - ${esc(h.sdt || "")}</div>
              <div><b>Địa chỉ:</b> ${esc(h.diaChi || "")}</div>
              <div><b>Số mâm:</b> ${esc(soMam)}</div>
              <!-- Debug info commented out -->
            </div>

            <table class="inv-table">
              <thead>
                <tr>
                  <th style="width:26px; text-align:center;">#</th>
                  <th>Món</th>
                  <th style="width:56px; text-align:center;">ĐVT</th>
                  <th style="width:50px; text-align:right;">SL</th>
                  <th style="width:88px; text-align:right;">Đơn giá</th>
                  <th style="width:98px; text-align:right;">Thành tiền</th>
                </tr>
              </thead>
              <tbody>
                ${rowsHtml}
              </tbody>
            </table>

            <div class="inv-total">
              ${items.length > 0 ? `<div class="line"><span>Đơn giá 1 mâm</span><span><b>${money(h.donGiaMam || totalPerMam || 0)}</b></span></div>` : ''}
              <div class="line"><span>Tổng đơn${items.length > 0 ? ` (${soMam} mâm)` : ''}</span><span><b>${money(calculatedTotal)}</b></span></div>
              ${kmSoTien > 0 ? `<div class="line"><span>Khuyến mãi</span><span><b>- ${money(kmSoTien)}</b></span></div>` : ''}
              ${(includeTax && vat10 > 0) ? `<div class="line"><span>VAT 10% (Đồ uống có cồn)</span><span><b>${money(vat10)}</b></span></div>` : ``}
              ${datCoc > 0 ? `<div class="line" style="color:red"><span>Đã đặt cọc</span><span>- ${money(datCoc)}</span></div>` : ""}
              <!-- Debug info commented out -->
              <div class="line grand"><span>Tổng thanh toán</span><span>${money(grandTotal)}</span></div>
            </div>

            <div class="inv-footer">Cảm ơn quý khách!</div>
          </div>
        </body>
      </html>
    `;
    return html;
  }


  function getInvoiceHtmlByDebtRow(rowIndex, username, includeTax) {
  try {
    rowIndex = Number(rowIndex || 0);
    if (!rowIndex) return { ok: false, error: "Thiếu rowIndex." };

    const existed = getInvoiceNoByDebtRow(rowIndex);
    let invoiceNo = (existed && existed.ok) ? String(existed.invoiceNo || "").trim() : "";

    if (!invoiceNo) {
      invoiceNo = getNextInvoiceNo_();
      const rsv = reserveInvoiceNo(invoiceNo, username || "", String(rowIndex));
      if (!rsv || rsv.ok === false) {
        return { ok: false, error: (rsv && rsv.error) ? rsv.error : "Reserve số hoá đơn thất bại." };
      }
      invoiceNo = rsv.invoiceNo;
    }

    const od = getOrderDetailByDebtRow(rowIndex);
    if (!od || od.ok === false) {
      return { ok: false, error: (od && od.error) ? od.error : "Không lấy được chi tiết lịch sử." };
    }

    const tax = (typeof includeTax === "undefined" || includeTax === null) ? getEnableTax_() : !!includeTax;

    const html = buildInvoiceHtml(od.header, od.items || [], invoiceNo, tax);

    return jsonSafe_({ ok: true, invoiceNo: invoiceNo, includeTax: tax, html: html, header: od.header, items: od.items || [] });
  } catch (e) {
    return { ok: false, error: e && e.message ? e.message : String(e) };
  }
}


  function reserveInvoiceAndBuildHtml(rowIndex, invoiceNo, username) {
    try {
      rowIndex = Number(rowIndex || 0);
      if (!rowIndex) return { ok: false, error: "Thiếu rowIndex." };

      invoiceNo = String(invoiceNo || "").trim();
      if (!invoiceNo) invoiceNo = getNextInvoiceNo_();

      const rsv = reserveInvoiceNo(invoiceNo, username || "", String(rowIndex));
      if (!rsv || rsv.ok === false) return rsv;

      const od = getOrderDetailByDebtRow(rowIndex);
      if (!od || od.ok === false) return od;

      const html = buildInvoiceHtml(od.header, od.items, rsv.invoiceNo);
      return { ok: true, invoiceNo: rsv.invoiceNo, reused: !!rsv.reused, html: html, header: od.header };
    } catch (e) {
      return { ok: false, error: e && e.message ? e.message : String(e) };
    }
  }

  /* ✅ PENDING PRINT: GIỐNG CƠ CHẾ CŨ (log_print + cấp số HĐ), KHÔNG “TAM-…” */
  function getInvoiceHtmlByPendingRow(rowIndex, username, includeTax) {
  try {
    rowIndex = Number(rowIndex || 0);
    if (!rowIndex) return { ok: false, error: "Thiếu rowIndex." };

    const tagRow = "P:" + String(rowIndex);

    const existed = getInvoiceNoByDebtRow(tagRow);
    let invoiceNo = (existed && existed.ok) ? String(existed.invoiceNo || "").trim() : "";

    if (!invoiceNo) {
      invoiceNo = getNextInvoiceNo_();
      const rsv = reserveInvoiceNo(invoiceNo, username || "", tagRow);
      if (!rsv || rsv.ok === false) {
        return { ok: false, error: (rsv && rsv.error) ? rsv.error : "Reserve số hoá đơn thất bại." };
      }
      invoiceNo = rsv.invoiceNo;
    }

    const od = getOrderDetailByPendingRow(rowIndex);
    if (!od || od.ok === false) {
      return { ok: false, error: (od && od.error) ? od.error : "Không lấy được chi tiết đơn chờ." };
    }

    if (od && od.header) {
      od.header.trangThai = STATUS_UNPAID;
      od.header.status = STATUS_UNPAID;
    }

    const tax = (typeof includeTax === "undefined" || includeTax === null) ? getEnableTax_() : !!includeTax;

    const html = buildInvoiceHtml(od.header, od.items || [], invoiceNo, tax);

    return jsonSafe_({
      ok: true,
      invoiceNo: invoiceNo,
      includeTax: tax,
      html: html,
      header: od.header,
      items: od.items || []
    });
  } catch (e) {
    return { ok: false, error: e && e.message ? e.message : String(e) };
  }
}



  /* =====================================================================================
    17) DEBUG + GPT STUB + include (GIỮ NGUYÊN)
    ===================================================================================== */
  function debugPing() {
    return { ok: true, msg: "Ping OK from server", time: new Date().toISOString() };
  }

  function chatWithDebtData(messages) {
    return { ok: false, error: 'Chức năng GPT chưa được cấu hình trong script.' };
  }

  function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  }


  /* =====================================================================================
    18) AUTH + LOGIN + doGet + đổi mật khẩu (GIỮ NGUYÊN)
    ===================================================================================== */
  function getSecretKey_() {
    var props = PropertiesService.getScriptProperties();
    var key   = props.getProperty('AUTH_SECRET_KEY');
    var time  = props.getProperty('AUTH_SECRET_TIME');

    var now = new Date();
    var needNew = true;
    var twoHoursMs = 2 * 60 * 60 * 1000;

    if (key && time) {
      var created = new Date(time);
      if (now.getTime() - created.getTime() < twoHoursMs) needNew = false;
    }

    if (needNew) {
      var newKey = Utilities.getUuid() + '_' + now.getTime();
      props.setProperty('AUTH_SECRET_KEY', newKey);
      props.setProperty('AUTH_SECRET_TIME', now.toISOString());
      key = newKey;
    }

    return key;
  }
  function parseBool_(v) {
    if (v === true) return true;
    if (v === false) return false;
    var s = String(v == null ? '' : v).trim().toLowerCase();
    return s === 'true' || s === '1' || s === 'yes' || s === 'y' || s === 'on' || s === 'x';
  }

  function normalizeMenuStatus01_(v) {
    var s = String(v == null ? "" : v).trim();
    if (!s) return 1; // default active
    var sl = s.toLowerCase();

    if (sl === "active") return 1;
    if (sl === "inactive") return 0;

    // boolean-like
    if (parseBool_(sl) === true) return 1;

    // numeric-like
    var n = Number(s);
    if (!isNaN(n)) return n ? 1 : 0;

    // default: non-empty => active
    return 1;
  }

  function getAccountByUsername_(username) {
    username = String(username || '').trim();
    if (!username) return null;

    // cache account (giảm đọc sheet account mỗi lần gọi)
    const __accKey = "ACC_V1:" + username.toLowerCase();
    try {
      const cache = CacheService.getScriptCache();
      const cached = cache.get(__accKey);
      if (cached) {
        if (cached === "NULL") return null;
        try { return JSON.parse(cached); } catch(e) {}
      }
    } catch(e) {}

    var ss = getSpreadsheet_();
    var sh = ss.getSheetByName(SHEET_ACCOUNT);
    if (!sh) throw new Error('Không tìm thấy sheet account: ' + SHEET_ACCOUNT);

    ensureAccountHeader_(sh);

    var values = sh.getDataRange().getValues();
    if (values.length < 2) return null;

    var header = values[0].map(function(h) { return String(h || '').toLowerCase().trim(); });

    var idxUser   = header.indexOf('username');
    var idxPass   = header.indexOf('password');
    var idxDN     = header.indexOf('displayname');
    if (idxDN === -1) idxDN = header.indexOf('display_name');
    var idxStatus = header.indexOf('status');
    var idxRole   = header.indexOf('role');
    var idxMust   = header.indexOf('must_change_password');

    // Heuristic fallback: nếu header đang ghi "status" nhưng data cột 3 thực tế là displayName,
    // và status thật nằm ở cột 5 (E) (active/inactive/...)
    if ((idxDN === -1) && (idxStatus === 2) && (values.length >= 2)) {
      var sample = values[1] || [];
      var u0 = String(sample[idxUser] || '').trim();
      var vStatusCol = String(sample[idxStatus] || '').trim().toLowerCase();
      var vReal = (sample.length >= 5) ? String(sample[4] || '').trim().toLowerCase() : '';
      function isStatusLike__(v){
        return (v === '1' || v === '0' || v === 'active' || v === 'inactive' || v === 'on' || v === 'off' || v === 'true' || v === 'false' || v === 'yes' || v === 'no');
      }
      if ((u0 && String(sample[2] || '').trim() === u0) || (!isStatusLike__(vStatusCol) && isStatusLike__(vReal))) {
        idxDN = 2;
        idxStatus = 4;
        idxMust = 5;
      }
    }

    var idxRole   = header.indexOf('role');

    // ✅ fallback: role nằm cột D (index 3) nếu không có header "role"
    if (idxRole === -1 && header.length >= 4) idxRole = 3;

    if (idxUser === -1 || idxPass === -1) throw new Error('Sheet account thiếu header username/password');

    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      var u = String(row[idxUser] || '').trim();
      if (!u) continue;
      if (u !== username) continue;

      var st = idxStatus !== -1 ? String(row[idxStatus] || '').trim() : '';
      if (st) {
        var stNorm = st.toLowerCase();
        var isActive = (stNorm === '1' || stNorm === 'active' || stNorm === 'on' || stNorm === 'true' || stNorm === 'yes');
        if (!isActive) return null;
      }

      var mustChange = false;
      if (idxMust !== -1) {
        var raw = row[idxMust];
        if (typeof raw === 'boolean') mustChange = raw;
        else {
          var flagRaw = String(raw || '').trim().toUpperCase();
          mustChange = (flagRaw === 'TRUE' || flagRaw === 'T' || flagRaw === 'YES' || flagRaw === 'Y' || flagRaw === '1');
        }
      }

          var out = {
  rowIndex: i + 1,
        username: u,
        displayName: (function(){
          var idxDN = header.indexOf('displayname');
          if (idxDN === -1) idxDN = header.indexOf('display_name');
          if (idxDN === -1) idxDN = header.indexOf('name');
          return String((idxDN !== -1 ? row[idxDN] : '') || u);
        })(),
        password: String(row[idxPass] || ''),
        status: (function(){
          var rawSt = idxStatus !== -1 ? String(row[idxStatus] || '') : '';
          var rawRole = idxRole !== -1 ? String(row[idxRole] || '') : '';
          function isStatusLike__(v){
            v = String(v||'').trim().toLowerCase();
            if (!v) return false;
            if (v === 'active' || v === 'inactive' || v === 'on' || v === 'off') return true;
            if (v === '1' || v === '0' || v === 'true' || v === 'false') return true;
            if (v === 'yes' || v === 'no' || v === 'enable' || v === 'enabled' || v === 'disable' || v === 'disabled') return true;
            if (v.indexOf('act') === 0) return true;
            if (v.indexOf('inact') === 0) return true;
            return false;
          }
          function isRoleLike__(v){
            v = String(v||'').trim().toLowerCase();
            return v === 'admin' || v === 'manager' || v === 'cashier' || v === 'user' || v === 'staff';
          }
          if (isStatusLike__(rawRole) && isRoleLike__(rawSt)) {
            var tmp = rawSt; rawSt = rawRole; rawRole = tmp;
          }
          return normalizeAccountStatus_(rawSt);
        })(),
        role: (function(){
          var rawSt = idxStatus !== -1 ? String(row[idxStatus] || '') : '';
          var rawRole = idxRole !== -1 ? String(row[idxRole] || '') : '';
          function isStatusLike__(v){
            v = String(v||'').trim().toLowerCase();
            if (!v) return false;
            if (v === 'active' || v === 'inactive' || v === 'on' || v === 'off') return true;
            if (v === '1' || v === '0' || v === 'true' || v === 'false') return true;
            if (v === 'yes' || v === 'no' || v === 'enable' || v === 'enabled' || v === 'disable' || v === 'disabled') return true;
            if (v.indexOf('act') === 0) return true;
            if (v.indexOf('inact') === 0) return true;
            return false;
          }
          function isRoleLike__(v){
            v = String(v||'').trim().toLowerCase();
            return v === 'admin' || v === 'manager' || v === 'cashier' || v === 'user' || v === 'staff';
          }
          if (isStatusLike__(rawRole) && isRoleLike__(rawSt)) {
            var tmp = rawSt; rawSt = rawRole; rawRole = tmp;
          }
          return normalizeAccountRole_(rawRole);
        })(),
        mustChangePassword: mustChange,
        idxMustChange: idxMust,
        idxPassword: idxPass
      };
      try { CacheService.getScriptCache().put(__accKey, JSON.stringify(out), 600); } catch(e) {}
      return out;
    }
  try {
    CacheService.getScriptCache().put(__accKey, "NULL", 300);
  } catch (e) {}
  return null;
  }
  function changePasswordForCurrentUser(username, oldPw, newPw) {
    return updateUserPassword(username, oldPw, newPw);
  }



  function generateToken_(username, displayName) {
    var secret = getSecretKey_();
    var data = username + '|' + displayName;
    var sigBytes = Utilities.computeHmacSha256Signature(data, secret);
    var sig = Utilities.base64Encode(sigBytes);
    return data + '|' + sig;
  }

  function validateToken_(token) {
    if (!token) return null;
    try {
      var parts = token.split('|');
      if (parts.length !== 3) return null;

      var username    = parts[0];
      var displayName = parts[1];
      var sig         = parts[2];

      var secret = getSecretKey_();
      var data = username + '|' + displayName;
      var expectedBytes = Utilities.computeHmacSha256Signature(data, secret);
      var expectedSig   = Utilities.base64Encode(expectedBytes);

      if (sig !== expectedSig) return null;

      var acc = getAccountByUsername_(username);
      if (!acc) return null;

      return { username: acc.username, displayName: acc.displayName || displayName || acc.username };
    } catch (err) {
      return null;
    }
  }
  /* =====================================================================================
    UI: DONCHO (pending) + LỊCH SỬ (congno) + IN HÓA ĐƠN
    - Dán ở cuối file Code.gs
    ===================================================================================== */
function resetPendingPicker_() {
  try { PropertiesService.getScriptProperties().deleteProperty("PENDING_SHEET_NAME"); } catch(e) {}
  try { CacheService.getScriptCache().remove("PENDING_LIST_V1"); } catch(e) {}
}

  function ui_listPendingAll() {
  // cache ngắn hạn để tránh reload liên tục (tab đơn chờ)
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get("PENDING_LIST_V1");
    if (cached) {
      try { return JSON.parse(cached) || []; } catch(e) {}
    }
  } catch(e) {}

  // Đọc A2:Q (doncho) bằng Sheets API (nhanh hơn), fallback SpreadsheetApp nếu chưa bật service
  let vals = [];
  try {
    const v = sh_valuesGet_(SHEET_PENDING + "!A2:Q");
    if (v) vals = v;
  } catch(e) {}

  if (!vals || !vals.length) {
    const sh = getPendingSheet_();
    const lastRow = sh ? sh.getLastRow() : 0;
    if (lastRow < 2) return [];
    vals = sh.getRange(2, 1, lastRow - 1, 17).getValues(); // A:Q
  }
  const tz = Session.getScriptTimeZone();
  const out = [];

  for (let i = 0; i < vals.length; i++) {
    const r = vals[i] || [];
    const rowIndex = i + 2;

    const orderId = String((r[12] || "")).trim();
    // skip completely empty rows
    if (!orderId && !r[0] && !r[1]) continue;

    // date
    let dateObj = parseDateCell_(r[0]);
    const dateStr = dateObj ? Utilities.formatDate(dateObj, tz, "dd/MM/yyyy") : (r[0] ? String(r[0]) : "");
    const dateRaw = dateObj ? dateObj.getTime() : "";


    const soMon = Number(r[13] || 0) || 0; // N: Số món (lưu sẵn, không scan sheet data)
      const itemStart = Number(r[14] || 0) || 0; // O: start row data
      const itemEnd   = Number(r[15] || 0) || 0; // P: end row data
      const datCoc    = Number(r[16] || 0) || 0; // Q: Đặt cọc

    out.push({
      row: rowIndex,
      rowIndex: rowIndex,
      dateRaw: dateObj ? dateObj.getTime() : 0,
      dateStr: dateStr,
      date: dateStr,
      info: String(r[1] || ""),                 // B Thông tin
      soMam: Number(r[2] || 0) || 0,            // C Số mâm
      soMon: soMon,
      itemStart: itemStart,
      itemEnd: itemEnd,
      donGiaMam: Number(r[3] || 0) || 0,        // D Đơn giá mâm
      tongDon: Number(r[4] || 0) || 0,          // E Tổng đơn
      doanhSo: Number(r[7] || 0) || 0,          // H Doanh số
      datCoc: datCoc,                           // Q Đặt cọc
      trangThai: "Đơn chờ",
      orderId: orderId
    });
  }

  out.sort((a,b)=> (b.dateRaw||0)-(a.dateRaw||0));
  out.forEach(x=>{ delete x.dateRaw; });

  try {
    const cache = CacheService.getScriptCache();
    cache.put("PENDING_LIST_V1", JSON.stringify(out), 15);
  } catch(e) {}

  return out;
}




  // ==========================
// FAST ENDPOINTS (UI ONLY)
// - chỉ trả dữ liệu bảng, không tính KPI/dish report
// ==========================
function ui_getDebtFullFast() {
  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName(SHEET_DEBT);
  if (!sh) return { headers: [], rows: [], rowIndices: [] };

  const lastRow = sh.getLastRow();
  const lastCol = Math.max(16, sh.getLastColumn() || 16);

  if (lastRow < 1) return { headers: [], rows: [], rowIndices: [] };

  let values = [];
  try {
    if (hasSheetsApi_()) {
      // Read fixed A:P for speed + predictable schema
      const rng = SHEET_DEBT + "!A1:P" + lastRow;
      const res = Sheets.Spreadsheets.Values.get(SPREADSHEET_ID, rng);
      values = (res && res.values) ? res.values : [];
    }
  } catch (e) {}

  if (!values || !values.length) {
    values = sh.getRange(1, 1, lastRow, 16).getValues();
  } else {
    // pad rows to 16 cols
    for (let i = 0; i < values.length; i++) {
      values[i] = values[i] || [];
      while (values[i].length < 16) values[i].push("");
    }
  }

  const headers = (values[0] || []).map(v => String(v == null ? "" : v));
  const rows = [];
  const rowIndices = [];
  for (let r = 1; r < values.length; r++) {
    rows.push(values[r]);
    rowIndices.push(r + 1);
  }
  return { headers: headers, rows: rows, rowIndices: rowIndices };
}

function ui_getCustomersFullFast() {
  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName(SHEET_KH);
  if (!sh) return { headers: [], rows: [], agg: {} };

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn() || 1;
  if (lastRow < 1) return { headers: [], rows: [], agg: {} };

  let values = [];
  try {
    if (hasSheetsApi_()) {
      const rng = SHEET_KH + "!A1:" + colToA1_(lastCol) + lastRow;
      const res = Sheets.Spreadsheets.Values.get(SPREADSHEET_ID, rng);
      values = (res && res.values) ? res.values : [];
    }
  } catch (e) {}

  if (!values || !values.length) {
    values = sh.getRange(1, 1, lastRow, lastCol).getValues();
  } else {
    for (let i = 0; i < values.length; i++) {
      values[i] = values[i] || [];
      while (values[i].length < lastCol) values[i].push("");
    }
  }

  const rawHeaders = (values[0] || []).map(v => String(v == null ? "" : v));
  const hasIndexCol = rawHeaders.length && !String(rawHeaders[0] || "").trim();

  // drop index column so UI assumes col0 = SĐT
  const headers = hasIndexCol ? rawHeaders.slice(1) : rawHeaders;

  const rows = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r] || [];
    rows.push(hasIndexCol ? row.slice(1) : row);
  }

  // build agg by scanning debt sheet
  const agg = {};
  try {
    const shDebt = ss.getSheetByName(SHEET_DEBT);
    if (shDebt) {
      const off = getLeadingIndexOffset_(shDebt);
      const baseCol = 1 + off;
      const lastDebtRow = shDebt.getLastRow();
      if (lastDebtRow >= DEBT_DATA_START) {
        const debtVals = shDebt.getRange(DEBT_DATA_START, baseCol, lastDebtRow - DEBT_DATA_START + 1, DEBT_LAST_COL).getValues();
        for (let i = 0; i < debtVals.length; i++) {
          const r = debtVals[i] || [];
          const info = String(r[DEBT_COL_INFO - 1] || "");
          const phone = normalizePhone_(extractPhoneFromInfo_(info));
          if (!phone) continue;

          const doanhSo = toMoneyNumber_(r[DEBT_COL_DOANHSO - 1] || 0);
          const congNo  = toMoneyNumber_(r[DEBT_COL_CONGNO - 1] || 0);
          if (!agg[phone]) agg[phone] = { totalOrdered: 0, totalDebt: 0 };
          agg[phone].totalOrdered += doanhSo;
          agg[phone].totalDebt += congNo;
        }
      }
    }
  } catch (e) {}

  // derive paid
  Object.keys(agg).forEach(k => {
    const a = agg[k];
    a.totalOrdered = Math.round(a.totalOrdered || 0);
    a.totalDebt = Math.round(a.totalDebt || 0);
    a.totalPaid = Math.max(0, a.totalOrdered - a.totalDebt);
  });

  return { headers: headers, rows: rows, agg: agg };
}

function colToA1_(n) {
  n = Number(n || 0);
  if (!n || n < 1) return "A";
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function ui_deletePending(rowIndex, actorOrPayload) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      // Chỉ admin được xoá (cashier/manager bị chặn)
      assertCanDelete_(actorOrPayload);

      rowIndex = Number(rowIndex || 0);
      if (!rowIndex || rowIndex < DEBT_DATA_START) throw new Error("Row không hợp lệ.");

      const ss = getSpreadsheet_();
      const shP = getPendingSheet_();
      const shData = ss.getSheetByName(SHEET_DATA);
      if (!shData) throw new Error("Không tìm thấy sheet '" + SHEET_DATA + "'.");

      const orderId = String(shP.getRange(rowIndex, DEBT_COL_ORDER_ID).getValue() || "").trim();

      // Hard delete: xóa hoàn toàn thông tin đặt hàng theo orderId
      if (orderId) {
        const span = findOrderIdRowSpan_(shData, getDataSheetMap_(shData).idxOid, orderId);
        if (span && span.startRow && span.endRow) {
          const n = span.endRow - span.startRow + 1;
          shData.deleteRows(span.startRow, n);
        }
      }

      // Xóa hàng trong sheet doncho
      sh_deleteRowFast_(SHEET_PENDING, rowIndex);
      try { CacheService.getScriptCache().remove("PENDING_LIST_V1"); } catch(e) {}
      return { ok: true, deletedRow: rowIndex };
    } catch (e) {
      return { ok: false, error: e && e.message ? e.message : String(e) };
    } finally {
      try { lock.releaseLock(); } catch (e2) {}
    }
  }

  function getPendingOrderByRow(rowIndex) {
    return getOrderDetailByPendingRow(rowIndex);
  }

  function ui_getPendingDetail(rowIndex) {
  try {
    rowIndex = Number(rowIndex || 0);
    if (!rowIndex || rowIndex < 2) throw new Error("Row không hợp lệ.");

    let row = [];
    try {
      if (hasSheetsApi_()) row = sh_readRowA1_(SHEET_PENDING, rowIndex, 17); // A:Q
    } catch (e) {}

    if (!row || !row.length) {
      const sh = getPendingSheet_();
      if (!sh) throw new Error("Không tìm thấy sheet '" + SHEET_PENDING + "'.");
      row = (sh.getRange(rowIndex, 1, 1, 17).getValues()[0] || []);
    }

    const tz = Session.getScriptTimeZone();
    const ngay = parseDateCell_(row[0]) || new Date();
    const info = String(row[1] || "");
    const tenKH = getNameFromInfo_(info) || "";
    const sdt = normalizePhone_(extractPhoneFromInfo_(info));
    const diaChi = getAddressFromInfo_(info) || "";const soMam = Number(row[2] || 0) || 0;
    const donGiaMam = Number(row[3] || 0) || 0;
    const tongDon = Number(row[4] || 0) || 0;
    const kmNoiDung = String(row[5] || "");
    const kmSoTien = Number(row[6] || 0) || 0;
    const doanhSo = Number(row[7] || 0) || 0;
    const nguoiLap = String(row[11] || "");
    const orderId = String(row[12] || "").trim();

    const itemStart = Number(row[14] || 0) || 0; // O
    const itemEnd   = Number(row[15] || 0) || 0; // P
    const datCoc    = Number(row[16] || 0) || 0; // Q

    const ss = getSpreadsheet_();
    const shData = ss.getSheetByName(SHEET_DATA);
    if (!shData) throw new Error("Không tìm thấy sheet '" + SHEET_DATA + "'.");

    // Ưu tiên lấy theo orderId để đảm bảo chính xác, span chỉ là backup
    const items = orderId ? getItemsByOrderId_(shData, orderId)
                : (itemStart && itemEnd ? getItemsBySpan_(shData, itemStart, itemEnd) : []);

    return {
      ok: true,
      header: {
        ngay: Utilities.formatDate(ngay, tz, "yyyy-MM-dd"),
        dateStr: Utilities.formatDate(ngay, tz, "dd/MM/yyyy"),
        ngayStr: Utilities.formatDate(ngay, tz, "dd/MM/yyyy"),
        tenKH: tenKH,
        sdt: sdt,
        diaChi: diaChi,
        soMam: soMam,
        donGiaMam: donGiaMam,
        tongDon: tongDon,
        kmNoiDung: kmNoiDung,
        kmSoTien: kmSoTien,
        doanhSo: doanhSo,
        datCoc: datCoc,
        nguoiLap: nguoiLap,
        orderId: orderId,
        itemStart: itemStart,
        itemEnd: itemEnd
      },
      items: items || []
    };
  } catch (e) {
    return { ok: false, error: e && e.message ? e.message : String(e) };
  }
}

 function ui_getDebtDetail(rowIndex) {
  try {
    rowIndex = Number(rowIndex || 0);
    if (!rowIndex || rowIndex < 2) throw new Error("Row không hợp lệ.");

    // Read A:P (16 cột)
    let row = [];
    try {
      if (hasSheetsApi_()) row = sh_readRowA1_(SHEET_DEBT, rowIndex, 16);
    } catch (e) {}

    if (!row || !row.length) {
      const ss = getSpreadsheet_();
      const sh = ss.getSheetByName(SHEET_DEBT);
      if (!sh) throw new Error("Không tìm thấy sheet '" + SHEET_DEBT + "'.");
      row = sh.getRange(rowIndex, 1, 1, 16).getValues()[0] || [];
    }

    const tz = Session.getScriptTimeZone();
    const dateObj = parseDateCell_(row[DEBT_COL_NGAYTT - 1]);
    const dateISO = dateObj ? Utilities.formatDate(dateObj, tz, "yyyy-MM-dd") : "";
    const dateStr = dateObj ? Utilities.formatDate(dateObj, tz, "dd/MM/yyyy") : "";

    const info = String(row[DEBT_COL_INFO - 1] || "").trim();
    const tenKH = getNameFromInfo_(info);
    const sdt   = String(extractPhoneFromInfo_(info) || "").trim();
    const diaChi = getAddressFromInfo_(info);

    const soMam     = Number(row[DEBT_COL_SOMAM - 1] || 0) || 0;
    const donGiaMam = toMoneyNumber_(row[DEBT_COL_DONGIA_MAM - 1] || 0);
    const tongDon   = toMoneyNumber_(row[DEBT_COL_TONG_DON - 1] || 0);

    const kmNoiDung = String(row[DEBT_COL_KM_NOTE - 1] || "").trim();
    const kmSoTien  = toMoneyNumber_(row[DEBT_COL_KM_AMOUNT - 1] || 0);

    const doanhSo = toMoneyNumber_(row[DEBT_COL_DOANHSO - 1] || 0);
    const congNo  = toMoneyNumber_(row[DEBT_COL_CONGNO - 1] || 0);

    const trangThai = String(row[DEBT_COL_STATUS - 1] || "").trim();
    const ngayThanhToan = row[DEBT_COL_NGAYTT - 1] || "";
    const thuNgan = String(row[DEBT_COL_THUNGAN - 1] || "").trim();

    const orderId = String(row[DEBT_COL_ORDER_ID - 1] || "").trim();
    const itemCount = Number(row[DEBT_COL_ITEM_COUNT - 1] || 0) || 0;
    const itemStart = Number(row[DEBT_COL_ITEM_START - 1] || 0) || 0;
    const itemEnd   = Number(row[DEBT_COL_ITEM_END - 1] || 0) || 0;

    const ss = getSpreadsheet_();
    const shData = ss.getSheetByName(SHEET_DATA);

    let items = [];
    if (shData && itemStart && itemEnd && itemEnd >= itemStart) {
      items = getItemsByRowSpan_(shData, itemStart, itemEnd, orderId);
    } else {
      items = orderId ? getItemsByOrderId_(shData, orderId) : [];
    }

    return {
      ok: true,
      header: {
        rowIndex: rowIndex,
        row: rowIndex,

        // UI đang dùng h.date
        date: dateStr,
        dateISO: dateISO,

        // giữ tương thích chỗ khác
        ngay: dateStr,
        ngayStr: dateStr,
        dateStr: dateStr,

        tenKH: tenKH,
        sdt: sdt,
        diaChi: diaChi,

        soMam: soMam,
        donGiaMam: donGiaMam,
        tongDon: tongDon,

        kmNoiDung: kmNoiDung,
        kmSoTien: kmSoTien,

        doanhSo: doanhSo,
        congNo: congNo,

        trangThai: trangThai,
        ngayThanhToan: ngayThanhToan,
        thuNgan: thuNgan,

        orderId: orderId,
        itemCount: itemCount,
        itemStart: itemStart,
        itemEnd: itemEnd
      },
      items: items || []
    };
  } catch (e) {
    return { ok: false, error: e && e.message ? e.message : String(e) };
  }
}



  function ui_savePendingEdit(payload) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);

    try {
      payload = payload || {};
      const meta = payload.meta || {};
      const ctx = assertCanEdit_(payload);
      if (!meta.username) meta.username = ctx.username;
      if (!meta.role) meta.role = ctx.role;
      const h = payload.header || {};
      const items = payload.items || [];

      const rowIndex = Number(meta.row || meta.rowIndex || payload.rowIndex || payload.row || payload.pendingRow || 0);
      if (!rowIndex || rowIndex < 2) throw new Error("Row pending không hợp lệ.");

      const ss = getSpreadsheet_();
      const shP = getPendingSheet_();
      const shDebt = ss.getSheetByName(SHEET_DEBT);
      const shData = ss.getSheetByName(SHEET_DATA);
      if (!shDebt) throw new Error("Không tìm thấy sheet '" + SHEET_DEBT + "'.");
      if (!shData) throw new Error("Không tìm thấy sheet '" + SHEET_DATA + "'.");

      const ngayRaw = h.ngay || h.date || shP.getRange(rowIndex, 1).getValue() || new Date();
      const ngay = parseDateCell_(ngayRaw) || new Date();
      ngay.setHours(0,0,0,0);


      const tenKH = String(h.tenKH || "").trim();
      const sdt = String(h.sdt || "").trim();
      const diaChi = String(h.diaChi || "").trim();

      const soMam = Number(h.soMam || 1) || 1;
      const kmNoiDung = String(h.kmNoiDung || "").trim();
      const kmSoTien = Number(h.kmSoTien || 0) || 0;
      const trangThai = String(h.trangThai || h.status || "").trim();
      const nguoiLapRaw = String(h.nguoiLap || meta.username || "").trim();
      const nguoiLap = formatDisplayName_(nguoiLapRaw);

      const orderIdCell = String(shP.getRange(rowIndex, 13).getValue() || "").trim();
      let orderId = String(meta.orderId || orderIdCell || "").trim();

      if (!orderId) {
        orderId = makeOrderId_(ngay, sdt);
        try { shP.getRange(rowIndex, 13).setValue(orderId); } catch(e) {}
      }
      if (!orderId) throw new Error("Thiếu orderId.");

      // tính totals trước để có donGiaMam đúng
      const calc = calcTotalsFromItems_(items, soMam, kmSoTien);
      const doanhSo = calc.doanhSo;

      // update items: xoá cũ + ghi mới (append) và lấy start/end row
      deleteItemsByOrderId_(shData, orderId);
      const span = appendItemsToData_(shData, ngay, tenKH, sdt, diaChi, items, orderId, calc.donGiaMam);
      const itemCount = span && span.count ? span.count : (Array.isArray(items) ? items.length : 0);

      let congNo = "";
      let ngayTT = "";
      let statusWrite = "";

      if (trangThai) {
        if (trangThai !== STATUS_PAID && trangThai !== STATUS_DEBT && trangThai !== STATUS_UNPAID) {
          throw new Error("Trạng thái không hợp lệ.");
        }
        statusWrite = trangThai;
        congNo = (trangThai === STATUS_PAID) ? 0 : doanhSo;
        ngayTT = (trangThai === STATUS_PAID) ? new Date() : "";
      }

      const info = tenKH + " - " + sdt + " - " + diaChi;

      const rowValuesDebt = [
        new Date(ngay),     // A
        info,               // B
        soMam,              // C
        calc.donGiaMam,     // D
        calc.tongDon,       // E
        kmNoiDung,          // F
        kmSoTien,           // G
        doanhSo,            // H
        congNo,             // I
        statusWrite,        // J
        ngayTT,             // K
        nguoiLap,           // L
        orderId,            // M
        itemCount           // N
      ];

      if (!statusWrite) {
        // vẫn là pending -> update doncho A:P (16 cols)
        const rowValuesPending = rowValuesDebt.concat([
          span && span.startRow ? span.startRow : "", // O
          span && span.endRow ? span.endRow : ""      // P
        ]);
        const offP = getLeadingIndexOffset_(shP);
        const baseColP = 1 + offP;
        if (offP) {
          try { shP.getRange(rowIndex, 1).setValue(rowIndex - PENDING_DATA_START + 1); } catch(e) {}
        }
        shP.getRange(rowIndex, baseColP, 1, rowValuesPending.length).setValues([rowValuesPending]);

        try { CacheService.getScriptCache().remove("PENDING_LIST_V1"); } catch(e) {}

        const tz = Session.getScriptTimeZone();
        const updatedRow = {
          row: rowIndex,
          rowIndex: rowIndex,
          dateStr: Utilities.formatDate(new Date(ngay), tz, "dd/MM/yyyy"),
          info: info,
          soMam: soMam,
          soMon: itemCount,
          donGiaMam: calc.donGiaMam,
          tongDon: calc.tongDon,
          trangThai: "Đơn chờ",
          orderId: orderId
        };

        return { ok: true, moved: false, orderId: orderId, updatedRow: updatedRow };
      }

      // chuyển sang congno: append + xoá doncho
      shDebt.appendRow(rowValuesDebt);
      sh_deleteRowFast_(SHEET_PENDING, rowIndex);

      try { CacheService.getScriptCache().remove("PENDING_LIST_V1"); } catch(e) {}
      return { ok: true, moved: true, orderId: orderId, deletedRow: rowIndex };
    } catch (e) {
      return { ok: false, error: e && e.message ? e.message : String(e) };
    } finally {
      try { lock.releaseLock(); } catch (e2) {}
    }
  }

  function ui_listDebtHistoryByPhone(phone) {
    try {
      const pNorm = normalizePhone_(phone);
      if (!pNorm) return [];

      const __key = "DEBT_HIST_V1:" + pNorm;
      try {
        const cached = CacheService.getScriptCache().get(__key);
        if (cached) {
          try { return JSON.parse(cached) || []; } catch(e) {}
        }
      } catch(e) {}

      let vals = [];
      try {
        const v = sh_valuesGet_(SHEET_DEBT + "!A2:P");
        if (v) vals = v;
      } catch(e) {}

      if (!vals || !vals.length) {
        const ss = getSpreadsheet_();
        const sh = ss.getSheetByName(SHEET_DEBT);
        if (!sh) return [];
        const last = sh.getLastRow();
        if (last < 2) return [];
        vals = sh.getRange(2, 1, last - 1, 13).getValues();
      }

      const tz = Session.getScriptTimeZone();
      const out = [];

      for (let i = 0; i < vals.length; i++) {
        const rowIndex = i + 2;
        const r = vals[i] || [];

        const phoneCell = normalizePhone_(r[DEBT_COL_SDT - 1]);
        if (phoneCell !== pNorm) continue;

        const info = String(r[DEBT_COL_INFO - 1] || "").trim();
        const dateObj = parseDateCell_(r[DEBT_COL_DATE - 1]);
        const orderId = String(r[DEBT_COL_ORDER_ID - 1] || "").trim();
        const statusVal = String(r[DEBT_COL_STATUS - 1] || "").trim();

        out.push({
          row: rowIndex,
          dateRaw: dateObj ? dateObj.getTime() : 0,
          dateStr: dateObj ? Utilities.formatDate(dateObj, tz, "dd/MM/yyyy") : "",
          info: info,
          soMam: Number(r[DEBT_COL_SOMAM - 1] || 0) || 0,
          tongDon: Number(r[DEBT_COL_TONG_DON - 1] || 0) || 0,
          trangThai: statusVal,
          status: statusVal,
          ngayTT: r[DEBT_COL_DATE_TT - 1] || "",
          thuNgan: String(r[DEBT_COL_THU_NGAN - 1] || ""),
          orderId: orderId
        });
      }

      // sort mới → cũ
      out.sort((a, b) => (b.dateRaw - a.dateRaw) || (b.row - a.row));

      try { CacheService.getScriptCache().put(__key, JSON.stringify(out), 30); } catch(e) {}
      return out;
    } catch(e) {
      return [];
    }
  }



  function ui_getInvoiceHtmlByDebtRow(rowIndex) {
    const detail = ui_getDebtDetail(rowIndex);
    if (!detail || !detail.ok) return "<div style='padding:12px'>Không lấy được dữ liệu hóa đơn.</div>";
    return ui_buildInvoiceHtml_(detail);
  }

  function ui_buildInvoiceHtml_(detail) {
    const h = detail.header || {};
    const items = detail.items || [];

    const fmt = (n) => {
      n = Number(n || 0) || 0;
      return n.toLocaleString("vi-VN");
    };

    let sum = 0;
    let rows = "";
    for (let i = 0; i < items.length; i++) {
      const it = items[i];
      const tt = Number(it.thanhTien || (Number(it.sl || 0) * Number(it.donGia || 0)) || 0) || 0;
      sum += tt;
      rows += `
        <tr>
          <td style="padding:6px 6px;border-bottom:1px solid #eee">${i + 1}</td>
          <td style="padding:6px 6px;border-bottom:1px solid #eee">${String(it.tenMon || "")}</td>
          <td style="padding:6px 6px;border-bottom:1px solid #eee;text-align:right">${fmt(it.sl)}</td>
          <td style="padding:6px 6px;border-bottom:1px solid #eee;text-align:right">${fmt(it.donGia)}</td>
          <td style="padding:6px 6px;border-bottom:1px solid #eee;text-align:right">${fmt(tt)}</td>
        </tr>
      `;
    }

    const soMam = Number(h.soMam || 1) || 1;
    const km = Number(h.kmSoTien || 0) || 0;
    const tongDon = sum * soMam;
    const doanhSo = Math.max(0, tongDon - km);

    return `
    <div style="font-family:Arial, sans-serif; padding:12px">
      <div style="text-align:center; font-weight:700; font-size:16px">PHIẾU / HÓA ĐƠN</div>
      <div style="margin-top:8px; font-size:13px">
        <div><b>Ngày:</b> ${h.ngay || ""}</div>
        <div><b>Mã đơn:</b> ${h.orderId || ""}</div>
        <div><b>Khách:</b> ${h.tenKH || ""} - ${h.sdt || ""}</div>
        <div><b>Địa chỉ:</b> ${h.diaChi || ""}</div>
        <div><b>Số mâm:</b> ${soMam}</div>
        <div><b>Trạng thái:</b> ${h.trangThai || ""}</div>
      </div>

      <table style="width:100%; border-collapse:collapse; margin-top:10px; font-size:13px">
        <thead>
          <tr>
            <th style="text-align:left; padding:6px 6px; border-bottom:1px solid #ddd">#</th>
            <th style="text-align:left; padding:6px 6px; border-bottom:1px solid #ddd">Món</th>
            <th style="text-align:right; padding:6px 6px; border-bottom:1px solid #ddd">SL</th>
            <th style="text-align:right; padding:6px 6px; border-bottom:1px solid #ddd">Đơn giá</th>
            <th style="text-align:right; padding:6px 6px; border-bottom:1px solid #ddd">Thành tiền</th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>

      <div style="margin-top:10px; font-size:13px">
        <div style="display:flex; justify-content:space-between"><span>Tổng 1 mâm:</span><b>${fmt(sum)}</b></div>
        <div style="display:flex; justify-content:space-between"><span>Tổng đơn:</span><b>${fmt(tongDon)}</b></div>
        <div style="display:flex; justify-content:space-between"><span>Khuyến mãi:</span><b>${fmt(km)}</b></div>
        <div style="display:flex; justify-content:space-between"><span>Doanh số:</span><b>${fmt(doanhSo)}</b></div>
      </div>
    </div>`;
  }

  function login(form) {
    var username = String((form && form.username) || '').trim();
    var password = String((form && form.password) || '').trim();

    if (!username || !password) return { ok: false, message: 'Vui lòng nhập đầy đủ tên đăng nhập và mật khẩu.' };

    var acc = getAccountByUsername_(username);
    if (!acc || acc.password !== password) return { ok: false, message: 'Sai tên đăng nhập hoặc mật khẩu.' };

    var token = generateToken_(acc.username, acc.displayName || acc.username);
    var url   = ScriptApp.getService().getUrl() + '?token=' + encodeURIComponent(token);

    return { ok: true, url: url, token: token, mustChangePassword: !!acc.mustChangePassword };
  }

  function doGet(e) {
    // 🚫 TẠM TẮT LOGIN - Luôn vào thẳng app với user mặc định
    var t = HtmlService.createTemplateFromFile('index');
    t.username           = 'senque'; // User mặc định
    t.displayName        = 'senque';
    t.token              = 'bypass_login'; // Token giả
    t.mustChangePassword = false;
    t.role               = 'admin'; // Role admin

    return t.evaluate()
      .setTitle('Phiếu nhập & Công nợ (Web)')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    /*
    // 🚫 ĐÃ TẮT - Code login cũ:
    var token = e && e.parameter && e.parameter.token;
    var user  = validateToken_(token);

    if (!user) {
      var tLogin = HtmlService.createTemplateFromFile('login');
      tLogin.webAppUrl = ScriptApp.getService().getUrl();
      return tLogin.evaluate()
        .setTitle('Báo cáo khách đoàn 2')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    var acc = getAccountByUsername_(user.username);
    var mustChange = acc && acc.mustChangePassword;

    var t = HtmlService.createTemplateFromFile('index');
    t.username           = user.username;
    t.displayName        = user.displayName;
    t.token              = token;
    t.mustChangePassword = mustChange ? true : false;
    t.role               = acc && acc.role ? acc.role : ''; // ✅ NEW

    return t.evaluate()
      .setTitle('Phiếu nhập & Công nợ (Web)')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    */
  }

  function updateUserPassword(username, oldPassword, newPassword) {
    username    = String(username || '').trim();
    oldPassword = String(oldPassword || '');
    newPassword = String(newPassword || '');

    if (!username || !oldPassword || !newPassword) return { ok: false, message: 'Thiếu thông tin đổi mật khẩu.' };

    var ss = getSpreadsheet_();
    var sh = ss.getSheetByName(SHEET_ACCOUNT);
    if (!sh) return { ok: false, message: 'Không tìm thấy sheet account.' };

    var values = sh.getDataRange().getValues();
    if (values.length < 2) return { ok: false, message: 'Sheet account đang trống.' };

    var header = values[0].map(function(h) { return String(h || '').toLowerCase().trim(); });

    var idxUser = header.indexOf('username');
    var idxPass = header.indexOf('password');
    var idxMust = header.indexOf('must_change_password');

    if (idxUser === -1 || idxPass === -1) return { ok: false, message: 'Sheet account thiếu cột username/password.' };

    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      var u = String(row[idxUser] || '').trim();
      if (u === username) {
        var currentPass = String(row[idxPass] || '');
        if (currentPass !== oldPassword) return { ok: false, message: 'Mật khẩu hiện tại không đúng.' };

        sh.getRange(i + 1, idxPass + 1).setValue(newPassword);

        if (idxMust !== -1) sh.getRange(i + 1, idxMust + 1).setValue(false);

        return { ok: true, message: 'Đổi mật khẩu thành công.' };
      }
    }

    return { ok: false, message: 'Không tìm thấy tài khoản.' };
  }


  /* =====================================================================================
    19) CÁC HÀM SỬA ĐƠN CŨ (GIỮ NGUYÊN)
    ===================================================================================== */
  function updateOrderByDebtRow(rowIndex, payload) {
    return { ok: false, error: "Chức năng sửa đơn đã bị vô hiệu hóa. Vui lòng xoá & lập lại." };
  }

  function unlockPaidDebtRow(rowIndex, newStatus) {
    return { ok: false, error: "Chức năng sửa đơn đã bị vô hiệu hóa." };
  }




  /* =====================================================================================
    20) UI: QUẢN LÝ TÀI KHOẢN (ADMIN)
    - Phục vụ overlay "Tài khoản" ở index.html
    ===================================================================================== */

  function normalizeAccountRole_(v) {
    var s = String(v || '').trim().toLowerCase();
    if (s === 'admin' || s === 'manager' || s === 'cashier') return s;
    return 'cashier';
  }

  function normalizeAccountStatus_(v) {
    var s = String(v || '').trim().toLowerCase();
    if (!s) return 'active';
    if (s === 'active' || s === 'on' || s === '1' || s === 'true' || s === 'yes' || s === 'enable' || s === 'enabled') return 'active';
    if (s === 'inactive' || s === 'off' || s === '0' || s === 'false' || s === 'no' || s === 'disable' || s === 'disabled') return 'inactive';
    if (s.indexOf('act') === 0) return 'active';      // ACTIVE
    if (s.indexOf('inact') === 0) return 'inactive';  // INACTIVE
    return 'active';
  }

  function ensureAccountHeader_(sh) {
    var lastCol = Math.max(6, sh.getLastColumn() || 6);
    var firstRow = sh.getRange(1, 1, 1, lastCol).getValues()[0] || [];
    var lower = firstRow.map(function(x){ return String(x || '').toLowerCase().trim(); });

    var hasUser = lower.indexOf('username') !== -1;
    var hasPass = lower.indexOf('password') !== -1;
    var hasDN   = (lower.indexOf('displayname') !== -1) || (lower.indexOf('display_name') !== -1) || (lower.indexOf('name') !== -1);
    var hasRole = lower.indexOf('role') !== -1;
    var hasSt   = lower.indexOf('status') !== -1;
    var hasMust = lower.indexOf('must_change_password') !== -1;

    var isEmptyRow = lower.slice(0,6).every(function(x){ return x === ''; });

    var headerWanted = ['username','password','displayName','role','status','must_change_password'];

    // Auto-fix trường hợp header có đủ cột nhưng đặt nhầm vị trí role/status (ví dụ: ... displayName,status,role ...)
    if (hasUser && hasPass && hasDN && hasRole && hasSt && hasMust) {
      var row2Fix = (sh.getLastRow() || 0) >= 2 ? (sh.getRange(2,1,1,Math.max(6,lastCol)).getValues()[0] || []) : [];
      var c4 = String(row2Fix[3] || '').trim().toLowerCase();
      var c5 = String(row2Fix[4] || '').trim().toLowerCase();
      function isStatusLike__(v){
        if (!v) return false;
        if (v === 'active' || v === 'inactive' || v === 'on' || v === 'off') return true;
        if (v === '1' || v === '0' || v === 'true' || v === 'false') return true;
        if (v === 'yes' || v === 'no' || v.indexOf('act') === 0 || v.indexOf('inact') === 0) return true;
        return false;
      }
      function isRoleLike__(v){
        return v === 'admin' || v === 'manager' || v === 'cashier' || v === 'user' || v === 'staff';
      }
      var idxRoleHdr = lower.indexOf('role');
      var idxStHdr = lower.indexOf('status');
      // Nếu header ghi status ở D nhưng data D lại là role -> đổi nhãn cột D/E cho đúng
      if (idxStHdr === 3 && idxRoleHdr === 4 && isRoleLike__(c4) && isStatusLike__(c5)) {
        sh.getRange(1,1,1,6).setValues([headerWanted]);
        return;
      }
      // Nếu header ghi role ở D nhưng data D lại là status -> đổi nhãn để phù hợp với data
      if (idxRoleHdr === 3 && idxStHdr === 4 && isStatusLike__(c4) && isRoleLike__(c5)) {
        sh.getRange(1,1,1,6).setValues([['username','password','displayName','status','role','must_change_password']]);
        return;
      }
    }

    // helper: detect status-like values
    function isStatusLike_(v) {
      v = String(v || '').trim().toLowerCase();
      if (!v) return false;
      return (v === '1' || v === '0' || v === 'active' || v === 'inactive' || v === 'on' || v === 'off' || v === 'true' || v === 'false' || v === 'yes' || v === 'no');
    }

    // Trường hợp sheet rỗng
    if (isEmptyRow && (sh.getLastRow() || 0) <= 1) {
      sh.getRange(1,1,1,6).setValues([headerWanted]);
      return;
    }

    // Nếu dòng 1 đang là data -> chèn header lên trên để KHÔNG ghi đè.
    if (!(hasUser && hasPass)) {
      sh.insertRows(1, 1);
      sh.getRange(1,1,1,6).setValues([headerWanted]);
      return;
    }

    // Nếu header bị "lệch tên cột" kiểu: username,password,status,role,must_change_password,(blank)
    // nhưng data thực tế đang theo thứ tự: username,password,displayName,role,status,must_change_password
    // => chỉ cần sửa lại header (KHÔNG đụng data).
    if (!hasDN && hasRole && hasSt) {
      var row2 = (sh.getLastRow() || 0) >= 2 ? (sh.getRange(2,1,1,6).getValues()[0] || []) : [];
      var u  = String(row2[0] || '').trim();
      var c3 = String(row2[2] || '').trim(); // đang là displayName
      var c5 = String(row2[4] || '').trim(); // đang là status (active/inactive/...)
      // Heuristic an toàn: cột 3 giống username (displayName mặc định) HOẶC cột 5 là status-like
      if ((u && c3 === u) || isStatusLike_(c5)) {
        sh.getRange(1,1,1,6).setValues([headerWanted]);
        return;
      }
    }

    // Nếu thiếu cột must_change_password ở cột F (thường bị "Unnamed")
    if (!hasMust) {
      // chỉ sửa tên nếu đủ 6 cột
      if (lastCol >= 6) {
        var r = sh.getRange(1,1,1,6).getValues()[0];
        // nếu cột F đang trống/unnamed => set lại
        var f = String(r[5] || '').toLowerCase().trim();
        if (!f || f.indexOf('unnamed') === 0) {
          r[5] = 'must_change_password';
          // nếu có displayName nhưng tên không đúng case thì cũng normalize nhẹ
          if (String(r[2] || '').toLowerCase().trim() === 'status' && String(r[4] || '').toLowerCase().trim() === 'must_change_password') {
            // trường hợp cực đoan: header đúng 5 cột, lệch như trên
            r = headerWanted;
          }
          sh.getRange(1,1,1,6).setValues([r]);
        }
      }
    }
  }

  function ui_listAccounts(requesterUsername) {
    try {
      var req = getAccountByUsername_(requesterUsername);
      if (!req || String(req.role || '').toLowerCase() !== 'admin') throw new Error('Không có quyền admin.');

      var ss = getSpreadsheet_();
      var sh = ss.getSheetByName(SHEET_ACCOUNT);
      if (!sh) throw new Error('Không tìm thấy sheet account.');

      ensureAccountHeader_(sh);

      var values = sh.getDataRange().getValues();
      if (!values || values.length < 2) return { ok: true, rows: [] };

      var header = values[0].map(function(h){ return String(h||'').toLowerCase().trim(); });

      var idxUser = header.indexOf('username');
      var idxPass = header.indexOf('password');
      var idxDN   = header.indexOf('displayname');
      if (idxDN === -1) idxDN = header.indexOf('display_name');
      var idxRole = header.indexOf('role');
      var idxSt   = header.indexOf('status');
      var idxMust = header.indexOf('must_change_password');

      // fallback nếu không có role: cột D (index 3)
      if (idxRole === -1 && header.length >= 4) idxRole = 3;

      if (idxUser === -1 || idxPass === -1) throw new Error('Sheet account thiếu header username/password');

      var rows = [];
      for (var i = 1; i < values.length; i++) {
        var r = values[i];
        var u = String(r[idxUser] || '').trim();
        if (!u) continue;

        var dn = String((idxDN !== -1 ? r[idxDN] : '') || u);
        var rawRole = (idxRole !== -1 ? String(r[idxRole] || '') : '');
        var rawSt = (idxSt !== -1 ? String(r[idxSt] || '') : '');
        function isStatusLike__(v){
          v = String(v||'').trim().toLowerCase();
          if (!v) return false;
          if (v === 'active' || v === 'inactive' || v === 'on' || v === 'off') return true;
          if (v === '1' || v === '0' || v === 'true' || v === 'false') return true;
          if (v === 'yes' || v === 'no' || v === 'enable' || v === 'enabled' || v === 'disable' || v === 'disabled') return true;
          if (v.indexOf('act') === 0) return true;
          if (v.indexOf('inact') === 0) return true;
          return false;
        }
        function isRoleLike__(v){
          v = String(v||'').trim().toLowerCase();
          return v === 'admin' || v === 'manager' || v === 'cashier' || v === 'user' || v === 'staff';
        }
        if (isStatusLike__(rawRole) && isRoleLike__(rawSt)) {
          var tmp = rawSt; rawSt = rawRole; rawRole = tmp;
        }
        var role = normalizeAccountRole_(rawRole || '');
        var stOut = normalizeAccountStatus_(rawSt || '');

        var must = false;
        if (idxMust !== -1) must = parseBool_(r[idxMust]);

        rows.push({
          username: u,
          displayName: dn,
          role: role || 'cashier',
          status: stOut,
          mustChangePassword: must
        });
      }

      return { ok: true, rows: rows };
    } catch (e) {
      return { ok: false, error: e && e.message ? e.message : String(e) };
    }
  }

  function ui_upsertAccount(requesterUsername, payload) {
    var lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      var req = getAccountByUsername_(requesterUsername);
      if (!req || String(req.role || '').toLowerCase() !== 'admin') throw new Error('Không có quyền admin.');

      payload = payload || {};
      var username = String(payload.username || '').trim();
      var displayName = String(payload.displayName || '').trim();
      var password = String(payload.password || '');
      var roleProvided = Object.prototype.hasOwnProperty.call(payload, 'role');
      var statusProvided = Object.prototype.hasOwnProperty.call(payload, 'status');

      var role = roleProvided ? normalizeAccountRole_(payload.role) : '';
      var status = statusProvided ? normalizeAccountStatus_(payload.status) : '';
      var mustChange = parseBool_(payload.mustChangePassword);

      if (!username) throw new Error('Thiếu username.');
      if (!/^[A-Za-z0-9._-]{3,50}$/.test(username)) throw new Error('Username không hợp lệ (3-50 ký tự, chỉ chữ/số/._-).');

      // role/status đã được normalize (active/inactive)

      var ss = getSpreadsheet_();
      var sh = ss.getSheetByName(SHEET_ACCOUNT);
      if (!sh) throw new Error('Không tìm thấy sheet account.');

      ensureAccountHeader_(sh);

      var values = sh.getDataRange().getValues();
      var header = (values[0] || []).map(function(h){ return String(h||'').toLowerCase().trim(); });
      var idxUser = header.indexOf('username');
      var idxPass = header.indexOf('password');
      var idxDN   = header.indexOf('displayname');
      if (idxDN === -1) idxDN = header.indexOf('display_name');
      var idxRole = header.indexOf('role');
      var idxSt   = header.indexOf('status');
      var idxMust = header.indexOf('must_change_password');

      if (idxDN === -1) idxDN = 2;   // col C
      if (idxRole === -1) idxRole = 3; // col D
      if (idxSt === -1) idxSt = 4;   // col E
      if (idxMust === -1) idxMust = 5; // col F

      // tìm row theo username
      var foundRow = -1;
      for (var i = 1; i < values.length; i++) {
        var u = String(values[i][idxUser] || '').trim();
        if (u === username) { foundRow = i + 1; break; }
      }

      if (foundRow === -1) {
        // create
        if (!password) throw new Error('Tạo mới cần password.');
        var row = [];
        var width = Math.max(6, (values[0] || []).length || 6);
        for (var k = 0; k < width; k++) row[k] = '';

        row[idxUser] = username;
        row[idxPass] = password;
        row[idxDN]   = displayName || username;
        row[idxRole] = role || 'cashier';
        row[idxSt]   = status || 'active';
        row[idxMust] = mustChange ? true : false;

        sh.appendRow(row);
      } else {
        // update
        if (displayName) sh.getRange(foundRow, idxDN + 1).setValue(displayName);
        if (roleProvided) sh.getRange(foundRow, idxRole + 1).setValue(role || 'cashier');
        if (statusProvided) sh.getRange(foundRow, idxSt + 1).setValue(status || 'active');
        sh.getRange(foundRow, idxMust + 1).setValue(mustChange ? true : false);

        if (password) sh.getRange(foundRow, idxPass + 1).setValue(password);
      }

      try { CacheService.getScriptCache().remove("ACC_V1:" + username.toLowerCase()); } catch(e) {}
      return { ok: true };
    } catch (e) {
      return { ok: false, error: e && e.message ? e.message : String(e) };
    } finally {
      try { lock.releaseLock(); } catch (e2) {}
    }
  }

  function ui_bulkUpsertAccounts(requesterUsername, rows) {
    rows = rows || [];
    if (!Array.isArray(rows)) return { ok: false, error: 'rows phải là mảng.' };

    var results = { ok: true, inserted: 0, updated: 0, errors: [] };

    for (var i = 0; i < rows.length; i++) {
      var r = rows[i] || {};
      var res = ui_upsertAccount(requesterUsername, r);
      if (!res || res.ok === false) {
        results.ok = false;
        results.errors.push({ index: i, username: r.username || '', error: (res && res.error) ? res.error : 'Unknown' });
        continue;
      }
      // không phân biệt inserted/updated chính xác vì ui_upsertAccount không trả về loại.
      results.updated++;
    }

    return results;
  }



  /* =====================================================================================
    DASHBOARD (THỐNG KÊ)
    - dishReport: nhân SL + Thành tiền theo Số mâm (join theo Mã đơn: Debt cột M, Data cột K)
    ===================================================================================== */
  /**
   * Trả về HTML fragment của Dashboard (file dashboard_html.html)
   */
  function getDashboardHtml() {
    return HtmlService
      .createHtmlOutputFromFile('dashboard_html')
      .getContent();
  }

  /** ===== Helper parse date / range ===== */
  function _dashToDate_(v) {
    if (!v) return null;
    if (Object.prototype.toString.call(v) === '[object Date]') {
      return new Date(v.getFullYear(), v.getMonth(), v.getDate());
    }

    var s = String(v).trim();
    if (!s) return null;

    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
      var p = s.split('-').map(function (x) { return Number(x); });
      return new Date(p[0], p[1] - 1, p[2]);
    }

    var parts = s.split(/[\/\-]/);
    if (parts.length >= 3) {
      var d = Number(parts[0]);
      var m = Number(parts[1]);
      var yRaw = parts[2];
      var y = Number(yRaw.length === 2 ? ('20' + yRaw) : yRaw);
      if (y && m && d) return new Date(y, m - 1, d);
    }

    var dt = new Date(s);
    if (isNaN(dt.getTime())) return null;
    return new Date(dt.getFullYear(), dt.getMonth(), dt.getDate());
  }

  function _dashInRange_(dt, fromDt, toDt) {
    if (!dt) return false;
    if (fromDt && dt < fromDt) return false;
    if (toDt && dt > toDt) return false;
    return true;
  }

  /**
   * Lấy dữ liệu Dashboard theo filterStr:
   *  - ""            : default từ đầu tháng -> hôm nay
   *  - "ALL"         : tất cả dữ liệu
   *  - "FROM|TO"     : lọc theo khoảng ngày
   */
  function getDashboardDataForMonth(filterStr) {
    try {
      var ss     = getSpreadsheet_();
      var shDeb  = ss.getSheetByName(SHEET_DEBT);
      var shData = ss.getSheetByName(SHEET_DATA);
      var tz     = Session.getScriptTimeZone();

      if (!shDeb) throw new Error('Không tìm thấy sheet Công nợ: ' + SHEET_DEBT);

      var raw = (filterStr || '').toString().trim();
      var isAll = false;

      var now = new Date();
      now = new Date(now.getFullYear(), now.getMonth(), now.getDate());
      var defFrom = new Date(now.getFullYear(), now.getMonth(), 1);
      var defTo = now;

      var fromDt = null, toDt = null;

      if (!raw) {
        fromDt = defFrom;
        toDt = defTo;
      } else if (raw.toUpperCase() === 'ALL') {
        isAll = true;
        fromDt = null;
        toDt = null;
      } else if (raw.indexOf('|') >= 0) {
        var ps = raw.split('|');
        var fromStr = (ps[0] || '').trim();
        var toStr   = (ps[1] || '').trim();
        fromDt = _dashToDate_(fromStr);
        toDt   = _dashToDate_(toStr);

        if (!fromDt && !toDt) {
          fromDt = defFrom;
          toDt = defTo;
        } else if (fromDt && !toDt) {
          toDt = fromDt;
        } else if (!fromDt && toDt) {
          fromDt = toDt;
        }
      } else {
        isAll = true;
        fromDt = null;
        toDt = null;
      }

      if (fromDt && toDt && fromDt > toDt) {
        var tmp = fromDt; fromDt = toDt; toDt = tmp;
      }

      var labelStr = '';
      if (isAll) {
        labelStr = 'Tất cả dữ liệu';
      } else {
        var fStr = Utilities.formatDate(fromDt, tz, 'dd/MM/yyyy');
        var tStr = Utilities.formatDate(toDt, tz, 'dd/MM/yyyy');
        labelStr = fStr + ' – ' + tStr;
      }

      var lastRowDeb = shDeb.getLastRow();
      if (lastRowDeb < DEBT_DATA_START) {
        return {
          ok: true,
          isAll: isAll,
          label: labelStr,
          month: labelStr,
          kpi: { tongThanhTien: 0, tongKm: 0, tongDoanhSo: 0, tongCongNo: 0, prevTongThanhTien: 0, prevTongKm: 0, prevTongDoanhSo: 0, prevTongCongNo: 0 },
          topCustomers: [],
          statusPie: [],
          orders: [],
          dishReport: []
        };
      }

      var debVals = shDeb.getRange(
        DEBT_DATA_START, 1,
        lastRowDeb - DEBT_DATA_START + 1,
        DEBT_LAST_COL
      ).getValues();

      var curKpi = { tongThanhTien: 0, tongKm: 0, tongDoanhSo: 0, tongCongNo: 0 };
      var orders = [];
      var statusAgg = {};
      var customerAgg = {};

      // Join map: Debt(Mã đơn -> Số mâm). Fallback legacy theo (ngày|sđt)
      var soMamByOrderId = {};
      var soMamByKey = {};

      debVals.forEach(function (row) {
        var dateVal   = row[DEBT_COL_DATE - 1];
        var info      = String(row[DEBT_COL_INFO - 1] || '');
        var soMam     = Number(row[DEBT_COL_SOMAM - 1] || 0);
        var donGiaMam = Number(row[DEBT_COL_DONGIA_MAM - 1] || 0);
        var tongDon   = Number(row[DEBT_COL_TONG_DON - 1] || 0);
        var kmNote    = String(row[DEBT_COL_KM_NOTE - 1] || '');
        var kmAmount  = Number(row[DEBT_COL_KM_AMOUNT - 1] || 0);
        var doanhSo   = Number(row[DEBT_COL_DOANHSO - 1] || 0);
        var congNo    = Number(row[DEBT_COL_CONGNO - 1] || 0);
        var status    = String(row[DEBT_COL_STATUS - 1] || '');
        var orderId   = String(row[DEBT_COL_ORDER_ID - 1] || '').trim();

        if (!dateVal) return;

        var d = dateVal instanceof Date
          ? new Date(dateVal.getFullYear(), dateVal.getMonth(), dateVal.getDate())
          : _dashToDate_(dateVal);
        if (!d) return;

        if (!isAll && !_dashInRange_(d, fromDt, toDt)) return;

        if (!soMam || soMam <= 0) soMam = 1;

        if (orderId) soMamByOrderId[orderId] = soMam;

        // fallback theo key (ngày|sđt)
        var parts = info.split(' - ');
        var phone = (parts[1] || '').trim();
        var key = Utilities.formatDate(d, tz, 'yyyy-MM-dd') + '|' + normalizePhone_(phone);
        if (key) soMamByKey[key] = soMam;

        if (!doanhSo && (tongDon || kmAmount)) doanhSo = tongDon - kmAmount;

        var paidPart = doanhSo - congNo;
        if (paidPart < 0) paidPart = 0;
        var unpaidPart = congNo;

        curKpi.tongThanhTien += tongDon;
        curKpi.tongKm        += kmAmount;
        curKpi.tongDoanhSo   += doanhSo;
        curKpi.tongCongNo    += congNo;

        var dateStr = Utilities.formatDate(d, tz, 'dd/MM/yyyy');
        orders.push({
          date: dateStr,
          info: info,
          soMam: soMam,
          donGiaMam: donGiaMam,
          tongDon: tongDon,
          kmNote: kmNote,
          kmAmount: kmAmount,
          doanhSo: doanhSo,
          congNo: congNo,
          status: status,
          orderId: orderId
        });

        var stKey = status || 'Khác';
        statusAgg[stKey] = (statusAgg[stKey] || 0) + doanhSo;

        var tenKH = (parts[0] || '').trim();
        var cKey  = normalizePhone_(phone) || tenKH || 'N/A';

        if (!customerAgg[cKey]) customerAgg[cKey] = { tenKH: tenKH, phone: phone, paid: 0, unpaid: 0, doanhSo: 0 };
        customerAgg[cKey].paid    += paidPart;
        customerAgg[cKey].unpaid  += unpaidPart;
        customerAgg[cKey].doanhSo += doanhSo;
      });

      var topCustomers = Object.keys(customerAgg)
        .map(function (k) { return customerAgg[k]; })
        .sort(function (a, b) { return (b.doanhSo || 0) - (a.doanhSo || 0); })
        .slice(0, 5);

      var statusPie = Object.keys(statusAgg).map(function (st) {
        return { status: st, value: statusAgg[st] };
      });

      // ===== DISH REPORT: nhân theo số mâm, join theo Mã đơn (Debt M, Data K) =====
      var dishReport = [];
      if (shData && shData.getLastRow() >= 2) {
        var m = getDataSheetMap_(shData);
        var lastRowData = shData.getLastRow();
        var lastColData = shData.getLastColumn();
        var width = Math.max(lastColData, m.idxOid, m.idxNgay, m.idxSdt, m.idxTen, m.idxSl, m.idxTt, m.idxDg);

        var dataVals = shData.getRange(2, 1, lastRowData - 1, width).getValues();

        var dishMap = {};
        dataVals.forEach(function (r) {
          var oid = String(r[m.idxOid - 1] || '').trim();
          var tenMon = (r[m.idxTen - 1] || '').toString().trim();
          var slBase = Number(r[m.idxSl - 1] || 0);
          var dg = Number(r[m.idxDg - 1] || 0);
          var amtBase = Number(r[m.idxTt - 1] || 0);
          var dCell = r[m.idxNgay - 1];

          if (!tenMon || slBase <= 0) return;

          // xác định soMam theo id (ưu tiên) -> fallback legacy theo (ngày|sđt)
          var soMam = 0;
          if (oid && soMamByOrderId[oid]) soMam = Number(soMamByOrderId[oid] || 0) || 0;

          if (!soMam) {
            var dt = dCell instanceof Date ? new Date(dCell.getFullYear(), dCell.getMonth(), dCell.getDate()) : _dashToDate_(dCell);
            if (!dt) return;

            if (!isAll && !_dashInRange_(dt, fromDt, toDt)) return;

            var phone = String(r[m.idxSdt - 1] || "").trim();
            var key = Utilities.formatDate(dt, tz, 'yyyy-MM-dd') + '|' + normalizePhone_(phone);
            soMam = Number(soMamByKey[key] || 0) || 0;
          }

          if (!soMam || soMam <= 0) soMam = 1;

          var qty = slBase * soMam;
          var amt = amtBase * soMam;

          if (!dishMap[tenMon]) dishMap[tenMon] = { tenMon: tenMon, tongSoLuong: 0, thanhTien: 0, donGiaSample: 0 };
          dishMap[tenMon].tongSoLuong += qty;
          dishMap[tenMon].thanhTien   += amt;
          if (!dishMap[tenMon].donGiaSample && dg) dishMap[tenMon].donGiaSample = dg;
        });

        Object.keys(dishMap).forEach(function (name) {
          var it = dishMap[name];
          var qty = it.tongSoLuong || 0;
          var avgPrice = 0;
          if (qty > 0) avgPrice = it.thanhTien > 0 ? (it.thanhTien / qty) : (it.donGiaSample || 0);
          dishReport.push({ tenMon: it.tenMon, tongSoLuong: qty, thanhTien: it.thanhTien, donGiaBinhQuan: avgPrice });
        });

        dishReport.sort(function (a, b) {
          var qA = Number(a.tongSoLuong || 0), qB = Number(b.tongSoLuong || 0);
          var amtA = Number(a.thanhTien || 0), amtB = Number(b.thanhTien || 0);
          var pA = Number(a.donGiaBinhQuan || 0), pB = Number(b.donGiaBinhQuan || 0);
          if (qA !== qB) return qB - qA;
          if (amtA !== amtB) return amtB - amtA;
          return pB - pA;
        });
      }

      return {
        ok: true,
        isAll: isAll,
        label: labelStr,
        month: labelStr,
        kpi: {
          tongThanhTien: curKpi.tongThanhTien,
          tongKm: curKpi.tongKm,
          tongDoanhSo: curKpi.tongDoanhSo,
          tongCongNo: curKpi.tongCongNo,
          prevTongThanhTien: 0, prevTongKm: 0, prevTongDoanhSo: 0, prevTongCongNo: 0
        },
        topCustomers: topCustomers,
        statusPie: statusPie,
        orders: orders,
        dishReport: dishReport
      };

    } catch (e) {
      return { ok: false, error: e && e.message ? e.message : String(e) };
    }
  }

function resetPendingSheetCache_() {
  try { PropertiesService.getScriptProperties().deleteProperty("PENDING_SHEET_NAME"); } catch(e) {}
  try { CacheService.getScriptCache().remove("PENDING_LIST_V1"); } catch(e) {}
}


function getOrderDetailByPendingRowLegacy_(rowIndex, includeTax) {
  includeTax = (typeof includeTax === "undefined" || includeTax === null) ? true : !!includeTax;

  const ss = getSpreadsheet_();
  const shP = ss.getSheetByName(SHEET_PENDING);
  if (!shP) return { ok: false, message: "Không tìm thấy sheet '" + SHEET_PENDING + "'." };

  const lastRow = shP.getLastRow();
  if (!rowIndex || rowIndex < 2 || rowIndex > lastRow) return { ok: false, message: "Dòng đơn chờ không hợp lệ." };

  const off = getLeadingIndexOffset_(shP);
  const baseCol = 1 + off;
  const r = shP.getRange(rowIndex, baseCol, 1, PENDING_LAST_COL).getValues()[0] || [];

  const ngay = r[PENDING_COL_DATE - 1];
  const info = String(r[PENDING_COL_INFO - 1] || "").trim();
  const soMam = toMoneyNumber_(r[PENDING_COL_SOMAM - 1] || 0);
  const donGiaMam = toMoneyNumber_(r[PENDING_COL_DONGIA_MAM - 1] || 0);
  const tongDon = toMoneyNumber_(r[PENDING_COL_TONG_DON - 1] || 0);
  const kmNoiDung = String(r[PENDING_COL_KM_NOIDUNG - 1] || "").trim();
  const kmSoTien = toMoneyNumber_(r[PENDING_COL_KM_SOTIEN - 1] || 0);
  const thungan = String(r[PENDING_COL_CASHIER - 1] || "").trim();
  const orderId = String(r[PENDING_COL_ORDER_ID - 1] || "").trim();
  const datCoc = toMoneyNumber_(r[PENDING_COL_DEPOSIT - 1] || 0);
  if (!orderId) return { ok: false, message: "Đơn chờ thiếu Mã đơn (orderId)." };

  const shData = ss.getSheetByName(SHEET_DATA);
  if (!shData) return { ok: false, message: "Không tìm thấy sheet '" + SHEET_DATA + "'." };

  const items = getItemsByOrderId_(shData, orderId);
  if (!items || !items.length) return { ok: false, message: "Không lấy được dữ liệu in (không tìm thấy món theo mã đơn)." };

  const sdt = normalizePhone_(extractPhoneFromInfo_(info));
  const tenKH = getNameFromInfo_(info) || "";
  const diaChi = getAddressFromInfo_(info) || "";

  const header = {
    dateRaw: ngay,
    date: ngay instanceof Date ? Utilities.formatDate(ngay, Session.getScriptTimeZone(), "dd/MM/yyyy") : String(ngay || ""),
    tenKH: tenKH,
    sdt: sdt,
    diaChi: diaChi,
    soMam: soMam,
    donGiaMam: donGiaMam,
    tongDon: tongDon,
    kmNoiDung: kmNoiDung,
    kmSoTien: kmSoTien,
    datCoc: datCoc,
    thungan: thungan,
    trangThai: STATUS_DEBT,
    orderId: orderId
  };

  const invoiceNo = reserveInvoiceNo_("P:" + rowIndex);
  const html = buildInvoiceHtml(header, items, invoiceNo, includeTax);

  return { ok: true, html: html, invoiceNo: invoiceNo };
}

