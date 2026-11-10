/*********************************************************
 * lil bulle Inventory System V10.39 — CLEAN PRODUCTION
 * Modules: Stock / Sales / Return / Audit / Stock Count / Approval / Event Booth
 * SYSTEM_VERSION is the single source of truth for the current production version.
 *********************************************************/

var props = PropertiesService.getScriptProperties();

var ACCESS_TOKEN = props.getProperty("ACCESS_TOKEN");
var GROUP_ID     = props.getProperty("GROUP_ID");
var SS_ID        = props.getProperty("SS_ID");
var WEB_APP_URL  = props.getProperty("WEB_APP_URL");

var SYSTEM_VERSION = "V11.2-LOT-HOTFIX";

var ITEMS_MAP = {
  "บาม":   { kw: "balm",   full: "lil bulle baby rich balm" },
  "โลชั่น": { kw: "lotion", full: "lil bulle organic lotion" },
  "สเปรย์": { kw: "spray",  full: "lil bulle sanitizer spray" },
  "โฟม":   { kw: "foam",   full: "lil bulle baby foam" },
  "ยุง":    { kw: "mozzie", full: "lil bulle MOZZIE GUARD" }
};

var CHANNELS = ["LINE", "Facebook", "Shopee", "TikTok", "Event", "โรงพยาบาล", "ฝากขาย", "แจกฟรี/รีวิว (PR)", "อื่น ๆ"];

/*********************************************************
 * Web App
 *********************************************************/


/*********************************************************
 * LINE Webhook
 *********************************************************/
function doPost(e) {
  if (e && e.postData) {
    var probe; try { probe = JSON.parse(e.postData.contents); } catch(ignore) {}
    if (probe && probe.posApi === "lilbulle-v11") return lbApi_(probe);
  }
  try {
    if (!e || !e.postData || !e.postData.contents) return ContentService.createTextOutput("OK");

    var data = JSON.parse(e.postData.contents);
    if (!data.events || data.events.length === 0) return ContentService.createTextOutput("OK");

    var event = data.events[0];
    var msgText = "";
    if (event.message && event.message.text) msgText = String(event.message.text).trim();

    if (msgText === "ลงบันทึก") {
      sendPush("📝 ลิงก์ลงบันทึก lil bulle\n" + WEB_APP_URL);
    } else if (msgText === "เช็คสต๊อก" || msgText === "เช็กสต๊อก") {
      sendPush(getCurrentStockSummary());
    }

    return ContentService.createTextOutput("OK");
  } catch (err) {
    Logger.log("doPost ERROR: " + err);
    return ContentService.createTextOutput("OK");
  }
}

/*********************************************************
 * Utilities
 *********************************************************/
function normalizeData(val) {
  if (val == null) return "";
  return String(val).trim().toLowerCase();
}


function safeDateTimeString(val) {
  if (!val) return "-";
  if (val instanceof Date) {
    return Utilities.formatDate(val, "GMT+7", "dd/MM/yyyy HH:mm:ss");
  }
  return String(val).trim();
}

function nowDateTime() {
  return Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss");
}

function parseChannel(str) {
  var raw = String(str || "");
  var first = normalizeData(raw.split("|")[0]);
  var exact = {
    "line": "LINE",
    "facebook": "Facebook",
    "shopee": "Shopee",
    "tiktok": "TikTok",
    "event": "Event",
    "หน้าร้าน (event)": "Event",
    "โรงพยาบาล": "โรงพยาบาล",
    "ฝากขาย": "ฝากขาย",
    "แจกฟรี/รีวิว (pr)": "แจกฟรี/รีวิว (PR)",
    "อื่น ๆ": "อื่น ๆ"
  };
  if (exact[first]) return exact[first];

  var text = normalizeData(raw);
  if (text.indexOf("line") > -1) return "LINE";
  if (text.indexOf("facebook") > -1 || text.indexOf("fb") > -1) return "Facebook";
  if (text.indexOf("shopee") > -1) return "Shopee";
  if (text.indexOf("tiktok") > -1) return "TikTok";
  if (text.indexOf("แจกฟรี") > -1 || text.indexOf("pr") > -1 || text.indexOf("สนับสนุน") > -1 || text.indexOf("เบิกฟรี") > -1) return "แจกฟรี/รีวิว (PR)";
  if (text.indexOf("event") > -1 || text.indexOf("หน้าร้าน") > -1 || text.indexOf("บูท") > -1) return "Event";
  if (text.indexOf("โรงพยาบาล") > -1 || text.indexOf("รพ.") > -1 || text.indexOf("hospital") > -1) return "โรงพยาบาล";
  if (text.indexOf("ฝากขาย") > -1 || text.indexOf("consign") > -1) return "ฝากขาย";
  return "อื่น ๆ";
}

function parseLocation_(source) {
  var s = String(source || "");
  var m = s.match(/(?:สถานที่|จุดขาย)\s*:\s*([^|]+)/i);
  if (m && String(m[1] || "").trim()) return String(m[1]).trim();
  return "-";
}

function buildSourceText_(channel, location, note) {
  var parts = [String(channel || "อื่น ๆ").trim() || "อื่น ๆ"];
  location = String(location || "").trim();
  note = String(note || "").trim();
  if (location) parts.push("สถานที่:" + location);
  if (note) parts.push("หมายเหตุ:" + note);
  return parts.join(" | ");
}


function ensureHeaderColumn_(sheet, headerName) {
  var lastCol = Math.max(1, sheet.getLastColumn());
  var headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i] || "").trim() === headerName) return i + 1;
  }
  var col = lastCol + 1;
  sheet.getRange(1, col).setValue(headerName).setFontWeight("bold");
  return col;
}

function setDateTimeColumnFormat_(sheet, col) {
  if (!sheet || !col) return;
  var rows = Math.max(1, sheet.getMaxRows() - 1);
  sheet.getRange(2, col, rows, 1).setNumberFormat("dd/MM/yyyy HH:mm:ss");
}

function applyAllDateFormats_() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var specs = [
    ["SalesLog", [2]],
    ["StockIn", [2]],
    ["AuditLog", [1]],
    ["StockCount", [1]],
    ["StockAdjust", [1]],
    ["StockAdjustRequest", [1, 12]],
    ["EventIssue", [1]],
    ["EventReturn", [1]]
  ];
  for (var i = 0; i < specs.length; i++) {
    var sh = ss.getSheetByName(specs[i][0]);
    if (!sh) continue;
    for (var j = 0; j < specs[i][1].length; j++) setDateTimeColumnFormat_(sh, specs[i][1][j]);
  }
  ss.setSpreadsheetTimeZone("Asia/Bangkok");
  return true;
}

function getWarehouseDocNumbersToday_() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var day = Utilities.formatDate(new Date(), "GMT+7", "yyyyMMdd");
  var head = "STK-" + day + "-";
  var maxNo = 0;
  ["SalesLog", "StockIn"].forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) return;
    var vals = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues();
    for (var i = 0; i < vals.length; i++) {
      var v = String(vals[i][0] || "").trim();
      if (v.indexOf(head) !== 0) continue;
      var n = Number(v.substring(head.length));
      if (!isNaN(n) && n > maxNo) maxNo = n;
    }
  });
  return { head: head, next: maxNo + 1 };
}

function makeWarehouseDocNo_() {
  var r = getWarehouseDocNumbersToday_();
  var seq = String(r.next);
  while (seq.length < 3) seq = "0" + seq;
  return r.head + seq;
}

function getNextWarehouseDocumentNo() {
  try {
    return { success: true, docNo: makeWarehouseDocNo_(), version: SYSTEM_VERSION };
  } catch (err) {
    return { success: false, docNo: "ระบบจะออกเลขให้เมื่อบันทึก", msg: err.message };
  }
}

function ensureV1035Setup_() {
  var p = PropertiesService.getScriptProperties();
  if (p.getProperty("LILBULLE_V1035_SETUP") === "DONE") return;
  ensureControlSheets_();
  ensureAdjustRequestSheet_();
  ensureEventSheets_();
  normalizeAllExpirySheets_();
  applyAllDateFormats_();
  p.setProperty("LILBULLE_V1035_SETUP", "DONE");
}

function setupV1035() {
  ensureControlSheets_();
  ensureAdjustRequestSheet_();
  ensureEventSheets_();
  var expiry = normalizeAllExpirySheets_();
  applyAllDateFormats_();
  PropertiesService.getScriptProperties().setProperty("LILBULLE_V1035_SETUP", "DONE");
  return {
    success: true,
    version: SYSTEM_VERSION,
    msg: "ตั้งค่า V10.39 สำเร็จ — วันที่/เวลาเป็น วว/ดด/ค.ศ. และ Exp เป็น dd/MM/yyyy แล้ว",
    expiry: expiry
  };
}

function getSheetOrThrow(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) throw new Error("ไม่พบชีต " + name);
  return sh;
}

function getProductInfo(name) {
  var key = String(name || "").trim();
  if (ITEMS_MAP[key]) {
    return {
      short: key,
      kw: ITEMS_MAP[key].kw,
      full: ITEMS_MAP[key].full
    };
  }

  var norm = normalizeData(key);
  for (var shortName in ITEMS_MAP) {
    if (!ITEMS_MAP.hasOwnProperty(shortName)) continue;
    var item = ITEMS_MAP[shortName];
    if (
      normalizeData(item.full) === norm ||
      normalizeData(item.kw) === norm
    ) {
      return {
        short: shortName,
        kw: item.kw,
        full: item.full
      };
    }
  }

  throw new Error("ไม่พบรหัสสินค้า: " + key);
}

function toNumber(val) {
  var n = Number(val);
  if (isNaN(n)) return 0;
  return n;
}

/*********************************************************
 * LINE Messaging API
 *********************************************************/
function sendPush(message) {
  try {
    if (!ACCESS_TOKEN) {
      Logger.log("LINE ERROR: ACCESS_TOKEN ว่าง");
      return;
    }
    if (!GROUP_ID) {
      Logger.log("LINE ERROR: GROUP_ID ว่าง");
      return;
    }

    var url = "https://api.line.me/v2/bot/message/push";

    var payload = {
      to: GROUP_ID,
      messages: [
        {
          type: "text",
          text: String(message).substring(0, 4900)
        }
      ]
    };

    var options = {
      method: "post",
      headers: {
        Authorization: "Bearer " + ACCESS_TOKEN
      },
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    var response = UrlFetchApp.fetch(url, options);
    Logger.log("LINE RESPONSE: " + response.getResponseCode() + " / " + response.getContentText());
  } catch (err) {
    Logger.log("sendPush ERROR: " + err);
  }
}

function buildNotifyMessage(typeKey, docNo, staff, list) {
  var titleMap = {
    SALE: "🛒 ขายออก",
    FREE: "🎁 สนับสนุน / แจกฟรี / PR",
    INBOUND: "📦 รับเข้าคลัง",
    RETURN: "↩️ รับคืน / ลงผิด",
    EXPIRED: "⚠️ ตัดหมดอายุ"
  };

  var msg = "";
  msg += titleMap[typeKey] || "✅ บันทึกข้อมูล";
  msg += "\nเลขที่เอกสาร: " + docNo;
  msg += "\nผู้บันทึก: " + staff;
  msg += "\nเวลา: " + nowDateTime();
  msg += "\n\n";

  for (var i = 0; i < list.length; i++) {
    var it = list[i];
    msg += "• " + it.product + " จำนวน " + it.qty + " ชิ้น";
    if (it.exp && it.exp !== "-") msg += " | Exp: " + it.exp;
    if (it.channel) msg += " | " + it.channel;
    if (it.note) msg += " (" + it.note + ")";
    msg += "\n";
  }

  return msg;
}

function notifyByType(docNo, staff, itemsList) {
  var groups = {};

  for (var i = 0; i < itemsList.length; i++) {
    var t = itemsList[i].type || "SALE";
    if (!groups[t]) groups[t] = [];
    groups[t].push(itemsList[i]);
  }

  var order = ["SALE", "FREE", "INBOUND", "RETURN", "EXPIRED"];
  for (var j = 0; j < order.length; j++) {
    var key = order[j];
    if (groups[key] && groups[key].length > 0) {
      sendPush(buildNotifyMessage(key, docNo, staff, groups[key]));
      Utilities.sleep(300);
    }
  }
}

/*********************************************************
 * Duplicate Document Check
 *********************************************************/
function isDuplicateDoc(docNo) {
  var ss = SpreadsheetApp.openById(SS_ID);
  var salesSheet = ss.getSheetByName("SalesLog");
  var inSheet = ss.getSheetByName("StockIn");

  docNo = String(docNo || "").trim();
  if (!docNo) return false;

  if (salesSheet && salesSheet.getLastRow() >= 2) {
    var sValues = salesSheet.getRange(2, 1, salesSheet.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < sValues.length; i++) {
      if (String(sValues[i][0]).trim() === docNo) return true;
    }
  }

  if (inSheet && inSheet.getLastRow() >= 2) {
    var iValues = inSheet.getRange(2, 1, inSheet.getLastRow() - 1, 1).getValues();
    for (var j = 0; j < iValues.length; j++) {
      if (String(iValues[j][0]).trim() === docNo) return true;
    }
  }

  return false;
}

/*********************************************************
 * Stock Helpers
 *********************************************************/


function getTotalStockByProduct(stockRows, full) {
  var total = 0;
  for (var i = 0; i < stockRows.length; i++) {
    if (normalizeData(stockRows[i].full) === normalizeData(full)) {
      total += toNumber(stockRows[i].qty);
    }
  }
  return total;
}

function findStockRow(stockRows, full, exp) {
  for (var i = 0; i < stockRows.length; i++) {
    if (
      normalizeData(stockRows[i].full) === normalizeData(full) &&
      normalizeExpKey(stockRows[i].exp) === normalizeExpKey(exp)
    ) {
      return stockRows[i];
    }
  }
  return null;
}


function planDeductFromStock(stockRows, full, qty, selectedExp) {
  qty = Number(qty);
  var remain = qty;
  var plan = [];

  if (qty <= 0) throw new Error("จำนวนต้องมากกว่า 0");

  // 1) ใช้ Exp ที่พนักงานเลือกก่อน
  if (selectedExp && selectedExp !== "-") {
    for (var i = 0; i < stockRows.length; i++) {
      var r = stockRows[i];

      if (
        normalizeData(r.full) === normalizeData(full) &&
        normalizeExpKey(r.exp) === normalizeExpKey(selectedExp) &&
        r.qty > 0
      ) {
        var cut = Math.min(r.qty, remain);
        plan.push({
          row: r,
          cut: cut,
          before: r.qty,
          after: r.qty - cut
        });

        r.qty -= cut;
        remain -= cut;

        if (remain <= 0) break;
      }
    }
  }

  // 2) ถ้า Exp ที่เลือกไม่พอ ค่อย FIFO ล็อตอื่นตามวันหมดอายุ
  if (remain > 0) {
    var candidates = [];

    for (var j = 0; j < stockRows.length; j++) {
      var r2 = stockRows[j];
      if (
        normalizeData(r2.full) === normalizeData(full) &&
        r2.qty > 0 &&
        normalizeExpKey(r2.exp) !== normalizeExpKey(selectedExp || "")
      ) {
        candidates.push(r2);
      }
    }

    candidates.sort(function(a, b) {
      return a.expSort - b.expSort;
    });

    for (var k = 0; k < candidates.length; k++) {
      var r3 = candidates[k];
      var cut2 = Math.min(r3.qty, remain);

      plan.push({
        row: r3,
        cut: cut2,
        before: r3.qty,
        after: r3.qty - cut2
      });

      r3.qty -= cut2;
      remain -= cut2;

      if (remain <= 0) break;
    }
  }

  if (remain > 0) {
    throw new Error(full + " สต๊อกไม่พอ ขาดอีก " + remain + " ชิ้น");
  }

  return plan;
}


/*********************************************************
 * Main Save Function
 *********************************************************/


/*********************************************************
 * Compatibility functions
 * เผื่อมีส่วนอื่นเรียกใช้ชื่อเดิม
 *********************************************************/
function processInboundManual(docNo, kw, full, qty, exp, source, staff) {
  var ss = SpreadsheetApp.openById(SS_ID);
  var stockSheet = getSheetOrThrow(ss, "Stock");
  var inSheet = getSheetOrThrow(ss, "StockIn");

  var stockRows = getStockRows(stockSheet);
  addStockSimulate(stockRows, kw, full, Number(qty), exp || "-");
  applyStockRowsToSheet(stockSheet, stockRows);

  inSheet.appendRow([docNo, new Date(), full, Number(qty), exp || "-", staff, source || "-"]);
}

function processReturnManual(docNo, kw, full, qty, source, staff, exp) {
  var ss = SpreadsheetApp.openById(SS_ID);
  var stockSheet = getSheetOrThrow(ss, "Stock");
  var salesSheet = getSheetOrThrow(ss, "SalesLog");

  var stockRows = getStockRows(stockSheet);
  var returnInfo = addStockSimulate(stockRows, kw, full, Number(qty), exp || "-");
  applyStockRowsToSheet(stockSheet, stockRows);

  salesSheet.appendRow([docNo, new Date(), full, -Number(qty), "รับคืน/ลงผิด: " + source, returnInfo.before, returnInfo.after, returnInfo.exp || exp || "-", staff]);
}

function processOrderManual(docNo, full, qty, source, staff, exp) {
  var ss = SpreadsheetApp.openById(SS_ID);
  var stockSheet = getSheetOrThrow(ss, "Stock");
  var salesSheet = getSheetOrThrow(ss, "SalesLog");

  var stockRows = getStockRows(stockSheet);
  var plan = planDeductFromStock(stockRows, full, Number(qty), exp || "-");
  applyStockRowsToSheet(stockSheet, stockRows);

  for (var i = 0; i < plan.length; i++) {
    var seg = plan[i];
    salesSheet.appendRow([docNo, new Date(), full, seg.cut, source, seg.before, seg.after, seg.row.exp || exp || "-", staff]);
  }
}

function processExpiredManual(docNo, full, qty, source, staff, exp) {
  var ss = SpreadsheetApp.openById(SS_ID);
  var stockSheet = getSheetOrThrow(ss, "Stock");
  var salesSheet = getSheetOrThrow(ss, "SalesLog");

  var stockRows = getStockRows(stockSheet);
  var plan = planDeductFromStock(stockRows, full, Number(qty), exp || "-");
  applyStockRowsToSheet(stockSheet, stockRows);

  for (var i = 0; i < plan.length; i++) {
    var seg = plan[i];
    salesSheet.appendRow([docNo, new Date(), full, seg.cut, "หมดอายุ: " + source, seg.before, seg.after, seg.row.exp || exp || "-", staff]);
  }
}

/*********************************************************
 * Expiry Map for Frontend
 *********************************************************/


/*********************************************************
 * Recent Logs
 *********************************************************/
function getRecentTransactions() {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var salesSheet = ss.getSheetByName("SalesLog");
    var inSheet = ss.getSheetByName("StockIn");

    var outbound = [];
    var inbound = [];

    if (salesSheet && salesSheet.getLastRow() >= 2) {
      var sData = salesSheet.getDataRange().getValues();

      for (var i = sData.length - 1; i >= 1 && outbound.length + inbound.length < 120; i--) {
        var docNo = sData[i][0];
        var dateVal = sData[i][1];
        var product = sData[i][2];
        var qty = toNumber(sData[i][3]);
        var source = String(sData[i][4] || "");
        var staff = sData[i][8] || "-";

        if (!product) continue;

        var timeStr = dateVal instanceof Date
          ? Utilities.formatDate(dateVal, "GMT+7", "dd/MM HH:mm")
          : String(dateVal);

        var ts = dateVal instanceof Date ? dateVal.getTime() : new Date(dateVal).getTime();
        if (isNaN(ts)) ts = 0;

        if (qty < 0) {
          inbound.push({
            time: timeStr,
            type: source.indexOf("EVENT_RETURN:") > -1 ? "คืนจากบูท" : "รับคืน",
            product: product,
            qty: Math.abs(qty),
            channel: source || "-",
            staff: staff,
            timestamp: ts
          });
        } else {
          var typeLabel = "ขายออก";
          if (source.indexOf("หมดอายุ") > -1) typeLabel = "หมดอายุ";
          else if (source.indexOf("สนับสนุน") > -1 || source.indexOf("แจกฟรี") > -1 || source.indexOf("PR") > -1) typeLabel = "สนับสนุน";

          outbound.push({
            time: timeStr,
            type: typeLabel,
            product: product,
            qty: qty,
            channel: source || "-",
            staff: staff,
            timestamp: ts
          });
        }
      }
    }

    if (inSheet && inSheet.getLastRow() >= 2) {
      var iData = inSheet.getDataRange().getValues();

      for (var j = iData.length - 1; j >= 1 && inbound.length < 120; j--) {
        var dVal = iData[j][1];
        var prod = iData[j][2];
        var q = toNumber(iData[j][3]);
        var staffIn = iData[j][5] || "-";
        var sourceIn = iData[j][6] || "-";

        if (!prod) continue;

        var timeStr2 = dVal instanceof Date
          ? Utilities.formatDate(dVal, "GMT+7", "dd/MM HH:mm")
          : String(dVal);

        var ts2 = dVal instanceof Date ? dVal.getTime() : new Date(dVal).getTime();
        if (isNaN(ts2)) ts2 = 0;

        inbound.push({
          time: timeStr2,
          type: "รับเข้า",
          product: prod,
          qty: q,
          channel: sourceIn || "-",
          staff: staffIn,
          timestamp: ts2
        });
      }
    }

    inbound.sort(function(a, b) { return b.timestamp - a.timestamp; });
    outbound.sort(function(a, b) { return b.timestamp - a.timestamp; });

    return {
      inbound: inbound.slice(0, 10),
      outbound: outbound.slice(0, 10)
    };
  } catch (err) {
    Logger.log("getRecentTransactions ERROR: " + err);
    return { inbound: [], outbound: [] };
  }
}

/*********************************************************
 * Dashboard
 *********************************************************/


function buildAiAdvice(donutData, lowStockItems) {
  var aiAdvice = "💡 ยินดีต้อนรับครับบอส! ระบบ V10.39 ตรวจข้อมูลคลังและยอดขายล่าสุดเรียบร้อยแล้ว\n\n";

  if (lowStockItems && lowStockItems.length > 0) {
    aiAdvice += "🚨 **สินค้าใกล้หมด:**\n";
    for (var i = 0; i < lowStockItems.length; i++) {
      aiAdvice += "• " + lowStockItems[i].name + " เหลือ " + lowStockItems[i].qty + " ชิ้น\n";
    }
    aiAdvice += "\n👉 แนะนำให้วางแผนเติมสต๊อกสินค้ากลุ่มนี้ก่อน เพื่อไม่ให้กระทบยอดขายครับ\n\n";
  }

  var topProduct = "";
  var topQty = 0;

  for (var k in donutData) {
    if (donutData[k] > topQty) {
      topQty = donutData[k];
      topProduct = k;
    }
  }

  if (topProduct) {
    aiAdvice += "🎯 **สินค้าในคลังสูงสุด:** " + topProduct + " จำนวน " + topQty + " ชิ้น\n";
    aiAdvice += "👉 แนะนำทำคอนเทนต์/โปรโมชันเพื่อหมุนสต๊อก โดยเฉพาะช่องทาง LINE OA, Facebook และ Shopee\n";
  } else {
    aiAdvice += "✨ ยังไม่มีข้อมูลสต๊อกสำหรับวิเคราะห์ครับ";
  }

  return aiAdvice;
}

/*********************************************************
 * Stock Summary for LINE
 *********************************************************/
function getCurrentStockSummary() {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var stockSheet = getSheetOrThrow(ss, "Stock");
    var data = stockSheet.getDataRange().getValues();

    var totals = {};

    for (var i = 1; i < data.length; i++) {
      var full = data[i][1];
      var qty = toNumber(data[i][2]);

      if (!full) continue;
      totals[full] = (totals[full] || 0) + qty;
    }

    var msg = "📦 สรุปสต๊อก lil bulle\nเวลา: " + nowDateTime() + "\n\n";

    Object.keys(totals).sort().forEach(function(name) {
      msg += "• " + name + ": " + totals[name] + " ชิ้น\n";
    });

    return msg;
  } catch (err) {
    return "ERROR: " + err.message;
  }
}

/*********************************************************
 * Test LINE
 * ใช้รันเองใน Apps Script เพื่อลองส่ง LINE
 *********************************************************/


var AUDIT_SHEET_NAME = "AuditLog";
var COUNT_SHEET_NAME = "StockCount";
var ADJUST_SHEET_NAME = "StockAdjust";
var LOW_LOT_THRESHOLD = 50;

function verifyBossPin_(pin) {
  var correctPin = PropertiesService.getScriptProperties().getProperty("BOSS_PIN");
  if (!correctPin) {
    throw new Error("ยังไม่ได้ตั้งค่า BOSS_PIN ใน Script Properties");
  }
  if (String(pin || "").trim() !== String(correctPin).trim()) {
    throw new Error("PIN ไม่ถูกต้อง");
  }
}

function getOrCreateSheetWithHeader_(ss, sheetName, headers) {
  var sh = ss.getSheetByName(sheetName);
  if (!sh) {
    sh = ss.insertSheet(sheetName);
  }
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  } else {
    var firstRow = sh.getRange(1, 1, 1, Math.max(headers.length, sh.getLastColumn())).getValues()[0];
    var empty = true;
    for (var i = 0; i < firstRow.length; i++) {
      if (String(firstRow[i] || "").trim() !== "") {
        empty = false;
        break;
      }
    }
    if (empty) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
      sh.setFrozenRows(1);
    }
  }
  return sh;
}

function ensureControlSheets_() {
  var ss = SpreadsheetApp.openById(SS_ID);

  var auditSheet = getOrCreateSheetWithHeader_(ss, AUDIT_SHEET_NAME, [
    "วันที่-เวลา",
    "เลขที่เอกสาร",
    "ประเภท",
    "ชื่อสินค้า",
    "วันหมดอายุ",
    "จำนวนเปลี่ยนแปลง",
    "สต๊อกก่อน",
    "สต๊อกหลัง",
    "ผู้บันทึก",
    "ช่องทาง/ที่มา",
    "หมายเหตุ"
  ]);

  var countSheet = getOrCreateSheetWithHeader_(ss, COUNT_SHEET_NAME, [
    "วันที่นับ",
    "ชื่อสินค้า",
    "วันหมดอายุ",
    "จำนวนในระบบ",
    "จำนวนจริง",
    "ส่วนต่าง",
    "ผู้ตรวจ",
    "หมายเหตุ"
  ]);

  var adjustSheet = getOrCreateSheetWithHeader_(ss, ADJUST_SHEET_NAME, [
    "วันที่ปรับ",
    "ชื่อสินค้า",
    "วันหมดอายุ",
    "สต๊อกเดิม",
    "จำนวนจริง/ยอดใหม่",
    "ปรับเพิ่ม/ลด",
    "เหตุผล",
    "ผู้อนุมัติ"
  ]);

  return {
    ss: ss,
    auditSheet: auditSheet,
    countSheet: countSheet,
    adjustSheet: adjustSheet
  };
}

function appendRowsIfAny_(sheet, rows) {
  if (!rows || rows.length === 0) return;
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

function getSystemQtyByFullExp_(stockRows, full, exp) {
  var total = 0;
  for (var i = 0; i < stockRows.length; i++) {
    if (
      normalizeData(stockRows[i].full) === normalizeData(full) &&
      normalizeExpKey(stockRows[i].exp) === normalizeExpKey(exp)
    ) {
      total += toNumber(stockRows[i].qty);
    }
  }
  return total;
}

function findExactStockRows_(stockRows, full, exp) {
  var list = [];
  for (var i = 0; i < stockRows.length; i++) {
    if (
      normalizeData(stockRows[i].full) === normalizeData(full) &&
      normalizeExpKey(stockRows[i].exp) === normalizeExpKey(exp)
    ) {
      list.push(stockRows[i]);
    }
  }
  return list;
}

function getRecentSheetRows_(sheet, limit) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var start = Math.max(2, lastRow - limit + 1);
  var values = sheet.getRange(start, 1, lastRow - start + 1, lastCol).getValues();
  var out = [];
  for (var i = values.length - 1; i >= 0; i--) {
    out.push(values[i]);
  }
  return out;
}

function getStockControlSnapshot_() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var stockSheet = getSheetOrThrow(ss, "Stock");
  var stockRows = getStockRows(stockSheet);

  var lowStockLots = [];
  var anomalies = [];
  var lotKeyMap = {};

  for (var i = 0; i < stockRows.length; i++) {
    var r = stockRows[i];
    var key = normalizeData(r.full) + "|" + normalizeExpKey(r.exp);

    if (!lotKeyMap[key]) lotKeyMap[key] = [];
    lotKeyMap[key].push(r);

    if (r.qty <= LOW_LOT_THRESHOLD) {
      lowStockLots.push({
        name: r.full,
        exp: r.exp,
        qty: r.qty,
        level: r.qty <= 0 ? "หมดแล้ว" : "ต่ำกว่า " + LOW_LOT_THRESHOLD
      });
    }

    if (r.qty < 0) {
      anomalies.push({
        type: "สต๊อกติดลบ",
        product: r.full,
        exp: r.exp,
        detail: "จำนวนคงเหลือ " + r.qty + " ชิ้น",
        sheet: "Stock",
        row: r.sheetRow,
        cell: "Stock!C" + r.sheetRow,
        docNo: "-"
      });
    }

    if (!r.exp || r.exp === "-") {
      anomalies.push({
        type: "ไม่มีวันหมดอายุ",
        product: r.full,
        exp: "-",
        detail: "ควรระบุ Exp เพื่อคุมล็อตเครื่องสำอาง",
        sheet: "Stock",
        row: r.sheetRow,
        cell: "Stock!E" + r.sheetRow,
        docNo: "-"
      });
    }
  }

  for (var k in lotKeyMap) {
    if (lotKeyMap[k].length > 1) {
      var same = lotKeyMap[k];
      anomalies.push({
        type: "พบล็อตซ้ำ",
        product: same[0].full,
        exp: same[0].exp,
        detail: "มี " + same.length + " แถวใน Stock สำหรับสินค้าและ Exp เดียวกัน",
        sheet: "Stock",
        row: same[0].sheetRow,
        cell: same.map(function(x) { return "Stock!A" + x.sheetRow + ":E" + x.sheetRow; }).join(", "),
        docNo: "-"
      });
    }
  }

  var sheets = ensureControlSheets_();
  var recentAuditRaw = getRecentSheetRows_(sheets.auditSheet, 15);
  var recentCountsRaw = getRecentSheetRows_(sheets.countSheet, 10);
  var recentAdjustsRaw = getRecentSheetRows_(sheets.adjustSheet, 10);

  var latestAuditAfter = {};
  if (sheets.auditSheet.getLastRow() >= 2) {
    var auditData = sheets.auditSheet.getRange(2, 1, sheets.auditSheet.getLastRow() - 1, 11).getValues();
    for (var a = 0; a < auditData.length; a++) {
      var product = auditData[a][3];
      var exp = auditData[a][4];
      var afterQty = toNumber(auditData[a][7]);
      if (!product) continue;
      latestAuditAfter[normalizeData(product) + "|" + normalizeExpKey(exp)] = {
        product: product,
        exp: formatExpForSheet_(exp),
        after: afterQty,
        row: a + 2,
        docNo: String(auditData[a][1] || "-")
      };
    }
  }

  for (var auditKey in latestAuditAfter) {
    var auditItem = latestAuditAfter[auditKey];
    var exactRows = findExactStockRows_(stockRows, auditItem.product, auditItem.exp);
    var currentQty = getSystemQtyByFullExp_(stockRows, auditItem.product, auditItem.exp);
    if (currentQty !== auditItem.after) {
      var swappedExp = swapExpiryDayMonth_(auditItem.exp);
      var swappedRows = normalizeExpKey(swappedExp) !== normalizeExpKey(auditItem.exp)
        ? findExactStockRows_(stockRows, auditItem.product, swappedExp) : [];

      if (exactRows.length === 0 && swappedRows.length > 0) {
        anomalies.push({
          type: "สงสัยวัน/เดือน Exp สลับ",
          product: auditItem.product,
          exp: auditItem.exp,
          detail: "AuditLog ใช้ " + auditItem.exp + " แต่ใน Stock พบ " + swappedExp + " — ตรวจวัน/เดือนก่อนปรับยอด",
          sheet: "AuditLog / Stock",
          row: auditItem.row,
          cell: "AuditLog!E" + auditItem.row + " ↔ Stock!E" + swappedRows[0].sheetRow,
          docNo: auditItem.docNo
        });
      } else {
        anomalies.push({
          type: "ยอด Stock ไม่ตรง Audit ล่าสุด",
          product: auditItem.product,
          exp: auditItem.exp,
          detail: "Stock ปัจจุบัน " + currentQty + " แต่ Audit ล่าสุดหลังทำรายการคือ " + auditItem.after,
          sheet: "AuditLog / Stock",
          row: auditItem.row,
          cell: "AuditLog!H" + auditItem.row + (exactRows.length ? " ↔ Stock!C" + exactRows[0].sheetRow : " ↔ ไม่พบล็อตใน Stock"),
          docNo: auditItem.docNo
        });
      }
    }
  }

  var anomalyTypeCount = {};
  for (var ah = 0; ah < anomalies.length; ah++) {
    var at = String(anomalies[ah].type || "อื่น ๆ");
    anomalyTypeCount[at] = (anomalyTypeCount[at] || 0) + 1;
  }
  var openEventCount = 0;
  try { openEventCount = (getOpenEventIssues().rows || []).length; } catch (e) {}
  var systemHealth = [
    { check: "รูปแบบวัน/เวลา", status: "OK", detail: "แสดง วว/ดด/ค.ศ. เวลา HH:mm:ss" },
    { check: "Exp สลับวัน/เดือน", status: anomalyTypeCount["สงสัยวัน/เดือน Exp สลับ"] ? "WARN" : "OK", detail: (anomalyTypeCount["สงสัยวัน/เดือน Exp สลับ"] || 0) + " รายการ" },
    { check: "ล็อตซ้ำ", status: anomalyTypeCount["พบล็อตซ้ำ"] ? "WARN" : "OK", detail: (anomalyTypeCount["พบล็อตซ้ำ"] || 0) + " รายการ" },
    { check: "Stock ติดลบ", status: anomalyTypeCount["สต๊อกติดลบ"] ? "ERROR" : "OK", detail: (anomalyTypeCount["สต๊อกติดลบ"] || 0) + " รายการ" },
    { check: "Stock เทียบ Audit", status: anomalyTypeCount["ยอด Stock ไม่ตรง Audit ล่าสุด"] ? "ERROR" : "OK", detail: (anomalyTypeCount["ยอด Stock ไม่ตรง Audit ล่าสุด"] || 0) + " รายการ" },
    { check: "งาน Event ยังเปิด", status: openEventCount ? "INFO" : "OK", detail: openEventCount + " งาน" }
  ];
  try {
    systemHealth = systemHealth.concat(getOperationalHealthChecks_());
  } catch (healthErr) {
    systemHealth.push({ check: "Operational Health", status: "ERROR", detail: String(healthErr) });
  }

  var recentAudit = [];
  for (var ra = 0; ra < recentAuditRaw.length; ra++) {
    recentAudit.push({
      time: safeDateTimeString(recentAuditRaw[ra][0]),
      docNo: recentAuditRaw[ra][1] || "-",
      type: recentAuditRaw[ra][2] || "-",
      product: recentAuditRaw[ra][3] || "-",
      exp: safeDateString(recentAuditRaw[ra][4]),
      qtyChange: recentAuditRaw[ra][5],
      before: recentAuditRaw[ra][6],
      after: recentAuditRaw[ra][7],
      staff: recentAuditRaw[ra][8] || "-",
      source: recentAuditRaw[ra][9] || "-",
      note: recentAuditRaw[ra][10] || "-"
    });
  }

  var recentCounts = [];
  for (var rc = 0; rc < recentCountsRaw.length; rc++) {
    recentCounts.push({
      date: safeDateTimeString(recentCountsRaw[rc][0]),
      product: recentCountsRaw[rc][1] || "-",
      exp: safeDateString(recentCountsRaw[rc][2]),
      systemQty: recentCountsRaw[rc][3],
      actualQty: recentCountsRaw[rc][4],
      diff: recentCountsRaw[rc][5],
      inspector: recentCountsRaw[rc][6] || "-",
      note: recentCountsRaw[rc][7] || "-"
    });
  }

  var recentAdjusts = [];
  for (var rd = 0; rd < recentAdjustsRaw.length; rd++) {
    recentAdjusts.push({
      date: safeDateTimeString(recentAdjustsRaw[rd][0]),
      product: recentAdjustsRaw[rd][1] || "-",
      exp: safeDateString(recentAdjustsRaw[rd][2]),
      before: recentAdjustsRaw[rd][3],
      after: recentAdjustsRaw[rd][4],
      diff: recentAdjustsRaw[rd][5],
      reason: recentAdjustsRaw[rd][6] || "-",
      approver: recentAdjustsRaw[rd][7] || "-"
    });
  }

  lowStockLots.sort(function(a, b) {
    var n = a.name.localeCompare(b.name, "th");
    if (n !== 0) return n;
    return parseExpForSort(a.exp) - parseExpForSort(b.exp);
  });

  return {
    lowStockLots: lowStockLots,
    anomalies: anomalies,
    recentAudit: recentAudit,
    recentCounts: recentCounts,
    recentAdjusts: recentAdjusts,
    anomalyCount: anomalies.length,
    lowLotCount: lowStockLots.length,
    systemHealth: systemHealth,
    workflowProgress: getWorkflowProgressV1038_(ss)
  };
}

function getStockControlData(pin) {
  try {
    verifyBossPin_(pin);
    return {
      success: true,
      data: getStockControlSnapshot_()
    };
  } catch (err) {
    Logger.log("getStockControlData ERROR: " + err);
    return {
      success: false,
      msg: err.message
    };
  }
}




function addStockSimulate(stockRows, kw, full, qty, exp) {
  qty = Number(qty);
  exp = formatExpForSheet_(exp || "-");

  var row = findStockRow(stockRows, full, exp);

  if (row) {
    var before = toNumber(row.qty);
    row.qty = before + qty;
    row.exp = formatExpForSheet_(row.exp);
    return { before: before, after: row.qty, exp: row.exp || exp };
  }

  var newRow = {
    sheetRow: null,
    kw: kw,
    full: full,
    qty: qty,
    unit: "ชิ้น",
    exp: exp,
    expSort: parseExpForSort(exp),
    newRow: true
  };

  stockRows.push(newRow);

  return { before: 0, after: qty, exp: exp };
}


function saveStockAdjust(pin, data) {
  var lock = LockService.getScriptLock();

  try {
    verifyBossPin_(pin);
    lock.waitLock(30000);

    if (!data) throw new Error("ไม่พบข้อมูลปรับสต๊อก");
    var info = getProductInfo(data.product);
    var exp = formatExpForSheet_(data.exp || "-");
    var actualQty = Number(data.actualQty);
    var reason = String(data.reason || "").trim();
    var approver = String(data.approver || "").trim();

    if (isNaN(actualQty) || actualQty < 0) throw new Error("ยอดใหม่ไม่ถูกต้อง");
    if (!reason) throw new Error("กรุณาระบุเหตุผล");
    if (!approver) throw new Error("กรุณาระบุผู้อนุมัติ");

    var ss = SpreadsheetApp.openById(SS_ID);
    var stockSheet = getSheetOrThrow(ss, "Stock");
    var stockRows = getStockRows(stockSheet);
    var exactRows = findExactStockRows_(stockRows, info.full, exp);

    if (exactRows.length > 1) {
      throw new Error("พบสินค้าและ Exp เดียวกันซ้ำหลายแถว กรุณาตรวจ Stock ก่อนปรับยอด");
    }

    var beforeQty = 0;
    if (exactRows.length === 1) {
      beforeQty = toNumber(exactRows[0].qty);
      stockSheet.getRange(exactRows[0].sheetRow, 3).setValue(actualQty);
      stockSheet.getRange(exactRows[0].sheetRow, 5).setValue(exp);
    } else {
      if (actualQty > 0) {
        stockSheet.appendRow([info.kw, info.full, actualQty, "ชิ้น", exp]);
      }
      beforeQty = 0;
    }

    var diff = actualQty - beforeQty;
    var sheets = ensureControlSheets_();

    sheets.adjustSheet.appendRow([new Date(), info.full, exp, beforeQty, actualQty, diff, reason, approver]);
    sheets.auditSheet.appendRow([new Date(), "ADJUST", "ปรับยอดสต๊อก", info.full, exp, diff, beforeQty, actualQty, approver, "StockAdjust", reason]);

    SpreadsheetApp.flush();

    sendPush(
      "🛠 ปรับยอดสต๊อก\n" +
      "สินค้า: " + info.full + "\n" +
      "Exp: " + exp + "\n" +
      "เดิม: " + beforeQty + " ชิ้น\n" +
      "ใหม่: " + actualQty + " ชิ้น\n" +
      "ส่วนต่าง: " + diff + " ชิ้น\n" +
      "เหตุผล: " + reason + "\n" +
      "ผู้อนุมัติ: " + approver
    );

    return { success: true, msg: "ปรับยอดสต๊อกสำเร็จ", before: beforeQty, after: actualQty, diff: diff };
  } catch (err) {
    Logger.log("saveStockAdjust V10.39 ERROR: " + err);
    return { success: false, msg: err.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function getStockMovementHistory(pin, filters) {
  try {
    verifyBossPin_(pin);
    filters = filters || {};

    var ss = SpreadsheetApp.openById(SS_ID);
    var sheets = ensureControlSheets_();
    var sheet = sheets.auditSheet;

    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, rows: [], total: 0 };
    }

    var lastRow = sheet.getLastRow();
    var data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
    var rows = [];

    var q = normalizeData(filters.q || "");
    var typeFilter = normalizeData(filters.type || "ALL");
    var productFilter = normalizeData(filters.product || "ALL");
    var expFilter = normalizeExpKey(filters.exp || "ALL");
    var limit = Number(filters.limit || 80);
    if (!limit || limit < 1) limit = 80;
    if (limit > 300) limit = 300;

    for (var i = data.length - 1; i >= 0; i--) {
      var r = data[i];
      var time = r[0];
      var docNo = r[1] || "-";
      var type = r[2] || "-";
      var product = r[3] || "-";
      var exp = formatExpForSheet_(r[4] || "-");
      var qtyChange = r[5];
      var before = r[6];
      var after = r[7];
      var staff = r[8] || "-";
      var source = r[9] || "-";
      var note = r[10] || "-";

      if (typeFilter !== "all" && normalizeData(type) !== typeFilter) continue;
      if (productFilter !== "all" && normalizeData(product) !== productFilter) continue;
      if (expFilter !== "all" && normalizeExpKey(exp) !== expFilter) continue;

      if (q) {
        var blob = normalizeData([docNo, type, product, exp, staff, source, note].join(" "));
        if (blob.indexOf(q) === -1) continue;
      }

      rows.push({
        time: safeDateTimeString(time),
        docNo: docNo,
        type: type,
        product: product,
        exp: exp,
        qtyChange: qtyChange,
        before: before,
        after: after,
        staff: staff,
        source: source,
        note: note
      });

      if (rows.length >= limit) break;
    }

    return { success: true, rows: rows, total: rows.length };
  } catch (err) {
    Logger.log("getStockMovementHistory ERROR: " + err);
    return { success: false, msg: err.message, rows: [] };
  }
}

function testLinePush() {
  sendPush("✅ ทดสอบ LINE lil bulle V10.39\nเวลา: " + nowDateTime());
}





function testCreateControlSheets() {
  ensureControlSheets_();
  sendPush("✅ ทดสอบสร้างชีตตรวจสอบ V10.39 สำเร็จ\nเวลา: " + nowDateTime());
}









function getProductTotalMap_() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var stockSheet = getSheetOrThrow(ss, "Stock");
  var data = stockSheet.getDataRange().getValues();
  var map = {};

  for (var i = 1; i < data.length; i++) {
    var full = String(data[i][1] || "").trim();
    var qty = toNumber(data[i][2]);
    if (!full) continue;
    map[full] = (map[full] || 0) + qty;
  }

  return map;
}


function getDashboardData(pin, monthFilter) {
  try {
    verifyBossPin_(pin);

    var ss = SpreadsheetApp.openById(SS_ID);
    var stockSheet = getSheetOrThrow(ss, "Stock");
    var salesSheet = getSheetOrThrow(ss, "SalesLog");

    var stockData = stockSheet.getDataRange().getValues();
    var salesData = salesSheet.getDataRange().getValues();
    var issueDateMap = getEventIssueDateMap_();

    var totalItems = 0;
    var donutData = {};
    var channelSales = {
      "LINE": 0,
      "Facebook": 0,
      "Shopee": 0,
      "TikTok": 0,
      "Event": 0,
      "โรงพยาบาล": 0,
      "ฝากขาย": 0,
      "แจกฟรี/รีวิว (PR)": 0,
      "อื่น ๆ": 0
    };
    var channelLocationSales = {};
    var monthlyChannelRowsMap = {};

    var lowStockItems = [];
    var lowStockLots = [];
    var allStockItems = [];
    var productTotals = {};
    var monthlySalesMap = {};
    var availableMonthsMap = {};

    for (var i = 1; i < stockData.length; i++) {
      var full = String(stockData[i][1] || "").trim();
      var qty = toNumber(stockData[i][2]);
      var unit = String(stockData[i][3] || "ชิ้น").trim();
      var exp = formatExpForSheet_(stockData[i][4]);
      if (!full) continue;

      totalItems += qty;
      donutData[full] = (donutData[full] || 0) + qty;
      if (!productTotals[full]) productTotals[full] = { qty: 0 };
      productTotals[full].qty += qty;

      allStockItems.push({
        name: full,
        qty: qty,
        unit: unit,
        exp: exp,
        sortKey: parseExpForSort(exp)
      });
    }

    for (var li = 1; li < stockData.length; li++) {
      var lotFull = String(stockData[li][1] || "").trim();
      var lotQty = toNumber(stockData[li][2]);
      var lotExp = formatExpForSheet_(stockData[li][4]);
      if (!lotFull) continue;

      if (lotQty <= LOW_LOT_THRESHOLD) {
        lowStockLots.push({
          name: lotFull,
          qty: lotQty,
          totalQty: toNumber((productTotals[lotFull] && productTotals[lotFull].qty) || 0),
          exp: lotExp,
          level: lotQty <= 0 ? "หมด" : "ต่ำกว่า " + LOW_LOT_THRESHOLD
        });
      }
    }

    allStockItems.sort(function(a, b) {
      var n = String(a.name).localeCompare(String(b.name), "th");
      if (n !== 0) return n;
      return toNumber(a.sortKey) - toNumber(b.sortKey);
    });
    for (var s = 0; s < allStockItems.length; s++) delete allStockItems[s].sortKey;

    lowStockLots.sort(function(a, b) {
      var n = String(a.name).localeCompare(String(b.name), "th");
      if (n !== 0) return n;
      return parseExpForSort(a.exp) - parseExpForSort(b.exp);
    });

    var lowStockCount = 0;
    for (var pFull in productTotals) {
      if (productTotals[pFull].qty <= LOW_LOT_THRESHOLD) {
        lowStockItems.push({
          name: pFull,
          qty: toNumber(productTotals[pFull].qty),
          exp: "รวมทุกล็อต"
        });
        lowStockCount++;
      }
    }

    var filteredSalesQty = 0;

    for (var j = 1; j < salesData.length; j++) {
      var sQty = toNumber(salesData[j][3]);
      var rawDate = salesData[j][1];
      var sDate = rawDate instanceof Date ? rawDate : new Date(rawDate);
      var sSource = String(salesData[j][4] || "");

      if (isNaN(sDate.getTime()) || sQty === 0) continue;

      var isEventReturn = sQty < 0 && sSource.indexOf("EVENT_RETURN:") > -1;
      if (sQty < 0 && !isEventReturn) continue;

      var effectiveDate = sDate;
      var chan = parseChannel(sSource);

      if (isEventReturn) {
        var refIssueNo = extractEventIssueNo_(sSource);
        if (refIssueNo && issueDateMap[refIssueNo]) {
          effectiveDate = issueDateMap[refIssueNo];
        }
        chan = "Event";
      }

      if (
        chan !== "แจกฟรี/รีวิว (PR)" &&
        sSource.indexOf("สนับสนุน") === -1 &&
        sSource.indexOf("แจกฟรี") === -1 &&
        sSource.indexOf("เบิกฟรี") === -1 &&
        sSource.indexOf("หมดอายุ") === -1
      ) {
        var sortKey = Utilities.formatDate(effectiveDate, "GMT+7", "yyyy-MM");
        var displayKey = Utilities.formatDate(effectiveDate, "GMT+7", "MMM yyyy");
        availableMonthsMap[sortKey] = displayKey;

        if (!monthlySalesMap[sortKey]) {
          monthlySalesMap[sortKey] = { display: displayKey, qty: 0 };
        }

        monthlySalesMap[sortKey].qty += sQty;

        var location = parseLocation_(sSource);
        var rowKey = sortKey + "||" + chan + "||" + location;
        if (!monthlyChannelRowsMap[rowKey]) {
          monthlyChannelRowsMap[rowKey] = { sortKey: sortKey, month: displayKey, channel: chan, location: location, qty: 0 };
        }
        monthlyChannelRowsMap[rowKey].qty += sQty;

        if (!monthFilter || monthFilter === "ALL" || monthFilter === sortKey) {
          filteredSalesQty += sQty;
          if (channelSales[chan] !== undefined) channelSales[chan] += sQty;
          else channelSales["อื่น ๆ"] += sQty;
          var clKey = chan + "||" + location;
          channelLocationSales[clKey] = (channelLocationSales[clKey] || 0) + sQty;
        }
      }
    }

    var monthList = Object.keys(availableMonthsMap).map(function(k) {
      return { val: k, txt: availableMonthsMap[k] };
    });
    monthList.sort(function(a, b) { return b.val.localeCompare(a.val); });

    var monthlyFinal = Object.keys(monthlySalesMap).sort().map(function(key) {
      return [String(monthlySalesMap[key].display), toNumber(monthlySalesMap[key].qty)];
    });

    var channelFinal = Object.keys(channelSales).map(function(key) {
      return [String(key), toNumber(channelSales[key])];
    });

    var channelLocationFinal = Object.keys(channelLocationSales).map(function(key) {
      var parts = key.split("||");
      return [parts[0] || "อื่น ๆ", parts[1] || "-", toNumber(channelLocationSales[key])];
    }).sort(function(a, b) {
      if (a[0] !== b[0]) return String(a[0]).localeCompare(String(b[0]), "th");
      return toNumber(b[2]) - toNumber(a[2]);
    });

    var monthlyChannelFinal = Object.keys(monthlyChannelRowsMap).map(function(key) {
      var r = monthlyChannelRowsMap[key];
      return [r.sortKey, r.month, r.channel, r.location, toNumber(r.qty)];
    }).sort(function(a, b) {
      if (a[0] !== b[0]) return String(b[0]).localeCompare(String(a[0]));
      if (a[2] !== b[2]) return String(a[2]).localeCompare(String(b[2]), "th");
      return toNumber(b[4]) - toNumber(a[4]);
    });

    var control = getSafeStockControlSnapshot_();
    control.lowStockLots = lowStockLots;
    control.lowLotCount = lowStockLots.length;

    var aiAdvice = buildAiAdvice(donutData, lowStockItems);
    if (control.anomalyCount > 0) {
      aiAdvice += "\n\n⚠️ **ระบบตรวจสอบพบความผิดปกติ:** " + control.anomalyCount + " รายการ ควรตรวจในกล่อง Stock Control";
    }

    return {
      success: true,
      version: SYSTEM_VERSION,
      metrics: {
        totalQty: toNumber(totalItems),
        alertCount: toNumber(lowStockCount),
        filteredSales: toNumber(filteredSalesQty),
        lowLotCount: toNumber(control.lowLotCount || 0),
        anomalyCount: toNumber(control.anomalyCount || 0)
      },
      donut: Object.keys(donutData).map(function(k) { return [String(k), toNumber(donutData[k])]; }),
      channels: channelFinal,
      channelLocations: channelLocationFinal,
      monthlyChannels: monthlyChannelFinal,
      monthly: monthlyFinal,
      lowStock: lowStockItems,
      lowStockLots: lowStockLots,
      allStock: allStockItems,
      months: monthList,
      aiAdvice: String(aiAdvice || ""),
      stockControl: control
    };
  } catch (err) {
    Logger.log("getDashboardData V10.39 ERROR: " + err);
    return { success: false, msg: "โหลด Dashboard บอสไม่สำเร็จ: " + err.toString() };
  }
}

function getSafeStockControlSnapshot_() {
  try {
    ensureControlSheets_();
    var raw = getStockControlSnapshot_() || {};
    return sanitizeStockControlSnapshot_(raw);
  } catch (err) {
    Logger.log("StockControl Snapshot V10.39 WARNING: " + err);
    return {
      lowLotCount: 0,
      anomalyCount: 1,
      lowStockLots: [],
      anomalies: [
        {
          type: "Stock Control Error",
          product: "-",
          exp: "-",
          detail: "โหลด Stock Control ไม่สำเร็จ แต่ Dashboard หลักยังใช้งานได้: " + err.toString()
        }
      ],
      recentAudit: [],
      systemHealth: [{check:"Stock Control", status:"ERROR", detail:String(err)}],
      workflowProgress: {
        stages:[{step:1,name:"Stock Control",status:"ERROR",detail:String(err),lastTime:"-"}],
        total:1,ready:0,errorCount:1,warnCount:0,percent:0,
        summary:"Stock Control โหลดไม่สำเร็จ"
      }
    };
  }
}

function sanitizeStockControlSnapshot_(raw) {
  raw = raw || {};
  var productTotalMap = {};

  try {
    productTotalMap = getProductTotalMap_();
  } catch (err) {
    Logger.log("Product total map V10.39 WARNING: " + err);
    productTotalMap = {};
  }

  var lowLots = [];
  var rawLowLots = raw.lowStockLots || [];
  for (var i = 0; i < rawLowLots.length && i < 100; i++) {
    var lotName = String(rawLowLots[i].name || rawLowLots[i].product || "-");
    lowLots.push({
      name: lotName,
      exp: formatExpForSheet_(rawLowLots[i].exp || "-"),
      qty: toNumber(rawLowLots[i].qty),
      totalQty: toNumber(rawLowLots[i].totalQty || productTotalMap[lotName] || 0),
      level: String(rawLowLots[i].level || (toNumber(rawLowLots[i].qty) <= 0 ? "หมด" : "ต่ำ"))
    });
  }

  var anomalies = [];
  var rawAnomalies = raw.anomalies || [];
  for (var j = 0; j < rawAnomalies.length && j < 100; j++) {
    anomalies.push({
      type: String(rawAnomalies[j].type || "-"),
      product: String(rawAnomalies[j].product || rawAnomalies[j].name || "-"),
      exp: formatExpForSheet_(rawAnomalies[j].exp || "-"),
      detail: String(rawAnomalies[j].detail || rawAnomalies[j].note || "-"),
      sheet: String(rawAnomalies[j].sheet || "-"),
      row: toNumber(rawAnomalies[j].row || 0),
      cell: String(rawAnomalies[j].cell || "-"),
      docNo: String(rawAnomalies[j].docNo || "-")
    });
  }

  var audit = [];
  var rawAudit = raw.recentAudit || [];
  for (var k = 0; k < rawAudit.length && k < 80; k++) {
    audit.push({
      time: String(rawAudit[k].time || "-"),
      docNo: String(rawAudit[k].docNo || "-"),
      type: String(rawAudit[k].type || "-"),
      product: String(rawAudit[k].product || "-"),
      exp: formatExpForSheet_(rawAudit[k].exp || "-"),
      qtyChange: toNumber(rawAudit[k].qtyChange),
      before: toNumber(rawAudit[k].before),
      after: toNumber(rawAudit[k].after),
      staff: String(rawAudit[k].staff || "-")
    });
  }

  var health = [];
  var rawHealth = raw.systemHealth || [];
  for (var h = 0; h < rawHealth.length && h < 20; h++) {
    health.push({
      check: String(rawHealth[h].check || "-"),
      status: String(rawHealth[h].status || "INFO"),
      detail: String(rawHealth[h].detail || "-")
    });
  }

  var workflowRaw = raw.workflowProgress || {};
  var workflowStages = [];
  var rawStages = workflowRaw.stages || [];
  for (var ws=0; ws<rawStages.length && ws<20; ws++) {
    workflowStages.push({
      step: toNumber(rawStages[ws].step || (ws+1)),
      name: String(rawStages[ws].name || "-"),
      status: String(rawStages[ws].status || "INFO"),
      detail: String(rawStages[ws].detail || "-"),
      lastTime: String(rawStages[ws].lastTime || "-")
    });
  }

  return {
    lowLotCount: toNumber(raw.lowLotCount || lowLots.length),
    anomalyCount: toNumber(raw.anomalyCount || anomalies.length),
    lowStockLots: lowLots,
    anomalies: anomalies,
    recentAudit: audit,
    systemHealth: health,
    workflowProgress: {
      stages: workflowStages,
      total: toNumber(workflowRaw.total || workflowStages.length),
      ready: toNumber(workflowRaw.ready || 0),
      errorCount: toNumber(workflowRaw.errorCount || 0),
      warnCount: toNumber(workflowRaw.warnCount || 0),
      percent: toNumber(workflowRaw.percent || 0),
      summary: String(workflowRaw.summary || "-")
    }
  };
}




function pad2_(n) {
  n = Number(n);
  return n < 10 ? "0" + n : String(n);
}

function isBlankExpiry_(val) {
  var s = String(val == null ? "" : val).trim();
  return s === "" || s === "-";
}

function parseThaiExpiryParts_(text) {
  if (isBlankExpiry_(text)) return null;

  var s = String(text).trim();
  s = s.replace(/^'/, "");
  s = s.replace(/-/g, "/");
  s = s.replace(/\s+/g, "");

  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})$/);
  if (!m) return null;

  var d = Number(m[1]);
  var mo = Number(m[2]);
  var y = Number(m[3]);

  if (y < 100) y = 2000 + y;
  if (d < 1 || d > 31 || mo < 1 || mo > 12 || y < 1900) return null;

  return { d: d, m: mo, y: y };
}

function formatExpiryParts_(parts) {
  if (!parts) return "-";
  return pad2_(parts.d) + "/" + pad2_(parts.m) + "/" + String(parts.y);
}

function swapExpiryDayMonth_(exp) {
  var p = parseThaiExpiryParts_(exp);
  if (!p) return formatExpForSheet_(exp);
  return pad2_(p.m) + "/" + pad2_(p.d) + "/" + String(p.y);
}

function formatExpForSheet_(val, displayVal) {
  if (!isBlankExpiry_(displayVal)) {
    var pDisplay = parseThaiExpiryParts_(displayVal);
    if (pDisplay) return formatExpiryParts_(pDisplay);
  }

  if (isBlankExpiry_(val)) return "-";

  if (val instanceof Date) {
    return Utilities.formatDate(val, "GMT+7", "dd/MM/yyyy");
  }

  var pValue = parseThaiExpiryParts_(val);
  if (pValue) return formatExpiryParts_(pValue);

  return String(val).trim();
}

function safeDateString(val) {
  return formatExpForSheet_(val);
}

function normalizeExpKey(val) {
  var exp = formatExpForSheet_(val);
  var p = parseThaiExpiryParts_(exp);
  if (!p) return normalizeData(exp);
  return String(p.y) + "-" + pad2_(p.m) + "-" + pad2_(p.d);
}

function parseExpForSort(exp) {
  if (isBlankExpiry_(exp)) return new Date(9999, 11, 31).getTime();

  if (exp instanceof Date) return exp.getTime();

  var p = parseThaiExpiryParts_(exp);
  if (p) return new Date(p.y, p.m - 1, p.d).getTime();

  return new Date(9999, 11, 31).getTime();
}

function getSheetValuesAndDisplays_(sheet) {
  var range = sheet.getDataRange();
  return {
    values: range.getValues(),
    displays: range.getDisplayValues()
  };
}

function getExpFromRow_(values, displays, rowIndex, expColIndex) {
  var displayVal = displays && displays[rowIndex] ? displays[rowIndex][expColIndex] : "";
  var rawVal = values && values[rowIndex] ? values[rowIndex][expColIndex] : "";
  return formatExpForSheet_(rawVal, displayVal);
}

function forceTextValue_(sheet, row, col, value) {
  sheet.getRange(row, col).setNumberFormat("@").setValue(String(value == null ? "" : value));
}

function forceExpiryTextInRange_(sheet, startRow, expCol, expValues) {
  if (!expValues || expValues.length === 0) return;
  var range = sheet.getRange(startRow, expCol, expValues.length, 1);
  range.setNumberFormat("@");
  range.setValues(expValues.map(function(v) { return [String(formatExpForSheet_(v || "-"))]; }));
}

function appendRowsWithExpiryText_(sheet, rows, expColIndexOneBased) {
  if (!rows || rows.length === 0) return;
  var startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
  if (expColIndexOneBased) {
    var expValues = [];
    for (var i = 0; i < rows.length; i++) expValues.push(rows[i][expColIndexOneBased - 1]);
    forceExpiryTextInRange_(sheet, startRow, expColIndexOneBased, expValues);
  }
}

function getStockRows(stockSheet) {
  var vd = getSheetValuesAndDisplays_(stockSheet);
  var data = vd.values;
  var disp = vd.displays;
  var rows = [];

  for (var i = 1; i < data.length; i++) {
    var kw = data[i][0];
    var full = data[i][1];
    var qty = toNumber(data[i][2]);
    var unit = data[i][3] || "ชิ้น";
    var exp = getExpFromRow_(data, disp, i, 4);

    if (!full) continue;

    rows.push({
      sheetRow: i + 1,
      kw: kw,
      full: String(full).trim(),
      qty: qty,
      unit: unit,
      exp: exp,
      expSort: parseExpForSort(exp)
    });
  }

  return rows;
}

function applyStockRowsToSheet(stockSheet, stockRows) {
  for (var i = 0; i < stockRows.length; i++) {
    var r = stockRows[i];
    var cleanExp = formatExpForSheet_(r.exp || "-");

    if (r.sheetRow) {
      stockSheet.getRange(r.sheetRow, 3).setValue(r.qty);
      forceTextValue_(stockSheet, r.sheetRow, 5, cleanExp);
    } else if (r.newRow) {
      var newRow = stockSheet.getLastRow() + 1;
      stockSheet.getRange(newRow, 1, 1, 5).setValues([[
        r.kw,
        r.full,
        r.qty,
        r.unit || "ชิ้น",
        cleanExp
      ]]);
      forceTextValue_(stockSheet, newRow, 5, cleanExp);
    }
  }
}

function getExpiryMap() {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var sheet = getSheetOrThrow(ss, "Stock");
    var vd = getSheetValuesAndDisplays_(sheet);
    var data = vd.values;
    var disp = vd.displays;
    var map = {};

    for (var i = 1; i < data.length; i++) {
      var full = String(data[i][1] || "").trim();
      var qty = toNumber(data[i][2]);
      var exp = getExpFromRow_(data, disp, i, 4);

      if (full && qty > 0) {
        if (!map[full]) map[full] = [];
        map[full].push({ exp: exp, qty: qty, sort: parseExpForSort(exp) });
      }
    }

    Object.keys(map).forEach(function(key) {
      map[key].sort(function(a, b) { return a.sort - b.sort; });
      map[key] = map[key].map(function(item) {
        return { exp: item.exp, qty: item.qty };
      });
    });

    return map;
  } catch (err) {
    Logger.log("getExpiryMap V10.39 ERROR: " + err);
    return {};
  }
}

function saveMultiple(masterData, itemsList) {
  var lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    if (!masterData) throw new Error("ไม่พบข้อมูลเอกสาร");
    if (!itemsList || itemsList.length === 0) throw new Error("ไม่มีรายการสินค้า");

    var docNo = String(masterData.docNo || "").trim();
    var staff = String(masterData.staff || "").trim();

    if (!staff) throw new Error("กรุณาระบุผู้บันทึก");
    if (!docNo || docNo === "AUTO") docNo = makeWarehouseDocNo_();

    if (isDuplicateDoc(docNo)) {
      throw new Error("เลขที่เอกสาร " + docNo + " ถูกใช้ไปแล้ว ห้ามบันทึกซ้ำ");
    }

    var ss = SpreadsheetApp.openById(SS_ID);
    var stockSheet = getSheetOrThrow(ss, "Stock");
    var salesSheet = getSheetOrThrow(ss, "SalesLog");
    var inSheet = getSheetOrThrow(ss, "StockIn");

    var controlSheets = ensureControlSheets_();
    var stockRows = getStockRows(stockSheet);

    var salesLogs = [];
    var inboundLogs = [];
    var auditLogs = [];

    for (var i = 0; i < itemsList.length; i++) {
      var item = itemsList[i];
      var info = getProductInfo(item.product);

      var qty = Number(item.qty);
      var exp = formatExpForSheet_(item.exp || "-");
      var channel = item.channel || "อื่น ๆ";
      var location = String(item.location || "").trim();
      var note = item.note || "";
      var combinedSource = buildSourceText_(channel, location, note);

      if (!qty || qty <= 0) {
        throw new Error("จำนวนสินค้าไม่ถูกต้อง: " + item.product);
      }

      if (item.type === "INBOUND") {
        var inboundInfo = addStockSimulate(stockRows, info.kw, info.full, qty, exp);
        var inboundExp = formatExpForSheet_(inboundInfo.exp || exp || "-");

        inboundLogs.push([docNo, new Date(), info.full, qty, inboundExp, staff, combinedSource]);
        auditLogs.push([new Date(), docNo, "รับเข้าคลัง", info.full, inboundExp, qty, inboundInfo.before, inboundInfo.after, staff, combinedSource, note || "-"]);
      }

      else if (item.type === "RETURN") {
        var returnInfo = addStockSimulate(stockRows, info.kw, info.full, qty, exp);
        var returnExp = formatExpForSheet_(returnInfo.exp || exp || "-");

        salesLogs.push([docNo, new Date(), info.full, -qty, "รับคืน/ลงผิด: " + combinedSource, returnInfo.before, returnInfo.after, returnExp, staff]);
        auditLogs.push([new Date(), docNo, "รับคืน/ลงผิด", info.full, returnExp, qty, returnInfo.before, returnInfo.after, staff, combinedSource, note || "-"]);
      }

      else {
        var plan = planDeductFromStock(stockRows, info.full, qty, exp);
        var multiLotText = plan.length > 1 ? " | ระบบแยกตัดหลายล็อต" : "";
        var sourceText = combinedSource;
        var auditType = "ขายออก";

        if (item.type === "EXPIRED") {
          sourceText = "หมดอายุ: " + combinedSource;
          auditType = "หมดอายุ/เสียหาย";
        } else if (item.type === "FREE") {
          sourceText = "สนับสนุน/แจกฟรี: " + combinedSource;
          auditType = "สนับสนุน/แจกฟรี";
        }

        for (var p = 0; p < plan.length; p++) {
          var seg = plan[p];
          var deductExp = formatExpForSheet_(seg.row.exp || exp || "-");

          salesLogs.push([docNo, new Date(), info.full, seg.cut, sourceText + multiLotText, seg.before, seg.after, deductExp, staff]);
          auditLogs.push([new Date(), docNo, auditType, info.full, deductExp, -seg.cut, seg.before, seg.after, staff, combinedSource, note || "-"]);
        }
      }
    }

    applyStockRowsToSheet(stockSheet, stockRows);

    appendRowsWithExpiryText_(salesSheet, salesLogs, 8);
    appendRowsWithExpiryText_(inSheet, inboundLogs, 5);
    appendRowsWithExpiryText_(controlSheets.auditSheet, auditLogs, 5);

    SpreadsheetApp.flush();
    notifyByType(docNo, staff, itemsList);

    return { success: true, msg: "SUCCESS", docNo: docNo };
  } catch (err) {
    Logger.log("saveMultiple V10.39 ERROR: " + err);
    return { success: false, msg: "ERROR : " + err.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function buildStockExpiryLookup_() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sheet = getSheetOrThrow(ss, "Stock");
  var vd = getSheetValuesAndDisplays_(sheet);
  var data = vd.values;
  var disp = vd.displays;
  var lookup = {};

  for (var i = 1; i < data.length; i++) {
    var product = String(data[i][1] || "").trim();
    if (!product) continue;
    var exp = getExpFromRow_(data, disp, i, 4);
    if (!lookup[product]) lookup[product] = {};
    lookup[product][normalizeExpKey(exp)] = exp;
  }

  return lookup;
}

function resolveExpiryAgainstStock_(product, exp, stockLookup) {
  exp = formatExpForSheet_(exp || "-");
  if (!product || !stockLookup || !stockLookup[product]) return exp;

  var productExps = stockLookup[product];
  var key = normalizeExpKey(exp);
  if (productExps[key]) return productExps[key];

  var swapped = swapExpiryDayMonth_(exp);
  var swappedKey = normalizeExpKey(swapped);
  if (productExps[swappedKey]) return productExps[swappedKey];

  return exp;
}

function normalizeExpiryColumnBySheet_(sheetName, productCol, expCol, stockLookup, mode) {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { sheet: sheetName, updated: 0, skipped: true };

  var vd = getSheetValuesAndDisplays_(sheet);
  var data = vd.values;
  var disp = vd.displays;
  var updates = 0;

  for (var i = 1; i < data.length; i++) {
    var product = productCol ? String(data[i][productCol - 1] || "").trim() : "";
    var oldExp = getExpFromRow_(data, disp, i, expCol - 1);
    var newExp = oldExp;

    if (mode === "stock") {
      newExp = formatExpForSheet_(oldExp);
    } else {
      newExp = resolveExpiryAgainstStock_(product, oldExp, stockLookup);
    }

    if (!isBlankExpiry_(newExp)) {
      forceTextValue_(sheet, i + 1, expCol, newExp);
      if (normalizeData(String(oldExp)) !== normalizeData(String(newExp))) updates++;
    }
  }

  return { sheet: sheetName, updated: updates, skipped: false };
}

function normalizeAllExpirySheets_() {
  var results = [];

  // 1) Stock ต้องทำก่อน เพื่อให้เป็นฐานอ้างอิงหลัก
  results.push(normalizeExpiryColumnBySheet_("Stock", 2, 5, {}, "stock"));

  var stockLookup = buildStockExpiryLookup_();

  // 2) ชีตอื่นเทียบกับ Stock ถ้าวัน/เดือนสลับ ระบบจะดึงกลับเป็น Exp เดียวกับ Stock
  results.push(normalizeExpiryColumnBySheet_("SalesLog", 3, 8, stockLookup, "matchStock"));
  results.push(normalizeExpiryColumnBySheet_("StockIn", 3, 5, stockLookup, "matchStock"));
  results.push(normalizeExpiryColumnBySheet_("AuditLog", 4, 5, stockLookup, "matchStock"));
  results.push(normalizeExpiryColumnBySheet_("StockCount", 2, 3, stockLookup, "matchStock"));
  results.push(normalizeExpiryColumnBySheet_("StockAdjust", 2, 3, stockLookup, "matchStock"));
  results.push(normalizeExpiryColumnBySheet_("StockAdjustRequest", 3, 4, stockLookup, "matchStock"));
  results.push(normalizeExpiryColumnBySheet_("EventIssue", 4, 5, stockLookup, "matchStock"));
  results.push(normalizeExpiryColumnBySheet_("EventReturn", 5, 6, stockLookup, "matchStock"));

  SpreadsheetApp.flush();
  return results;
}

function testNormalizeExpiry() {
  var results = normalizeAllExpirySheets_();
  var msg = "✅ ปรับรูปแบบวันหมดอายุ V10.39 สำเร็จ\n";
  msg += "รูปแบบมาตรฐาน: dd/MM/yyyy เช่น 05/06/2028\n";
  msg += "เวลา: " + nowDateTime() + "\n\n";

  for (var i = 0; i < results.length; i++) {
    msg += "• " + results[i].sheet + ": ปรับ " + results[i].updated + " ช่อง\n";
  }

  sendPush(msg);
  Logger.log(msg);
  return msg;
}

function getTeamStockViewerData() {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var stockSheet = getSheetOrThrow(ss, "Stock");
    var vd = getSheetValuesAndDisplays_(stockSheet);
    var data = vd.values;
    var disp = vd.displays;
    var totalMap = {};
    var rows = [];

    for (var t = 1; t < data.length; t++) {
      var totalName = String(data[t][1] || "").trim();
      if (!totalName) continue;
      totalMap[totalName] = (totalMap[totalName] || 0) + toNumber(data[t][2]);
    }

    for (var i = 1; i < data.length; i++) {
      var kw = data[i][0] || "";
      var full = String(data[i][1] || "").trim();
      var qty = toNumber(data[i][2]);
      var unit = data[i][3] || "ชิ้น";
      var exp = getExpFromRow_(data, disp, i, 4);

      if (!full) continue;

      var status = "ปกติ";
      var level = "normal";
      if (qty <= 0) {
        status = "หมด / 0 ชิ้น";
        level = "danger";
      } else if (qty <= LOW_LOT_THRESHOLD) {
        status = "ต่ำกว่า " + LOW_LOT_THRESHOLD;
        level = "warning";
      }

      rows.push({
        kw: String(kw),
        product: full,
        qty: qty,
        totalQty: toNumber(totalMap[full] || 0),
        unit: String(unit || "ชิ้น"),
        exp: exp,
        status: status,
        level: level,
        sort: parseExpForSort(exp)
      });
    }

    rows.sort(function(a, b) {
      var n = String(a.product).localeCompare(String(b.product), "th");
      if (n !== 0) return n;
      return toNumber(a.sort) - toNumber(b.sort);
    });

    for (var r = 0; r < rows.length; r++) delete rows[r].sort;

    return { success: true, rows: rows, version: SYSTEM_VERSION };
  } catch (err) {
    Logger.log("getTeamStockViewerData V10.39 ERROR: " + err);
    return { success: false, msg: err.toString(), rows: [] };
  }
}

function testBossDashboard() {
  var pin = PropertiesService.getScriptProperties().getProperty("BOSS_PIN");
  if (!pin) throw new Error("ยังไม่ได้ตั้งค่า BOSS_PIN ใน Script Properties");
  var res = getDashboardData(pin, "ALL");
  Logger.log(JSON.stringify(res).substring(0, 5000));
  return res;
}


var ADJUST_REQUEST_SHEET_NAME = "StockAdjustRequest";

function doGet(e) {
  if(e && e.parameter && e.parameter.page==='pos') return HtmlService.createHtmlOutput('<!doctype html><html lang="th"><meta charset="utf-8"><title>ย้าย POS ไปเว็บใหม่แล้ว</title><body style="font-family:Tahoma;padding:40px;line-height:1.8"><h2>ย้ายหน้าขายไปเว็บใหม่แล้ว</h2><p>ประวัติเดิมยังเก็บอยู่ ไม่มีการลบข้อมูล</p><a href="https://lilbulle-booth-pos.netlify.app/" target="_blank" rel="noopener">เปิดเว็บ lilbulle Booth POS</a></body></html>');
  var isPos=!!(e && e.parameter && e.parameter.page==='pos');
  // The POS entry does not run the inventory setup.
  if(!isPos){try { ensureV1035Setup_(); } catch (err) { Logger.log("V10.39 setup warning: " + err); }}
  var template=HtmlService.createTemplateFromFile(isPos ? "POS" : "Index");
  template.webAppUrl=ScriptApp.getService().getUrl() || '';
  return template.evaluate()
    .setTitle(isPos ? "Lil Bulle · POS V10.43" : "lil bulle Corporate Inventory " + SYSTEM_VERSION)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function ensureAdjustRequestSheet_() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sheet = ss.getSheetByName(ADJUST_REQUEST_SHEET_NAME);

  var headers = [
    "เวลา",
    "เลขคำขอ",
    "สินค้า",
    "Exp",
    "ยอดในระบบตอนนับ",
    "ยอดนับจริง",
    "ส่วนต่าง",
    "เหตุผลทีมงาน",
    "ผู้ขอ",
    "สถานะ",
    "ผู้อนุมัติ",
    "เวลาอนุมัติ",
    "หมายเหตุบอส"
  ];

  if (!sheet) {
    sheet = ss.insertSheet(ADJUST_REQUEST_SHEET_NAME);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    sheet.autoResizeColumns(1, headers.length);
  }

  return sheet;
}

function makeAdjustRequestId_() {
  var now = new Date();
  var stamp = Utilities.formatDate(now, "GMT+7", "yyyyMMdd-HHmmss");
  var rand = Math.floor(100 + Math.random() * 900);
  return "SAR-" + stamp + "-" + rand;
}

function appendAdjustRequest_(data) {
  var sheet = ensureAdjustRequestSheet_();
  var requestId = makeAdjustRequestId_();
  var exp = formatExpForSheet_(data.exp || "-");

  var row = sheet.getLastRow() + 1;
  sheet.getRange(row, 1, 1, 13).setValues([[
    new Date(),
    requestId,
    data.product,
    exp,
    Number(data.systemQty),
    Number(data.actualQty),
    Number(data.diff),
    data.reason || "ผลนับจริงไม่ตรงกับระบบ",
    data.requester || "-",
    "รออนุมัติ",
    "",
    "",
    ""
  ]]);
  forceTextValue_(sheet, row, 4, exp);
  return requestId;
}

function saveStockCount(pinOrData, dataMaybe) {
  try {
    var data = dataMaybe || pinOrData;
    if (!data) throw new Error("ไม่พบข้อมูลนับสต๊อก");

    var info = getProductInfo(data.product);
    var exp = formatExpForSheet_(data.exp || "-");
    var actualQty = Number(data.actualQty);
    var inspector = String(data.inspector || "").trim();
    var note = String(data.note || "").trim();

    if (isNaN(actualQty) || actualQty < 0) throw new Error("จำนวนจริงไม่ถูกต้อง");
    if (!inspector) throw new Error("กรุณาระบุผู้ตรวจ");

    var ss = SpreadsheetApp.openById(SS_ID);
    var stockSheet = getSheetOrThrow(ss, "Stock");
    var stockRows = getStockRows(stockSheet);
    var systemQty = getSystemQtyByFullExp_(stockRows, info.full, exp);
    var diff = actualQty - systemQty;
    var sheets = ensureControlSheets_();

    var countRow = sheets.countSheet.getLastRow() + 1;
    sheets.countSheet.getRange(countRow, 1, 1, 8).setValues([[
      new Date(),
      info.full,
      exp,
      systemQty,
      actualQty,
      diff,
      inspector,
      note || "-"
    ]]);
    forceTextValue_(sheets.countSheet, countRow, 3, exp);

    var auditRow = sheets.auditSheet.getLastRow() + 1;
    sheets.auditSheet.getRange(auditRow, 1, 1, 11).setValues([[
      new Date(),
      "COUNT",
      "ตรวจนับจริง",
      info.full,
      exp,
      diff,
      systemQty,
      systemQty,
      inspector,
      "StockCount",
      note || "-"
    ]]);
    forceTextValue_(sheets.auditSheet, auditRow, 5, exp);

    var requestId = "";
    if (diff !== 0) {
      requestId = appendAdjustRequest_({
        product: info.full,
        exp: exp,
        systemQty: systemQty,
        actualQty: actualQty,
        diff: diff,
        reason: note || "ผลนับจริงไม่ตรงกับระบบ",
        requester: inspector
      });

      sendPush(
        "🧮 พบผลนับสต๊อกไม่ตรง — รออนุมัติ\n" +
        "เลขคำขอ: " + requestId + "\n" +
        "สินค้า: " + info.full + "\n" +
        "Exp: " + exp + "\n" +
        "ในระบบ: " + systemQty + " ชิ้น\n" +
        "นับจริง: " + actualQty + " ชิ้น\n" +
        "ส่วนต่าง: " + diff + " ชิ้น\n" +
        "ผู้ตรวจ: " + inspector + "\n\n" +
        "สถานะ: รอบอสอนุมัติ — Stock ยังไม่ถูกปรับ"
      );
    }

    SpreadsheetApp.flush();

    return {
      success: true,
      msg: diff === 0 ? "บันทึกผลนับจริงสำเร็จ ยอดตรงกับระบบ" : "บันทึกผลนับจริงสำเร็จ และส่งคำขอรออนุมัติแล้ว",
      systemQty: systemQty,
      actualQty: actualQty,
      diff: diff,
      requestId: requestId,
      pendingApproval: diff !== 0
    };
  } catch (err) {
    Logger.log("saveStockCount V10.39 ERROR: " + err);
    return { success: false, msg: err.message };
  }
}

function getPendingAdjustRequests(pin) {
  try {
    verifyBossPin_(pin);

    var sheet = ensureAdjustRequestSheet_();
    var range = sheet.getDataRange();
    var values = range.getValues();
    var displays = range.getDisplayValues();
    var rows = [];

    var ss = SpreadsheetApp.openById(SS_ID);
    var stockSheet = getSheetOrThrow(ss, "Stock");
    var stockRows = getStockRows(stockSheet);

    // ใช้ตรวจคำขอซ้ำที่เคยอนุมัติไปแล้ว
    var approvedSignatureMap = {};
    for (var a = 1; a < values.length; a++) {
      var approvedStatus = String(values[a][9] || "").trim();
      if (approvedStatus !== "อนุมัติแล้ว") continue;

      var approvedProduct = String(values[a][2] || "").trim();
      var approvedExp = formatExpForSheet_(values[a][3], displays[a][3]);
      var approvedSystem = Number(values[a][4] || 0);
      var approvedActual = Number(values[a][5] || 0);
      var approvedDiff = Number(values[a][6] || 0);
      var approvedKey =
        normalizeData(approvedProduct) + "|" +
        normalizeExpKey(approvedExp) + "|" +
        approvedSystem + "|" + approvedActual + "|" + approvedDiff;

      approvedSignatureMap[approvedKey] = String(values[a][1] || "").trim();
    }

    for (var i = 1; i < values.length; i++) {
      var status = String(values[i][9] || "").trim();
      if (status !== "รออนุมัติ") continue;

      var product = String(values[i][2] || "").trim();
      var exp = formatExpForSheet_(values[i][3], displays[i][3]);
      var systemQty = Number(values[i][4] || 0);
      var actualQty = Number(values[i][5] || 0);
      var diff = Number(values[i][6] || 0);
      var currentQty = getSystemQtyByFullExp_(stockRows, product, exp);

      var signature =
        normalizeData(product) + "|" +
        normalizeExpKey(exp) + "|" +
        systemQty + "|" + actualQty + "|" + diff;

      var duplicateApprovedId = approvedSignatureMap[signature] || "";
      var approvalState = "READY";
      var approvalStateText = "พร้อมอนุมัติ";
      var approvalDetail = "Stock ปัจจุบันยังตรงกับยอดในระบบตอนทีมงานนับ";
      var canApprove = true;

      if (duplicateApprovedId) {
        approvalState = "DUPLICATE";
        approvalStateText = "คำขอซ้ำ";
        approvalDetail = "ข้อมูลชุดเดียวกันเคยอนุมัติแล้วใน " + duplicateApprovedId + " — ห้ามอนุมัติซ้ำ";
        canApprove = false;
      } else if (currentQty !== systemQty) {
        approvalState = "STALE";
        approvalStateText = "Stock เปลี่ยนแล้ว";
        approvalDetail =
          "ตอนทีมงานนับระบบมี " + systemQty +
          " / ปัจจุบัน " + currentQty +
          " — ต้องนับของจริงใหม่ก่อนอนุมัติ";
        canApprove = false;
      }

      rows.push({
        sheetRow: i + 1,
        time: displays[i][0] || "-",
        requestId: String(values[i][1] || ""),
        product: product,
        exp: exp,
        systemQty: systemQty,
        actualQty: actualQty,
        diff: diff,
        currentQty: currentQty,
        reason: String(values[i][7] || "-"),
        requester: String(values[i][8] || "-"),
        canApprove: canApprove,
        approvalState: approvalState,
        approvalStateText: approvalStateText,
        approvalDetail: approvalDetail,
        duplicateApprovedId: duplicateApprovedId
      });
    }

    rows.sort(function(a, b) { return b.sheetRow - a.sheetRow; });
    return { success: true, rows: rows };
  } catch (err) {
    Logger.log("getPendingAdjustRequests V10.39 ERROR: " + err);
    return { success: false, msg: err.message, rows: [] };
  }
}

function findPendingRequestRow_(sheet, requestId) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1] || "").trim() === String(requestId || "").trim()) {
      return { row: i + 1, values: data[i] };
    }
  }
  return null;
}

function approveStockAdjustRequest(pin, requestId, approver, bossNote) {
  var lock = LockService.getScriptLock();

  try {
    verifyBossPin_(pin);
    lock.waitLock(30000);

    requestId = String(requestId || "").trim();
    approver = String(approver || "").trim();
    bossNote = String(bossNote || "").trim();

    if (!requestId) throw new Error("ไม่พบเลขคำขอ");
    if (!approver) throw new Error("กรุณาระบุชื่อผู้อนุมัติ");

    var requestSheet = ensureAdjustRequestSheet_();
    var found = findPendingRequestRow_(requestSheet, requestId);
    if (!found) throw new Error("ไม่พบคำขอนี้");

    var row = found.row;
    var r = found.values;
    var status = String(r[9] || "").trim();
    if (status !== "รออนุมัติ") throw new Error("คำขอนี้ไม่ได้อยู่ในสถานะรออนุมัติแล้ว");

    var full = String(r[2] || "").trim();
    var exp = formatExpForSheet_(r[3] || "-");
    var countedSystemQty = Number(r[4] || 0);
    var actualQty = Number(r[5] || 0);
    var requestReason = String(r[7] || "-");
    var requester = String(r[8] || "-");

    var ss = SpreadsheetApp.openById(SS_ID);
    var stockSheet = getSheetOrThrow(ss, "Stock");
    var stockRows = getStockRows(stockSheet);
    var exactRows = findExactStockRows_(stockRows, full, exp);

    if (exactRows.length > 1) {
      throw new Error("พบสินค้าและ Exp เดียวกันซ้ำหลายแถว กรุณาตรวจ Stock ก่อนอนุมัติ");
    }

    var currentQty = exactRows.length === 1 ? toNumber(exactRows[0].qty) : 0;
    if (currentQty !== countedSystemQty) {
      throw new Error(
        "ยอด Stock เปลี่ยนหลังจากทีมงานนับแล้ว (ตอนนับ " + countedSystemQty +
        " / ปัจจุบัน " + currentQty + ") กรุณาให้ทีมงานนับใหม่ก่อนอนุมัติ"
      );
    }

    if (exactRows.length === 1) {
      stockSheet.getRange(exactRows[0].sheetRow, 3).setValue(actualQty);
      forceTextValue_(stockSheet, exactRows[0].sheetRow, 5, exp);
    } else if (actualQty > 0) {
      var info = getProductInfo(full);
      var newRow = stockSheet.getLastRow() + 1;
      stockSheet.getRange(newRow, 1, 1, 5).setValues([[
        info.kw,
        info.full,
        actualQty,
        "ชิ้น",
        exp
      ]]);
      forceTextValue_(stockSheet, newRow, 5, exp);
    }

    var diff = actualQty - currentQty;
    var sheets = ensureControlSheets_();

    var adjustRow = sheets.adjustSheet.getLastRow() + 1;
    sheets.adjustSheet.getRange(adjustRow, 1, 1, 8).setValues([[
      new Date(),
      full,
      exp,
      currentQty,
      actualQty,
      diff,
      "อนุมัติคำขอ " + requestId + ": " + requestReason,
      approver
    ]]);
    forceTextValue_(sheets.adjustSheet, adjustRow, 3, exp);

    var auditRow = sheets.auditSheet.getLastRow() + 1;
    sheets.auditSheet.getRange(auditRow, 1, 1, 11).setValues([[
      new Date(),
      requestId,
      "อนุมัติปรับยอดสต๊อก",
      full,
      exp,
      diff,
      currentQty,
      actualQty,
      approver,
      "StockAdjustRequest",
      "ผู้ขอ: " + requester + " | เหตุผล: " + requestReason + (bossNote ? " | หมายเหตุบอส: " + bossNote : "")
    ]]);
    forceTextValue_(sheets.auditSheet, auditRow, 5, exp);

    requestSheet.getRange(row, 10, 1, 4).setValues([[
      "อนุมัติแล้ว",
      approver,
      new Date(),
      bossNote || "-"
    ]]);

    SpreadsheetApp.flush();

    sendPush(
      "✅ บอสอนุมัติปรับยอดสต๊อกแล้ว\n" +
      "เลขคำขอ: " + requestId + "\n" +
      "สินค้า: " + full + "\n" +
      "Exp: " + exp + "\n" +
      "เดิม: " + currentQty + " ชิ้น\n" +
      "ใหม่: " + actualQty + " ชิ้น\n" +
      "ส่วนต่าง: " + diff + " ชิ้น\n" +
      "ผู้ขอ: " + requester + "\n" +
      "ผู้อนุมัติ: " + approver
    );

    return {
      success: true,
      msg: "อนุมัติและปรับยอด Stock สำเร็จ",
      before: currentQty,
      after: actualQty,
      diff: diff
    };
  } catch (err) {
    Logger.log("approveStockAdjustRequest V10.39 ERROR: " + err);
    return { success: false, msg: err.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function rejectStockAdjustRequest(pin, requestId, approver, bossNote) {
  var lock = LockService.getScriptLock();

  try {
    verifyBossPin_(pin);
    lock.waitLock(30000);

    requestId = String(requestId || "").trim();
    approver = String(approver || "").trim();
    bossNote = String(bossNote || "").trim();

    if (!requestId) throw new Error("ไม่พบเลขคำขอ");
    if (!approver) throw new Error("กรุณาระบุชื่อผู้พิจารณา");
    if (!bossNote) throw new Error("กรุณาระบุเหตุผลที่ไม่อนุมัติ");

    var requestSheet = ensureAdjustRequestSheet_();
    var found = findPendingRequestRow_(requestSheet, requestId);
    if (!found) throw new Error("ไม่พบคำขอนี้");

    var row = found.row;
    var r = found.values;
    var status = String(r[9] || "").trim();
    if (status !== "รออนุมัติ") throw new Error("คำขอนี้ไม่ได้อยู่ในสถานะรออนุมัติแล้ว");

    var full = String(r[2] || "").trim();
    var exp = formatExpForSheet_(r[3] || "-");
    var systemQty = Number(r[4] || 0);
    var actualQty = Number(r[5] || 0);
    var diff = Number(r[6] || 0);
    var requester = String(r[8] || "-");

    requestSheet.getRange(row, 10, 1, 4).setValues([[
      "ไม่อนุมัติ",
      approver,
      new Date(),
      bossNote
    ]]);

    var sheets = ensureControlSheets_();
    var auditRow = sheets.auditSheet.getLastRow() + 1;
    sheets.auditSheet.getRange(auditRow, 1, 1, 11).setValues([[
      new Date(),
      requestId,
      "ไม่อนุมัติปรับยอดสต๊อก",
      full,
      exp,
      0,
      systemQty,
      systemQty,
      approver,
      "StockAdjustRequest",
      "ผู้ขอ: " + requester + " | นับจริง: " + actualQty + " | ต่าง: " + diff + " | เหตุผลไม่อนุมัติ: " + bossNote
    ]]);
    forceTextValue_(sheets.auditSheet, auditRow, 5, exp);

    SpreadsheetApp.flush();

    sendPush(
      "❌ ไม่อนุมัติคำขอปรับยอดสต๊อก\n" +
      "เลขคำขอ: " + requestId + "\n" +
      "สินค้า: " + full + "\n" +
      "Exp: " + exp + "\n" +
      "ผู้ขอ: " + requester + "\n" +
      "ผู้พิจารณา: " + approver + "\n" +
      "เหตุผล: " + bossNote
    );

    return { success: true, msg: "บันทึกไม่อนุมัติเรียบร้อย Stock ไม่ถูกเปลี่ยน" };
  } catch (err) {
    Logger.log("rejectStockAdjustRequest V10.39 ERROR: " + err);
    return { success: false, msg: err.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/*********************************************************
 * V10.39 — BATCH APPROVAL
 * เลือกหลายคำขอแล้วอนุมัติ/ปฏิเสธพร้อมกัน
 * - ตรวจ PIN ครั้งเดียว
 * - Lock ครั้งเดียว
 * - คง stale-stock protection ของแต่ละคำขอ
 * - ถ้าบางรายการไม่ผ่าน จะไม่หยุดรายการอื่น
 * - LINE แจ้งสรุปครั้งเดียว ลดข้อความซ้ำ
 *********************************************************/
function batchProcessStockAdjustRequests(pin, requestIds, action, approver, note) {
  var lock = LockService.getScriptLock();

  try {
    verifyBossPin_(pin);
    lock.waitLock(30000);

    action = String(action || "").toUpperCase().trim();
    approver = String(approver || "").trim();
    note = String(note || "").trim();

    if (action !== "APPROVE" && action !== "REJECT") {
      throw new Error("คำสั่งอนุมัติไม่ถูกต้อง");
    }
    if (!approver) {
      throw new Error(action === "APPROVE" ? "กรุณาระบุชื่อผู้อนุมัติ" : "กรุณาระบุชื่อผู้พิจารณา");
    }
    if (action === "REJECT" && !note) {
      throw new Error("กรุณาระบุเหตุผลที่ไม่อนุมัติ");
    }

    if (!Array.isArray(requestIds)) requestIds = [];
    var uniqueIds = [];
    var seen = {};
    for (var i = 0; i < requestIds.length; i++) {
      var id = String(requestIds[i] || "").trim();
      if (!id || seen[id]) continue;
      seen[id] = true;
      uniqueIds.push(id);
    }

    if (!uniqueIds.length) throw new Error("กรุณาเลือกรายการอย่างน้อย 1 รายการ");
    if (uniqueIds.length > 100) throw new Error("เลือกได้สูงสุด 100 รายการต่อครั้ง");

    var requestSheet = ensureAdjustRequestSheet_();
    var requestData = requestSheet.getDataRange().getValues();
    var requestMap = {};

    for (var r = 1; r < requestData.length; r++) {
      var rid = String(requestData[r][1] || "").trim();
      if (rid) requestMap[rid] = { row: r + 1, values: requestData[r] };
    }

    var ss = SpreadsheetApp.openById(SS_ID);
    var stockSheet = getSheetOrThrow(ss, "Stock");
    var stockRows = getStockRows(stockSheet);
    var sheets = ensureControlSheets_();

    var results = [];
    var successCount = 0;
    var failedCount = 0;

    for (var n = 0; n < uniqueIds.length; n++) {
      var requestId = uniqueIds[n];

      try {
        var found = requestMap[requestId];
        if (!found) throw new Error("ไม่พบคำขอนี้");

        var row = found.row;
        var v = found.values;
        var status = String(v[9] || "").trim();
        if (status !== "รออนุมัติ") {
          throw new Error("รายการไม่ได้อยู่ในสถานะรออนุมัติแล้ว");
        }

        var full = String(v[2] || "").trim();
        var exp = formatExpForSheet_(v[3] || "-");
        var countedSystemQty = Number(v[4] || 0);
        var actualQty = Number(v[5] || 0);
        var requestDiff = Number(v[6] || 0);
        var requestReason = String(v[7] || "-");
        var requester = String(v[8] || "-");

        if (action === "APPROVE") {
          var exactRows = findExactStockRows_(stockRows, full, exp);

          if (exactRows.length > 1) {
            throw new Error("พบสินค้าและ Exp เดียวกันซ้ำหลายแถว กรุณาตรวจ Stock");
          }

          var currentQty = exactRows.length === 1 ? toNumber(exactRows[0].qty) : 0;

          // สำคัญ: ป้องกันการอนุมัติข้อมูลเก่าหลัง Stock เปลี่ยน
          if (currentQty !== countedSystemQty) {
            throw new Error(
              "Stock เปลี่ยนหลังนับ (ตอนนับ " + countedSystemQty +
              " / ปัจจุบัน " + currentQty + ") ต้องนับใหม่"
            );
          }

          if (exactRows.length === 1) {
            stockSheet.getRange(exactRows[0].sheetRow, 3).setValue(actualQty);
            forceTextValue_(stockSheet, exactRows[0].sheetRow, 5, exp);
            exactRows[0].qty = actualQty; // sync local cache for next selected request
          } else if (actualQty > 0) {
            var info = getProductInfo(full);
            var newRow = stockSheet.getLastRow() + 1;
            stockSheet.getRange(newRow, 1, 1, 5).setValues([[
              info.kw, info.full, actualQty, "ชิ้น", exp
            ]]);
            forceTextValue_(stockSheet, newRow, 5, exp);
            stockRows.push({
              sheetRow: newRow,
              kw: info.kw,
              full: info.full,
              qty: actualQty,
              unit: "ชิ้น",
              exp: exp,
              expSort: parseExpForSort(exp)
            });
          }

          var diff = actualQty - currentQty;

          var adjustRow = sheets.adjustSheet.getLastRow() + 1;
          sheets.adjustSheet.getRange(adjustRow, 1, 1, 8).setValues([[
            new Date(),
            full,
            exp,
            currentQty,
            actualQty,
            diff,
            "อนุมัติแบบหลายรายการ " + requestId + ": " + requestReason,
            approver
          ]]);
          forceTextValue_(sheets.adjustSheet, adjustRow, 3, exp);

          var auditRow = sheets.auditSheet.getLastRow() + 1;
          sheets.auditSheet.getRange(auditRow, 1, 1, 11).setValues([[
            new Date(),
            requestId,
            "อนุมัติปรับยอดสต๊อก (หลายรายการ)",
            full,
            exp,
            diff,
            currentQty,
            actualQty,
            approver,
            "StockAdjustRequest",
            "ผู้ขอ: " + requester +
            " | เหตุผล: " + requestReason +
            (note ? " | หมายเหตุบอส: " + note : "")
          ]]);
          forceTextValue_(sheets.auditSheet, auditRow, 5, exp);

          requestSheet.getRange(row, 10, 1, 4).setValues([[
            "อนุมัติแล้ว",
            approver,
            new Date(),
            note || "-"
          ]]);

          results.push({
            requestId: requestId,
            success: true,
            action: "APPROVE",
            before: currentQty,
            after: actualQty,
            diff: diff
          });

        } else {
          // REJECT — ไม่เปลี่ยน Stock
          requestSheet.getRange(row, 10, 1, 4).setValues([[
            "ไม่อนุมัติ",
            approver,
            new Date(),
            note
          ]]);

          var rejectAuditRow = sheets.auditSheet.getLastRow() + 1;
          sheets.auditSheet.getRange(rejectAuditRow, 1, 1, 11).setValues([[
            new Date(),
            requestId,
            "ไม่อนุมัติปรับยอดสต๊อก (หลายรายการ)",
            full,
            exp,
            0,
            countedSystemQty,
            countedSystemQty,
            approver,
            "StockAdjustRequest",
            "ผู้ขอ: " + requester +
            " | นับจริง: " + actualQty +
            " | ต่าง: " + requestDiff +
            " | เหตุผลไม่อนุมัติ: " + note
          ]]);
          forceTextValue_(sheets.auditSheet, rejectAuditRow, 5, exp);

          results.push({
            requestId: requestId,
            success: true,
            action: "REJECT"
          });
        }

        successCount++;

      } catch (itemErr) {
        failedCount++;
        results.push({
          requestId: requestId,
          success: false,
          action: action,
          msg: itemErr.message
        });
      }
    }

    SpreadsheetApp.flush();

    var actionThai = action === "APPROVE" ? "อนุมัติ" : "ไม่อนุมัติ";
    var okIds = results.filter(function(x) { return x.success; })
      .map(function(x) { return x.requestId; });
    var failRows = results.filter(function(x) { return !x.success; });

    var pushMsg =
      (action === "APPROVE" ? "✅" : "❌") +
      " บอส" + actionThai + "คำขอปรับยอดแบบหลายรายการ\n" +
      "สำเร็จ: " + successCount + " รายการ\n" +
      "ไม่สำเร็จ: " + failedCount + " รายการ\n" +
      "ผู้พิจารณา: " + approver;

    if (okIds.length) {
      pushMsg += "\nรายการ: " + okIds.slice(0, 20).join(", ");
      if (okIds.length > 20) pushMsg += " และอีก " + (okIds.length - 20) + " รายการ";
    }
    if (failRows.length) {
      pushMsg += "\nรายการที่ต้องตรวจใหม่: " +
        failRows.slice(0, 5).map(function(x) { return x.requestId; }).join(", ");
    }

    sendPush(pushMsg);

    return {
      success: true,
      action: action,
      requested: uniqueIds.length,
      successCount: successCount,
      failedCount: failedCount,
      results: results,
      msg: actionThai + "สำเร็จ " + successCount + " / " + uniqueIds.length + " รายการ"
    };

  } catch (err) {
    Logger.log("batchProcessStockAdjustRequests V10.39 ERROR: " + err);
    return {
      success: false,
      msg: err.message,
      successCount: 0,
      failedCount: 0,
      results: []
    };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function testCreateAdjustRequestSheet() {
  var sheet = ensureAdjustRequestSheet_();
  return "✅ สร้าง/ตรวจชีต " + sheet.getName() + " สำเร็จ — " + SYSTEM_VERSION;
}

/*********************************************************
 * lil bulle Inventory V10.39 — Event Booth Module
 * - เลขเบิกออกบูทสร้างอัตโนมัติ
 * - คืนจากบูทต้องเลือกเลขเบิกเดิม
 * - สินค้า/Exp ใช้รายการจากระบบ ไม่พิมพ์ชื่อเอง
 * - คืนได้ไม่เกินยอดคงค้างของเลขเบิกนั้น
 * - Dashboard หักคืนกลับไปยังเดือนของงานเดิม
 *********************************************************/
var EVENT_ISSUE_SHEET_NAME = "EventIssue";
var EVENT_RETURN_SHEET_NAME = "EventReturn";

function ensureEventSheets_() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var issue = getOrCreateSheetWithHeader_(ss, EVENT_ISSUE_SHEET_NAME, [
    "เวลา", "เลขเบิก", "ชื่องาน/บูท", "สินค้า", "Exp", "จำนวนเบิก", "ผู้เบิก", "สถานะ"
  ]);
  var ret = getOrCreateSheetWithHeader_(ss, EVENT_RETURN_SHEET_NAME, [
    "เวลา", "เลขคืน", "อ้างอิงเลขเบิก", "ชื่องาน/บูท", "สินค้า", "Exp", "จำนวนคืน", "ผู้คืน"
  ]);
  ensureHeaderColumn_(issue, "สถานที่");
  ensureHeaderColumn_(ret, "สถานที่");
  setDateTimeColumnFormat_(issue, 1);
  setDateTimeColumnFormat_(ret, 1);
  return { issueSheet: issue, returnSheet: ret };
}



function getProductShortName_(name) {
  try {
    return getProductInfo(name).short;
  } catch (e) {
    return String(name || "");
  }
}

function makeDailyRunningNo_(sheet, col, prefix) {
  var day = Utilities.formatDate(new Date(), "GMT+7", "yyyyMMdd");
  var head = prefix + "-" + day + "-";
  var last = 0;
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var vals = sheet.getRange(2, col, lastRow - 1, 1).getDisplayValues();
    for (var i = 0; i < vals.length; i++) {
      var s = String(vals[i][0] || "").trim();
      if (s.indexOf(head) !== 0) continue;
      var n = Number(s.substring(head.length));
      if (!isNaN(n) && n > last) last = n;
    }
  }
  var seq = String(last + 1);
  while (seq.length < 3) seq = "0" + seq;
  return head + seq;
}

function appendRowsWithTextExp_(sheet, rows, expCol) {
  if (!rows || !rows.length) return;
  var start = sheet.getLastRow() + 1;
  sheet.getRange(start, 1, rows.length, rows[0].length).setValues(rows);
  if (expCol) {
    sheet.getRange(start, expCol, rows.length, 1).setNumberFormat("@");
    for (var i = 0; i < rows.length; i++) {
      sheet.getRange(start + i, expCol).setValue(String(rows[i][expCol - 1] || "-"));
    }
  }
}

function getEventOutstandingMap_() {
  var sheets = ensureEventSheets_();
  var issueData = sheets.issueSheet.getDataRange().getValues();
  var issueDisp = sheets.issueSheet.getDataRange().getDisplayValues();
  var returnData = sheets.returnSheet.getDataRange().getValues();
  var returnDisp = sheets.returnSheet.getDataRange().getDisplayValues();
  var map = {};

  for (var i = 1; i < issueData.length; i++) {
    var issueNo = String(issueData[i][1] || "").trim();
    var eventName = String(issueData[i][2] || "").trim();
    var full = String(issueData[i][3] || "").trim();
    var exp = formatExpForSheet_(issueData[i][4], issueDisp[i][4]);
    var qty = toNumber(issueData[i][5]);
    var staff = String(issueData[i][6] || "-");
    var status = String(issueData[i][7] || "เปิด").trim();
    var location = String(issueData[i][8] || "").trim();
    if (!issueNo || !full || qty <= 0) continue;

    if (!map[issueNo]) {
      map[issueNo] = {
        issueNo: issueNo,
        eventName: eventName,
        location: location || "-",
        status: status || "เปิด",
        issueDate: issueData[i][0] instanceof Date ? issueData[i][0] : new Date(issueData[i][0]),
        staff: staff,
        items: {}
      };
    } else if (status === "ปิดแล้ว") {
      map[issueNo].status = "ปิดแล้ว";
    }
    var key = normalizeData(full) + "||" + normalizeExpKey(exp);
    if (!map[issueNo].items[key]) {
      map[issueNo].items[key] = { full: full, exp: exp, issued: 0, returned: 0 };
    }
    map[issueNo].items[key].issued += qty;
  }

  for (var r = 1; r < returnData.length; r++) {
    var ref = String(returnData[r][2] || "").trim();
    var rFull = String(returnData[r][4] || "").trim();
    var rExp = formatExpForSheet_(returnData[r][5], returnDisp[r][5]);
    var rQty = toNumber(returnData[r][6]);
    if (!map[ref] || !rFull || rQty <= 0) continue;

    var rKey = normalizeData(rFull) + "||" + normalizeExpKey(rExp);
    if (!map[ref].items[rKey]) {
      map[ref].items[rKey] = { full: rFull, exp: rExp, issued: 0, returned: 0 };
    }
    map[ref].items[rKey].returned += rQty;
  }

  return map;
}

function getOpenEventIssues() {
  try {
    var map = getEventOutstandingMap_();
    var rows = [];

    Object.keys(map).forEach(function(issueNo) {
      var g = map[issueNo];
      var items = [];
      var totalIssued = 0;
      var totalReturned = 0;
      var totalPending = 0;

      Object.keys(g.items).forEach(function(key) {
        var it = g.items[key];
        var pending = Math.max(0, toNumber(it.issued) - toNumber(it.returned));
        totalIssued += toNumber(it.issued);
        totalReturned += toNumber(it.returned);
        totalPending += pending;
        if (pending > 0) {
          items.push({
            product: getProductShortName_(it.full),
            full: it.full,
            exp: it.exp,
            issued: toNumber(it.issued),
            returned: toNumber(it.returned),
            pending: pending
          });
        }
      });

      if (g.status !== "ปิดแล้ว") {
        rows.push({
          issueNo: g.issueNo,
          eventName: g.eventName,
          location: g.location || "-",
          issueDate: g.issueDate instanceof Date && !isNaN(g.issueDate.getTime())
            ? Utilities.formatDate(g.issueDate, "GMT+7", "dd/MM/yyyy HH:mm") : "-",
          staff: g.staff,
          totalIssued: totalIssued,
          totalReturned: totalReturned,
          totalPending: totalPending,
          status: g.status || "เปิด",
          items: items
        });
      }
    });

    rows.sort(function(a, b) { return String(b.issueNo).localeCompare(String(a.issueNo)); });
    return { success: true, rows: rows, version: SYSTEM_VERSION };
  } catch (err) {
    Logger.log("getOpenEventIssues V10.39 ERROR: " + err);
    return { success: false, msg: err.message, rows: [] };
  }
}


function getEventHistory() {
  try {
    var map = getEventOutstandingMap_();
    var ss = SpreadsheetApp.openById(SS_ID);
    var salesSheet = ss.getSheetByName("SalesLog");
    var finalMap = {};
    if (salesSheet && salesSheet.getLastRow() >= 2) {
      var sData = salesSheet.getDataRange().getValues();
      for (var si = 1; si < sData.length; si++) {
        var src = String(sData[si][4] || "");
        var m = src.match(/EVENT_FINAL:([^|\s]+)/);
        if (m && m[1]) finalMap[String(m[1]).trim()] = true;
      }
    }

    var rows = [];
    Object.keys(map).forEach(function(issueNo) {
      var g = map[issueNo];
      var items = [];
      var totalIssued = 0;
      var totalReturned = 0;
      var totalSold = 0;

      Object.keys(g.items || {}).forEach(function(key) {
        var it = g.items[key];
        var issued = toNumber(it.issued);
        var returned = toNumber(it.returned);
        var sold = Math.max(0, issued - returned);
        totalIssued += issued;
        totalReturned += returned;
        totalSold += sold;
        items.push({
          product: getProductShortName_(it.full),
          full: it.full,
          exp: formatExpForSheet_(it.exp || "-"),
          issued: issued,
          returned: returned,
          sold: sold
        });
      });

      rows.push({
        issueNo: g.issueNo,
        eventName: g.eventName || "-",
        location: g.location || "-",
        issueDate: g.issueDate instanceof Date && !isNaN(g.issueDate.getTime())
          ? Utilities.formatDate(g.issueDate, "GMT+7", "dd/MM/yyyy HH:mm") : "-",
        staff: g.staff || "-",
        status: g.status || "เปิด",
        totalIssued: totalIssued,
        totalReturned: totalReturned,
        totalSold: totalSold,
        finalSalesRecorded: !!finalMap[g.issueNo],
        items: items
      });
    });

    rows.sort(function(a, b) {
      return String(b.issueNo).localeCompare(String(a.issueNo));
    });
    return { success: true, rows: rows, version: SYSTEM_VERSION };
  } catch (err) {
    Logger.log("getEventHistory " + SYSTEM_VERSION + " ERROR: " + err);
    return { success: false, msg: err.message, rows: [] };
  }
}


/*********************************************************
 * V10.39 — WORKFLOW PROGRESS / SYSTEM TRACE
 * ตรวจแบบ Read-only ว่าระบบทำงานมาถึงขั้นตอนไหน
 *********************************************************/
function getWorkflowProgressV1038_(ss) {
  var stages = [];

  function addStage(step, name, status, detail, lastTime) {
    stages.push({
      step: step,
      name: name,
      status: status || "INFO",
      detail: detail || "-",
      lastTime: lastTime || "-"
    });
  }

  function latestRow_(sheetName) {
    var sh = ss.getSheetByName(sheetName);
    if (!sh || sh.getLastRow() < 2) return null;
    var row = sh.getRange(sh.getLastRow(), 1, 1, sh.getLastColumn()).getValues()[0];
    return { sheet: sh, row: row, sheetRow: sh.getLastRow() };
  }

  try {
    addStage(1, "Web App / Server", "OK", "Code V10.39 ตอบสนองและเรียกฟังก์ชันตรวจระบบได้", safeDateTimeString(new Date()));
  } catch (e1) {
    addStage(1, "Web App / Server", "ERROR", String(e1), "-");
  }

  try {
    var stock = getSheetOrThrow(ss, "Stock");
    addStage(2, "เชื่อม Google Sheets", "OK", "เปิดไฟล์และอ่าน Stock ได้ " + Math.max(0, stock.getLastRow() - 1) + " ล็อต", safeDateTimeString(new Date()));
  } catch (e2) {
    addStage(2, "เชื่อม Google Sheets", "ERROR", String(e2), "-");
  }

  try {
    var sales = latestRow_("SalesLog");
    var stockIn = latestRow_("StockIn");
    var candidates = [];
    if (sales) candidates.push({kind:"ขาย/ตัดออก", data:sales});
    if (stockIn) candidates.push({kind:"รับเข้า", data:stockIn});

    if (candidates.length) {
      candidates.sort(function(a,b) {
        var ad = a.data.row[1] instanceof Date ? a.data.row[1].getTime() : 0;
        var bd = b.data.row[1] instanceof Date ? b.data.row[1].getTime() : 0;
        return bd - ad;
      });
      var lastTxn = candidates[0];
      addStage(
        3,
        "บันทึกรายการคลัง",
        "OK",
        "ล่าสุด " + lastTxn.kind + " เอกสาร " + String(lastTxn.data.row[0] || "-"),
        safeDateTimeString(lastTxn.data.row[1])
      );
    } else {
      addStage(3, "บันทึกรายการคลัง", "INFO", "ยังไม่พบรายการใน SalesLog/StockIn", "-");
    }
  } catch (e3) {
    addStage(3, "บันทึกรายการคลัง", "ERROR", String(e3), "-");
  }

  try {
    var count = latestRow_("StockCount");
    if (count) {
      addStage(
        4,
        "ทีมงานตรวจนับจริง",
        "OK",
        "ล่าสุด " + String(count.row[1] || "-") + " Exp " + formatExpForSheet_(count.row[2] || "-") +
        " | ระบบ " + toNumber(count.row[3]) + " / นับจริง " + toNumber(count.row[4]),
        safeDateTimeString(count.row[0])
      );
    } else {
      addStage(4, "ทีมงานตรวจนับจริง", "INFO", "ยังไม่พบประวัติ StockCount", "-");
    }
  } catch (e4) {
    addStage(4, "ทีมงานตรวจนับจริง", "ERROR", String(e4), "-");
  }

  try {
    var reqSheet = ss.getSheetByName("StockAdjustRequest");
    var pending = 0, approved = 0, rejected = 0, latestReq = null;
    if (reqSheet && reqSheet.getLastRow() >= 2) {
      var req = reqSheet.getDataRange().getValues();
      for (var i=1; i<req.length; i++) {
        var st = String(req[i][9] || "").trim();
        if (st === "รออนุมัติ") pending++;
        else if (st === "อนุมัติแล้ว") approved++;
        else if (st === "ไม่อนุมัติ") rejected++;
        latestReq = req[i];
      }
    }
    addStage(
      5,
      "คำขอปรับยอด / บอสอนุมัติ",
      pending ? "INFO" : "OK",
      "รออนุมัติ " + pending + " | อนุมัติแล้ว " + approved + " | ปฏิเสธ " + rejected,
      latestReq ? safeDateTimeString(latestReq[0]) : "-"
    );
  } catch (e5) {
    addStage(5, "คำขอปรับยอด / บอสอนุมัติ", "ERROR", String(e5), "-");
  }

  try {
    var audit = latestRow_("AuditLog");
    if (audit) {
      addStage(
        6,
        "AuditLog",
        "OK",
        "ล่าสุด " + String(audit.row[2] || "-") + " | " + String(audit.row[1] || "-") +
        " | " + String(audit.row[3] || "-"),
        safeDateTimeString(audit.row[0])
      );
    } else {
      addStage(6, "AuditLog", "WARN", "ยังไม่มี AuditLog", "-");
    }
  } catch (e6) {
    addStage(6, "AuditLog", "ERROR", String(e6), "-");
  }

  try {
    var issue = latestRow_("EventIssue");
    var ret = latestRow_("EventReturn");
    var eventDetail = [];
    var eventTime = "-";
    if (issue) {
      eventDetail.push("เบิกล่าสุด " + String(issue.row[1] || "-"));
      eventTime = safeDateTimeString(issue.row[0]);
    }
    if (ret) {
      eventDetail.push("คืนล่าสุดอ้างอิง " + String(ret.row[2] || "-"));
      eventTime = safeDateTimeString(ret.row[0]);
    }
    addStage(7, "Event Booth", "OK", eventDetail.length ? eventDetail.join(" | ") : "พร้อมใช้งาน ยังไม่มีรายการ Event", eventTime);
  } catch (e7) {
    addStage(7, "Event Booth", "ERROR", String(e7), "-");
  }

  try {
    var stockSheet = getSheetOrThrow(ss, "Stock");
    var salesSheet = getSheetOrThrow(ss, "SalesLog");
    addStage(
      8,
      "Boss Dashboard",
      "OK",
      "อ่าน Stock " + Math.max(0, stockSheet.getLastRow()-1) + " ล็อต และ SalesLog " + Math.max(0, salesSheet.getLastRow()-1) + " รายการได้",
      safeDateTimeString(new Date())
    );
  } catch (e8) {
    addStage(8, "Boss Dashboard", "ERROR", String(e8), "-");
  }

  addStage(
    9,
    "PDF / Print",
    "INFO",
    "Server เตรียมข้อมูลเอกสารได้ แต่การเปิดหน้าพิมพ์/Save as PDF ต้องทดสอบที่ Browser เพราะ Pop-up permission ตรวจจาก Server ไม่ได้",
    "-"
  );

  var errorCount = stages.filter(function(x){ return x.status === "ERROR"; }).length;
  var warnCount = stages.filter(function(x){ return x.status === "WARN"; }).length;
  var readyCount = stages.length - errorCount;
  var percent = stages.length ? Math.round((readyCount / stages.length) * 100) : 0;

  var firstProblem = null;
  for (var s=0; s<stages.length; s++) {
    if (stages[s].status === "ERROR") {
      firstProblem = stages[s];
      break;
    }
  }

  return {
    stages: stages,
    total: stages.length,
    ready: readyCount,
    errorCount: errorCount,
    warnCount: warnCount,
    percent: percent,
    summary: firstProblem
      ? "พบจุดหยุดที่ขั้น " + firstProblem.step + ": " + firstProblem.name
      : "ระบบผ่านจุดตรวจฝั่ง Server/Sheet ครบ ไม่มี ERROR"
  };
}

function getPendingApprovalHealthV1038_(ss) {
  var result = { pending:0, ready:0, stale:0, duplicate:0 };
  try {
    var requestSheet = ss.getSheetByName("StockAdjustRequest");
    var stockSheet = ss.getSheetByName("Stock");
    if (!requestSheet || !stockSheet || requestSheet.getLastRow() < 2) return result;

    var values = requestSheet.getDataRange().getValues();
    var displays = requestSheet.getDataRange().getDisplayValues();
    var stockRows = getStockRows(stockSheet);
    var approvedMap = {};

    for (var a=1; a<values.length; a++) {
      if (String(values[a][9] || "").trim() !== "อนุมัติแล้ว") continue;
      var key =
        normalizeData(values[a][2]) + "|" +
        normalizeExpKey(formatExpForSheet_(values[a][3], displays[a][3])) + "|" +
        Number(values[a][4] || 0) + "|" +
        Number(values[a][5] || 0) + "|" +
        Number(values[a][6] || 0);
      approvedMap[key] = true;
    }

    for (var i=1; i<values.length; i++) {
      if (String(values[i][9] || "").trim() !== "รออนุมัติ") continue;
      result.pending++;

      var product = String(values[i][2] || "");
      var exp = formatExpForSheet_(values[i][3], displays[i][3]);
      var systemQty = Number(values[i][4] || 0);
      var actualQty = Number(values[i][5] || 0);
      var diff = Number(values[i][6] || 0);
      var currentQty = getSystemQtyByFullExp_(stockRows, product, exp);
      var sig =
        normalizeData(product) + "|" +
        normalizeExpKey(exp) + "|" +
        systemQty + "|" + actualQty + "|" + diff;

      if (approvedMap[sig]) result.duplicate++;
      else if (currentQty !== systemQty) result.stale++;
      else result.ready++;
    }
  } catch (e) {}
  return result;
}

function getOperationalHealthChecks_() {
  var checks = [];
  try {
    var props = PropertiesService.getScriptProperties();
    checks.push({
      check: "Boss PIN",
      status: props.getProperty("BOSS_PIN") ? "OK" : "ERROR",
      detail: props.getProperty("BOSS_PIN") ? "ตั้งค่า BOSS_PIN แล้ว" : "ยังไม่ได้ตั้ง BOSS_PIN ใน Script Properties"
    });

    var ss = SpreadsheetApp.openById(SS_ID);
    var required = ["Stock","SalesLog","StockIn","AuditLog","StockCount","StockAdjust","StockAdjustRequest","EventIssue","EventReturn"];
    var missing = required.filter(function(n){ return !ss.getSheetByName(n); });
    checks.push({
      check: "ชีตระบบ",
      status: missing.length ? "ERROR" : "OK",
      detail: missing.length ? "ไม่พบ: " + missing.join(", ") : "พบชีตหลักครบ " + required.length + " ชีต"
    });

    // V10.39: ตรวจ dependency ที่ Dashboard ใช้จริงแบบ read-only
    try {
      var dashStock = getSheetOrThrow(ss, "Stock");
      var dashSales = getSheetOrThrow(ss, "SalesLog");
      var stockRowsHealth = dashStock.getLastRow();
      var salesRowsHealth = dashSales.getLastRow();
      checks.push({
        check: "Boss Dashboard",
        status: "OK",
        detail: "อ่าน Stock และ SalesLog ได้ (Stock " + stockRowsHealth + " แถว / SalesLog " + salesRowsHealth + " แถว)"
      });
    } catch (dashErr) {
      checks.push({check:"Boss Dashboard", status:"ERROR", detail:"Dashboard dependency ผิดพลาด: " + String(dashErr)});
    }

    // ตรวจ EventReturn ที่อ้างอิงเลขเบิกที่ไม่มีใน EventIssue
    var es = ensureEventSheets_();
    var issueVals = es.issueSheet.getDataRange().getValues();
    var returnVals = es.returnSheet.getDataRange().getValues();
    var issueSet = {};
    for (var i=1;i<issueVals.length;i++) {
      var ino = String(issueVals[i][1] || "").trim();
      if (ino) issueSet[ino] = true;
    }
    var orphan = 0;
    for (var r=1;r<returnVals.length;r++) {
      var ref = String(returnVals[r][2] || "").trim();
      if (ref && !issueSet[ref]) orphan++;
    }
    checks.push({
      check: "Event อ้างอิงใบเบิก",
      status: orphan ? "ERROR" : "OK",
      detail: orphan ? "พบ EventReturn อ้างอิงเลขเบิกที่ไม่มี " + orphan + " รายการ" : "EventReturn อ้างอิง EventIssue ถูกต้อง"
    });

    var hist = getEventHistory();
    var closedNoFinal = 0;
    var openCount = 0;
    if (hist && hist.success) {
      (hist.rows || []).forEach(function(x){
        if (x.status === "ปิดแล้ว" && !x.finalSalesRecorded) closedNoFinal++;
        if (x.status !== "ปิดแล้ว") openCount++;
      });
    }
    checks.push({
      check: "Event Settlement",
      status: closedNoFinal ? "WARN" : "OK",
      detail: closedNoFinal ? "งานปิดแล้วแต่ยังไม่พบ EVENT_FINAL " + closedNoFinal + " งาน" : "งานปิดมีบันทึกยอดขายจริงครบ"
    });
    checks.push({
      check: "งาน Event ที่ยังเปิด",
      status: openCount ? "INFO" : "OK",
      detail: openCount + " งาน"
    });

    var approvalHealth = getPendingApprovalHealthV1038_(ss);
    checks.push({
      check: "คำขอปรับยอดรออนุมัติ",
      status: approvalHealth.stale || approvalHealth.duplicate ? "WARN" : (approvalHealth.pending ? "INFO" : "OK"),
      detail:
        "รอ " + approvalHealth.pending +
        " | พร้อมอนุมัติ " + approvalHealth.ready +
        " | Stock เปลี่ยนหลังนับ " + approvalHealth.stale +
        " | คำขอซ้ำที่เคยอนุมัติ " + approvalHealth.duplicate
    });

    checks.push({
      check: "บันทึกรายการคลัง",
      status: "OK",
      detail: "saveMultiple พร้อม SalesLog / StockIn / AuditLog"
    });
    checks.push({
      check: "เอกสาร Event / PDF",
      status: "OK",
      detail: "ใบเบิกและใบสรุปอ่านย้อนหลังจากประวัติงานได้"
    });
  } catch (err) {
    checks.push({check:"Operational Health", status:"ERROR", detail:String(err)});
  }
  return checks;
}


/** V11.2: เบิกสินค้าออกบูทตามล็อตที่เลือกเท่านั้น (no cross-EXP fallback). */
function planExactEventLot_(stockRows, full, qty, exp) {
  qty = Number(qty);
  if (!Number.isSafeInteger(qty) || qty <= 0) throw new Error('จำนวนเบิกต้องเป็นจำนวนเต็มบวก');
  exp = formatExpForSheet_(exp);
  if (!exp || exp === '-') throw new Error('กรุณาเลือก EXP');
  var rows = stockRows.filter(function(r) {
    return normalizeData(r.full) === normalizeData(full) &&
      normalizeExpKey(r.exp) === normalizeExpKey(exp) && Number(r.qty) > 0;
  });
  var available = rows.reduce(function(sum, r) { return sum + Number(r.qty); }, 0);
  if (qty > available) throw new Error(full + ' EXP ' + exp + ' มี ' + available + ' ชิ้น (ขอเบิก ' + qty + ') กรุณาเลือกลॉटอื่นเป็นรายการใหม่');
  var remain = qty, plan = [];
  rows.forEach(function(r) {
    if (!remain) return;
    var cut = Math.min(remain, Number(r.qty)), before = Number(r.qty);
    r.qty = before - cut;
    plan.push({ row:r, cut:cut, before:before, after:r.qty });
    remain -= cut;
  });
  return plan;
}

/** อ่านตรวจยอดแต่ละล็อต (ไม่แก้ Sheet/Stock) ใช้ตรวจบูทที่พบปัญหา */
function inspectEventLotsV112(issueNo) {
  try {
    issueNo = String(issueNo || '').trim();
    if (!issueNo) throw new Error('ระบุเลขใบเบิกก่อน เช่น inspectEventLotsV112("EVT-20261022-005")');
    var g = getEventOutstandingMap_()[issueNo];
    if (!g) throw new Error('ไม่พบเลขเบิก ' + issueNo);
    var rows = Object.keys(g.items).map(function(key) {
      var it = g.items[key], reserved = 0;
      try {
        var info = getProductInfo(it.full);
        if (typeof lbUnreturnableLot_ === 'function') reserved = Number(lbUnreturnableLot_(issueNo,info.kw,it.exp)) || 0;
        else if (typeof lbSoldLot_ === 'function') reserved = Number(lbSoldLot_(issueNo,info.kw,it.exp)) || 0;
      } catch(e) { throw new Error('ตรวจยอด POS ไม่สำเร็จ: '+e.message); }
      return {product:it.full,exp:it.exp,issued:it.issued,returned:it.returned,
        soldOrReserved:reserved,returnable:it.issued-it.returned-reserved};
    });
    return {success:true,issueNo:issueNo,status:g.status,rows:rows};
  } catch(e) { return {success:false,msg:String(e.message||e)}; }
}

function saveEventIssue(data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    if (!data) throw new Error("ไม่พบข้อมูลเบิกออกบูท");

    var eventName = String(data.eventName || "").trim();
    var location = String(data.location || "").trim();
    var staff = String(data.staff || "").trim();
    var items = data.items || [];
    if (!eventName) throw new Error("กรุณาระบุชื่องาน/บูท");
    if (!location) throw new Error("กรุณาระบุสถานที่");
    if (!staff) throw new Error("กรุณาระบุผู้เบิก");
    if (!items.length) throw new Error("กรุณาเลือกสินค้าอย่างน้อย 1 รายการ");

    var ss = SpreadsheetApp.openById(SS_ID);
    var stockSheet = getSheetOrThrow(ss, "Stock");
    var controlSheets = ensureControlSheets_();
    var eventSheets = ensureEventSheets_();
    var stockRows = getStockRows(stockSheet);
    var issueNo = makeDailyRunningNo_(eventSheets.issueSheet, 2, "EVT");

    var issueRows = [];
    var auditLogs = [];
    var notifyItems = [];

    for (var i = 0; i < items.length; i++) {
      var item = items[i];
      var info = getProductInfo(item.product);
      var qty = Number(item.qty);
      var exp = formatExpForSheet_(item.exp || "-");
      if (!Number.isSafeInteger(qty) || qty <= 0) throw new Error("จำนวนสินค้าไม่ถูกต้อง: " + info.short);
      if (!exp || exp === "-") throw new Error("กรุณาเลือก Exp ของ " + info.short);

      // Event ต้องตัดเฉพาะ EXP ที่ทีมงานเลือก ห้ามยืมยอดล็อตอื่นโดยอัตโนมัติ
      var plan = planExactEventLot_(stockRows, info.full, qty, exp);
      for (var p = 0; p < plan.length; p++) {
        var seg = plan[p];
        var actualExp = formatExpForSheet_(seg.row.exp || exp);
        var source = "Event | EVENT_ISSUE:" + issueNo + " | งาน:" + eventName + " | สถานที่:" + location;

        issueRows.push([new Date(), issueNo, eventName, info.full, actualExp, seg.cut, staff, "เปิด", location]);
        auditLogs.push([new Date(), issueNo, "เบิกออกบูท", info.full, actualExp, -seg.cut, seg.before, seg.after, staff, "Event | สถานที่:" + location, eventName]);
        notifyItems.push(info.short + " " + seg.cut + " ชิ้น Exp " + actualExp);
      }
    }

    applyStockRowsToSheet(stockSheet, stockRows);
    appendRowsWithTextExp_(eventSheets.issueSheet, issueRows, 5);
    appendRowsWithExpiryText_(controlSheets.auditSheet, auditLogs, 5);
    SpreadsheetApp.flush();

    sendPush(
      "🎪 เบิกออกบูท\n" +
      "เลขเบิก: " + issueNo + "\n" +
      "งาน: " + eventName + "\n" +
      "สถานที่: " + location + "\n" +
      "ผู้เบิก: " + staff + "\n" +
      notifyItems.map(function(x) { return "• " + x; }).join("\n")
    );

    return { success: true, msg: "เบิกออกบูทสำเร็จ", issueNo: issueNo, eventName: eventName, location: location };
  } catch (err) {
    Logger.log("saveEventIssue V10.39 ERROR: " + err);
    return { success: false, msg: err.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}


function getEventSettlementData_(issueNo) {
  var map = getEventOutstandingMap_();
  var g = map[String(issueNo || '').trim()];
  if (!g) throw new Error('ไม่พบเลขเบิก ' + issueNo);
  var lines = [];
  Object.keys(g.items || {}).forEach(function(k) {
    var it = g.items[k];
    lines.push({
      product: it.full,
      exp: formatExpForSheet_(it.exp || '-'),
      issued: toNumber(it.issued),
      returned: toNumber(it.returned),
      sold: Math.max(0, toNumber(it.issued) - toNumber(it.returned))
    });
  });
  return {issueNo:issueNo,eventName:g.eventName,location:g.location||'-',staff:g.staff||'-',lines:lines};
}

function writeEventFinalSales_(issueNo, closingStaff) {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sales = getSheetOrThrow(ss, 'SalesLog');
  var data = sales.getDataRange().getValues();
  var tag = 'EVENT_FINAL:' + issueNo;
  for (var i=1;i<data.length;i++) {
    if (String(data[i][4]||'').indexOf(tag) > -1) return getEventSettlementData_(issueNo);
  }
  var doc = getEventSettlementData_(issueNo);
  var rows=[];
  doc.lines.forEach(function(x){
    if (x.sold <= 0) return;
    var source='Event | '+tag+' | งาน:'+doc.eventName+' | สถานที่:'+doc.location;
    rows.push([issueNo,new Date(),x.product,x.sold,source,'','',x.exp,closingStaff]);
  });
  if (rows.length) appendRowsWithExpiryText_(sales, rows, 8);
  return doc;
}

function getEventDocumentData(issueNo) {
  try {
    var doc=getEventSettlementData_(issueNo);
    return {success:true,version:SYSTEM_VERSION,meta:{issueNo:doc.issueNo,eventName:doc.eventName,location:doc.location,staff:doc.staff},lines:doc.lines};
  } catch(err) { return {success:false,msg:err.message}; }
}

function saveEventReturn(data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    if (!data) throw new Error("ให้ทำรายการคืนจากหน้าเว็บและเลือกเลขเบิกก่อน (ไม่ใช่ข้อผิดพลาดสต๊อก)");

    var issueNo = String(data.issueNo || "").trim();
    var staff = String(data.staff || "").trim();
    var closeEvent = data.closeEvent === true;
    var items = data.items || [];
    if (!issueNo) throw new Error("กรุณาเลือกเลขเบิกงาน");
    if (!staff) throw new Error("กรุณาระบุผู้คืน");
    if (!items.length && !closeEvent) throw new Error("กรุณาระบุจำนวนที่คืน หรือเลือกปิดงาน");

    var outstandingMap = getEventOutstandingMap_();
    var eventGroup = outstandingMap[issueNo];
    if (!eventGroup) throw new Error("ไม่พบเลขเบิก " + issueNo);
    if (eventGroup.status === "ปิดแล้ว") throw new Error("ใบเบิก " + issueNo + " ปิดงานแล้ว ไม่สามารถคืนซ้ำได้");

    var ss = SpreadsheetApp.openById(SS_ID);
    var stockSheet = getSheetOrThrow(ss, "Stock");
    var controlSheets = ensureControlSheets_();
    var eventSheets = ensureEventSheets_();
    var stockRows = getStockRows(stockSheet);
    var returnNo = makeDailyRunningNo_(eventSheets.returnSheet, 2, "EVR");

    var returnRows = [];
    var auditLogs = [];
    var notifyItems = [];
    var returnInRequest = {};

    for (var i = 0; i < items.length; i++) {
      var item = items[i];
      var info = getProductInfo(item.product);
      var exp = formatExpForSheet_(item.exp || "-");
      var qty = Number(item.qty);
      if (item.qty === '' || item.qty === null || item.qty === undefined || qty === 0) continue;
      if (!Number.isSafeInteger(qty) || qty < 0) throw new Error('จำนวนคืนของ ' + info.short + ' ต้องเป็นจำนวนเต็มไม่ติดลบ');

      var key = normalizeData(info.full) + "||" + normalizeExpKey(exp);
      var sourceItem = eventGroup.items[key];
      if (!sourceItem) throw new Error(info.short + " Exp " + exp + " ไม่ได้อยู่ในเลขเบิกนี้");

      var pending = toNumber(sourceItem.issued) - toNumber(sourceItem.returned);
      pending -= returnInRequest[key] || 0;
      if (pending < 0) throw new Error('ตรวจพบยอดล็อต '+info.short+' EXP '+exp+' ติดลบ กรุณาตรวจรายการเบิก คืน และ POS ก่อนบันทึก');
      if (qty > pending) {
        throw new Error(info.short + " EXP " + exp + " คืนได้สูงสุด " + pending + " ชิ้น (เบิกสะสม " + sourceItem.issued + " / คืนแล้ว " + sourceItem.returned + ") กรุณาตรวจล็อตจริงก่อนคืน");
      }

      returnInRequest[key] = (returnInRequest[key] || 0) + qty;
      var retInfo = addStockSimulate(stockRows, info.kw, info.full, qty, exp);
      var actualExp = formatExpForSheet_(retInfo.exp || exp);
      var source = "Event | คืนจากบูท | EVENT_RETURN:" + issueNo + " | งาน:" + eventGroup.eventName + " | สถานที่:" + (eventGroup.location || "-");

      returnRows.push([new Date(), returnNo, issueNo, eventGroup.eventName, info.full, actualExp, qty, staff, eventGroup.location || "-"]);
      auditLogs.push([new Date(), returnNo, "คืนจากบูท", info.full, actualExp, qty, retInfo.before, retInfo.after, staff, "อ้างอิง " + issueNo, eventGroup.eventName]);
      notifyItems.push(info.short + " " + qty + " ชิ้น Exp " + actualExp);
    }

    if (!returnRows.length && !closeEvent) throw new Error("กรุณาใส่จำนวนคืนอย่างน้อย 1 รายการ");

    if (returnRows.length) applyStockRowsToSheet(stockSheet, stockRows);
    if (returnRows.length) appendRowsWithTextExp_(eventSheets.returnSheet, returnRows, 6);
    if (auditLogs.length) appendRowsWithExpiryText_(controlSheets.auditSheet, auditLogs, 5);
    SpreadsheetApp.flush();

    var settlement = null;
    if (closeEvent) settlement = writeEventFinalSales_(issueNo, staff);
    refreshEventIssueStatus_(issueNo, closeEvent);

    sendPush(
      "↩️ คืนจากบูท\n" +
      "เลขคืน: " + returnNo + "\n" +
      "อ้างอิงเลขเบิก: " + issueNo + "\n" +
      "งาน: " + eventGroup.eventName + "\n" +
      "สถานที่: " + (eventGroup.location || "-") + "\n" +
      "ผู้คืน: " + staff + "\n" +
      notifyItems.map(function(x) { return "• " + x; }).join("\n")
    );

    return {
      success: true,
      msg: "คืนจากบูทสำเร็จ",
      returnNo: returnNo,
      issueNo: issueNo,
      eventName: eventGroup.eventName,
      closed: closeEvent,
      settlement: settlement
    };
  } catch (err) {
    Logger.log("saveEventReturn V10.39 ERROR: " + err);
    return { success: false, msg: err.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function refreshEventIssueStatus_(issueNo, forceClose) {
  var sheets=ensureEventSheets_();
  var rows=sheets.issueSheet.getDataRange().getValues();
  var found=false,alreadyClosed=false;
  for(var i=1;i<rows.length;i++) if(String(rows[i][1]||'').trim()===String(issueNo||'').trim()) {
    found=true;
    if(String(rows[i][7]||'').trim()==='ปิดแล้ว')alreadyClosed=true;
  }
  if(!found)throw new Error('ไม่พบเลขเบิก '+issueNo);
  var status=(forceClose||alreadyClosed)?'ปิดแล้ว':'เปิด';
  for(var j=1;j<rows.length;j++) if(String(rows[j][1]||'').trim()===String(issueNo||'').trim()) {
    sheets.issueSheet.getRange(j+1,8).setValue(status);
  }
  return status;
}

function extractEventIssueNo_(source) {
  var m = String(source || "").match(/EVENT_RETURN:([^|\s]+)/);
  return m ? String(m[1] || "").trim() : "";
}

function getEventIssueDateMap_() {
  try {
    var sheets = ensureEventSheets_();
    var data = sheets.issueSheet.getDataRange().getValues();
    var map = {};
    for (var i = 1; i < data.length; i++) {
      var issueNo = String(data[i][1] || "").trim();
      var dt = data[i][0] instanceof Date ? data[i][0] : new Date(data[i][0]);
      if (issueNo && !map[issueNo] && !isNaN(dt.getTime())) map[issueNo] = dt;
    }
    return map;
  } catch (e) {
    Logger.log("getEventIssueDateMap_ ERROR: " + e);
    return {};
  }
}

function testEventBoothSetup() {
  ensureEventSheets_();
  return "✅ Event Booth V10.39 พร้อมใช้งาน";
}



/*********************************************************
 * V10.39 PDF ALL DOCUMENTS
 * - Warehouse document lookup / reprint
 * - Staff can print latest or historical STK documents
 * - Event PDFs remain available from Event History
 * - Stock count printable receipt data
 *********************************************************/
function classifyWarehouseTypeV1037_(qty, source, inbound) {
  source = String(source || '');
  if (inbound) return 'รับเข้าคลัง';
  if (qty < 0) return source.indexOf('EVENT_RETURN:') > -1 ? 'คืนจากบูท' : 'รับคืน/ลงผิด';
  if (source.indexOf('หมดอายุ') > -1) return 'หมดอายุ/เสียหาย';
  if (source.indexOf('สนับสนุน') > -1 || source.indexOf('แจกฟรี') > -1 || source.indexOf('PR') > -1) return 'สนับสนุน/แจกฟรี';
  return 'ขายออก';
}

function getWarehouseDocumentData(docNo) {
  try {
    docNo = String(docNo || '').trim();
    if (!docNo) throw new Error('กรุณาระบุเลขเอกสาร');
    var ss = SpreadsheetApp.openById(SS_ID);
    var sales = ss.getSheetByName('SalesLog');
    var inbound = ss.getSheetByName('StockIn');
    var lines = [];
    var meta = {docNo:docNo,date:'-',staff:'-',source:'-'};

    if (sales && sales.getLastRow() >= 2) {
      var s = sales.getDataRange().getValues();
      for (var i=1;i<s.length;i++) {
        if (String(s[i][0]||'').trim() !== docNo) continue;
        if (meta.date === '-') meta.date = safeDateTimeString(s[i][1]);
        if (meta.staff === '-') meta.staff = String(s[i][8]||'-');
        var src = String(s[i][4]||'-');
        if (meta.source === '-') meta.source = src;
        var q = toNumber(s[i][3]);
        lines.push({
          type: classifyWarehouseTypeV1037_(q,src,false),
          product:String(s[i][2]||'-'), qty:Math.abs(q), exp:formatExpForSheet_(s[i][7]||'-'), source:src,
          before:toNumber(s[i][5]), after:toNumber(s[i][6])
        });
      }
    }

    if (inbound && inbound.getLastRow() >= 2) {
      var b = inbound.getDataRange().getValues();
      for (var j=1;j<b.length;j++) {
        if (String(b[j][0]||'').trim() !== docNo) continue;
        if (meta.date === '-') meta.date = safeDateTimeString(b[j][1]);
        if (meta.staff === '-') meta.staff = String(b[j][5]||'-');
        var src2 = String(b[j][6]||'-');
        if (meta.source === '-') meta.source = src2;
        lines.push({type:'รับเข้าคลัง',product:String(b[j][2]||'-'),qty:Math.abs(toNumber(b[j][3])),exp:formatExpForSheet_(b[j][4]||'-'),source:src2,before:null,after:null});
      }
    }

    if (!lines.length) throw new Error('ไม่พบเลขเอกสาร '+docNo);
    return {success:true,version:SYSTEM_VERSION,meta:meta,lines:lines};
  } catch(err) {
    return {success:false,msg:err.message};
  }
}

function getRecentWarehouseDocumentsV1037(limit) {
  try {
    limit = Math.max(1, Math.min(Number(limit||20), 50));
    var ss=SpreadsheetApp.openById(SS_ID), map={};
    var sales=ss.getSheetByName('SalesLog'), inbound=ss.getSheetByName('StockIn');
    function addRows(sh,isIn){
      if(!sh || sh.getLastRow()<2) return;
      var d=sh.getDataRange().getValues();
      for(var i=1;i<d.length;i++){
        var doc=String(d[i][0]||'').trim(); if(!doc || doc.indexOf('STK-')!==0) continue;
        var dt=d[i][1] instanceof Date?d[i][1]:new Date(d[i][1]); var ts=isNaN(dt.getTime())?0:dt.getTime();
        if(!map[doc]) map[doc]={docNo:doc,ts:ts,date:safeDateTimeString(d[i][1]),staff:String(d[i][isIn?5:8]||'-'),count:0};
        map[doc].count++;
      }
    }
    addRows(sales,false); addRows(inbound,true);
    var rows=Object.keys(map).map(function(k){return map[k];}).sort(function(a,b){return b.ts-a.ts;}).slice(0,limit);
    return {success:true,rows:rows};
  }catch(err){return {success:false,msg:err.message,rows:[]};}
}

function getStockCountPrintDataV1037(payload, result) {
  payload = payload || {}; result = result || {};
  try {
    var info=getProductInfo(payload.product);
    return {success:true,meta:{docNo:String(result.requestId||'COUNT-'+Utilities.formatDate(new Date(),'GMT+7','yyyyMMdd-HHmmss')),date:safeDateTimeString(new Date()),staff:String(payload.inspector||'-'),title:'ใบตรวจนับสต๊อก'},lines:[{type:'ตรวจนับจริง',product:info.full,qty:toNumber(result.actualQty),exp:formatExpForSheet_(payload.exp||'-'),source:'ในระบบ '+toNumber(result.systemQty)+' / ต่าง '+toNumber(result.diff)+(result.requestId?' / เลขคำขอ '+result.requestId:'')} ]};
  } catch(err) { return {success:false,msg:err.message}; }
}


/*********************************************************
 * V10.39 — Consolidated Inventory & Non-Sale Stock
 * Non-sale inventory is deliberately separate from Stock/SalesLog.
 *********************************************************/
var NONSALE_MASTER_V1039 = 'InventoryNonSale';
var NONSALE_LOG_V1039 = 'InventoryNonSaleLog';

function getNonSaleInventoryV1039() {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var sh = getSheetOrThrow(ss, NONSALE_MASTER_V1039);
    var v = sh.getDataRange().getValues();
    var display = sh.getDataRange().getDisplayValues();
    var rows = [];
    for (var i=1;i<v.length;i++) {
      var id = String(v[i][0] || '').trim();
      if (!id) continue;
      var qtyCell = v[i][5];
      var unknown = qtyCell === '' || qtyCell === null || typeof qtyCell === 'undefined';
      rows.push({
        id:id, category:String(v[i][1] || ''), name:String(v[i][2] || ''),
        detail:String(v[i][3] || ''), exp:String(display[i][4] || '').trim(),
        qty:unknown ? null : Number(qtyCell), unit:String(v[i][6] || 'ชิ้น'),
        status:String(v[i][7] || ''), note:String(v[i][8] || '')
      });
    }
    return {success:true, rows:rows};
  } catch(e) { return {success:false,msg:e.message,rows:[]}; }
}

function saveNonSaleMovementV1039(payload) {
  var lock = LockService.getScriptLock();
  try {
    payload = payload || {};
    var id = String(payload.id || '').trim();
    var kind = String(payload.kind || '').trim().toUpperCase();
    var staff = String(payload.staff || '').trim();
    var note = String(payload.note || '').trim();
    var qty = Number(payload.qty);
    if (!id || !staff) throw new Error('กรุณาระบุรายการและผู้ทำรายการ');
    if (kind !== 'IN' && kind !== 'OUT') throw new Error('เลือกประเภทรับเข้า/เบิกออก');
    if (!Number.isSafeInteger(qty) || qty <= 0 || qty > 1000000) throw new Error('จำนวนต้องเป็นจำนวนเต็มบวก');
    lock.waitLock(30000);
    var ss = SpreadsheetApp.openById(SS_ID);
    var sh = getSheetOrThrow(ss, NONSALE_MASTER_V1039);
    var log = getSheetOrThrow(ss, NONSALE_LOG_V1039);
    var values = sh.getDataRange().getValues();
    var target = -1;
    for (var i=1;i<values.length;i++) {
      if (String(values[i][0] || '').trim() === id) {
        if (target !== -1) throw new Error('พบรหัสซ้ำในทะเบียน กรุณาตรวจสอบ');
        target = i;
      }
    }
    if (target < 0) throw new Error('ไม่พบรหัสรายการในทะเบียน');
    var current = values[target][5];
    if (current === '' || current === null) throw new Error('รายการนี้ยังไม่ระบุยอดตั้งต้น กรุณาให้บอสยืนยันก่อน');
    var before = Number(current);
    if (!Number.isFinite(before) || before < 0) throw new Error('ยอดสต๊อกเดิมไม่ถูกต้อง');
    var after = before + (kind === 'IN' ? qty : -qty);
    if (after < 0) throw new Error('ยอดไม่พอสำหรับเบิกออก (คงเหลือ '+before+')');
    var docNo = 'NS-' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyyMMdd-HHmmss') + '-' + Utilities.getUuid().slice(0,6).toUpperCase();
    var delta = kind === 'IN' ? qty : -qty;
    // Append audit first. If stock write fails, the log still records the attempted operation for reconciliation.
    log.appendRow([new Date(),docNo,id,String(values[target][2] || ''),kind,delta,before,after,staff,note]);
    sh.getRange(target+1,6).setValue(after);
    sh.getRange(target+1,8).setValue('อัปเดตแล้ว');
    SpreadsheetApp.flush();
    return {success:true,docNo:docNo,before:before,after:after};
  } catch(e) { return {success:false,msg:e.message}; }
  finally {try{lock.releaseLock();}catch(ignore){}}
}

function getCompleteInventoryReportV1039(pin) {
  try {
    verifyBossPin_(pin);
    var ss = SpreadsheetApp.openById(SS_ID);
    var stock = getStockRows(getSheetOrThrow(ss, 'Stock'));
    var nonSaleResult = getNonSaleInventoryV1039();
    if (!nonSaleResult.success) throw new Error(nonSaleResult.msg);
    var now = new Date();
    var today = new Date(now.getFullYear(),now.getMonth(),now.getDate()).getTime();
    var dayMs = 86400000;
    var rows = [], productTotals = {}, categoryCounts = {}, expiring = 0, expired = 0, low = 0, unknown = 0;
    function add(item) {
      var qty = item.qty;
      var exp = String(item.exp || '').trim();
      var expTs = exp ? parseExpForSort(exp) : null;
      var expDays = (expTs !== null && expTs < new Date(9998,0,1).getTime()) ? Math.round((expTs - today) / dayMs) : null;
      var expiryStatus = expDays === null ? 'ไม่ระบุ EXP' : (expDays < 0 ? 'หมดอายุแล้ว' : (expDays <= 90 ? 'ใกล้หมดอายุ' : 'ปกติ'));
      if (expiryStatus === 'หมดอายุแล้ว' && qty > 0) expired++;
      if (expiryStatus === 'ใกล้หมดอายุ' && qty > 0) expiring++;
      if (qty === null) unknown++;
      if (qty !== null && qty <= 50) low++;
      if (!categoryCounts[item.category]) categoryCounts[item.category] = {lines:0,known:0,unknown:0};
      categoryCounts[item.category].lines++;
      if (qty === null) categoryCounts[item.category].unknown++;
      else categoryCounts[item.category].known++;
      if (item.category === 'สินค้าพร้อมขาย') {
        productTotals[item.name] = (productTotals[item.name] || 0) + (qty || 0);
      }
      rows.push({id:item.id,category:item.category,name:item.name,detail:item.detail || '',qty:qty,unit:item.unit || 'ชิ้น',exp:exp || '-',expiryStatus:expiryStatus,daysToExpiry:expDays,stockStatus:qty === null ? 'ไม่ทราบจำนวน' : (qty === 0 ? 'หมดสต๊อก' : (qty <= 50 ? 'ใกล้หมด' : 'ปกติ'))});
    }
    stock.forEach(function(x){add({id:'Stock!'+x.sheetRow,category:'สินค้าพร้อมขาย',name:x.full,detail:'',qty:x.qty,unit:x.unit,exp:x.exp});});
    nonSaleResult.rows.forEach(function(x){add(x);});
    return {success:true,at:Utilities.formatDate(now,'Asia/Bangkok','dd/MM/yyyy HH:mm:ss'),rows:rows,productTotals:productTotals,categoryCounts:categoryCounts,alerts:{expired:expired,expiring:expiring,low:low,unknown:unknown},source:'Stock + InventoryNonSale'};
  } catch(e) {return {success:false,msg:e.message,rows:[]};}
}


/*********************************************************
 * V10.41 — Booth POS (financial ledger only)
 * STRICT ISOLATION: no writes to Stock, SalesLog, EventIssue,
 * EventReturn, StockIn, StockAdjust or inventory sheets.
 *********************************************************/
var POS_SHEETS_ = {
  products:'BoothPOS_Products', booths:'BoothPOS_Booths',
  bills:'BoothPOS_Bills', items:'BoothPOS_BillItems', audit:'BoothPOS_Audit'
};
function posGetSheet_(ss, name, heads) {
  var sh=ss.getSheetByName(name);
  if(!sh){sh=ss.insertSheet(name);sh.appendRow(heads);sh.setFrozenRows(1);}
  else if(sh.getLastRow()===0){sh.appendRow(heads);}
  return sh;
}
function posSetup_(){
  var ss=SpreadsheetApp.openById(SS_ID);
  var p=posGetSheet_(ss,POS_SHEETS_.products,['SKU','ชื่อสินค้า','ราคาแนะนำ','URL รูปสินค้าจริง','ใช้งาน']);
  posGetSheet_(ss,POS_SHEETS_.booths,['BoothID','วันที่เปิด','ชื่องาน','สถานที่','พนักงานเปิดงาน','ค่าเช่า','ค่าใช้จ่ายอื่น','เลขเบิกอ้างอิง (เพื่อเทียบเท่านั้น)','สถานะ']);
  posGetSheet_(ss,POS_SHEETS_.bills,['BillID','RequestID','BoothID','เวลา','พนักงานขาย','วิธีชำระ','ยอดขาย','หมายเหตุ','สถานะ']);
  posGetSheet_(ss,POS_SHEETS_.items,['BillID','SKU','ชื่อสินค้า','จำนวน','ราคาต่อชิ้น','รวม']);
  posGetSheet_(ss,POS_SHEETS_.audit,['เวลา','ประเภท','เลขบิล','ผู้ดำเนินการ','เหตุผล']);
  if(p.getLastRow()===1){
    p.getRange(2,1,5,5).setValues([
      ['balm','Baby Rich Balm','','','YES'],
      ['lotion','Organic Lotion','','','YES'],
      ['foam','Baby Foam','','','YES'],
      ['mozzie','Mozzie Guard','','','YES'],
      ['spray','Sanitizer Spray','','','YES']
    ]);
  }
  return ss;
}
function posRows_(sh){return sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues():[];}
function posText_(x,max){return String(x===undefined||x===null?'':x).trim().slice(0,max||150);}
function posMoney_(x){var n=Number(x);if(!isFinite(n)||n<0||n>100000000)throw new Error('จำนวนเงินไม่ถูกต้อง');return Math.round(n*100)/100;}
function posInitV1040(){
  var ss=posSetup_();
  return {products:posRows_(ss.getSheetByName(POS_SHEETS_.products)).filter(function(r){return String(r[4]).toUpperCase()!=='NO';}).map(function(r){return {sku:String(r[0]),name:String(r[1]),price:Number(r[2])||0,imageSource:String(r[3]||''),image:String(r[3]||'')};}),
  booths:posRows_(ss.getSheetByName(POS_SHEETS_.booths)).map(function(r){return {id:String(r[0]),date:posText_(r[1]),name:String(r[2]),location:String(r[3]),staff:String(r[4]),rent:Number(r[5])||0,other:Number(r[6])||0,issueNo:String(r[7]||''),status:String(r[8]||'OPEN')};}).reverse().slice(0,100)};
}
function posCreateBoothV1040(payload){
  throw new Error('POS เก่าย้ายไปเว็บ lilbulle-booth-pos.netlify.app แล้ว');
  payload=payload||{};
  var name=posText_(payload.name,120), staff=posText_(payload.staff,80);
  if(!name||!staff)throw new Error('กรุณากรอกชื่องานและพนักงาน');
  var rent=posMoney_(payload.rent||0), other=posMoney_(payload.other||0);
  var lock=LockService.getScriptLock();lock.waitLock(20000);
  try{
    var ss=posSetup_(), sh=ss.getSheetByName(POS_SHEETS_.booths);
    var id='BP-'+Utilities.formatDate(new Date(),'Asia/Bangkok','yyyyMMdd-HHmmss')+'-'+Utilities.getUuid().slice(0,6).toUpperCase();
    sh.appendRow([id,new Date(),name,posText_(payload.location,150),staff,rent,other,posText_(payload.issueNo,100),'OPEN']);
    return {success:true,id:id};
  }finally{lock.releaseLock();}
}
function posSaveBillV1040(payload){
  throw new Error('POS เก่าย้ายไปเว็บ lilbulle-booth-pos.netlify.app แล้ว');
  payload=payload||{};
  var boothId=posText_(payload.boothId,100), staff=posText_(payload.staff,80);
  var requestId=posText_(payload.requestId,120);
  if(!boothId||!staff||!requestId)throw new Error('ข้อมูลบูท / ผู้ขาย / รหัสคำขอไม่ครบ');
  var pay=posText_(payload.payment,30);
  if(['CASH','TRANSFER','CARD','OTHER'].indexOf(pay)<0)throw new Error('วิธีชำระไม่ถูกต้อง');
  var items=payload.items||[];
  if(!Array.isArray(items)||items.length<1||items.length>50)throw new Error('ต้องมีรายการสินค้า 1–50 รายการ');
  var lock=LockService.getScriptLock();lock.waitLock(25000);
  try{
    var ss=posSetup_(), bsh=ss.getSheetByName(POS_SHEETS_.booths), bills=ss.getSheetByName(POS_SHEETS_.bills), ish=ss.getSheetByName(POS_SHEETS_.items);
    var allBills=posRows_(bills);
    for(var a=0;a<allBills.length;a++)if(String(allBills[a][1])===requestId)return {success:true,id:String(allBills[a][0]),duplicate:true,total:Number(allBills[a][6])||0};
    var booths=posRows_(bsh), booth=booths.filter(function(r){return String(r[0])===boothId;})[0];
    if(!booth||String(booth[8])!=='OPEN')throw new Error('ไม่พบบูทที่เปิดขาย');
    if(booth[7] && typeof lbState_==='function' && lbState_(String(booth[7])).received) throw new Error('ใบเบิกนี้ใช้ POS V11 แล้ว กรุณาขายในเว็บใหม่');
    var products={};posRows_(ss.getSheetByName(POS_SHEETS_.products)).forEach(function(r){if(String(r[4]).toUpperCase()!=='NO')products[String(r[0])]=String(r[1]);});
    var normalized=[],total=0;
    items.forEach(function(x){
      var sku=posText_(x.sku,60), name=products[sku];
      if(!name)throw new Error('SKU ไม่ถูกต้อง: '+sku);
      var qty=Number(x.qty),price=posMoney_(x.price);
      if(price<=0)throw new Error('กรุณาใส่ราคาขายจริงมากกว่า 0 บาท');
      if(!Number.isInteger(qty)||qty<1||qty>10000)throw new Error('จำนวนสินค้าไม่ถูกต้อง');
      var line=Math.round(qty*price*100)/100;total+=line;
      normalized.push([sku,name,qty,price,line]);
    });
    total=Math.round(total*100)/100;
    var id='POS-'+Utilities.formatDate(new Date(),'Asia/Bangkok','yyyyMMdd-HHmmss')+'-'+Utilities.getUuid().slice(0,6).toUpperCase();
    // Append line items first, then commit bill header; orphan lines never count as a sale.
    ish.getRange(ish.getLastRow()+1,1,normalized.length,6).setValues(normalized.map(function(x){return [id].concat(x);}));
    bills.appendRow([id,requestId,boothId,new Date(),staff,pay,total,posText_(payload.note,250),'PAID']);
    return {success:true,id:id,total:total};
  }finally{lock.releaseLock();}
}
function posVoidBillV1040(pin,billId,reason){
  throw new Error('POS เก่าย้ายไปเว็บ lilbulle-booth-pos.netlify.app แล้ว');
  verifyBossPin_(pin);
  billId=posText_(billId,100);reason=posText_(reason,250);
  if(reason.length<5)throw new Error('ระบุเหตุผลยกเลิกอย่างน้อย 5 ตัวอักษร');
  var lock=LockService.getScriptLock();lock.waitLock(20000);
  try{
    var ss=posSetup_(),sh=ss.getSheetByName(POS_SHEETS_.bills),r=posRows_(sh),idx=-1;
    for(var i=0;i<r.length;i++)if(String(r[i][0])===billId){idx=i;break;}
    if(idx<0)throw new Error('ไม่พบบิล');
    var parent=posRows_(ss.getSheetByName(POS_SHEETS_.booths)).filter(function(b){return String(b[0])===String(r[idx][2]);})[0];
    if(parent&&String(parent[8])==='CLOSED')throw new Error('งาน POS ปิดแล้ว ไม่สามารถยกเลิกบิลได้');
    if(String(r[idx][8])==='VOID')throw new Error('บิลถูกยกเลิกแล้ว');
    sh.getRange(idx+2,9).setValue('VOID');
    ss.getSheetByName(POS_SHEETS_.audit).appendRow([new Date(),'VOID',billId,'BOSS',reason]);
    return {success:true};
  }finally{lock.releaseLock();}
}
function posReportV1040(pin,boothId){
  verifyBossPin_(pin);boothId=posText_(boothId,100);
  var ss=posSetup_(),booth=posRows_(ss.getSheetByName(POS_SHEETS_.booths)).filter(function(r){return String(r[0])===boothId;})[0];
  if(!booth)throw new Error('ไม่พบบูท');
  var bills=posRows_(ss.getSheetByName(POS_SHEETS_.bills)).filter(function(r){return String(r[2])===boothId;});
  var active=bills.filter(function(r){return String(r[8])==='PAID';});
  var byId={},byPay={},total=0;
  active.forEach(function(r){byId[String(r[0])]=true;var p=String(r[5]);byPay[p]=(byPay[p]||0)+Number(r[6]||0);total+=Number(r[6]||0);});
  var productMap={};
  posRows_(ss.getSheetByName(POS_SHEETS_.items)).forEach(function(r){
    if(!byId[String(r[0])])return;
    var sku=String(r[1]);if(!productMap[sku])productMap[sku]={sku:sku,name:String(r[2]),qty:0,amount:0};
    productMap[sku].qty+=Number(r[3])||0;productMap[sku].amount+=Number(r[5])||0;
  });
  var products=Object.keys(productMap).map(function(k){return productMap[k];});
  var compare=null,issueNo=String(booth[7]||'').trim();
  if(issueNo){
    try{
      var issueSh=ss.getSheetByName('EventIssue'),returnSh=ss.getSheetByName('EventReturn');
      if(!issueSh||!returnSh)throw new Error('ไม่พบชีตเบิก/คืน');
      var issueRows=posRows_(issueSh),returnRows=posRows_(returnSh),mapped={};
      var translate={'lil bulle baby rich balm':'balm','lil bulle organic lotion':'lotion','lil bulle baby foam':'foam','lil bulle mozzie guard':'mozzie','lil bulle sanitizer spray':'spray'};
      issueRows.forEach(function(line){if(String(line[1])!==issueNo)return;var sku=translate[String(line[3]||'').toLowerCase().trim()];if(sku)mapped[sku]=(mapped[sku]||0)+(Number(line[5])||0);});
      returnRows.forEach(function(line){if(String(line[2])!==issueNo)return;var sku=translate[String(line[4]||'').toLowerCase().trim()];if(sku)mapped[sku]=(mapped[sku]||0)-(Number(line[6])||0);});
      compare=Object.keys(Object.assign({},mapped,productMap)).map(function(sku){var pos=productMap[sku]?productMap[sku].qty:0;return {sku:sku,posQty:pos,eventDifference:mapped[sku]||0,difference:pos-(mapped[sku]||0)};});
    }catch(e){compare={error:'ไม่สามารถอ่านเลขเบิกอ้างอิง: '+e.message};}
  }
  var dailyMap={};
  active.forEach(function(r){
    var when=r[3];
    var date=when instanceof Date?Utilities.formatDate(when,'Asia/Bangkok','yyyy-MM-dd'):posDateKey_(when);
    if(!dailyMap[date])dailyMap[date]={date:date,total:0,billCount:0,payments:{}};
    var d=dailyMap[date], amount=Number(r[6])||0, method=String(r[5]);
    d.total=Math.round((d.total+amount)*100)/100;
    d.billCount++;
    d.payments[method]=Math.round(((d.payments[method]||0)+amount)*100)/100;
  });
  var daily=Object.keys(dailyMap).sort().map(function(k){return dailyMap[k];});
  var audit=posRows_(ss.getSheetByName(POS_SHEETS_.audit));
  var closing=audit.filter(function(r){return String(r[1])==='CLOSE'&&String(r[2])===boothId;}).slice(-1)[0];
  var status=String(booth[8]||'OPEN');
  var rent=Number(booth[5])||0,other=Number(booth[6])||0;
  return {success:true,booth:{id:boothId,name:String(booth[2]),location:String(booth[3]),issueNo:issueNo,status:status,closedAt:closing?(closing[0] instanceof Date?Utilities.formatDate(closing[0],'Asia/Bangkok','dd/MM/yyyy HH:mm'):String(closing[0])):''},daily:daily,total:total,rent:rent,other:other,net:total-rent-other,billCount:active.length,voidCount:bills.length-active.length,payments:byPay,products:products,comparison:compare,bills:bills.slice(-80).reverse().map(function(r){return {id:String(r[0]),date:posText_(r[3]),staff:String(r[4]),payment:String(r[5]),amount:Number(r[6])||0,status:String(r[8])};})};
}


/** V10.41: Date label for historical text values. Live timestamps are Date objects. */
function posDateKey_(value){
  var s=String(value||'').trim();
  var iso=s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if(iso)return iso[1]+'-'+iso[2]+'-'+iso[3];
  var thai=s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if(thai)return thai[3]+'-'+('0'+thai[2]).slice(-2)+'-'+('0'+thai[1]).slice(-2);
  return 'ไม่ทราบวันที่';
}

/** Close only POS sales for one booth; never close EventIssue or touch stock. */
function posCloseBoothV1041(pin,boothId,ack){
  throw new Error('POS เก่าย้ายไปเว็บ lilbulle-booth-pos.netlify.app แล้ว');
  verifyBossPin_(pin);
  boothId=posText_(boothId,100);
  if(!boothId)throw new Error('กรุณาเลือกงาน POS');
  var lock=LockService.getScriptLock();lock.waitLock(25000);
  try{
    var ss=posSetup_(),sh=ss.getSheetByName(POS_SHEETS_.booths),rows=posRows_(sh),index=-1;
    for(var i=0;i<rows.length;i++)if(String(rows[i][0])===boothId){index=i;break;}
    if(index<0)throw new Error('ไม่พบงาน POS');
    if(String(rows[index][8])==='CLOSED')return {success:true,alreadyClosed:true,boothId:boothId};
    if(String(rows[index][8])!=='OPEN')throw new Error('สถานะงาน POS ไม่ถูกต้อง');
    posSimBeforeClose_(boothId,ack);
    sh.getRange(index+2,9).setValue('CLOSED');
    ss.getSheetByName(POS_SHEETS_.audit).appendRow([new Date(),'CLOSE',boothId,'BOSS','ปิดงาน POS เท่านั้น ไม่แตะสต๊อกหรือใบเบิกคืน']);
    return {success:true,boothId:boothId};
  }finally{lock.releaseLock();}
}

/** V10.42: team view, limited to the selected POS event. No costs or inventory data. */
function posTeamHistoryV1042(boothId){
  boothId=posText_(boothId,100);
  var ss=posSetup_();
  var booth=posRows_(ss.getSheetByName(POS_SHEETS_.booths)).filter(function(r){return String(r[0])===boothId;})[0];
  if(!booth)throw new Error('ไม่พบงาน POS');
  var today=Utilities.formatDate(new Date(),'Asia/Bangkok','yyyy-MM-dd');
  var rows=posRows_(ss.getSheetByName(POS_SHEETS_.bills)).filter(function(r){return String(r[2])===boothId;});
  var total=0,count=0;
  rows.forEach(function(r){var key=r[3] instanceof Date?Utilities.formatDate(r[3],'Asia/Bangkok','yyyy-MM-dd'):posDateKey_(r[3]);if(key===today&&String(r[8])==='PAID'){total+=Number(r[6])||0;count++;}});
  return {boothId:boothId,status:String(booth[8]),today:today,total:Math.round(total*100)/100,count:count,
    bills:rows.slice(-80).reverse().map(function(r){return {id:String(r[0]),date:r[3] instanceof Date?Utilities.formatDate(r[3],'Asia/Bangkok','dd/MM/yyyy HH:mm'):String(r[3]),staff:String(r[4]),payment:String(r[5]),amount:Number(r[6])||0,status:String(r[8])};})};
}

// V10.43: independent booth forecast and immutable end-of-day reconciliation.
function posSimSetup_(){
  var ss=posSetup_();
  posGetSheet_(ss,'BoothPOS_SimMoves',['MoveID','RequestID','BoothID','เวลา','ประเภท','SKU','จำนวน','เตือนเมื่อเหลือ','ผู้บันทึก','หมายเหตุ']);
  posGetSheet_(ss,'BoothPOS_Reconcile',['CheckID','RequestID','BoothID','เวลา','วันที่ไทย','ผู้ตรวจ','SnapshotJSON']);
  return ss;
}
function posSimState_(ss,boothId){
  var booth=posRows_(ss.getSheetByName(POS_SHEETS_.booths)).filter(function(r){return String(r[0])===boothId;})[0];
  if(!booth)throw new Error('ไม่พบงาน POS');
  var map={},moves=posRows_(ss.getSheetByName('BoothPOS_SimMoves')).filter(function(r){return String(r[2])===boothId;});
  posRows_(ss.getSheetByName(POS_SHEETS_.products)).forEach(function(r){if(String(r[4]).toUpperCase()!=='NO')map[String(r[0])]={sku:String(r[0]),name:String(r[1]),ready:false,opening:0,added:0,out:0,sold:0,threshold:5};});
  moves.forEach(function(r){var sku=String(r[5]);if(!map[sku])map[sku]={sku:sku,name:sku,ready:false,opening:0,added:0,out:0,sold:0,threshold:5};var p=map[sku],q=Number(r[6])||0;
    if(r[4]==='OPEN'){p.ready=true;p.opening+=q;p.threshold=Number(r[7])||0;}
    else if(r[4]==='ADD')p.added+=q;else p.out+=q;
  });
  var bills=posRows_(ss.getSheetByName(POS_SHEETS_.bills)).filter(function(r){return String(r[2])===boothId;}),paid={};
  var today=Utilities.formatDate(new Date(),'Asia/Bangkok','yyyy-MM-dd'),payments={CASH:0,TRANSFER:0,CARD:0,OTHER:0},todayBills=0,allTotal=0;
  bills.forEach(function(r){if(String(r[8])!=='PAID')return;paid[String(r[0])]=true;allTotal+=Number(r[6])||0;
    var day=r[3] instanceof Date?Utilities.formatDate(r[3],'Asia/Bangkok','yyyy-MM-dd'):posDateKey_(r[3]);
    if(day===today){payments[String(r[5])]=(payments[String(r[5])]||0)+(Number(r[6])||0);todayBills++;}
  });
  var items=posRows_(ss.getSheetByName(POS_SHEETS_.items)).filter(function(r){return !!paid[String(r[0])];});
  items.forEach(function(r){var sku=String(r[1]);if(!map[sku])map[sku]={sku:sku,name:String(r[2]),ready:false,opening:0,added:0,out:0,sold:0,threshold:5};map[sku].sold+=Number(r[3])||0;});
  var products=Object.keys(map).sort().map(function(sku){var p=map[sku];p.expected=p.ready?p.opening+p.added-p.out-p.sold:null;p.low=p.ready&&p.expected<=p.threshold;return p;});
  Object.keys(payments).forEach(function(k){payments[k]=Math.round(payments[k]*100)/100;});
  // Fingerprint includes bill status and quantities as well as moves, so later voids/sales invalidate a count.
  var raw=JSON.stringify({day:today,products:products,payments:payments,bills:bills.map(function(r){return [String(r[0]),String(r[8]),Number(r[6])];}),moves:moves.map(function(r){return String(r[0]);})});
  var token=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,raw,Utilities.Charset.UTF_8).map(function(b){return ('0'+((b+256)%256).toString(16)).slice(-2);}).join('');
  return {boothId:boothId,name:String(booth[2]),status:String(booth[8]),today:today,products:products,payments:payments,todayBills:todayBills,allTotal:Math.round(allTotal*100)/100,token:token};
}
function posSimDashboardV1043(boothId){
  var ss=posSimSetup_(),state=posSimState_(ss,posText_(boothId,100));
  var checks=posRows_(ss.getSheetByName('BoothPOS_Reconcile')).filter(function(r){return String(r[2])===state.boothId;});
  state.checks=checks.slice(-10).reverse().map(function(r){var s=JSON.parse(String(r[6]));return {id:String(r[0]),date:String(r[4]),staff:String(r[5]),at:r[3] instanceof Date?Utilities.formatDate(r[3],'Asia/Bangkok','dd/MM/yyyy HH:mm'):String(r[3]),fresh:s.token===state.token,snapshot:s};});
  return state;
}
function posSimQty_(v,zero){
  if(v===null||v===undefined||String(v).trim()==='')throw new Error('กรุณากรอกจำนวนให้ครบ ใส่ 0 ได้');
  var q=Number(v);if(!Number.isInteger(q)||q<(zero?0:1)||q>1000000)throw new Error('จำนวนต้องเป็นจำนวนเต็มที่ถูกต้อง');return q;
}
function posSimMoveV1043(payload){
  throw new Error('POS เก่าย้ายไปเว็บ lilbulle-booth-pos.netlify.app แล้ว');
  payload=payload||{};var boothId=posText_(payload.boothId,100),requestId=posText_(payload.requestId,120),staff=posText_(payload.staff,80),kind=String(payload.kind),note=posText_(payload.note,250);
  if(!boothId||!requestId||!staff)throw new Error('กรอกงาน ผู้บันทึก และรหัสคำขอ');
  if(['OPEN','ADD','FREE','TESTER','DAMAGE'].indexOf(kind)<0)throw new Error('ประเภทรายการไม่ถูกต้อง');
  if(!Array.isArray(payload.items)||payload.items.length<1||payload.items.length>100)throw new Error('รายการไม่ถูกต้อง');
  var lock=LockService.getScriptLock();lock.waitLock(25000);
  try{
    var ss=posSimSetup_(),sh=ss.getSheetByName('BoothPOS_SimMoves'),rows=posRows_(sh);
    var previous=rows.filter(function(r){return String(r[1])===requestId;});
    if(previous.length)return {success:true,duplicate:true,id:String(previous[0][0])};
    var state=posSimState_(ss,boothId);if(state.status!=='OPEN')throw new Error('งานปิดแล้ว แก้ยอดจำลองไม่ได้');
    var seen={},normalized=payload.items.map(function(x){
      var p=state.products.filter(function(p){return p.sku===String(x.sku);})[0];if(!p||seen[p.sku])throw new Error('สินค้าไม่ถูกต้องหรือซ้ำ');seen[p.sku]=true;
      if(kind==='OPEN'&&p.ready)throw new Error('ตั้งยอดเริ่มต้นสินค้านี้แล้ว หากเติมของให้ใช้รายการเติม');
      if(kind!=='OPEN'&&!p.ready)throw new Error('ตั้งยอดเริ่มต้นก่อนบันทึกรายการ');
      return [p.sku,posSimQty_(x.qty,kind==='OPEN'),kind==='OPEN'?posSimQty_(x.threshold,true):''];
    });
    var id='SM-'+Utilities.getUuid(),at=new Date();
    sh.getRange(sh.getLastRow()+1,1,normalized.length,10).setValues(normalized.map(function(x){return [id,requestId,boothId,at,kind,x[0],x[1],x[2],staff,note];}));
    return {success:true,id:id};
  }finally{lock.releaseLock();}
}
function posSimReconcileV1043(payload){
  throw new Error('POS เก่าย้ายไปเว็บ lilbulle-booth-pos.netlify.app แล้ว');
  payload=payload||{};var requestId=posText_(payload.requestId,120),staff=posText_(payload.staff,80),boothId=posText_(payload.boothId,100);
  if(!requestId||!staff||!boothId)throw new Error('กรอกงานและชื่อผู้ตรวจ');
  var lock=LockService.getScriptLock();lock.waitLock(25000);
  try{
    var ss=posSimSetup_(),sh=ss.getSheetByName('BoothPOS_Reconcile');
    var previous=posRows_(sh).filter(function(r){return String(r[1])===requestId;})[0];
    if(previous)return {success:true,id:String(previous[0]),duplicate:true};
    var state=posSimState_(ss,boothId);
    if(String(payload.token)!==state.token)throw new Error('ยอดเปลี่ยนแล้ว กรุณาโหลดใหม่และนับตรวจอีกครั้ง');
    if(state.products.some(function(p){return !p.ready;}))throw new Error('ตั้งยอดเริ่มต้นทุกสินค้าก่อน ใส่ 0 สำหรับสินค้าที่ไม่ได้นำไป');
    if(!Array.isArray(payload.counts)||payload.counts.length!==state.products.length)throw new Error('กรอกจำนวนจริงครบทุกสินค้า');
    var counts={},seen={};payload.counts.forEach(function(x){var sku=String(x.sku);if(seen[sku]||!state.products.some(function(p){return p.sku===sku;}))throw new Error('รายการนับไม่ถูกต้อง');seen[sku]=true;counts[sku]=posSimQty_(x.actual,true);});
    var stock=state.products.map(function(p){return {sku:p.sku,name:p.name,expected:p.expected,actual:counts[p.sku],difference:counts[p.sku]-p.expected};});
    var money=Object.keys(state.payments).map(function(k){var actual=(payload.money||{})[k];if(actual===null||actual===undefined||String(actual).trim()==='')throw new Error('กรอกยอดรับจริงทุกช่องทาง ใส่ 0 ถ้าไม่มี');actual=posMoney_(actual);return {method:k,expected:state.payments[k],actual:actual,difference:Math.round((actual-state.payments[k])*100)/100};});
    var matched=stock.every(function(p){return p.difference===0;})&&money.every(function(p){return p.difference===0;});
    var note=posText_(payload.note,500);if(!matched&&note.length<5)throw new Error('ยอดต่างกัน กรุณาอธิบายอย่างน้อย 5 ตัวอักษร');
    var expectedTotal=Math.round(money.reduce(function(s,m){return s+m.expected;},0)*100)/100;
    var actualTotal=Math.round(money.reduce(function(s,m){return s+m.actual;},0)*100)/100;
    var snapshot={token:state.token,stock:stock,money:money,matched:matched,note:note,todayBills:state.todayBills,allTotal:state.allTotal,expectedTotal:expectedTotal,actualTotal:actualTotal,moneyDifference:Math.round((actualTotal-expectedTotal)*100)/100};
    var json=JSON.stringify(snapshot);if(json.length>45000)throw new Error('ข้อมูลตรวจนับมากเกินไป กรุณาติดต่อผู้ดูแล');
    var id='RC-'+Utilities.getUuid();sh.appendRow([id,requestId,boothId,new Date(),state.today,staff,json]);
    return {success:true,id:id,matched:matched};
  }finally{lock.releaseLock();}
}
// Called inside the existing close lock. A mismatch is allowed with a note and explicit boss confirmation.
function posSimBeforeClose_(boothId,ack){
  var ss=posSimSetup_(),state=posSimState_(ss,boothId);
  var checks=posRows_(ss.getSheetByName('BoothPOS_Reconcile')).filter(function(r){return String(r[2])===boothId;});
  if(!checks.length)throw new Error('บันทึกตรวจนับสินค้าและเงินวันนี้ก่อนปิดงาน');
  var snapshot=JSON.parse(String(checks[checks.length-1][6]));
  if(snapshot.token!==state.token)throw new Error('ยอดเปลี่ยนหลังตรวจนับ กรุณาบันทึกตรวจยอดใหม่ก่อนปิดงาน');
  if(!snapshot.matched&&ack!==true)throw new Error('มีส่วนต่าง กรุณาให้บอสยืนยันรับทราบก่อนปิดงาน');
}


/** V11.3 read-only lot return diagnostic: warehouse-only, never uses POS counts. */
function previewEventReturnV113(issueNo) {
  var g=getEventOutstandingMap_()[String(issueNo||'').trim()];
  if(!g)return {success:false,msg:'ไม่พบใบเบิก '+issueNo,rows:[]};
  return {success:true,issueNo:g.issueNo,status:g.status,rows:Object.keys(g.items).map(function(k){
    var x=g.items[k],available=Number(x.issued)-Number(x.returned);
    return {product:x.full,exp:x.exp,issued:Number(x.issued),returned:Number(x.returned),returnable:available,warning:available<0?'พบยอดคืนเกินเบิก':''};
  })};
}
