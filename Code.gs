// 貓咪血壓記錄 — Google Apps Script
// 貼到「綁定試算表」的 Apps Script（試算表 → 擴充功能 → Apps Script）
// 部署：部署 → 新增部署 → 網頁應用程式｜執行身分：我｜存取權：任何人

var HEADERS = ['日期', '時間', '收縮壓', '舒張壓', '脈搏'];

function cleanName_(name) {
  return String(name || '')
    .replace(/[\[\]\*\?\:\/\\]/g, '')   // 試算表分頁名稱禁用字元
    .trim()
    .slice(0, 30);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// 回傳現有貓咪（分頁）名單，供網頁下拉選單使用
function doGet() {
  var cats = SpreadsheetApp.getActiveSpreadsheet().getSheets().map(function (s) {
    return s.getName();
  });
  return json_({ ok: true, cats: cats });
}

// 寫入一筆血壓紀錄
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var d = JSON.parse(e.postData.contents);
    var cat = cleanName_(d.cat);
    if (!cat) return json_({ ok: false, error: '請輸入貓咪名字' });

    var sys = Number(d.systolic), dia = Number(d.diastolic), pulse = Number(d.pulse);
    if (!d.date || !d.time || !(sys > 0) || !(dia > 0) || !(pulse > 0)) {
      return json_({ ok: false, error: '資料不完整' });
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(cat);
    if (!sheet) {
      sheet = ss.insertSheet(cat);
      sheet.appendRow(HEADERS);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
      // 預設空白的「工作表1」若還是空的，順手刪掉
      var blank = ss.getSheetByName('工作表1') || ss.getSheetByName('Sheet1');
      if (blank && blank.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(blank);
    }

    var row = sheet.getLastRow() + 1;
    sheet.getRange(row, 1, 1, 2).setNumberFormat('@');   // 日期、時間存成文字，避免被轉格式
    sheet.getRange(row, 1, 1, HEADERS.length).setValues([[d.date, d.time, sys, dia, pulse]]);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
