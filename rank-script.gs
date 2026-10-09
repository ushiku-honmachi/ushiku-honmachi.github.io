// 本町ゲームのランキング用スクリプト(Google Apps Script)
// スプレッドシートの「拡張機能」→「Apps Script」に、これを全部貼ります。
var SHEET_NAME = 'scores';
var MAX_SCORE = 30000;  // これより大きい点数は受け付けません
var MAX_ROWS = 2000;    // 記録が増えすぎたときの上限

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) sh.appendRow(['日時', '名前', '点数', '非表示にするなら x']);
  return sh;
}

// 上位3人を返す(「非表示」の列に何か書いた行は、数えません)
function top_() {
  var sh = sheet_(), last = sh.getLastRow();
  if (last < 2) return [];
  var rows = sh.getRange(2, 1, last - 1, 4).getValues();
  var list = [];
  rows.forEach(function (r, i) {
    var sc = Number(r[2]);
    if (r[1] === '' || isNaN(sc) || String(r[3]).trim() !== '') return;
    list.push({ name: String(r[1]), score: sc, order: i });
  });
  list.sort(function (a, b) { return b.score - a.score || a.order - b.order; });
  return list.slice(0, 3).map(function (x) { return { name: x.name, score: x.score }; });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return json_({ ranks: top_() });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var d = JSON.parse(e.postData.contents);
    var name = String(d.name || '').replace(/[<>&"'\\]/g, '').replace(/\s+/g, ' ').trim().slice(0, 12);
    var score = parseInt(d.score, 10);
    if (!name || isNaN(score) || score < 1 || score > MAX_SCORE) return json_({ ok: false });
    var sh = sheet_();
    sh.appendRow([new Date(), name, score, '']);
    if (sh.getLastRow() > MAX_ROWS + 1) sh.deleteRow(2);   // 古い記録から消す
    return json_({ ok: true, ranks: top_() });
  } catch (err) {
    return json_({ ok: false });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}
