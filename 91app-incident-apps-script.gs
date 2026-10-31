/**
 * 91APP 資安事件｜來電登記表 後台程式
 * 用法：貼到 Google 試算表的「擴充功能 → Apps Script」，部署成「網頁應用程式」
 * 部署設定：執行身分＝我；存取權＝所有人
 * 網頁上看到的姓名、電話會自動遮蔽，完整資料只在試算表裡。
 */
var SHEET_NAME = '91APP來電登記';
var HEADERS = ['登記時間', '來電時間', '登記同仁', '顧客姓名', '聯絡電話', '官網訂單編號', '來電類型', '是否有財損', '損失金額', '處理方式', '備註'];

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
  }
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#B81C22').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}

function clean_(v, max) {
  return String(v == null ? '' : v).replace(/^[=+\-@]/, "'$&").slice(0, max || 300);
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var d = JSON.parse(e.postData.contents || '{}');
    if (!d.by || !d.name || !d.phone || !d.type) {
      return json_({ ok: false, error: '必填欄位不完整' });
    }
    var sh = getSheet_();
    var now = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');
    sh.appendRow([
      now,
      "'" + clean_(d.time, 20).replace(/-/g, '/'),
      clean_(d.by, 30),
      clean_(d.name, 50),
      "'" + clean_(d.phone, 20),
      "'" + clean_(d.order, 40),
      clean_(d.type, 40),
      clean_(d.loss, 20),
      clean_(d.amount, 15),
      clean_(d.action, 40),
      clean_(d.note, 1000)
    ]);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function doGet() {
  var sh = getSheet_();
  var last = sh.getLastRow();
  var rows = [];
  if (last > 1) {
    var vals = sh.getRange(2, 1, last - 1, HEADERS.length).getDisplayValues();
    for (var i = vals.length - 1; i >= 0; i--) {
      var r = vals[i];
      rows.push({
        time: r[1] || r[0],
        by: r[2],
        name: maskName_(r[3]),
        phone: maskPhone_(r[4]),
        type: r[6],
        loss: r[7],
        amount: r[8],
        action: r[9],
        note: r[10]
      });
    }
  }
  return json_({ ok: true, rows: rows });
}

function maskName_(s) {
  s = String(s || '').trim();
  if (s.length <= 1) return s;
  if (s.length === 2) return s.charAt(0) + '○';
  return s.charAt(0) + '○'.repeat(s.length - 2) + s.charAt(s.length - 1);
}

function maskPhone_(s) {
  s = String(s || '').replace(/\s/g, '');
  if (s.length <= 6) return s.replace(/.(?=.{2})/g, '*');
  return s.slice(0, 4) + '***' + s.slice(-3);
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
