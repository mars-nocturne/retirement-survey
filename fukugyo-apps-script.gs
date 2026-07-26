/*
 * =====================================================================
 * 副業OK？NG？リアル調査 — 共有データベース用 Google Apps Script
 * =====================================================================
 *
 * ■ これは何？
 *   fukugyo.html の投稿を、Google スプレッドシートに保存して
 *   全員で共有するためのバックエンド（無料・サーバー不要）です。
 *
 * ■ セットアップ手順
 *   1. https://sheets.google.com で新しいスプレッドシートを作成
 *   2. メニュー「拡張機能」→「Apps Script」を開く
 *   3. 既存のコードを全部消して、このファイルの中身を貼り付けて保存
 *   4. 右上「デプロイ」→「新しいデプロイ」→ 種類「ウェブアプリ」
 *        - 次のユーザーとして実行： 自分
 *        - アクセスできるユーザー： 全員
 *      →「デプロイ」。表示された「ウェブアプリのURL」（.../exec）をコピー
 *   5. fukugyo.html の  const GAS_URL = '';  にそのURLを貼り付け
 *   6. 保存してページを開くと「✅ 共有データベースと接続中」と表示されればOK
 *
 * ■ 注意
 *   - シート「fukugyo」は初回アクセス時に自動作成されます
 *   - コードを修正したら「デプロイを管理」から再デプロイしてください
 *   - 誰でも投稿できる仕様です。荒らし対策が必要なら handlePost に
 *     フィルタ（NGワード等）を追加してください
 */

var SHEET_NAME = 'fukugyo';
var HEADERS = ['id', 'company', 'status', 'industry', 'source', 'year', 'url', 'note', 'ts'];

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'get';
  if (action === 'post') return handlePost(e);
  return handleGet();
}

function handleGet() {
  var sheet = getSheet();
  var values = sheet.getDataRange().getValues();
  var data = [];
  for (var i = 1; i < values.length; i++) { // 0行目はヘッダー
    var row = values[i];
    var obj = {};
    for (var j = 0; j < HEADERS.length; j++) obj[HEADERS[j]] = row[j];
    if (!obj.id) continue;
    obj.ts = Number(obj.ts) || 0;
    data.push(obj);
  }
  return json({ status: 'ok', data: data });
}

function handlePost(e) {
  try {
    var entry = JSON.parse(e.parameter.body);
    if (!entry || !entry.id || !entry.company || !entry.status) {
      return json({ status: 'error', message: 'invalid entry' });
    }
    var sheet = getSheet();
    // 同一idの重複防止
    var ids = sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (ids[i][0] === entry.id) return json({ status: 'ok', duplicate: true });
    }
    var newRow = [];
    for (var k = 0; k < HEADERS.length; k++) {
      var v = entry[HEADERS[k]];
      newRow.push(v == null ? '' : v);
    }
    sheet.appendRow(newRow);
    return json({ status: 'ok' });
  } catch (err) {
    return json({ status: 'error', message: String(err) });
  }
}

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
  }
  if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS);
  return sheet;
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
