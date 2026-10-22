/**
 * @OnlyCurrentDoc
 *
 * Caddie｜Web集客・AI対策診断 — 回答受け取り用スクリプト（Google Apps Script）
 *
 * 権限を「このスプレッドシートだけ」に絞っているため、Googleにブロックされにくい作りです。
 *
 * 使い方:
 *   1. 回答を貯めたいGoogleアカウントで、新しいGoogleスプレッドシートを作成（https://sheets.new）
 *   2. メニュー「拡張機能」→「Apps Script」を開き、このコードを貼り付けて保存
 *   3. 関数「setup」を選んで実行し、権限を許可
 *      （「Googleで確認されていません」と出たら「詳細」→「（安全ではないページ）に移動」）
 *   4. 右上「デプロイ」→「新しいデプロイ」→ 種類「ウェブアプリ」
 *        次のユーザーとして実行: 自分
 *        アクセスできるユーザー: 全員
 *      → 発行された「ウェブアプリURL」（https://script.google.com/macros/s/…/exec）を
 *        index.html の ENDPOINT に貼り付けます
 *
 * 動き:
 *   - 結果画面が表示された時点で、点数を1行記録します（無料診断希望＝いいえ）
 *   - 無料診断に申し込まれると、同じ行に連絡先を追記します
 *   - 申し込みのメール通知は、スプレッドシートの「ツール」→「通知設定」→「通知を編集」で設定できます
 */

const SHEET_NAME = '診断結果';
const QUESTION_COUNT = 10;
const VALID_POINTS = [0, 5, 10];

const HEADERS = [
  '受付ID', '日時', '合計スコア', '判定', '無料診断希望',
  '会社名', 'お名前', 'ご連絡先', '流入元パラメータ', '参照元',
  ...Array.from({ length: QUESTION_COUNT }, (_, i) => `Q${String(i + 1).padStart(2, '0')}`),
];

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0].setName(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  }
  Logger.log('準備完了: ' + ss.getUrl());
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const data = JSON.parse(e.postData.contents);

    // ハニーポットが埋まっていたらボットとみなして無視
    if (data.website) return ok_();

    const scores = Array.isArray(data.scores) ? data.scores.map(Number) : [];
    if (scores.length !== QUESTION_COUNT || !scores.every(s => VALID_POINTS.includes(s))) return ok_();
    const id = clean_(data.id, 64);
    if (!id) return ok_();

    const total = scores.reduce((a, b) => a + b, 0);
    const consult = data.consult === true;
    const row = [
      id,
      new Date(),
      total,
      judge_(total),
      consult ? 'はい' : 'いいえ',
      consult ? clean_(data.company, 100) : '',
      consult ? clean_(data.name, 100) : '',
      consult ? clean_(data.contact, 200) : '',
      clean_(data.source, 300),
      clean_(data.referrer, 300),
      ...scores,
    ];

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    const found = sheet.getRange('A:A').createTextFinder(id).matchEntireCell(true).findNext();
    if (found) {
      sheet.getRange(found.getRow(), 1, 1, row.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }
    return ok_();
  } finally {
    lock.releaseLock();
  }
}

function judge_(total) {
  if (total < 50) return '要改善';
  if (total < 80) return '改善余地あり';
  return '運用充実';
}

// 長さを制限し、数式として解釈されないようにする
function clean_(value, max) {
  let s = String(value == null ? '' : value).trim().slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function ok_() {
  return ContentService.createTextOutput('ok');
}
