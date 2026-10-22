/**
 * Caddie｜Web集客・AI対策診断 — 回答受け取り用スクリプト（Google Apps Script）
 *
 * 使い方:
 *   1. https://script.google.com で「新しいプロジェクト」を作成し、このコードを貼り付けて保存
 *   2. 関数「setup」を選んで実行（初回は権限の許可が必要）
 *      → 回答用スプレッドシートが作られ、実行ログにURLが出ます
 *   3. 右上「デプロイ」→「新しいデプロイ」→ 種類「ウェブアプリ」
 *        次のユーザーとして実行: 自分
 *        アクセスできるユーザー: 全員
 *      → 発行された「ウェブアプリURL」（https://script.google.com/macros/s/…/exec）を
 *        index.html の ENDPOINT に貼り付けます
 *
 * 動き:
 *   - 結果画面が表示された時点で、点数を1行記録します（無料診断希望＝いいえ）
 *   - 無料診断に申し込まれると、同じ行に連絡先を追記し、通知メールを送ります
 */

const SHEET_NAME = '診断結果';
const NOTIFY = true; // 申し込み時に自分宛てへメール通知する
const QUESTION_COUNT = 10;
const VALID_POINTS = [0, 5, 10];

const HEADERS = [
  '受付ID', '日時', '合計スコア', '判定', '無料診断希望',
  '会社名', 'お名前', 'ご連絡先', '流入元パラメータ', '参照元',
  ...Array.from({ length: QUESTION_COUNT }, (_, i) => `Q${String(i + 1).padStart(2, '0')}`),
];

function setup() {
  const ss = SpreadsheetApp.create('Caddie｜Web集客・AI対策診断（回答）');
  const sheet = ss.getSheets()[0].setName(SHEET_NAME);
  sheet.appendRow(HEADERS);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  PropertiesService.getScriptProperties().setProperty('SHEET_ID', ss.getId());
  Logger.log('回答シート: ' + ss.getUrl());
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

    const sheet = SpreadsheetApp
      .openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID'))
      .getSheetByName(SHEET_NAME);

    const found = sheet.getRange('A:A').createTextFinder(id).matchEntireCell(true).findNext();
    if (found) {
      sheet.getRange(found.getRow(), 1, 1, row.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }

    if (consult && NOTIFY) notify_(row, sheet.getParent().getUrl());
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

function notify_(row, sheetUrl) {
  const [, , total, band, , company, name, contact] = row;
  MailApp.sendEmail({
    to: Session.getEffectiveUser().getEmail(),
    subject: `【Web集客診断】無料診断の申し込み：${company} ${name}様（${total}点）`,
    body: [
      'Web集客・AI対策診断から無料診断の申し込みがありました。',
      '',
      `会社名：${company}`,
      `お名前：${name}`,
      `ご連絡先：${contact}`,
      `スコア：${total}点（${band}）`,
      '',
      `回答シート：${sheetUrl}`,
    ].join('\n'),
  });
}

function ok_() {
  return ContentService.createTextOutput('ok');
}
