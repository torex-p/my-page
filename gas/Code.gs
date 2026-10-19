/**
 * エアコンクリーニング予約フォーム 受付スクリプト
 * すまいのお掃除専門店
 *
 * 予約フォームから送られた内容を
 *   1. このスプレッドシートの「予約一覧」シートに1行追加
 *   2. 社内向けに通知メールを送信
 *   3. お客様に受付確認メールを送信
 * します。
 */

// ===== 設定 =====
const SETTINGS = {
  // 予約を記録するスプレッドシートのID（URLの /d/ と /edit の間の文字列）
  spreadsheetId: '1qWYeaYNrox6sZADi71P81OsZLH76mq51zrtGnjCLCVs',
  sheetName: '予約一覧',
  // 通知メールの送り先（空欄ならこのスクリプトをデプロイしたアカウント宛て）
  notifyTo: '',
  companyName: 'すまいのお掃除専門店',
  // TODO: 自社の電話番号に差し替え
  companyTel: '00-0000-0000'
};

// 予約一覧シートの列（左から順）。
// 列の並び替えや追加をしても、見出し名で書き込み先を探すので動きます。
const HEADERS = [
  '受付番号', '受付日時', 'お名前', '電話', 'メール', 'エリア', '郵便番号', '住所',
  '通常台数', 'お掃除機能付き台数', '3点セット', '合計 (税込)', '通常価格合計',
  '第1希望日', '第2希望日', '時間帯', '支払方法', '駐車場', '備考', '紹介コード',
  '紹介料 (20%)', '支払い状況', '作業確定日', 'PayPal 請求書ID', 'PayPal 状態', 'メモ',
  '3点セット台数', '防カビコート台数', '室外機台数', 'ドレン台数', '住居タイプ',
  '住居タイプ根拠', '重複チェック', '同期状況 (スケジュール表)', '紹介の判定',
  '紹介料 お知らせ日', '紹介料 締め月'
];

const REFERRAL_RATE = 0.2;

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const data = JSON.parse(e.postData.contents);
    data.receivedAt = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');

    appendRow_(data);
    sendNotify_(data);
    if (data.email) sendConfirmation_(data);

    return json_({ ok: true });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// 動作確認用（ブラウザでURLを開くと表示される）
function doGet() {
  return json_({ ok: true, message: '予約受付スクリプトは動作しています' });
}

function getSheet_() {
  const ss = SpreadsheetApp.openById(SETTINGS.spreadsheetId);
  let sheet = ss.getSheetByName(SETTINGS.sheetName);
  if (!sheet) sheet = ss.insertSheet(SETTINGS.sheetName);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#e8f2fb');
  }
  return sheet;
}

const digits_ = v => String(v || '').replace(/[^\d]/g, '');

// 受付番号：受付日（yyMMdd）＋その日の通し番号（例：261019-001）
function makeReceiptNo_(sheet, col) {
  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyMMdd');
  let n = 0;
  if (col && sheet.getLastRow() > 1) {
    sheet.getRange(2, col, sheet.getLastRow() - 1, 1).getValues().forEach(r => {
      if (String(r[0]).indexOf(today + '-') === 0) n++;
    });
  }
  return today + '-' + String(n + 1).padStart(3, '0');
}

// 同じ電話番号・メールの予約がすでにあれば、その受付番号を返す
function findDuplicates_(sheet, index, d) {
  if (sheet.getLastRow() < 2) return '';
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const tel = digits_(d.tel), mail = String(d.email || '').trim().toLowerCase();
  const hits = [];
  rows.forEach(r => {
    const sameTel = index['電話'] && tel && digits_(r[index['電話'] - 1]) === tel;
    const sameMail = index['メール'] && mail && String(r[index['メール'] - 1]).trim().toLowerCase() === mail;
    if (sameTel || sameMail) hits.push(index['受付番号'] ? r[index['受付番号'] - 1] : '（番号なし）');
  });
  return hits.length ? '重複の可能性あり：' + hits.join('、') : '';
}

function buildValues_(d, sheet, index) {
  const units = Number(d.units) || 0;
  const autoUnits = Number(d.autoUnits) || 0;
  const standardUnits = d.standardUnits !== undefined ? Number(d.standardUnits) : units - autoUnits;
  const total = Number(d.total) || 0;

  const noteParts = [];
  if (d.note) noteParts.push(d.note);
  if (d.date3 || d.time3) noteParts.push(`【第3希望】${d.date3 || ''} ${d.time3 || ''}`.trim());
  if (d.extra) noteParts.push(`【エアコン以外のお掃除】${d.extra}`);

  const times = [];
  if (d.time1) times.push(`第1：${d.time1}`);
  if (d.time2) times.push(`第2：${d.time2}`);

  return {
    '受付番号': makeReceiptNo_(sheet, index['受付番号']),
    '受付日時': d.receivedAt,
    'お名前': d.name,
    '電話': d.tel,
    'メール': d.email,
    'エリア': d.pref,
    '郵便番号': d.zip,
    '住所': d.address,
    '通常台数': standardUnits,
    'お掃除機能付き台数': autoUnits,
    '3点セット': '',
    '合計 (税込)': total,
    '通常価格合計': total,
    '第1希望日': d.date1,
    '第2希望日': d.date2,
    '時間帯': times.join(' / '),
    '支払方法': d.payment,
    '駐車場': d.parking,
    '備考': noteParts.join('\n'),
    '紹介コード': d.referral,
    '紹介料 (20%)': d.referral ? Math.round(total * REFERRAL_RATE) : '',
    '3点セット台数': 0,
    '防カビコート台数': Number(d.antiMold) || 0,
    '室外機台数': Number(d.outdoor) || 0,
    'ドレン台数': Number(d.drain) || 0,
    '住居タイプ': d.houseType,
    '住居タイプ根拠': d.houseType ? 'お客様選択（予約フォーム）' : '',
    '重複チェック': findDuplicates_(sheet, index, d)
  };
}

// 先頭が 0 の番号（電話・郵便番号）や、=,+,-,@ で始まる文字列を文字のまま残す
const TEXT_COLUMNS = ['電話', '郵便番号', '受付番号'];
function safe_(header, v) {
  if (v === undefined || v === null) return '';
  if (typeof v !== 'string') return v;
  if (TEXT_COLUMNS.indexOf(header) >= 0 && v !== '') return "'" + v;
  if (/^[=+\-@]/.test(v)) return "'" + v;
  return v;
}

function appendRow_(data) {
  const sheet = getSheet_();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const index = {};
  headers.forEach((h, i) => { if (h) index[h] = i + 1; });

  const values = buildValues_(data, sheet, index);
  data.receiptNo = values['受付番号'];
  const row = headers.map(h => safe_(h, values[h]));
  sheet.appendRow(row);
}

function summary_(d) {
  const wish = [1, 2, 3]
    .map(i => `第${['一', '二', '三'][i - 1]}希望：${d['date' + i] || ''} ${d['time' + i] || ''}`)
    .join('\n');
  return [
    `受付番号：${d.receiptNo || ''}`,
    '',
    '■ ご依頼内容',
    `エアコン台数：${d.units}台（スタンダード ${d.units - d.autoUnits}台／お掃除機能付 ${d.autoUnits}台）`,
    `内訳：${d.breakdown}`,
    `概算金額：${Number(d.total).toLocaleString('ja-JP')}円（税込・駐車場代別）`,
    '',
    '■ 作業日のご希望',
    wish,
    '',
    '■ お客様情報',
    `お名前：${d.name}`,
    `電話番号：${d.tel}`,
    `メールアドレス：${d.email}`,
    `ご住所：〒${d.zip || ''} ${d.pref} ${d.address}`,
    `お住まいのタイプ：${d.houseType}`,
    `お支払い方法：${d.payment}`,
    `駐車場：${d.parking}`,
    `エアコン以外のお掃除：${d.extra || 'なし'}`,
    `ご要望・備考：${d.note || 'なし'}`,
    `紹介コード：${d.referral || 'なし'}`
  ].join('\n');
}

function sendNotify_(d) {
  const to = SETTINGS.notifyTo || Session.getEffectiveUser().getEmail();
  MailApp.sendEmail({
    to: to,
    replyTo: d.email,
    subject: `【新規予約】${d.name}様 エアコン${d.units}台（${d.pref}）`,
    body: `予約フォームから新しいご予約がありました。\n受付日時：${d.receivedAt}\n\n${summary_(d)}\n`
  });
}

function sendConfirmation_(d) {
  const body = [
    `${d.name} 様`,
    '',
    `このたびは${SETTINGS.companyName}のエアコンクリーニングをご予約いただき、誠にありがとうございます。`,
    '以下の内容でご予約を受け付けました。',
    '',
    summary_(d),
    '',
    '作業日は、ご予約から5営業日以内に担当者よりご連絡し、確定いたします。',
    'ご希望日での対応が難しい場合は、日程調整のご相談をさせていただきます。',
    '',
    '※本メールは送信専用です。ご不明な点はお電話にてお問い合わせください。',
    '',
    '――――――――――――――――',
    SETTINGS.companyName,
    `TEL：${SETTINGS.companyTel}`,
    '――――――――――――――――'
  ].join('\n');

  MailApp.sendEmail({
    to: d.email,
    subject: `【${SETTINGS.companyName}】エアコンクリーニングのご予約を受け付けました`,
    body: body,
    name: SETTINGS.companyName
  });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
