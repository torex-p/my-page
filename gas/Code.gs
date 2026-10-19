/**
 * エアコンクリーニング予約フォーム 受付スクリプト
 * 株式会社トレックスパートナーズ
 *
 * 予約フォームから送られた内容を
 *   1. このスプレッドシートの「予約一覧」シートに1行追加
 *   2. 社内向けに通知メールを送信
 *   3. お客様に受付確認メールを送信
 * します。
 */

// ===== 設定 =====
const SETTINGS = {
  sheetName: '予約一覧',
  // 通知メールの送り先（空欄ならこのスクリプトをデプロイしたアカウント宛て）
  notifyTo: '',
  companyName: '株式会社トレックスパートナーズ',
  // TODO: 自社の電話番号に差し替え
  companyTel: '00-0000-0000'
};

const COLUMNS = [
  ['受付日時', 'receivedAt'],
  ['対応状況', 'status'],
  ['お名前', 'name'],
  ['電話番号', 'tel'],
  ['メールアドレス', 'email'],
  ['エリア', 'pref'],
  ['郵便番号', 'zip'],
  ['ご住所', 'address'],
  ['お住まいのタイプ', 'houseType'],
  ['第一希望日', 'date1'],
  ['第一希望時間', 'time1'],
  ['第二希望日', 'date2'],
  ['第二希望時間', 'time2'],
  ['第三希望日', 'date3'],
  ['第三希望時間', 'time3'],
  ['台数', 'units'],
  ['うちお掃除機能付き', 'autoUnits'],
  ['防カビ・抗菌コート', 'antiMold'],
  ['室外機クリーニング', 'outdoor'],
  ['ドレンホース洗浄', 'drain'],
  ['内訳', 'breakdown'],
  ['概算金額（円）', 'total'],
  ['お支払い方法', 'payment'],
  ['駐車場', 'parking'],
  ['エアコン以外のお掃除', 'extra'],
  ['ご要望・備考', 'note'],
  ['紹介コード', 'referral']
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const data = JSON.parse(e.postData.contents);
    data.receivedAt = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
    data.status = '未対応';

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

function appendRow_(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SETTINGS.sheetName);
  if (!sheet) sheet = ss.insertSheet(SETTINGS.sheetName);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(COLUMNS.map(c => c[0]));
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, COLUMNS.length).setFontWeight('bold').setBackground('#e8f2fb');
  }
  // 数式として解釈されないよう、=,+,-,@ で始まる文字列の先頭に ' を付ける
  const row = COLUMNS.map(c => {
    const v = data[c[1]];
    if (v === undefined || v === null) return '';
    if (typeof v === 'string' && /^[=+\-@]/.test(v)) return "'" + v;
    return v;
  });
  sheet.appendRow(row);
}

function summary_(d) {
  const wish = [1, 2, 3]
    .map(i => `第${['一', '二', '三'][i - 1]}希望：${d['date' + i] || ''} ${d['time' + i] || ''}`)
    .join('\n');
  return [
    '■ ご依頼内容',
    `エアコン台数：${d.units}台（うちお掃除機能付き ${d.autoUnits}台）`,
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
