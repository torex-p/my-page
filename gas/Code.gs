/**
 * エアコンクリーニング予約フォーム 受付スクリプト
 * すまいのお掃除専門店
 *
 * 予約フォームから送られた内容を
 *   1. このスプレッドシートの「予約一覧」シートに1行追加
 *   2. 社内向けに通知メールを送信
 *   3. お客様に受付確認メールを送信
 *   4. 「顧客」シートを計算し直す（顧客台帳）
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
    try { rebuildCustomers(); } catch (err) { console.error('顧客台帳の更新に失敗', err); }
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

/* =====================================================================
 * 顧客台帳（「顧客」シート）
 *
 * 予約一覧の全行を見て、同じお客様の予約を1人分にまとめ、
 * 「顧客」シートを毎回まるごと計算し直します。
 *   - 同じお客様の判断：電話番号（数字だけで比較）か、
 *     メールアドレス（大文字・小文字は無視）のどちらかが一致
 *   - 来店回数・初回／最終利用日・累計金額：
 *     「作業確定日」に日付が入っている予約だけを数える
 *   - 顧客IDとメモ欄は、計算し直しても引き継ぐ
 *   - 「経過日数」列はシートの計算式が担当するので書き込まない
 * ===================================================================== */

const CUSTOMER_SHEET = '顧客';
const CUSTOMER_HEADERS = [
  '顧客ID', 'お名前', '電話', 'メール', 'エリア', '住所', '住居タイプ',
  '初回利用日', '最終利用日', '来店回数', '累計金額 (税込)', '最終受付番号',
  '経過日数', 'メモ'
];
const CUSTOMER_FORMULA_COLUMNS = ['経過日数'];

const normEmail_ = v => String(v || '').trim().toLowerCase();

function toDate_(v) {
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  const s = String(v || '').trim();
  if (!s) return null;
  const m = s.match(/^(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})(?:\s+(\d{1,2}):(\d{2}))?/);
  if (!m) return null;
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), Number(m[4] || 0), Number(m[5] || 0));
}

/**
 * 予約の一覧と、いまの顧客シートの中身から、新しい顧客シートの中身を作る。
 * （シートを読み書きしない純粋な計算部分。テストしやすいように分けてある）
 * bookings: [{receiptNo, receivedAt, name, tel, email, area, address, houseType, total, confirmedAt}]
 * existing: [{id, tel, email, memo}]
 * 戻り値:   [{id, name, tel, email, area, address, houseType, first, last, visits, amount, lastReceipt, memo}]
 */
function buildCustomers_(bookings, existing) {
  // 1) 電話・メールでつながる予約をグループにまとめる（つながりを辿る）
  const parent = bookings.map((_, i) => i);
  const find = i => (parent[i] === i ? i : (parent[i] = find(parent[i])));
  const owner = {};
  bookings.forEach((b, i) => {
    [digits_(b.tel) && 'tel:' + digits_(b.tel), normEmail_(b.email) && 'mail:' + normEmail_(b.email)]
      .filter(Boolean)
      .forEach(k => {
        if (owner[k] === undefined) owner[k] = i;
        else parent[find(i)] = find(owner[k]);
      });
  });
  const groups = {};
  bookings.forEach((b, i) => {
    const g = find(i);
    (groups[g] = groups[g] || []).push(b);
  });

  // 2) 既存の顧客IDとメモを、電話・メールから引けるようにする
  const known = {};
  let maxNo = 0;
  existing.forEach(c => {
    const n = Number(String(c.id).replace(/\D/g, ''));
    if (n > maxNo) maxNo = n;
    if (digits_(c.tel)) known['tel:' + digits_(c.tel)] = known['tel:' + digits_(c.tel)] || c;
    if (normEmail_(c.email)) known['mail:' + normEmail_(c.email)] = known['mail:' + normEmail_(c.email)] || c;
  });

  const used = {};
  const result = [];
  Object.keys(groups)
    .map(k => groups[k])
    // 初めて予約した順に並べる（新しい顧客IDもこの順で振る）
    .sort((a, b) => firstTime_(a) - firstTime_(b))
    .forEach(list => {
      const byTime = list.slice().sort((a, b) => time_(a.receivedAt) - time_(b.receivedAt));
      const latest = byTime[byTime.length - 1];

      let prev = null;
      for (const b of byTime) {
        const keys = ['tel:' + digits_(b.tel), 'mail:' + normEmail_(b.email)];
        for (const k of keys) {
          if (known[k] && !used[known[k].id]) { prev = known[k]; break; }
        }
        if (prev) break;
      }
      const id = prev ? prev.id : 'C' + String(++maxNo).padStart(4, '0');
      used[id] = true;

      const done = list.filter(b => toDate_(b.confirmedAt));
      const dates = done.map(b => toDate_(b.confirmedAt)).sort((a, b) => a - b);

      result.push({
        id: id,
        name: latest.name || '',
        tel: latest.tel || '',
        email: latest.email || '',
        area: latest.area || '',
        address: latest.address || '',
        houseType: latest.houseType || '',
        first: dates.length ? dates[0] : '',
        last: dates.length ? dates[dates.length - 1] : '',
        visits: done.length,
        amount: done.reduce((s, b) => s + (Number(b.total) || 0), 0),
        lastReceipt: latest.receiptNo || '',
        memo: prev ? prev.memo || '' : ''
      });
    });

  // 予約一覧から消えた顧客も、名簿からは消さずに残す
  existing.forEach(c => {
    if (c.id && !used[c.id]) {
      used[c.id] = true;
      result.push(Object.assign({ carried: true }, c));
    }
  });

  return result.sort((a, b) => String(a.id).localeCompare(String(b.id)));
}

function time_(v) { const d = toDate_(v); return d ? d.getTime() : 0; }
function firstTime_(list) { return Math.min.apply(null, list.map(b => time_(b.receivedAt) || Infinity)); }

/** 「顧客」シートを、予約一覧から計算し直す */
function rebuildCustomers() {
  const ss = SpreadsheetApp.openById(SETTINGS.spreadsheetId);
  const bookingSheet = ss.getSheetByName(SETTINGS.sheetName);
  const custSheet = ss.getSheetByName(CUSTOMER_SHEET);
  if (!bookingSheet || !custSheet || bookingSheet.getLastRow() < 2) return;

  // 予約一覧を読む（見出し名で列を探す）
  const bv = bookingSheet.getDataRange().getValues();
  const bh = bv[0].map(h => String(h).trim());
  const col = name => bh.indexOf(name);
  const pick = (r, name) => (col(name) >= 0 ? r[col(name)] : '');
  const bookings = bv.slice(1)
    .filter(r => pick(r, 'お名前') || pick(r, '電話') || pick(r, 'メール'))
    .map(r => ({
      receiptNo: pick(r, '受付番号'),
      receivedAt: pick(r, '受付日時'),
      name: pick(r, 'お名前'),
      tel: pick(r, '電話'),
      email: pick(r, 'メール'),
      area: pick(r, 'エリア'),
      address: pick(r, '住所'),
      houseType: pick(r, '住居タイプ'),
      total: pick(r, '合計 (税込)'),
      confirmedAt: pick(r, '作業確定日')
    }));

  // いまの顧客シートを読む（IDとメモを引き継ぐため）
  const ch = custSheet.getRange(1, 1, 1, custSheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const cidx = {};
  ch.forEach((h, i) => { if (h) cidx[h] = i; });
  // 経過日数の列は計算式の結果なので、見出しが計算式から出ていても位置で特定する
  const cRows = custSheet.getLastRow() > 1
    ? custSheet.getRange(2, 1, custSheet.getLastRow() - 1, ch.length).getValues()
    : [];
  const existing = cRows
    .filter(r => r[cidx['顧客ID']])
    .map(r => ({
      id: r[cidx['顧客ID']], name: r[cidx['お名前']], tel: r[cidx['電話']], email: r[cidx['メール']],
      area: r[cidx['エリア']], address: r[cidx['住所']], houseType: r[cidx['住居タイプ']],
      first: r[cidx['初回利用日']], last: r[cidx['最終利用日']], visits: r[cidx['来店回数']],
      amount: r[cidx['累計金額 (税込)']], lastReceipt: r[cidx['最終受付番号']], memo: r[cidx['メモ']]
    }));

  const customers = buildCustomers_(bookings, existing);

  const field = {
    '顧客ID': 'id', 'お名前': 'name', '電話': 'tel', 'メール': 'email', 'エリア': 'area',
    '住所': 'address', '住居タイプ': 'houseType', '初回利用日': 'first', '最終利用日': 'last',
    '来店回数': 'visits', '累計金額 (税込)': 'amount', '最終受付番号': 'lastReceipt', 'メモ': 'memo'
  };

  // 計算式の列（経過日数）を避けて、列ごとに書き込む
  const oldCount = cRows.length;
  const rowsToWrite = Math.max(customers.length, oldCount);
  if (rowsToWrite === 0) return;
  ch.forEach((h, i) => {
    if (!field[h] || CUSTOMER_FORMULA_COLUMNS.indexOf(h) >= 0) return;
    const values = [];
    for (let r = 0; r < rowsToWrite; r++) {
      const c = customers[r];
      let v = c ? c[field[h]] : '';
      if (v === undefined || v === null) v = '';
      if (h === '電話' && v !== '') v = String(v);
      values.push([v]);
    }
    custSheet.getRange(2, i + 1, rowsToWrite, 1).setValues(values);
  });
}

/** 予約一覧が手で書き換えられたら、顧客シートを計算し直す（トリガーから呼ばれる） */
function onBookingEdit(e) {
  try {
    if (e && e.range && e.range.getSheet().getName() !== SETTINGS.sheetName) return;
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(20000)) return;
    try { rebuildCustomers(); } finally { lock.releaseLock(); }
  } catch (err) {
    console.error(err);
  }
}

/** 最初に1回だけ実行：予約一覧を書き換えたときに onBookingEdit が動くよう登録する */
function setupTriggers() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'onBookingEdit')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('onBookingEdit')
    .forSpreadsheet(SETTINGS.spreadsheetId)
    .onEdit()
    .create();
  rebuildCustomers();
}
