// ===== たすかる 予約・売上管理システム（全コード） =====

// ----- Config.gs -----
/**
 * 設定。屋号・連絡先・料金・通知先はここだけ変えれば全体に反映されます。
 */
const CONFIG = {
  shop: {
    name: 'たすかるエアコンお掃除専門店',
    brandShort: 'たすかる',   // 売上データの「案件元」「施工店」に入る自社名
    company: '株式会社じぶんデザイン',
    phone: '090-3839-7267',
    replyTo: '',            // 空ならスクリプト所有者のアドレスから返信
    replyName: 'たすかるエアコンお掃除専門店',
  },

  // 予約通知・レポートの送付先（複数可）
  adminEmails: ['●●●@gmail.com'],  // 公開用の控えなので伏せ字。本物は Apps Script 側にあります

  // 料金（税込）。予約フォームの概算と、確認メールの金額の両方に使います。
  prices: {
    unit: 8900,          // 壁掛け 1台
    cleaningFunc: 6000,  // お掃除機能付き 加算/台
    options: {
      set3:     { label: '3点セット（防カビ抗菌コート＋室外機＋ドレンホース）', price: 7000 },
      antibac:  { label: '防カビ・抗菌コート', price: 3000 },
      outdoor:  { label: '室外機クリーニング', price: 4000 },
      drain:    { label: 'ドレンホース洗浄', price: 2500 },
    },
  },

  areas: ['東京', '神奈川', '埼玉', '千葉', '愛知', '岐阜', '大阪', '兵庫', '和歌山', '福岡', '佐賀', '熊本', '宮城', 'その他'],

  // エリア（都府県）→ 売上表の地域区分
  regionOf: {
    '東京': '関東', '神奈川': '関東', '埼玉': '関東', '千葉': '関東',
    '愛知': '中部', '岐阜': '中部',
    '大阪': '関西', '兵庫': '関西', '和歌山': '関西',
    '福岡': '九州', '佐賀': '九州', '熊本': '九州',
    '宮城': '東北',
  },

  payments: ['現金', 'PayPay', 'クレジットカード', 'NP後払い', '楽天ペイ', '振込み', 'PayPal請求書'],
  media: ['おうちにプロ', 'ユアマイスター', 'ミツモア', 'くらしのマーケット', 'LP（ペライチ）', 'HP', '公式LINE', 'スタッフ直接紹介', '電話予約', 'その他'],
  statuses: ['新規受付', '日程調整中', '日程確定', '作業完了', 'キャンセル'],

  // エアコンクリーニングの推奨周期。顧客シートの「次回ご案内目安」に使います。
  repeatIntervalMonths: 12,

  sheets: {
    bookings: '予約',
    sales: '売上データ',
    customers: '顧客',
    partners: '提携先',
    log: 'エラーログ',
  },

  timezone: 'Asia/Tokyo',
};

// ----- Schema.gs -----
/**
 * 各シートの列。順番を変えるときは setup() を実行し直す前に既存データの列も合わせてください。
 */
const SCHEMA = {
  bookings: [
    '予約ID', '受付日時', '受付曜日', '受付時間帯', 'ステータス', '顧客ID', '初回/リピート',
    'お名前', '電話番号', 'メールアドレス', 'エリア', '地域', '郵便番号', 'ご住所', 'お住まいのタイプ',
    '台数', 'うちお掃除機能付き', '3点セット', '防カビ抗菌コート', '室外機', 'ドレンホース', 'オプション計（台）', '概算金額',
    '第一希望日', '第一希望時間', '第二希望日', '第二希望時間', '第三希望日', '第三希望時間',
    '支払方法', '駐車場', 'その他お掃除の相談', 'ご要望・備考', '紹介コード',
    '流入元', '流入媒体', 'キャンペーン', '参照元URL',
    '担当者', '確定作業日', '確定金額', '売上計上', '管理メモ',
  ],
  sales: [
    '売上ID', '作業日', '登録元', '案件元', '施工店', 'お客様名', 'エリア', '地域', '媒体', '支払方法',
    '台数', '売上金額', '担当者', '金額変更有無', '備考', '予約ID', '登録日時',
  ],
  customers: [
    '顧客ID', 'お名前', '電話番号', 'メールアドレス', 'エリア', '初回予約日', '最終予約日', '予約回数',
    '作業完了回数', '累計売上', '最終作業日', '次回ご案内目安', '初回の流入元', '初回の紹介コード',
  ],
  partners: ['提携先名', 'アクセスコード', '有効', '連絡先メール', '備考', '最終登録日時'],
  log: ['日時', '処理', '内容'],
};

// スタッフが手で入力する列（予約シートで色を変えます）
const BOOKING_STAFF_COLUMNS = ['ステータス', '担当者', '確定作業日', '確定金額', '管理メモ'];

// ----- Util.gs -----
function ss_() {
  return SpreadsheetApp.getActiveSpreadsheet() ||
    SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'));
}

function sheet_(key) {
  const sh = ss_().getSheetByName(CONFIG.sheets[key]);
  if (!sh) throw new Error('シート「' + CONFIG.sheets[key] + '」がありません。setup() を実行してください。');
  return sh;
}

/** 見出し名 → 列番号(1始まり) */
function colMap_(key) {
  const map = {};
  SCHEMA[key].forEach((h, i) => { map[h] = i + 1; });
  return map;
}

/** 見出し名をキーにしたオブジェクトを、列順の配列に変換 */
function toRow_(key, obj) {
  return SCHEMA[key].map(h => (obj[h] === undefined || obj[h] === null) ? '' : cell_(obj[h]));
}

/** シートの全データを見出し名キーのオブジェクト配列で返す（_row に行番号） */
function readAll_(key) {
  const sh = sheet_(key);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const values = sh.getRange(2, 1, last - 1, SCHEMA[key].length).getValues();
  return values.map((r, i) => {
    const o = { _row: i + 2 };
    SCHEMA[key].forEach((h, j) => { o[h] = r[j]; });
    return o;
  });
}

/** フォームから来た文字列を整える（制御文字の除去・前後の空白除去・長さ制限） */
function text_(v, maxLen) {
  if (v === undefined || v === null) return '';
  let s = String(v).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').trim();
  if (maxLen && s.length > maxLen) s = s.slice(0, maxLen);
  return s;
}

/**
 * シートに書く直前の処理。
 * ・=,+,-,@ で始まる値は数式として実行されないよう先頭に ' を付ける
 * ・電話番号や郵便番号の先頭0が消えないよう、数字だけの値も文字列として書く
 */
function cell_(s) {
  if (typeof s !== 'string') return s;
  if (/^[=+\-@]/.test(s) || /^[0-9][0-9\-\s]*$/.test(s)) return "'" + s;
  return s;
}

function int_(v, min, max) {
  const n = Math.floor(Number(v));
  if (!isFinite(n)) return min;
  return Math.min(max, Math.max(min, n));
}

function fmt_(d, pattern) {
  return Utilities.formatDate(d, CONFIG.timezone, pattern);
}

function yen_(n) {
  return Number(n || 0).toLocaleString('ja-JP');
}

function digits_(s) {
  return String(s || '').replace(/[^0-9]/g, '');
}

function newId_(prefix) {
  const rand = Utilities.getUuid().replace(/-/g, '').slice(0, 4).toUpperCase();
  return prefix + '-' + fmt_(new Date(), 'yyMMdd-HHmm') + '-' + rand;
}

function logError_(where, err) {
  try {
    sheet_('log').appendRow([new Date(), where, String(err && err.stack || err)]);
  } catch (e) {
    console.error(where, err);
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function escapeHtml_(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// ----- Api.gs -----
/**
 * Webアプリの入口。
 *   …/exec              → お客様用 予約フォーム
 *   …/exec?page=partner → 提携先用 売上登録ページ
 * GitHub Pages など外部に置いたフォームからは doPost(JSON) で受け付けます。
 */
function doGet(e) {
  const p = (e && e.parameter) || {};
  const isPartner = p.page === 'partner';
  const file = isPartner ? 'Partner' : 'Booking';
  const title = isPartner ? '提携先 売上登録｜' + CONFIG.shop.name : 'エアコンクリーニング ご予約フォーム｜' + CONFIG.shop.name;

  // GAS 上で表示するときは URL パラメータ（utm_source 等）が iframe 内から読めないため、サーバー側で渡す
  const boot = {
    params: p,
    shop: CONFIG.shop,
    prices: CONFIG.prices,
    areas: CONFIG.areas,
    payments: CONFIG.payments,
    media: CONFIG.media,
  };
  const html = HtmlService.createHtmlOutputFromFile(file).getContent()
    .replace('/*__BOOT__*/null', JSON.stringify(boot).replace(/</g, '\\u003c'));

  return HtmlService.createHtmlOutput(html)
    .setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return json_({ ok: false, error: '送信データを読み取れませんでした。' });
  }
  switch (body.action) {
    case 'booking':        return json_(submitBooking(body.data));
    case 'partnerLogin':   return json_(partnerLogin(body.code));
    case 'partnerSubmit':  return json_(partnerSubmit(body.code, body.rows));
    case 'partnerHistory': return json_(partnerHistory(body.code, body.month));
    default:               return json_({ ok: false, error: '不明な操作です。' });
  }
}

// ----- Booking.gs -----
const TIME_SLOTS = ['9時〜12時', '12時〜15時', '15時〜18時', '上記以外で相談したい'];
const HOUSING = ['戸建て', 'マンション・アパート', 'その他（店舗・事務所など）'];
const PARKING = ['敷地内に駐車可', '近隣コインパーキング利用', '相談したい'];
const BOOKING_PAYMENTS = ['作業当日にお支払い（現金）', '作業当日にお支払い（PayPay）', '作業当日にお支払い（クレジットカード）'];
const OTHER_CLEANING = ['お風呂・浴室', 'キッチン・水回り', '洗濯機', 'その他'];

/** 予約フォームの送信を受け付ける（google.script.run / doPost 共通） */
function submitBooking(raw) {
  try {
    raw = raw || {};
    // スパム対策：人には見えない入力欄に値が入っていたら、記録せずに成功を返す
    if (raw.website) return { ok: true, bookingId: '-', total: 0 };

    const v = validateBooking_(raw);
    if (v.errors.length) return { ok: false, error: v.errors.join('\n') };
    const d = v.data;

    // 二重送信の防止（同じ内容が10分以内に来たら同じ予約IDを返す）
    const cache = CacheService.getScriptCache();
    const dupKey = 'bk_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(
      Utilities.DigestAlgorithm.MD5, [d.email, digits_(d.phone), d.date1, d.units].join('|')));
    const prev = cache.get(dupKey);
    if (prev) return JSON.parse(prev);

    const price = calcPrice_(d);
    const now = new Date();
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    let bookingId, customer;
    try {
      customer = upsertCustomerOnBooking_(d, now);
      bookingId = newId_('TSK');
      const row = {
        '予約ID': bookingId,
        '受付日時': now,
        '受付曜日': '日月火水木金土'.charAt(now.getDay()),
        '受付時間帯': fmt_(now, 'H') + '時台',
        'ステータス': '新規受付',
        '顧客ID': customer.id,
        '初回/リピート': customer.isRepeat ? 'リピート' : '初回',
        'お名前': d.name, '電話番号': d.phone, 'メールアドレス': d.email,
        'エリア': d.area, '地域': CONFIG.regionOf[d.area] || 'その他',
        '郵便番号': d.zip, 'ご住所': d.address, 'お住まいのタイプ': d.housing,
        '台数': d.units, 'うちお掃除機能付き': d.cleaningFunc,
        '3点セット': d.opt.set3, '防カビ抗菌コート': d.opt.antibac, '室外機': d.opt.outdoor, 'ドレンホース': d.opt.drain,
        'オプション計（台）': d.opt.set3 + d.opt.antibac + d.opt.outdoor + d.opt.drain,
        '概算金額': price.total,
        '第一希望日': d.date1, '第一希望時間': d.time1,
        '第二希望日': d.date2, '第二希望時間': d.time2,
        '第三希望日': d.date3, '第三希望時間': d.time3,
        '支払方法': d.payment, '駐車場': d.parking,
        'その他お掃除の相談': d.otherCleaning.join('、'),
        'ご要望・備考': d.note, '紹介コード': d.referral,
        '流入元': d.utmSource, '流入媒体': d.utmMedium, 'キャンペーン': d.utmCampaign, '参照元URL': d.referrer,
      };
      sheet_('bookings').appendRow(toRow_('bookings', row));
    } finally {
      lock.releaseLock();
    }

    const result = { ok: true, bookingId: bookingId, total: price.total };
    cache.put(dupKey, JSON.stringify(result), 600);

    try { sendCustomerReply_(bookingId, d, price); } catch (err) { logError_('自動返信メール ' + bookingId, err); }
    try { sendAdminNotice_(bookingId, d, price, customer); } catch (err) { logError_('管理者通知 ' + bookingId, err); }
    return result;
  } catch (err) {
    logError_('submitBooking', err);
    return { ok: false, error: '送信中にエラーが発生しました。お手数ですがお電話（' + CONFIG.shop.phone + '）でご連絡ください。' };
  }
}

function validateBooking_(r) {
  const errors = [];
  const d = {
    name: text_(r.name, 60),
    phone: text_(r.phone, 20),
    email: text_(r.email, 120).toLowerCase(),
    area: text_(r.area, 20),
    zip: text_(r.zip, 10),
    address: text_(r.address, 200),
    housing: text_(r.housing, 30),
    units: int_(r.units, 0, 20),
    cleaningFunc: 0,
    opt: {},
    date1: text_(r.date1, 10), time1: text_(r.time1, 20),
    date2: text_(r.date2, 10), time2: text_(r.time2, 20),
    date3: text_(r.date3, 10), time3: text_(r.time3, 20),
    payment: text_(r.payment, 60),
    parking: text_(r.parking, 30),
    otherCleaning: (Array.isArray(r.otherCleaning) ? r.otherCleaning : []).map(x => text_(x, 30)).filter(x => OTHER_CLEANING.indexOf(x) >= 0),
    note: text_(r.note, 1000),
    referral: text_(r.referral, 40),
    utmSource: text_(r.utmSource, 60), utmMedium: text_(r.utmMedium, 60), utmCampaign: text_(r.utmCampaign, 80),
    referrer: text_(r.referrer, 300),
  };
  d.cleaningFunc = int_(r.cleaningFunc, 0, d.units);
  Object.keys(CONFIG.prices.options).forEach(k => { d.opt[k] = int_(r.opt && r.opt[k], 0, d.units); });

  if (!d.name) errors.push('お名前を入力してください。');
  if (digits_(d.phone).length < 10 || digits_(d.phone).length > 11) errors.push('電話番号を正しく入力してください。');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(d.email)) errors.push('メールアドレスを正しく入力してください。');
  if (CONFIG.areas.indexOf(d.area) < 0) errors.push('エリアを選択してください。');
  if (!d.address) errors.push('ご住所を入力してください。');
  if (HOUSING.indexOf(d.housing) < 0) errors.push('お住まいのタイプを選択してください。');
  if (d.units < 1) errors.push('エアコン台数を1台以上にしてください。');
  if (d.payment && BOOKING_PAYMENTS.indexOf(d.payment) < 0) errors.push('お支払い方法を選択し直してください。');
  if (d.parking && PARKING.indexOf(d.parking) < 0) errors.push('駐車場を選択し直してください。');

  const today = fmt_(new Date(), 'yyyy-MM-dd');
  [['第一', 'date1', 'time1'], ['第二', 'date2', 'time2'], ['第三', 'date3', 'time3']].forEach(([label, dk, tk]) => {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(d[dk])) errors.push('作業日' + label + '希望を選択してください。');
    else if (d[dk] < today) errors.push('作業日' + label + '希望に過去の日付が入っています。');
    if (TIME_SLOTS.indexOf(d[tk]) < 0) errors.push('作業日' + label + '希望時間を選択してください。');
  });
  return { errors: errors, data: d };
}

function calcPrice_(d) {
  const P = CONFIG.prices;
  const lines = [];
  lines.push({ label: 'エアコンクリーニング（壁掛け）', qty: d.units, unit: P.unit });
  if (d.cleaningFunc) lines.push({ label: 'お掃除機能付き 加算', qty: d.cleaningFunc, unit: P.cleaningFunc });
  Object.keys(P.options).forEach(k => {
    if (d.opt[k]) lines.push({ label: P.options[k].label, qty: d.opt[k], unit: P.options[k].price });
  });
  lines.forEach(l => { l.amount = l.qty * l.unit; });
  return { lines: lines, total: lines.reduce((s, l) => s + l.amount, 0) };
}

/** 電話番号で顧客を特定し、新規なら追加、既存なら予約回数を更新 */
function upsertCustomerOnBooking_(d, now) {
  const sh = sheet_('customers');
  const c = colMap_('customers');
  const all = readAll_('customers');
  const key = digits_(d.phone);
  const hit = all.find(x => digits_(x['電話番号']) === key);
  if (hit) {
    sh.getRange(hit._row, c['お名前']).setValue(cell_(d.name));
    sh.getRange(hit._row, c['メールアドレス']).setValue(cell_(d.email));
    sh.getRange(hit._row, c['エリア']).setValue(cell_(d.area));
    sh.getRange(hit._row, c['最終予約日']).setValue(now);
    sh.getRange(hit._row, c['予約回数']).setValue(Number(hit['予約回数'] || 0) + 1);
    return { id: hit['顧客ID'], isRepeat: true };
  }
  const maxNo = all.reduce((m, x) => Math.max(m, Number(String(x['顧客ID']).replace(/\D/g, '')) || 0), 0);
  const id = 'C' + ('00000' + (maxNo + 1)).slice(-5);
  sh.appendRow(toRow_('customers', {
    '顧客ID': id, 'お名前': d.name, '電話番号': d.phone, 'メールアドレス': d.email, 'エリア': d.area,
    '初回予約日': now, '最終予約日': now, '予約回数': 1, '作業完了回数': 0, '累計売上': 0,
    '初回の流入元': d.utmSource || (d.referral ? '紹介' : ''), '初回の紹介コード': d.referral,
  }));
  return { id: id, isRepeat: false };
}

function bookingSummaryText_(d, price) {
  const dateLine = (dt, tm) => {
    const x = new Date(dt + 'T00:00:00+09:00');
    return Utilities.formatDate(x, CONFIG.timezone, 'yyyy年M月d日') + '（' + '日月火水木金土'.charAt(x.getDay()) + '） ' + tm;
  };
  return [
    '■ ご依頼内容',
    ...price.lines.map(l => '・' + l.label + '　' + l.qty + '台 × ' + yen_(l.unit) + '円 ＝ ' + yen_(l.amount) + '円'),
    '概算金額：' + yen_(price.total) + '円（税込・駐車場代別）',
    '',
    '■ ご希望日時',
    '第一希望：' + dateLine(d.date1, d.time1),
    '第二希望：' + dateLine(d.date2, d.time2),
    '第三希望：' + dateLine(d.date3, d.time3),
    '',
    '■ お客様情報',
    'お名前：' + d.name + ' 様',
    '電話番号：' + d.phone,
    'ご住所：' + (d.zip ? '〒' + d.zip + ' ' : '') + d.address,
    'お住まいのタイプ：' + d.housing,
    'お支払い方法：' + (d.payment || '未選択'),
    '駐車場：' + (d.parking || '未選択'),
    d.otherCleaning.length ? 'その他のお掃除のご相談：' + d.otherCleaning.join('、') : '',
    d.note ? 'ご要望・備考：' + d.note : '',
  ].filter(x => x !== '').join('\n');
}

function sendCustomerReply_(bookingId, d, price) {
  const body = [
    d.name + ' 様',
    '',
    'このたびは' + CONFIG.shop.name + 'にご予約いただき、ありがとうございます。',
    '以下の内容でご予約を受け付けました。',
    '',
    '予約番号：' + bookingId,
    '',
    bookingSummaryText_(d, price),
    '',
    '■ 今後の流れ',
    '作業日は、ご予約から5営業日以内に担当者よりご連絡のうえ確定いたします。',
    'ご希望日での対応が難しい場合は、1〜2日以内に日程再調整のご連絡をいたします。',
    '',
    '※ 金額は概算です。汚れによる追加料金は基本的に発生しません。特殊な事情がある場合は、作業前に必ずご説明いたします。',
    '※ このメールは送信専用の自動返信です。ご不明点はお電話でお問い合わせください。',
    '',
    '────────────────',
    CONFIG.shop.name,
    '☎ ' + CONFIG.shop.phone,
    CONFIG.shop.company,
    '────────────────',
  ].join('\n');
  const opt = { name: CONFIG.shop.replyName };
  if (CONFIG.shop.replyTo) opt.replyTo = CONFIG.shop.replyTo;
  MailApp.sendEmail(d.email, '【' + CONFIG.shop.name + '】ご予約を受け付けました（予約番号 ' + bookingId + '）', body, opt);
}

function sendAdminNotice_(bookingId, d, price, customer) {
  if (!CONFIG.adminEmails.length) return;
  const body = [
    '新しい予約が入りました。',
    '',
    '予約番号：' + bookingId,
    '顧客ID：' + customer.id + '（' + (customer.isRepeat ? 'リピート' : '初回') + '）',
    'エリア：' + d.area,
    'メール：' + d.email,
    d.referral ? '紹介コード：' + d.referral : '',
    d.utmSource ? '流入元：' + d.utmSource + (d.utmMedium ? ' / ' + d.utmMedium : '') : '',
    '',
    bookingSummaryText_(d, price),
    '',
    'スプレッドシート：' + ss_().getUrl(),
  ].filter(x => x !== '').join('\n');
  MailApp.sendEmail(CONFIG.adminEmails.join(','),
    '【新規予約】' + d.name + '様 ' + d.area + ' ' + d.units + '台 ' + yen_(price.total) + '円', body);
}

// ----- Sales.gs -----
/**
 * 予約シートでステータスを「作業完了」にし、確定金額が入っていれば売上データに自動で計上します。
 * setup() がこの関数を「編集時」トリガーとして登録します。
 */
function onBookingEdit(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    if (sh.getName() !== CONFIG.sheets.bookings) return;
    const c = colMap_('bookings');
    const watch = [c['ステータス'], c['確定金額'], c['確定作業日'], c['担当者']];
    const firstCol = e.range.getColumn(), lastCol = firstCol + e.range.getNumColumns() - 1;
    if (!watch.some(col => col >= firstCol && col <= lastCol)) return;

    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      for (let r = e.range.getRow(); r < e.range.getRow() + e.range.getNumRows(); r++) {
        if (r < 2) continue;
        recordSaleFromBookingRow_(sh, r, c);
      }
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    logError_('onBookingEdit', err);
  }
}

function recordSaleFromBookingRow_(sh, r, c) {
  const vals = sh.getRange(r, 1, 1, SCHEMA.bookings.length).getValues()[0];
  const b = {};
  SCHEMA.bookings.forEach((h, i) => { b[h] = vals[i]; });
  if (b['ステータス'] !== '作業完了' || b['売上計上']) return;
  const amount = Number(b['確定金額']);
  if (!amount || amount <= 0) return;

  const workDate = b['確定作業日'] instanceof Date ? b['確定作業日'] : new Date();
  const saleId = newId_('S');
  const est = Number(b['概算金額']) || 0;
  sheet_('sales').appendRow(toRow_('sales', {
    '売上ID': saleId,
    '作業日': workDate,
    '登録元': '自社予約フォーム',
    '案件元': CONFIG.shop.brandShort,
    '施工店': CONFIG.shop.brandShort,
    'お客様名': String(b['お名前']),
    'エリア': String(b['エリア']),
    '地域': String(b['地域']),
    '媒体': String(b['流入元'] || '自社予約フォーム'),
    '支払方法': normalizePayment_(b['支払方法']),
    '台数': Number(b['台数']) || '',
    '売上金額': amount,
    '担当者': String(b['担当者']),
    '金額変更有無': amount !== est ? '有' : '無',
    '備考': amount !== est ? '概算 ' + yen_(est) + '円 → 確定 ' + yen_(amount) + '円' : '',
    '予約ID': String(b['予約ID']),
    '登録日時': new Date(),
  }));
  sh.getRange(r, c['売上計上']).setValue(saleId);
  updateCustomerOnComplete_(b['顧客ID'], amount, workDate);
}

function normalizePayment_(p) {
  p = String(p || '');
  if (p.indexOf('現金') >= 0) return '現金';
  if (p.indexOf('PayPay') >= 0) return 'PayPay';
  if (p.indexOf('クレジットカード') >= 0) return 'クレジットカード';
  if (p.indexOf('PayPal') >= 0) return 'PayPal請求書';
  return p;
}

function updateCustomerOnComplete_(customerId, amount, workDate) {
  if (!customerId) return;
  const sh = sheet_('customers');
  const c = colMap_('customers');
  const hit = readAll_('customers').find(x => x['顧客ID'] === customerId);
  if (!hit) return;
  const next = new Date(workDate);
  next.setMonth(next.getMonth() + CONFIG.repeatIntervalMonths);
  sh.getRange(hit._row, c['作業完了回数']).setValue(Number(hit['作業完了回数'] || 0) + 1);
  sh.getRange(hit._row, c['累計売上']).setValue(Number(hit['累計売上'] || 0) + amount);
  sh.getRange(hit._row, c['最終作業日']).setValue(workDate);
  sh.getRange(hit._row, c['次回ご案内目安']).setValue(next);
}

// ----- Partner.gs -----
/**
 * 提携先用の売上登録窓口。
 * 提携先ごとに「提携先」シートのアクセスコードを発行し、そのコードで登録します。
 */
const PARTNER_MAX_ROWS = 30;

function findPartner_(code) {
  code = text_(code, 40);
  if (!code) return null;
  // 総当たり対策：失敗が10分で30回を超えたら一時停止
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('partner_fail') || 0);
  if (fails > 30) throw new Error('しばらく時間をおいてからお試しください。');
  const p = readAll_('partners').find(x => String(x['アクセスコード']) === code && x['有効'] === true);
  if (!p) cache.put('partner_fail', String(fails + 1), 600);
  return p || null;
}

function partnerLogin(code) {
  try {
    const p = findPartner_(code);
    if (!p) return { ok: false, error: 'アクセスコードが正しくないか、無効になっています。' };
    return { ok: true, partner: p['提携先名'], areas: CONFIG.areas, payments: CONFIG.payments, media: CONFIG.media };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

function partnerSubmit(code, rows) {
  try {
    const p = findPartner_(code);
    if (!p) return { ok: false, error: 'アクセスコードが正しくないか、無効になっています。' };
    if (!Array.isArray(rows) || !rows.length) return { ok: false, error: '登録する売上がありません。' };
    if (rows.length > PARTNER_MAX_ROWS) return { ok: false, error: '一度に登録できるのは' + PARTNER_MAX_ROWS + '件までです。' };

    const errors = [];
    const clean = rows.map((r, i) => {
      const n = i + 1;
      const o = {
        date: text_(r.date, 10),
        customer: text_(r.customer, 60),
        area: text_(r.area, 20),
        media: text_(r.media, 40),
        payment: text_(r.payment, 30),
        units: int_(r.units, 0, 50),
        amount: int_(r.amount, -1, 10000000),
        staff: text_(r.staff, 40),
        changed: r.changed === '有' ? '有' : '無',
        note: text_(r.note, 500),
      };
      if (!/^\d{4}-\d{2}-\d{2}$/.test(o.date)) errors.push(n + '件目：作業日を入力してください。');
      if (!o.customer) errors.push(n + '件目：お客様名を入力してください。');
      if (CONFIG.areas.indexOf(o.area) < 0) errors.push(n + '件目：エリアを選択してください。');
      if (CONFIG.payments.indexOf(o.payment) < 0) errors.push(n + '件目：支払方法を選択してください。');
      if (o.amount < 0) errors.push(n + '件目：売上金額を入力してください。');
      if (!o.staff) errors.push(n + '件目：担当者名を入力してください。');
      return o;
    });
    if (errors.length) return { ok: false, error: errors.join('\n') };

    const now = new Date();
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    const ids = [];
    try {
      const out = clean.map(o => {
        const id = newId_('S');
        ids.push(id);
        return toRow_('sales', {
          '売上ID': id,
          '作業日': new Date(o.date + 'T00:00:00+09:00'),
          '登録元': '提携先ページ',
          '案件元': p['提携先名'],
          '施工店': CONFIG.shop.brandShort,
          'お客様名': o.customer,
          'エリア': o.area,
          '地域': CONFIG.regionOf[o.area] || 'その他',
          '媒体': o.media,
          '支払方法': o.payment,
          '台数': o.units || '',
          '売上金額': o.amount,
          '担当者': o.staff,
          '金額変更有無': o.changed,
          '備考': o.note,
          '登録日時': now,
        });
      });
      const sh = sheet_('sales');
      sh.getRange(sh.getLastRow() + 1, 1, out.length, SCHEMA.sales.length).setValues(out);
      const pc = colMap_('partners');
      sheet_('partners').getRange(p._row, pc['最終登録日時']).setValue(now);
    } finally {
      lock.releaseLock();
    }
    return { ok: true, count: ids.length, ids: ids };
  } catch (err) {
    logError_('partnerSubmit', err);
    return { ok: false, error: '登録中にエラーが発生しました。時間をおいて再度お試しください。' };
  }
}

/** 提携先が自社分の登録内容を確認するための一覧（指定月 yyyy-MM） */
function partnerHistory(code, month) {
  try {
    const p = findPartner_(code);
    if (!p) return { ok: false, error: 'アクセスコードが正しくないか、無効になっています。' };
    month = /^\d{4}-\d{2}$/.test(String(month)) ? month : fmt_(new Date(), 'yyyy-MM');
    const rows = readAll_('sales')
      .filter(s => s['案件元'] === p['提携先名'] && s['作業日'] instanceof Date && fmt_(s['作業日'], 'yyyy-MM') === month)
      .sort((a, b) => a['作業日'] - b['作業日'])
      .map(s => ({
        date: fmt_(s['作業日'], 'M/d'),
        customer: String(s['お客様名']),
        area: String(s['エリア']),
        payment: String(s['支払方法']),
        amount: Number(s['売上金額']) || 0,
        staff: String(s['担当者']),
      }));
    return { ok: true, month: month, rows: rows, total: rows.reduce((t, r) => t + r.amount, 0) };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

// ----- Report.gs -----
/**
 * 売上分析レポート。
 * 週次：毎週月曜 8時に先週（月〜日）分、月次：毎月1日 8時に前月分を作成し、管理者にメールします。
 * レポートの数字はすべて「売上データ」「予約」シートを参照する数式なので、
 * 後から売上を追加・修正してもレポートタブの数字は自動で更新されます。
 */

function weeklyReport() {
  const today = startOfDay_(new Date());
  const dow = (today.getDay() + 6) % 7;           // 月曜=0
  const start = addDays_(today, -dow - 7);        // 先週の月曜
  const end = addDays_(start, 6);
  buildReport_('週次', start, end, addDays_(start, -7), addDays_(start, -1));
}

function monthlyReport() {
  const today = startOfDay_(new Date());
  const start = new Date(today.getFullYear(), today.getMonth() - 1, 1);
  const end = new Date(today.getFullYear(), today.getMonth(), 0);
  const pStart = new Date(today.getFullYear(), today.getMonth() - 2, 1);
  const pEnd = new Date(today.getFullYear(), today.getMonth() - 1, 0);
  buildReport_('月次', start, end, pStart, pEnd);
}

/** 動作確認用：今月1日〜今日のレポートを作る（メールも送ります） */
function reportThisMonthSoFar() {
  const today = startOfDay_(new Date());
  const start = new Date(today.getFullYear(), today.getMonth(), 1);
  const pStart = new Date(today.getFullYear(), today.getMonth() - 1, 1);
  const pEnd = addDays_(pStart, today.getDate() - 1);
  buildReport_('今月途中', start, today, pStart, pEnd);
}

function startOfDay_(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
function addDays_(d, n) { const x = new Date(d); x.setDate(x.getDate() + n); return x; }

function colLetter_(key, header) {
  let n = SCHEMA[key].indexOf(header) + 1;
  if (n < 1) throw new Error(header + ' 列が見つかりません');
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

function buildReport_(kind, start, end, pStart, pEnd) {
  const book = ss_();
  const S = "'" + CONFIG.sheets.sales + "'!";
  const B = "'" + CONFIG.sheets.bookings + "'!";
  const sc = h => S + '$' + colLetter_('sales', h) + ':$' + colLetter_('sales', h);
  const bc = h => B + '$' + colLetter_('bookings', h) + ':$' + colLetter_('bookings', h);

  const name = kind === '月次'
    ? '月次レポート ' + fmt_(start, 'yyyy-MM')
    : kind + 'レポート ' + fmt_(start, 'MM/dd') + '-' + fmt_(end, 'MM/dd');
  const old = book.getSheetByName(name);
  if (old) book.deleteSheet(old);
  const sh = book.insertSheet(name, 0);

  // 期間（今期 B3:B4、比較期間 B5:B6）。終了日はその日を含みます。
  const inCur = (dateCol) => dateCol + ',">="&$B$3,' + dateCol + ',"<"&($B$4+1)';
  const inPrev = (dateCol) => dateCol + ',">="&$B$5,' + dateCol + ',"<"&($B$6+1)';
  const sDate = sc('作業日'), bDate = bc('受付日時');

  const rows = [];
  const fmtRows = { title: [], header: [], yen: [], pct: [], int: [], section: [], total: [] };
  const push = (r, type) => { rows.push(r.concat(Array(6 - r.length).fill(''))); if (type) fmtRows[type].push(rows.length); };

  push([kind + 'レポート（' + fmt_(start, 'yyyy/M/d') + '〜' + fmt_(end, 'M/d') + '）'], 'title');
  push(['数字は「売上データ」「予約」シートから数式で集計しています。売上を追加・修正すると自動で更新されます。']);
  push(['今期 開始日', start]);
  push(['今期 終了日', end]);
  push(['比較期間 開始日', pStart]);
  push(['比較期間 終了日', pEnd]);
  push(['']);

  // ---- KPI
  push(['指標', '今期', '比較期間', '増減', '増減率'], 'header');
  const kpiStart = rows.length + 1;
  const kpi = (label, cur, prev, type) => {
    const r = rows.length + 1;
    push([label, cur, prev, '=B' + r + '-C' + r, '=IF(C' + r + '=0,"-",B' + r + '/C' + r + '-1)'], type);
    return r;
  };
  const rSales = kpi('売上合計（円）', '=SUMIFS(' + sc('売上金額') + ',' + inCur(sDate) + ')', '=SUMIFS(' + sc('売上金額') + ',' + inPrev(sDate) + ')', 'yen');
  const rCnt = kpi('売上件数', '=COUNTIFS(' + inCur(sDate) + ')', '=COUNTIFS(' + inPrev(sDate) + ')', 'int');
  kpi('平均単価（円）', '=IF(B' + rCnt + '=0,0,B' + rSales + '/B' + rCnt + ')', '=IF(C' + rCnt + '=0,0,C' + rSales + '/C' + rCnt + ')', 'yen');
  kpi('台数', '=SUMIFS(' + sc('台数') + ',' + inCur(sDate) + ')', '=SUMIFS(' + sc('台数') + ',' + inPrev(sDate) + ')', 'int');
  kpi('金額変更ありの件数', '=COUNTIFS(' + sc('金額変更有無') + ',"有",' + inCur(sDate) + ')', '=COUNTIFS(' + sc('金額変更有無') + ',"有",' + inPrev(sDate) + ')', 'int');
  const rBk = kpi('予約受付件数（自社フォーム）', '=COUNTIFS(' + inCur(bDate) + ')', '=COUNTIFS(' + inPrev(bDate) + ')', 'int');
  kpi('うち初回のお客様', '=COUNTIFS(' + bc('初回/リピート') + ',"初回",' + inCur(bDate) + ')', '=COUNTIFS(' + bc('初回/リピート') + ',"初回",' + inPrev(bDate) + ')', 'int');
  kpi('うちリピートのお客様', '=COUNTIFS(' + bc('初回/リピート') + ',"リピート",' + inCur(bDate) + ')', '=COUNTIFS(' + bc('初回/リピート') + ',"リピート",' + inPrev(bDate) + ')', 'int');
  kpi('うちキャンセル', '=COUNTIFS(' + bc('ステータス') + ',"キャンセル",' + inCur(bDate) + ')', '=COUNTIFS(' + bc('ステータス') + ',"キャンセル",' + inPrev(bDate) + ')', 'int');
  kpi('予約の概算金額合計（円）', '=SUMIFS(' + bc('概算金額') + ',' + inCur(bDate) + ')', '=SUMIFS(' + bc('概算金額') + ',' + inPrev(bDate) + ')', 'yen');
  kpi('オプション付帯率（オプション台数÷台数）',
    '=IFERROR(SUMIFS(' + bc('オプション計（台）') + ',' + inCur(bDate) + ')/SUMIFS(' + bc('台数') + ',' + inCur(bDate) + '),0)',
    '=IFERROR(SUMIFS(' + bc('オプション計（台）') + ',' + inPrev(bDate) + ')/SUMIFS(' + bc('台数') + ',' + inPrev(bDate) + '),0)', 'pct');
  const kpiEnd = rows.length;
  push(['']);

  // ---- 内訳表（ラベルは今期・比較期間に出てきた値。数字は数式）
  const sales = readAll_('sales');
  const bookings = readAll_('bookings');
  const inRange = (d, a, b) => d instanceof Date && d >= a && d < addDays_(b, 1);
  const curSales = sales.filter(s => inRange(s['作業日'], start, end));
  const anySales = sales.filter(s => inRange(s['作業日'], pStart, pEnd) || inRange(s['作業日'], start, end));
  const anyBk = bookings.filter(b => inRange(b['受付日時'], pStart, pEnd) || inRange(b['受付日時'], start, end));

  const breakdowns = [];
  const salesTable = (title, header) => {
    const sumBy = {};
    curSales.forEach(s => { const k = String(s[header] || ''); sumBy[k] = (sumBy[k] || 0) + (Number(s['売上金額']) || 0); });
    const labels = Array.from(new Set(anySales.map(s => String(s[header] || '')).filter(Boolean)))
      .sort((a, b) => (sumBy[b] || 0) - (sumBy[a] || 0));
    push([title, '売上（円）', '件数', '構成比', '比較期間の売上', '前期比'], 'header');
    const first = rows.length + 1;
    labels.forEach(l => {
      const r = rows.length + 1;
      const crit = sc(header) + ',$A' + r;
      push([l,
        '=SUMIFS(' + sc('売上金額') + ',' + crit + ',' + inCur(sDate) + ')',
        '=COUNTIFS(' + crit + ',' + inCur(sDate) + ')',
        '=IF($B$' + rSales + '=0,0,B' + r + '/$B$' + rSales + ')',
        '=SUMIFS(' + sc('売上金額') + ',' + crit + ',' + inPrev(sDate) + ')',
        '=IF(E' + r + '=0,"-",B' + r + '/E' + r + '-1)']);
    });
    const r = rows.length + 1;
    const last = r - 1;
    const sumOf = col => labels.length ? 'SUM(' + col + first + ':' + col + last + ')' : '0';
    push(['未入力・その他',
      '=$B$' + rSales + '-' + sumOf('B'), '=$B$' + rCnt + '-' + sumOf('C'),
      '=IF($B$' + rSales + '=0,0,B' + r + '/$B$' + rSales + ')',
      '=$C$' + rSales + '-' + sumOf('E'), '=IF(E' + r + '=0,"-",B' + r + '/E' + r + '-1)']);
    breakdowns.push({ title: title, first: first, last: last });
    push(['']);
  };
  salesTable('案件元別', '案件元');
  salesTable('媒体別', '媒体');
  salesTable('地域別', '地域');
  salesTable('担当者別', '担当者');
  salesTable('支払方法別', '支払方法');

  // ---- 予約の傾向（営業・集客向け）
  const bookingTable = (title, header, fixedLabels) => {
    const labels = fixedLabels || Array.from(new Set(anyBk.map(b => String(b[header] || '')).filter(Boolean)));
    push([title, '予約件数', '比較期間', '初回', 'リピート', '概算金額合計（円）'], 'header');
    labels.forEach(l => {
      const r = rows.length + 1;
      const crit = bc(header) + ',$A' + r;
      push([l,
        '=COUNTIFS(' + crit + ',' + inCur(bDate) + ')',
        '=COUNTIFS(' + crit + ',' + inPrev(bDate) + ')',
        '=COUNTIFS(' + crit + ',' + bc('初回/リピート') + ',"初回",' + inCur(bDate) + ')',
        '=COUNTIFS(' + crit + ',' + bc('初回/リピート') + ',"リピート",' + inCur(bDate) + ')',
        '=SUMIFS(' + bc('概算金額') + ',' + crit + ',' + inCur(bDate) + ')']);
    });
    push(['']);
  };
  bookingTable('予約の流入元別（URLの utm_source）', '流入元');
  bookingTable('予約のエリア別', 'エリア');
  bookingTable('予約を受け付けた曜日', '受付曜日', ['月', '火', '水', '木', '金', '土', '日']);

  // ---- 日別推移
  push(['日別の売上', '売上（円）', '件数'], 'header');
  const dailyFirst = rows.length + 1;
  for (let d = new Date(start); d <= end; d = addDays_(d, 1)) {
    const r = rows.length + 1;
    push([new Date(d),
      '=SUMIFS(' + sc('売上金額') + ',' + sDate + ',">="&$A' + r + ',' + sDate + ',"<"&($A' + r + '+1))',
      '=COUNTIFS(' + sDate + ',">="&$A' + r + ',' + sDate + ',"<"&($A' + r + '+1))']);
  }
  const dailyLast = rows.length;

  // ---- 書き込みと書式
  sh.getRange(1, 1, rows.length, 6).setValues(rows.map(r => r.map(v => (typeof v === 'string' && v.charAt(0) !== '=') ? cell_(v) : v)));
  sh.getRange(1, 1, rows.length, 6).setFontFamily('Arial');
  sh.getRange('A1').setFontSize(14).setFontWeight('bold');
  sh.getRange('A2').setFontColor('#666666');
  sh.getRange('B3:B6').setNumberFormat('yyyy/mm/dd').setFontColor('#0000FF').setBackground('#FFFF00');
  fmtRows.header.forEach(r => sh.getRange(r, 1, 1, 6).setBackground('#1F4E78').setFontColor('#FFFFFF').setFontWeight('bold'));
  const YEN = '#,##0;(#,##0);"-"', INT = '#,##0;(#,##0);"-"', PCT = '0.0%;(0.0%);"-"';
  for (let r = kpiStart; r <= kpiEnd; r++) {
    const isPct = fmtRows.pct.indexOf(r) >= 0;
    sh.getRange(r, 2, 1, 3).setNumberFormat(isPct ? PCT : (fmtRows.yen.indexOf(r) >= 0 ? YEN : INT));
    sh.getRange(r, 5).setNumberFormat(PCT);
  }
  breakdowns.forEach(b => {
    const n = b.last - b.first + 2; // 未入力・その他 行を含む
    sh.getRange(b.first, 2, n, 1).setNumberFormat(YEN);
    sh.getRange(b.first, 3, n, 1).setNumberFormat(INT);
    sh.getRange(b.first, 4, n, 1).setNumberFormat(PCT);
    sh.getRange(b.first, 5, n, 1).setNumberFormat(YEN);
    sh.getRange(b.first, 6, n, 1).setNumberFormat(PCT);
  });
  sh.getRange(dailyFirst, 1, dailyLast - dailyFirst + 1, 1).setNumberFormat('m/d（ddd）');
  sh.getRange(dailyFirst, 2, dailyLast - dailyFirst + 1, 1).setNumberFormat(YEN);
  sh.setColumnWidth(1, 300);
  sh.setColumnWidths(2, 5, 120);
  sh.setFrozenRows(1);

  const chart = sh.newChart().setChartType(Charts.ChartType.COLUMN)
    .addRange(sh.getRange(dailyFirst - 1, 1, dailyLast - dailyFirst + 2, 2))
    .setNumHeaders(1)
    .setOption('title', '日別の売上（円）')
    .setOption('legend', { position: 'none' })
    .setPosition(kpiStart, 8, 0, 0)
    .build();
  sh.insertChart(chart);

  SpreadsheetApp.flush();
  sendReportMail_(sh, kind, start, end, kpiStart, kpiEnd, breakdowns);
  return sh.getName();
}

function sendReportMail_(sh, kind, start, end, kpiStart, kpiEnd, breakdowns) {
  if (!CONFIG.adminEmails.length) return;
  const v = sh.getDataRange().getDisplayValues();
  const url = ss_().getUrl() + '#gid=' + sh.getSheetId();
  const td = 'style="padding:4px 10px;border-bottom:1px solid #eee"';
  const tr = cells => '<tr>' + cells.map((c, i) => '<td ' + td + (i ? ' align="right"' : '') + '>' + escapeHtml_(c) + '</td>').join('') + '</tr>';

  let html = '<div style="font-family:Arial,sans-serif;font-size:14px;color:#222">';
  html += '<h2 style="font-size:18px">' + escapeHtml_(v[0][0]) + '</h2>';
  html += '<table style="border-collapse:collapse">' + tr(['指標', '今期', '比較期間', '増減率']);
  for (let r = kpiStart; r <= kpiEnd; r++) html += tr([v[r - 1][0], v[r - 1][1], v[r - 1][2], v[r - 1][4]]);
  html += '</table>';
  breakdowns.forEach(b => {
    html += '<h3 style="font-size:15px;margin-top:20px">' + escapeHtml_(b.title) + '（上位5件）</h3><table style="border-collapse:collapse">';
    html += tr(['', '売上（円）', '件数', '構成比']);
    for (let r = b.first; r <= Math.min(b.last, b.first + 4); r++) html += tr([v[r - 1][0], v[r - 1][1], v[r - 1][2], v[r - 1][3]]);
    html += '</table>';
  });
  html += '<p style="margin-top:20px"><a href="' + url + '">スプレッドシートでレポート全体を見る</a></p></div>';

  MailApp.sendEmail({
    to: CONFIG.adminEmails.join(','),
    subject: '【' + CONFIG.shop.brandShort + '】' + kind + '売上レポート ' + fmt_(start, 'yyyy/M/d') + '〜' + fmt_(end, 'M/d'),
    htmlBody: html,
    body: v[0][0] + '\n' + url,
  });
}

// ----- Setup.gs -----
/**
 * 初回に1回だけ実行してください（再実行しても既存データは消えません）。
 * ・各シートを作成して見出し・入力規則・書式を設定
 * ・提携先シートに提携先とアクセスコードを登録
 * ・編集時トリガー（売上の自動計上）と週次・月次レポートのトリガーを登録
 */
function setup() {
  const book = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', book.getId());
  book.setSpreadsheetTimeZone(CONFIG.timezone);

  Object.keys(SCHEMA).forEach(key => {
    let sh = book.getSheetByName(CONFIG.sheets[key]);
    if (!sh) sh = book.insertSheet(CONFIG.sheets[key]);
    sh.getRange(1, 1, 1, SCHEMA[key].length).setValues([SCHEMA[key]])
      .setBackground('#1F4E78').setFontColor('#FFFFFF').setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getDataRange().setFontFamily('Arial');
  });

  // 予約：スタッフ入力列を黄色に、ステータスはプルダウン、日付・金額の書式
  const bk = book.getSheetByName(CONFIG.sheets.bookings);
  const bcm = colMap_('bookings');
  BOOKING_STAFF_COLUMNS.forEach(h => bk.getRange(1, bcm[h]).setBackground('#BF9000'));
  bk.getRange(2, bcm['ステータス'], bk.getMaxRows() - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(CONFIG.statuses, true).build());
  bk.getRange(2, bcm['確定作業日'], bk.getMaxRows() - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build()).setNumberFormat('yyyy/mm/dd');
  bk.getRange(2, bcm['受付日時'], bk.getMaxRows() - 1, 1).setNumberFormat('yyyy/mm/dd hh:mm');
  ['概算金額', '確定金額'].forEach(h => bk.getRange(2, bcm[h], bk.getMaxRows() - 1, 1).setNumberFormat('#,##0'));
  ['電話番号', '郵便番号'].forEach(h => bk.getRange(2, bcm[h], bk.getMaxRows() - 1, 1).setNumberFormat('@'));

  // 売上データ
  const sa = book.getSheetByName(CONFIG.sheets.sales);
  const scm = colMap_('sales');
  sa.getRange(2, scm['作業日'], sa.getMaxRows() - 1, 1).setNumberFormat('yyyy/mm/dd');
  sa.getRange(2, scm['売上金額'], sa.getMaxRows() - 1, 1).setNumberFormat('#,##0');
  sa.getRange(2, scm['登録日時'], sa.getMaxRows() - 1, 1).setNumberFormat('yyyy/mm/dd hh:mm');

  // 顧客
  const cu = book.getSheetByName(CONFIG.sheets.customers);
  const ccm = colMap_('customers');
  ['初回予約日', '最終予約日', '最終作業日', '次回ご案内目安'].forEach(h =>
    cu.getRange(2, ccm[h], cu.getMaxRows() - 1, 1).setNumberFormat('yyyy/mm/dd'));
  cu.getRange(2, ccm['累計売上'], cu.getMaxRows() - 1, 1).setNumberFormat('#,##0');
  cu.getRange(2, ccm['電話番号'], cu.getMaxRows() - 1, 1).setNumberFormat('@');

  // 提携先
  const pa = book.getSheetByName(CONFIG.sheets.partners);
  const pcm = colMap_('partners');
  // チェックボックスは値 FALSE を持つため getLastRow() では判定できない。提携先名の列で判定する
  const hasPartners = pa.getRange(2, pcm['提携先名'], pa.getMaxRows() - 1, 1).getValues().some(r => r[0] !== '');
  if (!hasPartners) {
    const names = ['ACY', 'エアコンクリーニングさわやか', '笑がお吉日', 'ありがとう', 'BASE', 'DAIZEN', 'アロー電工',
      '生活空間', 'KSS', 'クリーンサポートJOYs', 'お助けsky', 'CLIMB D', 'ジグラットクリーン', '大日本美研',
      'ホームクリーニングBLESS YOU', 'magicart', 'CTW', 'クリーニングネスト'];
    pa.getRange(2, pcm['有効'], 200, 1).insertCheckboxes();
    pa.getRange(2, 1, names.length, 3).setValues(names.map(n => [n, '', true]));
  }
  issueMissingPartnerCodes();

  // 空の初期シートを削除
  ['シート1', 'Sheet1'].forEach(n => {
    const s = book.getSheetByName(n);
    if (s && s.getLastRow() === 0 && book.getSheets().length > 1) book.deleteSheet(s);
  });

  installTriggers_();
  SpreadsheetApp.getActive().toast('セットアップが完了しました。', 'タスカル予約システム', 5);
}

/** 提携先シートでアクセスコードが空の行にコードを発行（提携先を追加したら実行） */
function issueMissingPartnerCodes() {
  const sh = sheet_('partners');
  const c = colMap_('partners');
  const last = sh.getLastRow();
  if (last < 2) return;
  const range = sh.getRange(2, c['アクセスコード'], last - 1, 1);
  const vals = range.getValues();
  const used = new Set(vals.map(v => String(v[0])).filter(Boolean));
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  const names = sh.getRange(2, c['提携先名'], last - 1, 1).getValues();
  vals.forEach((v, i) => {
    if (v[0] || !names[i][0]) return;
    let code;
    do {
      code = '';
      const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid());
      for (let k = 0; k < 10; k++) code += chars.charAt((bytes[k] + 256) % chars.length);
    } while (used.has(code));
    used.add(code);
    v[0] = code;
  });
  range.setValues(vals);
}

function installTriggers_() {
  const wanted = ['onBookingEdit', 'weeklyReport', 'monthlyReport'];
  ScriptApp.getProjectTriggers()
    .filter(t => wanted.indexOf(t.getHandlerFunction()) >= 0)
    .forEach(t => ScriptApp.deleteTrigger(t));
  const book = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger('onBookingEdit').forSpreadsheet(book).onEdit().create();
  ScriptApp.newTrigger('weeklyReport').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).inTimezone(CONFIG.timezone).create();
  ScriptApp.newTrigger('monthlyReport').timeBased().onMonthDay(1).atHour(8).inTimezone(CONFIG.timezone).create();
}
