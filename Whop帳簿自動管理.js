// ============================================================
//  WHOP サブスク帳簿自動化スクリプト v4（完全版）
// ============================================================

// ===== 設定 =====
const CONFIG = {
  SHEET_LEDGER:  'サブスク管理･天底極致',
  SHEET_BY_USER: 'ユーザー別サマリー',
  SHEET_MONTHLY: '月次サマリー',
  SHEET_YEARLY:  '年次サマリー',

  SECRET_TOKEN: 'Kk505011919',   // ← 変更済みか確認

  WHOP_API_BASE: 'https://api.whop.com/api/v1',
  // APIキー → スクリプトプロパティ「WHOP_API_KEY」に保存

  PARTNER_SHARE: 0.5,

  PLAN_MONTHLY_IDS: ['plan_ZdjLLlHv81yqp'],
  PLAN_ANNUAL_IDS:  ['plan_YYYYYYYY'],      // ← 年額プランがあれば更新
};

// ===== 列定義 =====
const COL = {
  DATE_RECORD:    1,  // A
  DATE_PAYMENT:   2,  // B
  PAYMENT_ID:     3,  // C
  USERNAME:       4,  // D
  EMAIL:          5,  // E
  PLAN_TYPE:      6,  // F
  CURRENCY:       7,  // G
  GROSS_JPY:      8,  // H 売上（自分分50%）
  FEE_JPY:        9,  // I 手数料（自分分50%）
  NET_JPY:        10, // J 受取（自分分50%）
  USD_REF:        11, // K USD参考
  BILLING_REASON: 12, // L 請求理由
  STATUS:         13, // M
  USER_ID:        14, // N
};

// サマリーエリアの高さ（行1〜SUMMARY_HEIGHT）
// データヘッダー = SUMMARY_HEIGHT+1行目
// データ開始     = SUMMARY_HEIGHT+2行目
const SUMMARY_HEIGHT = 22;

// ============================================================
//  【設定チェック】最初に実行して設定が正しいか確認する
// ============================================================
function checkSetup() {
  const results = [];
  let allOk = true;

  const ok  = (msg) => { results.push('✅ ' + msg); };
  const err = (msg) => { results.push('❌ ' + msg); allOk = false; };
  const warn = (msg) => { results.push('⚠️ ' + msg); };

  // 1. SECRET_TOKEN
  if (CONFIG.SECRET_TOKEN === 'YOUR_SECRET_TOKEN_HERE') {
    err('SECRET_TOKEN がデフォルトのままです。変更してください。');
  } else {
    ok('SECRET_TOKEN が設定されています');
  }

  // 2. WHOP_API_KEY（スクリプトプロパティ）
  const apiKey = PropertiesService.getScriptProperties().getProperty('WHOP_API_KEY');
  if (!apiKey) {
    err('WHOP_API_KEY がスクリプトプロパティに未設定です');
  } else {
    ok('WHOP_API_KEY が設定されています（' + apiKey.substring(0, 8) + '...）');
  }

  // 3. PLAN_MONTHLY_IDS
  if (CONFIG.PLAN_MONTHLY_IDS.includes('plan_XXXXXXXX')) {
    warn('PLAN_MONTHLY_IDS にデフォルト値が含まれています');
  } else {
    ok('PLAN_MONTHLY_IDS: ' + CONFIG.PLAN_MONTHLY_IDS.join(', '));
  }

  // 4. スプレッドシート接続
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ok('スプレッドシートに接続できています: ' + ss.getName());

    // 5. 各シートの存在確認
    [CONFIG.SHEET_LEDGER, CONFIG.SHEET_BY_USER, CONFIG.SHEET_MONTHLY]
      .forEach(name => {
        const s = ss.getSheetByName(name);
        if (s) ok('シート「' + name + '」が存在します');
        else   err('シート「' + name + '」が見つかりません → initAllSheets を実行してください');
      });

    // 6. サマリーエリア確認
    const ls = ss.getSheetByName(CONFIG.SHEET_LEDGER);
    if (ls) {
      const marker = ls.getRange(1, 1).getValue();
      if (marker === '📊 年度サマリー') {
        ok('サマリーエリアがシート上部に設置されています');
        // データヘッダー行の確認
        const headerRow = SUMMARY_HEIGHT + 1;
        const headerVal = ls.getRange(headerRow, 1).getValue();
        if (headerVal === '記録日時') {
          ok('データヘッダーが行' + headerRow + 'に正しく設置されています');
        } else {
          warn('データヘッダーが見つかりません（行' + headerRow + 'が「' + headerVal + '」）');
        }
      } else {
        warn('サマリーエリア未設置です → setupSheetSummary を実行してください');
      }
    }
  } catch(e) {
    err('スプレッドシートへの接続エラー: ' + e.toString());
  }

  // 7. Whop API 疎通確認
  if (apiKey) {
    try {
      const res = UrlFetchApp.fetch(CONFIG.WHOP_API_BASE + '/me', {
        headers: { 'Authorization': 'Bearer ' + apiKey },
        muteHttpExceptions: true,
      });
      const code = res.getResponseCode();
      if (code === 200) {
        const json = JSON.parse(res.getContentText());
        ok('Whop API に接続できています（会社: ' + (json.title || json.id || '取得済み') + '）');
      } else if (code === 401) {
        err('Whop API 認証エラー（401）: APIキーを確認してください');
      } else {
        warn('Whop API レスポンス: ' + code);
      }
    } catch(e) {
      err('Whop API 接続エラー: ' + e.toString());
    }
  }

  // 結果表示
  Logger.log('========== 設定チェック結果 ==========');
  results.forEach(r => Logger.log(r));
  Logger.log('======================================');
  Logger.log(allOk ? '🎉 すべての設定が正常です！' : '⚠️ 上記の項目を修正してください');
}

// ============================================================
//  Webhook受信
// ============================================================
function doPost(e) {
  try {
    if (e.parameter.token !== CONFIG.SECRET_TOKEN) {
      return jsonRes({ status: 'error', message: 'Unauthorized' });
    }
    const body = JSON.parse(e.postData.contents);
    const eventType = body.type || body.action;
    Logger.log('Webhook: ' + eventType);

    if (eventType === 'payment.succeeded') {
      const paymentId = body.data && body.data.id;
      if (!paymentId) throw new Error('payment_id なし');
      const detail = fetchPaymentDetail(paymentId);
      processPayment(detail);
    }
    return jsonRes({ status: 'ok' });
  } catch (err) {
    Logger.log('エラー: ' + err.toString());
    return jsonRes({ status: 'error', message: err.toString() });
  }
}

// ============================================================
//  Whop APIから支払い詳細を取得
// ============================================================
function fetchPaymentDetail(paymentId) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('WHOP_API_KEY');
  if (!apiKey) throw new Error('WHOP_API_KEY がスクリプトプロパティに未設定です。');

  const url = CONFIG.WHOP_API_BASE + '/payments/' + paymentId;
  const res = UrlFetchApp.fetch(url, {
    method: 'GET',
    headers: { 'Authorization': 'Bearer ' + apiKey },
    muteHttpExceptions: true,
  });

  const code = res.getResponseCode();
  const text = res.getContentText();
  if (code !== 200) {
    // エラー内容をパースして分かりやすく表示
    try {
      const json = JSON.parse(text);
      throw new Error(`Whop API Error (${code}): ${json.error?.message || json.message || text}`);
    } catch (e) {
      throw new Error(`Whop API Error (${code}): ${text}`);
    }
  }
  return JSON.parse(text);
}

// ============================================================
//  支払い処理・シート書き込み
// ============================================================
function processPayment(data) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet(ss, CONFIG.SHEET_LEDGER);
  if (sheet.getLastRow() === 0) initLedgerHeaders(sheet);

  if (isDuplicate(sheet, data.id)) {
    Logger.log('重複スキップ: ' + data.id);
    return;
  }

  const user     = data.user || {};
  const userName = user.name || user.username || '不明';
  const email    = user.email || '';
  const userId   = user.id || '';
  const paidAt   = new Date(data.paid_at || data.created_at);
  const planId   = data.plan && data.plan.id ? data.plan.id : '';
  const planType = getPlanType(planId);
  const currency = (data.currency || 'usd').toUpperCase();

  const grossRaw = parseFloat(data.total || 0);
  const netRaw   = parseFloat(data.amount_after_fees || 0);
  const feeRaw   = Math.max(0, grossRaw - netRaw);
  const usdRef   = parseFloat(data.usd_total || 0);

  let grossFull, feeFull, netFull;
  if (currency === 'JPY') {
    grossFull = Math.round(grossRaw);
    netFull   = Math.round(netRaw);
    feeFull   = Math.round(feeRaw);
  } else {
    const rate = getUsdJpyRate();
    grossFull  = Math.round(grossRaw * rate);
    netFull    = Math.round(netRaw   * rate);
    feeFull    = Math.round(feeRaw   * rate);
  }

  // 50%（自分の取り分）
  const grossJpy = Math.round(grossFull * CONFIG.PARTNER_SHARE);
  const netJpy   = Math.round(netFull   * CONFIG.PARTNER_SHARE);
  const feeJpy   = Math.round(feeFull   * CONFIG.PARTNER_SHARE);

  const row = new Array(COL.USER_ID).fill('');
  row[COL.DATE_RECORD    - 1] = new Date();
  row[COL.DATE_PAYMENT   - 1] = paidAt;
  row[COL.PAYMENT_ID     - 1] = data.id;
  row[COL.USERNAME       - 1] = userName;
  row[COL.EMAIL          - 1] = email;
  row[COL.PLAN_TYPE      - 1] = planType;
  row[COL.CURRENCY       - 1] = currency;
  row[COL.GROSS_JPY      - 1] = grossJpy;
  row[COL.FEE_JPY        - 1] = feeJpy;
  row[COL.NET_JPY        - 1] = netJpy;
  row[COL.USD_REF        - 1] = usdRef;
  row[COL.BILLING_REASON - 1] = formatBillingReason(data.billing_reason || '');
  row[COL.STATUS         - 1] = '確定';
  row[COL.USER_ID        - 1] = userId;

  sheet.appendRow(row);
  SpreadsheetApp.flush();
  formatLedgerRow(sheet, sheet.getLastRow());

  try { updateUserSummary(ss, userName, email, paidAt, grossJpy, feeJpy, netJpy); }
  catch(e) { Logger.log('ユーザー別サマリーエラー: ' + e); }
  try { updateMonthlySummary(ss, paidAt, grossJpy, feeJpy, netJpy); }
  catch(e) { Logger.log('月次サマリーエラー: ' + e); }
  try { updateYearlySummary(ss, paidAt, grossJpy, feeJpy, netJpy); }
  catch(e) { Logger.log('年次サマリーエラー: ' + e); }

  Logger.log('✅ 記録: ' + userName + ' / ' + planType + ' / ¥' + grossJpy.toLocaleString() + '（自分分）');
}

// ============================================================
//  ヘルパー関数
// ============================================================
function getUsdJpyRate() {
  try {
    const res  = UrlFetchApp.fetch('https://open.er-api.com/v6/latest/USD', { muteHttpExceptions: true });
    const json = JSON.parse(res.getContentText());
    if (json.result === 'success') {
      const rate = Math.round(json.rates.JPY * 100) / 100;
      PropertiesService.getScriptProperties().setProperty('LAST_USDJPY_RATE', rate.toString());
      return rate;
    }
  } catch(e) { Logger.log('為替APIエラー: ' + e); }
  return parseFloat(PropertiesService.getScriptProperties().getProperty('LAST_USDJPY_RATE') || '150');
}

function getPlanType(planId) {
  if (CONFIG.PLAN_MONTHLY_IDS.includes(planId)) return '月額 ¥3,000';
  if (CONFIG.PLAN_ANNUAL_IDS.includes(planId))  return '年額 ¥30,000';
  return '不明 (' + planId + ')';
}

function formatBillingReason(r) {
  return { subscription_create: '初回購入', subscription_cycle: '自動更新',
           subscription_update: 'プラン変更', manual: '手動' }[r] || (r || '—');
}

function getHeaderRow(sheet) {
  return sheet.getRange(1, 1).getValue() === '📊 年度サマリー' ? SUMMARY_HEIGHT + 1 : 1;
}

function isDuplicate(sheet, paymentId) {
  if (!paymentId) return false;
  const dataStart = getHeaderRow(sheet) + 1;
  const lastRow   = sheet.getLastRow();
  if (lastRow < dataStart) return false;
  return sheet.getRange(dataStart, COL.PAYMENT_ID, lastRow - dataStart + 1, 1)
    .getValues().some(r => r[0] === paymentId);
}

function jsonRes(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

// ============================================================
//  ヘッダー・書式
// ============================================================
function initLedgerHeaders(sheet) {
  const headerRow = getHeaderRow(sheet);
  const h = ['記録日時','支払日時','Payment ID','ユーザー名','メール','プラン','通貨',
             '売上(円)','手数料(円)','受取(円)','USD参考','請求理由','ステータス','User ID'];
  sheet.getRange(headerRow, 1, 1, h.length).setValues([h]).setFontWeight('bold');
  sheet.setFrozenRows(headerRow);
  [145,145,170,120,180,130,60,100,100,100,90,90,70,160]
    .forEach((w, i) => sheet.setColumnWidth(i + 1, w));
}

function formatLedgerRow(sheet, rowNum) {
  sheet.getRange(rowNum, COL.DATE_RECORD, 1, 2).setNumberFormat('yyyy/MM/dd HH:mm:ss');
  sheet.getRange(rowNum, COL.GROSS_JPY, 1, 3).setNumberFormat('¥#,##0');
  sheet.getRange(rowNum, COL.USD_REF).setNumberFormat('#,##0.00');
}

// ============================================================
//  ユーザー別サマリー
// ============================================================
function updateUserSummary(ss, userName, email, payDate, grossJpy, feeJpy, netJpy) {
  const sheet = getOrCreateSheet(ss, CONFIG.SHEET_BY_USER);
  if (sheet.getLastRow() === 0) {
    const h = ['ユーザー名','メール','初回購入日','最終支払日','支払回数','累計売上(円)','累計手数料(円)','累計受取(円)'];
    sheet.appendRow(h);
    sheet.getRange(1, 1, 1, h.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    [120,180,120,120,70,120,120,120].forEach((w,i) => sheet.setColumnWidth(i+1,w));
  }
  const lastRow = sheet.getLastRow();
  let foundRow  = -1;
  if (lastRow > 1) {
    sheet.getRange(2, 2, lastRow - 1, 1).getValues()
      .forEach((r, i) => { if (r[0] === email) foundRow = i + 2; });
  }
  if (foundRow === -1) {
    sheet.appendRow([userName, email, payDate, payDate, 1, grossJpy, feeJpy, netJpy]);
    const nr = sheet.getLastRow();
    sheet.getRange(nr, 3, 1, 2).setNumberFormat('yyyy/MM/dd');
    sheet.getRange(nr, 6, 1, 3).setNumberFormat('¥#,##0');
  } else {
    const d = sheet.getRange(foundRow, 1, 1, 8).getValues()[0];
    sheet.getRange(foundRow, 4).setValue(payDate).setNumberFormat('yyyy/MM/dd');
    sheet.getRange(foundRow, 5).setValue(d[4] + 1);
    sheet.getRange(foundRow, 6).setValue(d[5] + grossJpy);
    sheet.getRange(foundRow, 7).setValue(d[6] + feeJpy);
    sheet.getRange(foundRow, 8).setValue(d[7] + netJpy);
  }
}

// ============================================================
//  月次サマリー
// ============================================================
function updateMonthlySummary(ss, payDate, grossJpy, feeJpy, netJpy) {
  const sheet = getOrCreateSheet(ss, CONFIG.SHEET_MONTHLY);
  if (sheet.getLastRow() === 0) {
    const h = ['年月','件数','売上合計(円)','手数料合計(円)','受取合計(円)'];
    sheet.appendRow(h);
    sheet.getRange(1, 1, 1, h.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    [100,60,120,120,120].forEach((w,i) => sheet.setColumnWidth(i+1,w));
  }
  const ym      = Utilities.formatDate(payDate, 'Asia/Tokyo', 'yyyy-MM');
  const lastRow = sheet.getLastRow();
  let foundRow  = -1;
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 1).getValues()
      .forEach((r, i) => { if (r[0] === ym) foundRow = i + 2; });
  }
  if (foundRow === -1) {
    sheet.appendRow([ym, 1, grossJpy, feeJpy, netJpy]);
    sheet.getRange(sheet.getLastRow(), 3, 1, 3).setNumberFormat('¥#,##0');
    if (lastRow > 1) sheet.getRange(2, 1, sheet.getLastRow()-1, 5).sort(1);
  } else {
    const d = sheet.getRange(foundRow, 1, 1, 5).getValues()[0];
    sheet.getRange(foundRow, 2).setValue(d[1] + 1);
    sheet.getRange(foundRow, 3).setValue(d[2] + grossJpy);
    sheet.getRange(foundRow, 4).setValue(d[3] + feeJpy);
    sheet.getRange(foundRow, 5).setValue(d[4] + netJpy);
  }
}

// ============================================================
//  年次サマリー
// ============================================================
function updateYearlySummary(ss, payDate, grossJpy, feeJpy, netJpy) {
  const sheet = getOrCreateSheet(ss, CONFIG.SHEET_YEARLY);
  if (sheet.getLastRow() === 0) {
    const h = ['年','件数','売上合計(円)','手数料合計(円)','受取合計(円)'];
    sheet.appendRow(h);
    sheet.getRange(1, 1, 1, h.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    [80,60,120,120,120].forEach((w,i) => sheet.setColumnWidth(i+1,w));
  }
  const yyyy    = Utilities.formatDate(payDate, 'Asia/Tokyo', 'yyyy');
  const lastRow = sheet.getLastRow();
  let foundRow  = -1;
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 1).getValues()
      .forEach((r, i) => { if (String(r[0]) === yyyy) foundRow = i + 2; });
  }
  if (foundRow === -1) {
    sheet.appendRow([yyyy, 1, grossJpy, feeJpy, netJpy]);
    sheet.getRange(sheet.getLastRow(), 3, 1, 3).setNumberFormat('¥#,##0');
    if (lastRow > 1) sheet.getRange(2, 1, sheet.getLastRow()-1, 5).sort(1);
  } else {
    const d = sheet.getRange(foundRow, 1, 1, 5).getValues()[0];
    sheet.getRange(foundRow, 2).setValue(d[1] + 1);
    sheet.getRange(foundRow, 3).setValue(d[2] + grossJpy);
    sheet.getRange(foundRow, 4).setValue(d[3] + feeJpy);
    sheet.getRange(foundRow, 5).setValue(d[4] + netJpy);
  }
}

// ============================================================
//  シート上部にサマリーエリアを設置（一度だけ実行）
// ============================================================
function setupSheetSummary() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_LEDGER);
  if (!sheet) { Logger.log('シートが見つかりません'); return; }

  if (sheet.getRange(1, 1).getValue() === '📊 年度サマリー') {
    Logger.log('サマリーは既に設置済みです');
    return;
  }

  // 既存データを下にシフト
  if (sheet.getLastRow() > 0) {
    sheet.insertRowsBefore(1, SUMMARY_HEIGHT + 1);
  }

  const A = 1, B = 2;
  const dataStartRow = SUMMARY_HEIGHT + 2;

  sheet.getRange(1, A).setValue('📊 年度サマリー').setFontWeight('bold').setFontSize(12);

  sheet.getRange(3, A).setValue('対象年度').setFontWeight('bold');
  sheet.getRange(3, B)
    .setValue(new Date().getFullYear())
    .setBackground('#fff2cc').setFontWeight('bold').setFontSize(13).setHorizontalAlignment('center');
  sheet.getRange(3, 3).setValue('← ここに年を入力（例: 2026）').setFontColor('#999999').setFontSize(9);

  sheet.getRange(4, A, 1, 4).setBorder(false, false, true, false, false, false);

  const items = [
    ['件数（件）',    `=COUNTIFS(B${dataStartRow}:B,">="&DATE(B3,1,1),B${dataStartRow}:B,"<"&DATE(B3+1,1,1))`],
    ['売上合計(円)',  `=SUMIFS(H${dataStartRow}:H,B${dataStartRow}:B,">="&DATE(B3,1,1),B${dataStartRow}:B,"<"&DATE(B3+1,1,1))`],
    ['手数料合計(円)',`=SUMIFS(I${dataStartRow}:I,B${dataStartRow}:B,">="&DATE(B3,1,1),B${dataStartRow}:B,"<"&DATE(B3+1,1,1))`],
    ['受取合計(円)',  `=SUMIFS(J${dataStartRow}:J,B${dataStartRow}:B,">="&DATE(B3,1,1),B${dataStartRow}:B,"<"&DATE(B3+1,1,1))`],
  ];
  items.forEach(([label, formula], i) => {
    const r = i + 5;
    sheet.getRange(r, A).setValue(label).setFontWeight('bold');
    const cell = sheet.getRange(r, B).setFormula(formula).setHorizontalAlignment('right');
    if (i > 0) cell.setNumberFormat('¥#,##0');
  });
  sheet.getRange(8, B).setFontSize(12).setFontWeight('bold');

  sheet.getRange(9, A, 1, 4).setBorder(false, false, true, false, false, false);
  sheet.getRange(10, A).setValue('月別内訳').setFontWeight('bold');
  sheet.getRange(10, B).setValue('受取(円)').setFontWeight('bold').setHorizontalAlignment('right');

  ['1月','2月','3月','4月','5月','6月','7月','8月','9月','10月','11月','12月']
    .forEach((m, i) => {
      const r     = i + 11;
      const mon   = i + 1;
      const mS    = `DATE(B3,${mon},1)`;
      const mE    = mon < 12 ? `DATE(B3,${mon+1},1)` : `DATE(B3+1,1,1)`;
      sheet.getRange(r, A).setValue(m);
      sheet.getRange(r, B)
        .setFormula(`=SUMIFS(J${dataStartRow}:J,B${dataStartRow}:B,">="&${mS},B${dataStartRow}:B,"<"&${mE})`)
        .setNumberFormat('¥#,##0').setHorizontalAlignment('right');
    });

  sheet.getRange(1, A, SUMMARY_HEIGHT, 4).setBorder(true, true, true, true, false, false);
  sheet.setColumnWidth(A, 130);
  sheet.setColumnWidth(B, 130);

  // データヘッダーを設置
  const headerRow = SUMMARY_HEIGHT + 1;
  if (sheet.getRange(headerRow, 1).getValue() === '') initLedgerHeaders(sheet);
  sheet.setFrozenRows(headerRow);

  Logger.log('✅ サマリーをシート上部に設置しました');
}

// ============================================================
//  初期設定（最初に一度だけ実行 → その後 setupSheetSummary）
// ============================================================
function initAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [CONFIG.SHEET_LEDGER, CONFIG.SHEET_BY_USER, CONFIG.SHEET_MONTHLY, CONFIG.SHEET_YEARLY]
    .forEach(name => getOrCreateSheet(ss, name));
  const ls = ss.getSheetByName(CONFIG.SHEET_LEDGER);
  if (ls.getLastRow() === 0) initLedgerHeaders(ls);
  Logger.log('✅ 初期化完了 → 次に setupSheetSummary を実行してください');
}

// ============================================================
//  テスト：実Payment IDで取得＆書き込み
// ============================================================
function testFetchAndWrite() {
  const paymentId = 'pay_beyoF2JebpXM5V';
  const detail = fetchPaymentDetail(paymentId);
  Logger.log(JSON.stringify(detail, null, 2));
  processPayment(detail);
  Logger.log('✅ 書き込み完了');
}

// ============================================================
//  テスト：ダミーデータで書き込み確認
// ============================================================
function testWithDummyData() {
  processPayment({
    id: 'pay_DUMMY_' + Date.now(),
    paid_at: new Date().toISOString(),
    plan: { id: CONFIG.PLAN_MONTHLY_IDS[0] },
    user: { id: 'user_test', name: 'テストユーザー', username: 'testuser', email: 'test@example.com' },
    currency: 'jpy',
    total: 3000,
    amount_after_fees: 2656,
    usd_total: 18.87,
    billing_reason: 'subscription_create',
  });
  Logger.log('✅ ダミーテスト完了');
}

function fixLedgerHeaders() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_LEDGER);
  const headerRow = SUMMARY_HEIGHT + 1; // 25行目

  // 25行目に行を挿入してヘッダーを書き込む
  sheet.insertRowBefore(headerRow);
  const h = ['記録日時','支払日時','Payment ID','ユーザー名','メール','プラン','通貨',
             '売上(円)','手数料(円)','受取(円)','USD参考','請求理由','ステータス','User ID'];
  sheet.getRange(headerRow, 1, 1, h.length).setValues([h]).setFontWeight('bold');
  sheet.setFrozenRows(headerRow);
  Logger.log('✅ ヘッダーを行' + headerRow + 'に挿入しました');
}