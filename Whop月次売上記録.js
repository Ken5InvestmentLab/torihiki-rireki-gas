const WHOP_MONTHLY_CONFIG = {
  sourceSheetName: "サブスク管理･天底極致",
  destSheetName: "インジ販売管理",
  cashflowSheetName: "入出金履歴",

  // この元データは取引明細のヘッダーが23行目
  sourceHeaderRow: 23,

  // 元データの列
  sourceColumns: {
    date: "B",   // 支払日時
    sales: "H",  // 売上(円)
    fee: "I"     // 手数料(円)
  },

  // インジ販売管理に入れる日付
  recordDateMode: "PREVIOUS_MONTH_END",

  // Whop → PayPal 出金手数料
  whopToPaypalFee: 31,

  // 入出金履歴に入れる日付 = 実行日の何日後か
  cashflowDepositOffsetDays: 4,

  // 入出金履歴の3列目に入れる値
  cashflowTypeCode: 10
};


/**
 * 毎月1日に実行する本体
 */
function recordMonthlyWhopSubscriptionSales() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(WHOP_MONTHLY_CONFIG.sourceSheetName);
  const destSheet = ss.getSheetByName(WHOP_MONTHLY_CONFIG.destSheetName);
  const cashflowSheet = ss.getSheetByName(WHOP_MONTHLY_CONFIG.cashflowSheetName);

  if (!sourceSheet) {
    throw new Error(`元シートが見つかりません: ${WHOP_MONTHLY_CONFIG.sourceSheetName}`);
  }
  if (!destSheet) {
    throw new Error(`記録先シートが見つかりません: ${WHOP_MONTHLY_CONFIG.destSheetName}`);
  }
  if (!cashflowSheet) {
    throw new Error(`入出金履歴シートが見つかりません: ${WHOP_MONTHLY_CONFIG.cashflowSheetName}`);
  }

  const today = new Date();
  const target = getPreviousMonthRange_(today);

  const sourceValues = sourceSheet.getDataRange().getValues();
  if (sourceValues.length <= WHOP_MONTHLY_CONFIG.sourceHeaderRow) {
    throw new Error("元シートに集計対象データがありません。");
  }

  const headerRowIndex = WHOP_MONTHLY_CONFIG.sourceHeaderRow - 1;
  const headers = sourceValues[headerRowIndex].map(v => normalizeHeader_(v));

  const dateCol = resolveSourceColumn_(
    WHOP_MONTHLY_CONFIG.sourceColumns.date,
    headers,
    ["日付", "年月日", "決済日", "売上日", "発生日", "購入日", "支払日", "タイムスタンプ", "created_at", "created", "date"]
  );

  const salesCol = resolveSourceColumn_(
    WHOP_MONTHLY_CONFIG.sourceColumns.sales,
    headers,
    ["売上", "売上金額", "売上合計", "販売額", "販売金額", "売上高", "gross", "revenue", "amount", "subtotal"]
  );

  const feeCol = resolveSourceColumn_(
    WHOP_MONTHLY_CONFIG.sourceColumns.fee,
    headers,
    ["手数料", "販売手数料", "whop手数料", "決済手数料", "fee", "fees", "platformfee", "transactionfee"]
  );

  if (dateCol === -1) throw new Error("日付列を特定できません。sourceColumns.date に列記号を指定してください。");
  if (salesCol === -1) throw new Error("売上列を特定できません。sourceColumns.sales に列記号を指定してください。");
  if (feeCol === -1) throw new Error("手数料列を特定できません。sourceColumns.fee に列記号を指定してください。");

  let salesTotal = 0;
  let feeTotal = 0;
  let count = 0;

  for (let r = headerRowIndex + 1; r < sourceValues.length; r++) {
    const row = sourceValues[r];
    const rowDate = toDate_(row[dateCol]);
    if (!rowDate) continue;

    if (rowDate >= target.start && rowDate < target.end) {
      salesTotal += toNumber_(row[salesCol]);
      feeTotal += toNumber_(row[feeCol]);
      count++;
    }
  }

  const recordDate =
    WHOP_MONTHLY_CONFIG.recordDateMode === "RUN_DATE"
      ? today
      : target.previousMonthEnd;

  const monthLabel = `${target.month}月分`;
  const note = `${monthLabel} 手数料=販売手数料 支出=Whop→Paypal出金手数料`;

  const whopToPaypalFee = WHOP_MONTHLY_CONFIG.whopToPaypalFee;
  const depositAmount = salesTotal - feeTotal - whopToPaypalFee;

  // 既存行は絶対に探さない。
  // インジ販売管理の最新行の次に必ず追加する。
  const writeRow = findNextWriteRow_(destSheet);

  ensureRowExists_(destSheet, writeRow);
  preserveFormulaInColumnJ_(destSheet, writeRow);

  // A:I を記録。J列は関数を残すため触らない。
  destSheet.getRange(writeRow, 1, 1, 9).setValues([[
    "販売",          // A
    recordDate,       // B
    "Whop",           // C
    "Whop",           // D
    "",               // E
    "サブスク販売",    // F
    salesTotal,        // G
    feeTotal,          // H
    whopToPaypalFee    // I
  ]]);

  // K:M を記録。J列は触らない。
  destSheet.getRange(writeRow, 11, 1, 3).setValues([[
    "",    // K
    "",    // L
    note   // M
  ]]);

  // 入出金履歴に追加
  appendWhopCashflowHistory_(cashflowSheet, today, depositAmount);

  Logger.log(JSON.stringify({
    targetMonth: monthLabel,
    count,
    salesTotal,
    feeTotal,
    whopToPaypalFee,
    depositAmount,
    salesManageWriteRow: writeRow
  }));
}


/**
 * 毎月1日の自動実行トリガーを作成
 * 初回だけ手動実行してください。
 */
function installMonthlyWhopSubscriptionTrigger() {
  const handlerName = "recordMonthlyWhopSubscriptionSales";

  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === handlerName)
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger(handlerName)
    .timeBased()
    .onMonthDay(1)
    .atHour(9)
    .create();
}


/**
 * 手動テスト用
 */
function testRecordMonthlyWhopSubscriptionSales() {
  recordMonthlyWhopSubscriptionSales();
}


/***** 以下、補助関数 *****/

function getPreviousMonthRange_(baseDate) {
  const y = baseDate.getFullYear();
  const m = baseDate.getMonth();

  const start = new Date(y, m - 1, 1, 0, 0, 0, 0);
  const end = new Date(y, m, 1, 0, 0, 0, 0);
  const previousMonthEnd = new Date(y, m, 0, 0, 0, 0, 0);

  return {
    start,
    end,
    previousMonthEnd,
    year: start.getFullYear(),
    month: start.getMonth() + 1
  };
}

function resolveSourceColumn_(manualColumnLetter, normalizedHeaders, candidates) {
  if (manualColumnLetter) {
    return columnLetterToIndex_(manualColumnLetter);
  }

  const normalizedCandidates = candidates.map(v => normalizeHeader_(v));

  for (let i = 0; i < normalizedHeaders.length; i++) {
    const h = normalizedHeaders[i];
    if (!h) continue;

    if (normalizedCandidates.includes(h)) {
      return i;
    }
  }

  for (let i = 0; i < normalizedHeaders.length; i++) {
    const h = normalizedHeaders[i];
    if (!h) continue;

    if (normalizedCandidates.some(c => h.includes(c))) {
      return i;
    }
  }

  return -1;
}

function normalizeHeader_(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/[ 　_\-ー・･（）()\[\]［］]/g, "");
}

function columnLetterToIndex_(letter) {
  const s = String(letter || "").trim().toUpperCase();
  let n = 0;

  for (let i = 0; i < s.length; i++) {
    const code = s.charCodeAt(i);
    if (code < 65 || code > 90) continue;
    n = n * 26 + (code - 64);
  }

  return n > 0 ? n - 1 : -1;
}

function toDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }

  const s = String(value || "").trim();
  if (!s) return null;

  const normalized = toHalfWidth_(s)
    .replace(/[年月]/g, "/")
    .replace(/[日]/g, "")
    .replace(/-/g, "/");

  const match = normalized.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})/);
  if (!match) return null;

  const y = Number(match[1]);
  const m = Number(match[2]);
  const d = Number(match[3]);

  const date = new Date(y, m - 1, d);
  return isNaN(date.getTime()) ? null : date;
}

function toNumber_(value) {
  if (typeof value === "number") return value;
  if (value === null || value === undefined || value === "") return 0;

  let s = toHalfWidth_(String(value)).trim();
  if (!s) return 0;

  let negative = false;

  if (/^\(.*\)$/.test(s) || s.includes("▲") || s.includes("△")) {
    negative = true;
  }

  s = s
    .replace(/[￥¥円,\s]/g, "")
    .replace(/[()▲△]/g, "")
    .replace(/[^\d.\-]/g, "");

  if (!s) return 0;

  const n = Number(s);
  if (isNaN(n)) return 0;

  return negative ? -Math.abs(n) : n;
}

function toHalfWidth_(str) {
  return String(str).replace(/[！-～]/g, ch =>
    String.fromCharCode(ch.charCodeAt(0) - 0xFEE0)
  );
}

function findExistingMonthlyRow_(sheet, monthLabel) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return -1;

  const notes = sheet.getRange(1, 13, lastRow, 1).getDisplayValues();

  for (let i = notes.length - 1; i >= 0; i--) {
    if (String(notes[i][0] || "").includes(monthLabel)) {
      return i + 1;
    }
  }

  return -1;
}

function findNextWriteRow_(sheet) {
  const lastRow = Math.max(sheet.getLastRow(), 1);
  const maxCols = Math.max(sheet.getLastColumn(), 13);

  const values = sheet.getRange(1, 1, lastRow, maxCols).getDisplayValues();

  // A, B, G, H, M のどれかに値がある最後の行を探す
  const checkCols = [1, 2, 7, 8, 13].map(n => n - 1);

  for (let r = values.length - 1; r >= 0; r--) {
    const hasValue = checkCols.some(c => String(values[r][c] || "").trim() !== "");
    if (hasValue) {
      return r + 2;
    }
  }

  return 1;
}

function ensureRowExists_(sheet, row) {
  if (row <= sheet.getMaxRows()) return;

  sheet.insertRowsAfter(sheet.getMaxRows(), row - sheet.getMaxRows());
}

function preserveFormulaInColumnJ_(sheet, row) {
  const target = sheet.getRange(row, 10);

  if (target.getFormula()) return;

  if (row <= 1) return;

  const source = sheet.getRange(row - 1, 10);
  const formulaR1C1 = source.getFormulaR1C1();

  if (formulaR1C1) {
    target.setFormulaR1C1(formulaR1C1);
  }
}

function appendWhopCashflowHistory_(sheet, baseDate, depositAmount) {
  const depositDate = addDays_(baseDate, WHOP_MONTHLY_CONFIG.cashflowDepositOffsetDays);
  const writeRow = findNextCashflowWriteRow_(sheet);

  ensureRowExists_(sheet, writeRow);

  // 入出金履歴
  // A: 日付
  // B: 種別
  // C: 金額
  // D: コード
  sheet.getRange(writeRow, 1, 1, 4).setValues([[
    depositDate,
    "入金",
    depositAmount,
    WHOP_MONTHLY_CONFIG.cashflowTypeCode
  ]]);

  Logger.log(JSON.stringify({
    cashflowWriteRow: writeRow,
    depositDate,
    type: "入金",
    depositAmount,
    code: WHOP_MONTHLY_CONFIG.cashflowTypeCode
  }));
}

function findNextCashflowWriteRow_(sheet) {
  const lastRow = Math.max(sheet.getLastRow(), 1);
  const values = sheet.getRange(1, 1, lastRow, 4).getDisplayValues();

  // A〜D列のどこかに値がある最後の行の次に追加
  for (let r = values.length - 1; r >= 0; r--) {
    const hasValue = values[r].some(v => String(v || "").trim() !== "");
    if (hasValue) {
      return r + 2;
    }
  }

  return 1;
}

function addDays_(date, days) {
  const d = new Date(date);
  d.setDate(d.getDate() + Number(days || 0));
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}