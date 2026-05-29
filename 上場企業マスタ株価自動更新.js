/************ 設定 ************/
const MASTER_SHEET   = '上場企業マスタ';
const PRIORITY_SHEET = '資産サマリ';

const SYMBOL_COL = 1; // A列
const PRICE_COL  = 7; // G列

const MAX_SYMBOLS_PER_RUN = 150;
const FETCH_SLEEP_MS = 120;
const TIME_LIMIT_MS = 300000; // 5分（GAS制限6分に対する安全策）

/************ メイン ************/
function updateStockPrices() {
  const START_TIME = new Date().getTime();
  Logger.log('==== 株価更新処理 開始 ====');

  if (!isMarketOpen(new Date())) {
    Logger.log('市場時間外のため処理を終了');
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName(MASTER_SHEET);
  const priority = ss.getSheetByName(PRIORITY_SHEET);

  if (!master || !priority) {
    throw new Error('シートが見つかりません');
  }

  const props = PropertiesService.getScriptProperties();

  /************ 1. 上場企業マスタから銘柄一覧と行番号を作成 ************/
  const masterLastRow = master.getLastRow();

  if (masterLastRow < 2) {
    Logger.log('上場企業マスタに銘柄がありません');
    return;
  }

  const masterValues = master
    .getRange(2, SYMBOL_COL, masterLastRow - 1, 1)
    .getValues();

  const allMasterSymbols = [];
  const rowMap = {};

  masterValues.forEach((row, i) => {
    const symbol = normalizeSymbol_(row[0]);

    if (!symbol) {
      return;
    }

    allMasterSymbols.push(symbol);
    rowMap[symbol] = i + 2; // 実際のシート行番号
  });

  const uniqueMasterSymbols = dedupe_(allMasterSymbols);
  const masterSymbolSet = new Set(uniqueMasterSymbols);

  if (uniqueMasterSymbols.length === 0) {
    Logger.log('有効な銘柄コードがありません');
    return;
  }

  /************ 2. 優先銘柄を取得 ************/
  const prioritySymbolsRaw = readPrioritySymbols_(priority);

  // マスタに存在する銘柄だけを優先更新対象にする
  const prioritySymbols = dedupe_(prioritySymbolsRaw)
    .filter(symbol => masterSymbolSet.has(symbol));

  const prioritySet = new Set(prioritySymbols);

  const normalSymbols = uniqueMasterSymbols
    .filter(symbol => !prioritySet.has(symbol));

  /************ 3. 今回更新する銘柄リストを作成 ************/
  const normalPtr = normalizePointer_(
    Number(props.getProperty('NORMAL_PTR') || 0),
    normalSymbols.length
  );

  const symbolsToUpdate = [];

  // 優先銘柄を先に入れる。ただし上限を超えないようにする。
  for (let i = 0; i < prioritySymbols.length && symbolsToUpdate.length < MAX_SYMBOLS_PER_RUN; i++) {
    symbolsToUpdate.push(prioritySymbols[i]);
  }

  // 残り枠に通常銘柄をローテーションで入れる
  const need = MAX_SYMBOLS_PER_RUN - symbolsToUpdate.length;

  for (let i = 0; i < need && i < normalSymbols.length; i++) {
    const symbol = normalSymbols[(normalPtr + i) % normalSymbols.length];
    symbolsToUpdate.push(symbol);
  }

  if (symbolsToUpdate.length === 0) {
    Logger.log('今回更新する銘柄がありません');
    return;
  }

  Logger.log('優先銘柄数: %d', prioritySymbols.length);
  Logger.log('通常銘柄数: %d', normalSymbols.length);
  Logger.log('今回の更新対象数: %d', symbolsToUpdate.length);
  Logger.log('通常銘柄ポインタ開始位置: %s', normalPtr);

  /************ 4. 株価取得ループ ************/
  const prices = {};

  let processedCount = 0;
  let successCount = 0;

  let normalAttemptedCount = 0;
  let normalSuccessCount = 0;

  for (let i = 0; i < symbolsToUpdate.length; i++) {
    if (new Date().getTime() - START_TIME > TIME_LIMIT_MS) {
      Logger.log('タイムリミットが近づいたため、途中で終了します');
      break;
    }

    const code = symbolsToUpdate[i];
    const isNormalSymbol = !prioritySet.has(code);

    if (isNormalSymbol) {
      normalAttemptedCount++;
    }

    try {
      const price = fetchGoogleFinancePrice(code);

      if (price != null) {
        prices[code] = price;
        successCount++;

        if (isNormalSymbol) {
          normalSuccessCount++;
        }
      }
    } catch (e) {
      Logger.log('取得失敗 %s : %s', code, e.message);
    }

    processedCount++;
    Utilities.sleep(FETCH_SLEEP_MS);
  }

  /************ 5. シートへ書き込み ************/
  const priceCodes = Object.keys(prices);

  if (priceCodes.length > 0) {
    priceCodes.forEach(code => {
      const row = rowMap[code];

      if (row) {
        master.getRange(row, PRICE_COL).setValue(prices[code]);
      }
    });

    Logger.log('書き込み件数: %d', priceCodes.length);
  } else {
    Logger.log('取得成功が0件のため、書き込みは行いません');
  }

  /************ 6. 次回のための通常銘柄ポインタ更新 ************/
  let nextNormalPtr = normalPtr;

  /*
    方針:
    - 通常銘柄が1件も成功していない場合は、ポインタを進めない
    - 通常銘柄に成功があった場合は、試行した通常銘柄数ぶん進める
      → 一部失敗があってもローテーション全体は詰まらない
  */
  if (normalSymbols.length === 0) {
    Logger.log('通常銘柄がないため、通常銘柄ポインタは更新しません');
  } else if (normalAttemptedCount === 0) {
    Logger.log('今回は通常銘柄を処理していないため、通常銘柄ポインタは更新しません');
  } else if (normalSuccessCount === 0) {
    Logger.log('通常銘柄が全件失敗のため、通常銘柄ポインタは更新しません');
  } else {
    nextNormalPtr = (normalPtr + normalAttemptedCount) % normalSymbols.length;
    props.setProperty('NORMAL_PTR', nextNormalPtr.toString());

    Logger.log('通常銘柄ポインタを更新しました: %s → %s', normalPtr, nextNormalPtr);
  }

  Logger.log('今回処理した件数: %d', processedCount);
  Logger.log('取得成功件数: %d', successCount);
  Logger.log('通常銘柄の試行件数: %d', normalAttemptedCount);
  Logger.log('通常銘柄の取得成功件数: %d', normalSuccessCount);
  Logger.log('次回の開始位置（通常銘柄ポインタ）: %s', nextNormalPtr);
  Logger.log('==== 処理終了 ====');
}

/************ Google Finance 株価取得 ************/
function fetchGoogleFinancePrice(symbol) {
  symbol = normalizeSymbol_(symbol);

  if (!symbol) {
    return null;
  }

  const url = `https://www.google.com/finance/beta/quote/${encodeURIComponent(symbol)}:TYO?hl=ja&gl=JP`;

  const res = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    followRedirects: true,
    headers: {
      'User-Agent': 'Mozilla/5.0',
      'Accept-Language': 'ja-JP,ja;q=0.9,en-US;q=0.8,en;q=0.7'
    }
  });

  const status = res.getResponseCode();
  const html = res.getContentText('UTF-8');

  if (status !== 200) {
    Logger.log('Google Finance HTTP失敗 %s : status=%s', symbol, status);
    return null;
  }

  /*
    旧Google Finance DOM用。
    今回は has YMlKec=false なので通常はここには入らないが、念のため残す。
  */
  const oldDomMatch = html.match(/<div[^>]*class=["'][^"']*\bYMlKec\b[^"']*["'][^>]*>\s*([^<]+?)\s*<\/div>/i);

  if (oldDomMatch) {
    return parsePriceText_(oldDomMatch[1]);
  }

  /*
    新しい finance/beta HTML向け。
    まずHTMLをテキスト化して、
    「9432:TYO」の近くにある「￥152.50」のような価格を拾う。
  */
  const text = htmlToPlainText_(html);
  const priceFromText = extractPriceNearSymbol_(text, symbol);

  if (priceFromText != null) {
    return priceFromText;
  }

  /*
    script内やエスケープ文字列内にだけ残っている場合の保険。
  */
  const decodedRaw = decodeGoogleFinanceHtml_(html);
  const priceFromRaw = extractPriceNearSymbol_(decodedRaw, symbol);

  if (priceFromRaw != null) {
    return priceFromRaw;
  }

  /*
    自動アクセス制限・同意画面などの可能性を軽く判定。
  */
  if (/captcha|unusual traffic|sorry|consent/i.test(html)) {
    Logger.log('Google側の制限または同意画面の可能性 %s', symbol);
    return null;
  }

  Logger.log(
    '株価取得失敗（価格未検出） %s status=%s length=%s markerIndexText=%s markerIndexRaw=%s',
    symbol,
    status,
    html.length,
    text.indexOf(`${symbol}:TYO`),
    decodedRaw.indexOf(`${symbol}:TYO`)
  );

  return null;
}

/************ 価格抽出補助 ************/
function extractPriceNearSymbol_(text, symbol) {
  const markerRegex = new RegExp(escapeRegExp_(symbol) + '\\s*:\\s*TYO', 'i');
  const marker = markerRegex.exec(text);

  if (!marker) {
    return null;
  }

  /*
    基本は銘柄コードの後ろ側を見る。
    例:
    9432:TYO ... NTT ... ￥152.50
  */
  const afterChunk = text.slice(marker.index, marker.index + 3000);
  let priceMatch = afterChunk.match(/[¥￥]\s*([0-9][0-9,]*(?:\.\d+)?)/);

  if (priceMatch) {
    return parsePriceText_(priceMatch[1]);
  }

  /*
    念のため前後も見る。
  */
  const start = Math.max(0, marker.index - 500);
  const aroundChunk = text.slice(start, marker.index + 3000);
  priceMatch = aroundChunk.match(/[¥￥]\s*([0-9][0-9,]*(?:\.\d+)?)/);

  if (priceMatch) {
    return parsePriceText_(priceMatch[1]);
  }

  return null;
}

function parsePriceText_(value) {
  const cleaned = String(value)
    .replace(/[^\d.]/g, '')
    .trim();

  if (!cleaned) {
    return null;
  }

  const price = Number(cleaned);

  return isFinite(price) ? price : null;
}

function htmlToPlainText_(html) {
  return decodeGoogleFinanceHtml_(html)
    .replace(/<script\b[\s\S]*?<\/script>/gi, ' ')
    .replace(/<style\b[\s\S]*?<\/style>/gi, ' ')
    .replace(/<[^>]+>/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function decodeGoogleFinanceHtml_(html) {
  return String(html)
    .replace(/\\u003c/gi, '<')
    .replace(/\\u003e/gi, '>')
    .replace(/\\u0026/gi, '&')
    .replace(/\\u003d/gi, '=')
    .replace(/\\u003a/gi, ':')
    .replace(/\\u002f/gi, '/')
    .replace(/\\u00a5/gi, '¥')
    .replace(/\\uffe5/gi, '￥')
    .replace(/&amp;/gi, '&')
    .replace(/&nbsp;|&#160;/gi, ' ')
    .replace(/&#165;|&#xA5;|&yen;/gi, '￥')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;|&apos;/gi, "'");
}

function escapeRegExp_(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/************ 優先銘柄取得 ************/
function readPrioritySymbols_(sheet) {
  const lastRow = sheet.getLastRow();

  if (lastRow < 6) {
    return [];
  }

  return sheet
    .getRange(6, 1, lastRow - 5, 1)
    .getValues()
    .flat()
    .map(value => normalizeSymbol_(value))
    .filter(value => value);
}

/************ 銘柄コード整形 ************/
function normalizeSymbol_(value) {
  if (value == null) {
    return '';
  }

  let s = String(value).trim();

  if (!s) {
    return '';
  }

  // 全角英数字を半角に変換
  s = s.replace(/[Ａ-Ｚａ-ｚ０-９]/g, ch => {
    return String.fromCharCode(ch.charCodeAt(0) - 0xFEE0);
  });

  s = s.toUpperCase();
  s = s.replace(/\s+/g, '');

  // 9432.T、9432:TYO、TYO:9432 のような入力にも軽く対応
  s = s.replace(/\.T$/i, '');
  s = s.replace(/:TYO$/i, '');
  s = s.replace(/^TYO:/i, '');

  return s;
}

function dedupe_(values) {
  const seen = new Set();
  const result = [];

  values.forEach(value => {
    const v = normalizeSymbol_(value);

    if (!v || seen.has(v)) {
      return;
    }

    seen.add(v);
    result.push(v);
  });

  return result;
}

function normalizePointer_(ptr, length) {
  if (!length || length <= 0) {
    return 0;
  }

  if (!isFinite(ptr) || ptr < 0) {
    return 0;
  }

  return Math.floor(ptr) % length;
}

/************ 市場時間判定 ************/
function isMarketOpen(date) {
  const day = date.getDay(); // 0=日, 6=土

  if (day === 0 || day === 6) {
    return false;
  }

  const h = date.getHours();
  const m = date.getMinutes();
  const t = h * 60 + m;

  // 9:00〜18:00
  return t >= 540 && t <= 1080;
}

/************ 単体テスト用 ************/
function testFetchGoogleFinancePrice() {
  const codes = ['9432', '3382', '7203'];

  codes.forEach(code => {
    const price = fetchGoogleFinancePrice(code);
    Logger.log('%s => %s', code, price);
  });
}

/************ デバッグ用 ************/
function debugGoogleFinance() {
  const symbol = '9432';
  const url = `https://www.google.com/finance/beta/quote/${symbol}:TYO?hl=ja&gl=JP`;

  const res = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    followRedirects: true,
    headers: {
      'User-Agent': 'Mozilla/5.0',
      'Accept-Language': 'ja-JP,ja;q=0.9,en-US;q=0.8,en;q=0.7'
    }
  });

  const html = res.getContentText('UTF-8');
  const text = htmlToPlainText_(html);
  const decodedRaw = decodeGoogleFinanceHtml_(html);

  Logger.log('status=%s', res.getResponseCode());
  Logger.log('length=%s', html.length);
  Logger.log('has YMlKec=%s', html.indexOf('YMlKec') >= 0);
  Logger.log('has fxKbKc=%s', html.indexOf('fxKbKc') >= 0);
  Logger.log('text marker index=%s', text.indexOf(`${symbol}:TYO`));
  Logger.log('raw marker index=%s', decodedRaw.indexOf(`${symbol}:TYO`));

  const idx = text.indexOf(`${symbol}:TYO`);

  if (idx >= 0) {
    Logger.log(text.slice(idx, idx + 1000));
  } else {
    Logger.log(html.slice(0, 1000));
  }
}