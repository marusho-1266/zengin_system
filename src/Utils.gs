/**
 * 全銀協システム共通ユーティリティ関数
 * 
 * このファイルは複数のファイルで重複していた関数を統一し、
 * コードの保守性と一貫性を向上させるために作成されました。
 * 
 * @author AIコードアシスタント
 * @version 1.0
 * @since 2025/06/21
 */

// ===== 数値・文字列正規化関数 =====

/**
 * 数値コードを適切な桁数に0埋めして正規化
 * @param {any} value - 元の値（数値または文字列）
 * @param {number} targetLength - 目標桁数
 * @return {string} 0埋めされた文字列
 */
function normalizeCode(value, targetLength) {
  if (value === null || value === undefined || value === '') {
    return '';
  }
  
  // 数値の場合は文字列に変換して0埋め
  if (typeof value === 'number') {
    return String(value).padStart(targetLength, '0');
  }
  
  // 文字列の場合はトリムして0埋め
  const strValue = String(value).trim();
  if (strValue === '') {
    return '';
  }
  
  // 数値文字列の場合は0埋め
  if (/^\d+$/.test(strValue)) {
    return strValue.padStart(targetLength, '0');
  }
  
  // その他の場合はそのまま返す
  return strValue;
}

/**
 * 数値コードの正規化（０埋め処理）
 * @param {any} value - 元の値
 * @param {number} length - 目標桁数
 * @return {string} ０埋めされた文字列
 * @deprecated normalizeCode関数を使用してください
 */
function normalizeNumericCode(value, length) {
  return normalizeCode(value, length);
}

/**
 * 銀行コードの正規化
 * @param {any} value - 銀行コード
 * @return {string} 4桁に正規化された銀行コード
 */
function normalizeBankCode(value) {
  return normalizeCode(value, VALIDATION_RULES.BANK_CODE_LENGTH);
}

/**
 * 支店コードの正規化
 * @param {any} value - 支店コード
 * @return {string} 3桁に正規化された支店コード
 */
function normalizeBranchCode(value) {
  return normalizeCode(value, VALIDATION_RULES.BRANCH_CODE_LENGTH);
}

/**
 * 口座番号の正規化
 * @param {any} value - 口座番号
 * @return {string} 7桁に正規化された口座番号
 */
function normalizeAccountNumber(value) {
  return normalizeCode(value, VALIDATION_RULES.ACCOUNT_NUMBER_MAX_LENGTH);
}

// ===== データ検証関数 =====

/**
 * 行が空かどうかチェック
 * @param {Array} row - 行データ
 * @return {boolean} 空行かどうか
 */
function isEmptyRow(row) {
  if (!row || !Array.isArray(row)) return true;
  return row.every(cell => !cell || String(cell).trim() === '');
}

/**
 * 有効なデータを含む行かどうかチェック
 * @param {Array} row - 行データ
 * @return {boolean} 有効なデータを含むかどうか
 */
function hasValidData(row) {
  return !isEmptyRow(row);
}

/**
 * セルが空かどうかチェック
 * @param {any} cell - セルデータ
 * @return {boolean} 空セルかどうか
 */
function isEmptyCell(cell) {
  return !cell || String(cell).trim() === '';
}

// ===== 文字列処理関数 =====

/**
 * プルダウン選択値の抽出（「1:普通」→「1」）
 * @param {any} value - 元の値
 * @return {string} 抽出された値
 */
function extractSelectValue(value) {
  const stringValue = String(value || '').trim();
  if (!stringValue) return '';
  
  // 「1:普通」形式の場合、「:」の前の値を抽出
  if (stringValue.includes(':')) {
    return stringValue.split(':')[0];
  }
  
  return stringValue;
}

/**
 * 左側を指定文字でパディング
 * @param {string} str - 対象文字列
 * @param {number} length - 目標長
 * @param {string} padChar - パディング文字
 * @return {string} パディング後の文字列
 */
function padLeft(str, length, padChar = ' ') {
  str = String(str);
  while (str.length < length) {
    str = padChar + str;
  }
  return str.substring(0, length); // 長すぎる場合は切り詰め
}

/**
 * 右側を指定文字でパディング
 * @param {string} str - 対象文字列
 * @param {number} length - 目標長
 * @param {string} padChar - パディング文字
 * @return {string} パディング後の文字列
 */
function padRight(str, length, padChar = ' ') {
  str = String(str);
  while (str.length < length) {
    str = str + padChar;
  }
  return str.substring(0, length); // 長すぎる場合は切り詰め
}

/**
 * 文字列を指定長に調整（長すぎる場合は切り詰め、短い場合はパディング）
 * @param {string} str - 対象文字列
 * @param {number} length - 目標長
 * @param {string} padChar - パディング文字
 * @param {string} direction - パディング方向 ('left' または 'right')
 * @return {string} 調整後の文字列
 */
function fixLength(str, length, padChar = ' ', direction = 'right') {
  str = String(str);
  if (str.length > length) {
    return str.substring(0, length);
  }
  
  return direction === 'left' ? 
    padLeft(str, length, padChar) : 
    padRight(str, length, padChar);
}

// ===== 配列処理関数 =====

/**
 * 配列をチャンクに分割
 * @param {Array} array - 分割対象の配列
 * @param {number} size - チャンクサイズ
 * @return {Array[]} チャンクに分割された配列
 */
function chunkArray(array, size) {
  const chunks = [];
  for (let i = 0; i < array.length; i += size) {
    chunks.push(array.slice(i, i + size));
  }
  return chunks;
}

/**
 * 配列から重複を除去
 * @param {Array} array - 対象配列
 * @param {Function} keyFunc - キー生成関数（オプション）
 * @return {Array} 重複が除去された配列
 */
function removeDuplicates(array, keyFunc = null) {
  if (!keyFunc) {
    return [...new Set(array)];
  }
  
  const seen = new Set();
  return array.filter(item => {
    const key = keyFunc(item);
    if (seen.has(key)) {
      return false;
    }
    seen.add(key);
    return true;
  });
}

// ===== 数値処理関数 =====

/**
 * 数値を安全に整数に変換
 * @param {any} value - 変換対象の値
 * @param {number} defaultValue - デフォルト値
 * @return {number} 整数値
 */
function safeParseInt(value, defaultValue = 0) {
  if (value === null || value === undefined || value === '') {
    return defaultValue;
  }
  
  const parsed = parseInt(value, 10);
  return isNaN(parsed) ? defaultValue : parsed;
}

/**
 * 数値を安全に浮動小数点数に変換
 * @param {any} value - 変換対象の値
 * @param {number} defaultValue - デフォルト値
 * @return {number} 浮動小数点数値
 */
function safeParseFloat(value, defaultValue = 0) {
  if (value === null || value === undefined || value === '') {
    return defaultValue;
  }
  
  const parsed = parseFloat(value);
  return isNaN(parsed) ? defaultValue : parsed;
}

/**
 * 金額を整数に変換（全銀協フォーマット対応）
 * @param {any} value - 金額
 * @return {number} 整数の金額
 */
function formatAmount(value) {
  return Math.floor(safeParseFloat(value, 0));
}

// ===== 日付処理関数 =====

/**
 * 日付を全銀協フォーマット用文字列に変換
 * @param {Date} date - 対象日付
 * @return {string} YYMMDD形式の文字列
 */
function formatDateToZengin(date = new Date()) {
  const year = date.getFullYear().toString().slice(-2);
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return year + month + day;
}

/**
 * 時刻を文字列に変換
 * @param {Date} date - 対象日付
 * @return {string} HHMM形式の文字列
 */
function formatTimeToString(date = new Date()) {
  const hours = String(date.getHours()).padStart(2, '0');
  const minutes = String(date.getMinutes()).padStart(2, '0');
  return hours + minutes;
}

// ===== エラーハンドリング関数 =====

/**
 * 安全に関数を実行（エラーハンドリング付き）
 * @param {Function} func - 実行する関数
 * @param {any} defaultValue - エラー時のデフォルト値
 * @param {Function} errorHandler - エラーハンドラ関数（オプション）
 * @return {any} 実行結果またはデフォルト値
 */
function safeExecute(func, defaultValue = null, errorHandler = null) {
  try {
    return func();
  } catch (error) {
    if (errorHandler) {
      errorHandler(error);
    } else {
      logError('safeExecute エラー: ' + error.toString());
    }
    return defaultValue;
  }
}

/**
 * 非同期処理の安全な実行
 * @param {Promise} promise - 実行するPromise
 * @param {any} defaultValue - エラー時のデフォルト値
 * @return {any} 実行結果またはデフォルト値
 */
function safeAsync(promise, defaultValue = null) {
  return promise.catch(error => {
    logError('safeAsync エラー: ' + error.toString());
    return defaultValue;
  });
}

// ===== バリデーション関数 =====

/**
 * 全銀協フォーマット対応文字列の検証
 * @param {string} str - 検証対象文字列
 * @return {boolean} 有効かどうか
 */
function isValidZenginFormat(str) {
  if (!str) return false;
  // 全銀協フォーマット対応文字: 半角カナ、英数字、記号（カンマ除く）、スペース
  const zenginFormatRegex = /^[A-Z0-9ｱ-ﾝﾞﾟｧ-ｯ ・ー().\-/]+$/;
  return zenginFormatRegex.test(str);
}

/**
 * 数値文字列の検証
 * @param {string} str - 検証対象文字列
 * @return {boolean} 有効かどうか
 */
function isValidNumber(str) {
  return /^\d+$/.test(str);
}

/**
 * 英数字文字列の検証
 * @param {string} str - 検証対象文字列
 * @return {boolean} 有効かどうか
 */
function isValidAlphanumeric(str) {
  return /^[A-Za-z0-9]+$/.test(str);
}

/**
 * 半角カナ文字列の検証
 * @param {string} str - 検証対象文字列
 * @return {boolean} 有効かどうか
 */
function isValidKana(str) {
  return /^[ｱ-ﾝﾞﾟｧ-ｯ ・ー]+$/.test(str);
}

// ===== デバッグ・テスト用関数 =====

/**
 * オブジェクトの詳細情報を文字列として出力
 * @param {any} obj - 対象オブジェクト
 * @param {number} depth - 階層深度制限
 * @return {string} オブジェクトの文字列表現
 */
function debugString(obj, depth = 2) {
  try {
    return JSON.stringify(obj, null, 2);
  } catch (error) {
    return `[デバッグ出力エラー: ${error.message}]`;
  }
}

/**
 * 実行時間を測定
 * @param {Function} func - 実行する関数
 * @param {string} label - ラベル（オプション）
 * @return {any} 関数の実行結果
 */
function measureTime(func, label = 'Processing') {
  const startTime = Date.now();
  const result = func();
  const endTime = Date.now();
  
  logInfo(`${label} 実行時間: ${endTime - startTime}ms`);
  return result;
} 