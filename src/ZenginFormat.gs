/**
 * 全銀協フォーマット生成機能
 * 120バイト固定長、Shift_JIS形式での全銀協フォーマットファイル生成
 */

/**
 * 全銀協フォーマットファイルの生成
 * @return {Object} 生成結果 { success: boolean, fileName: string, recordCount: number, error?: string }
 */
function createZenginFormatFile() {
  try {
    // 振込依頼人情報の取得
    const clientInfo = getClientInfo();
    if (!clientInfo.isValid) {
      return { 
        success: false, 
        error: '振込依頼人情報が正しく設定されていません:\n' + clientInfo.errors.join('\n') 
      };
    }
    
    // 振込データの取得
    const transferData = getTransferDataForFormat();
    if (transferData.length === 0) {
      return { 
        success: false, 
        error: '振込データが存在しません。' 
      };
    }
    
    // 全銀協フォーマットレコードの生成
    const records = [];
    
    // 1. ヘッダレコード
    records.push(createHeaderRecord(clientInfo.data, transferData.length));
    
    // 2. データレコード
    for (let i = 0; i < transferData.length; i++) {
      records.push(createDataRecord(transferData[i], i + 1, clientInfo.data.nameOutputMode));
    }
    
    // 3. トレーラレコード
    records.push(createTrailerRecord(transferData));
    
    // 4. エンドレコード
    records.push(createEndRecord(transferData.length + 3)); // +3はヘッダ、トレーラ、エンド自身
    
    // ファイル内容の結合
    const fileContent = records.join('\r\n') + '\r\n';
    
    // Shift_JIS互換性チェック
    const encodingCheck = validateShiftJISCompatibility(fileContent);
    if (!encodingCheck.isValidSJIS) {
      Logger.log('警告: ' + encodingCheck.message);
    }
    
    // ファイル名の生成
    const fileName = generateFileName(clientInfo.data);
    
    // ファイルのダウンロード準備（Shift_JISエンコーディング）
    const blob = createShiftJISBlob(fileContent, fileName);
    
    // DriveAPIを使用してファイルを作成（一時的）
    const file = DriveApp.createFile(blob);
    
    // ダウンロード用のURLを作成（注意：実際のダウンロードはGAS制限により制約あり）
    Logger.log(`全銀協フォーマットファイル生成完了: ${fileName}`);
    Logger.log(`ファイルID: ${file.getId()}`);
    Logger.log(`レコード数: ${records.length}件`);
    
    return {
      success: true,
      fileName: fileName,
      recordCount: transferData.length,
      fileId: file.getId(),
      totalRecords: records.length
    };
    
  } catch (error) {
    Logger.log('全銀協フォーマット生成エラー: ' + error.toString());
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 振込依頼人情報の取得
 * @return {Object} { isValid: boolean, data?: Object, errors?: string[] }
 */
function getClientInfo() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.CLIENT_INFO);
    if (!sheet) {
      return { isValid: false, errors: ['振込依頼人情報シートが見つかりません。'] };
    }
    
    const data = {
      clientCode: String(sheet.getRange(CLIENT_INFO_CELLS.CLIENT_CODE).getValue() || '').trim(),
      clientName: String(sheet.getRange(CLIENT_INFO_CELLS.CLIENT_NAME).getValue() || '').trim(),
      bankCode: normalizeNumericCode(sheet.getRange(CLIENT_INFO_CELLS.BANK_CODE).getValue(), 4),
      bankName: String(sheet.getRange(CLIENT_INFO_CELLS.BANK_NAME).getValue() || '').trim(),
      branchCode: normalizeNumericCode(sheet.getRange(CLIENT_INFO_CELLS.BRANCH_CODE).getValue(), 3),
      branchName: String(sheet.getRange(CLIENT_INFO_CELLS.BRANCH_NAME).getValue() || '').trim(),
      accountType: extractSelectValue(sheet.getRange(CLIENT_INFO_CELLS.ACCOUNT_TYPE).getValue()),
      accountNumber: normalizeNumericCode(sheet.getRange(CLIENT_INFO_CELLS.ACCOUNT_NUMBER).getValue(), 7),
      categoryCode: normalizeNumericCode(sheet.getRange(CLIENT_INFO_CELLS.CATEGORY_CODE).getValue(), 2),
      fileExtension: String(sheet.getRange(CLIENT_INFO_CELLS.FILE_EXTENSION).getValue() || '').trim(),
      nameOutputMode: String(sheet.getRange(CLIENT_INFO_CELLS.NAME_OUTPUT_MODE).getValue() || NAME_OUTPUT_MODES.STANDARD.value).trim()
    };
    
    // 基本的な必須項目チェック
    const errors = [];
    if (!data.clientCode) errors.push('委託者コードが設定されていません。');
    if (!data.clientName) errors.push('委託者名が設定されていません。');
    if (!data.bankCode) errors.push('取引銀行コードが設定されていません。');
    if (!data.branchCode) errors.push('取引支店コードが設定されていません。');
    if (!data.accountType) errors.push('預金種目が設定されていません。');
    if (!data.accountNumber) errors.push('口座番号が設定されていません。');
    if (!data.categoryCode) errors.push('種別コードが設定されていません。');
    
    if (errors.length > 0) {
      return { isValid: false, errors };
    }
    
    return { isValid: true, data };
    
  } catch (error) {
    Logger.log('振込依頼人情報取得エラー: ' + error.toString());
    return { isValid: false, errors: ['振込依頼人情報の取得に失敗しました: ' + error.message] };
  }
}

/**
 * フォーマット用振込データの取得
 * @return {Array} 振込データの配列
 */
function getTransferDataForFormat() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.TRANSFER_DATA);
    if (!sheet) {
      return [];
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return [];
    }
    
    const dataRange = sheet.getRange(2, 1, lastRow - 1, Object.keys(TRANSFER_DATA_COLUMNS).length);
    const values = dataRange.getValues();
    
    const result = [];
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      
      // 空行スキップ
      if (isEmptyRow(row)) continue;
      
      // 必須項目がある行のみ対象
      const bankCode = String(row[TRANSFER_DATA_COLUMNS.BANK_CODE - 1] || '').trim();
      const bankName = String(row[TRANSFER_DATA_COLUMNS.BANK_NAME - 1] || '').trim();
      const branchCode = String(row[TRANSFER_DATA_COLUMNS.BRANCH_CODE - 1] || '').trim();
      const branchName = String(row[TRANSFER_DATA_COLUMNS.BRANCH_NAME - 1] || '').trim();
      const accountNumber = String(row[TRANSFER_DATA_COLUMNS.ACCOUNT_NUMBER - 1] || '').trim();
      const recipientName = String(row[TRANSFER_DATA_COLUMNS.RECIPIENT_NAME - 1] || '').trim();
      const amount = row[TRANSFER_DATA_COLUMNS.AMOUNT - 1];
      
      if (bankCode && branchCode && accountNumber && recipientName && amount) {
        result.push({
          bankCode: normalizeNumericCode(bankCode, 4),
          bankName,
          branchCode: normalizeNumericCode(branchCode, 3),
          branchName,
          accountType: extractSelectValue(row[TRANSFER_DATA_COLUMNS.ACCOUNT_TYPE - 1]) || '1',
          accountNumber: normalizeNumericCode(accountNumber, 7),
          recipientName,
          amount: Number(amount),
          customerCode: String(row[TRANSFER_DATA_COLUMNS.CUSTOMER_CODE - 1] || '').trim(),
          identification: String(row[TRANSFER_DATA_COLUMNS.IDENTIFICATION - 1] || '').trim(),
          ediInfo: String(row[TRANSFER_DATA_COLUMNS.EDI_INFO - 1] || '').trim()
        });
      }
    }
    
    return result;
    
  } catch (error) {
    Logger.log('振込データ取得エラー: ' + error.toString());
    return [];
  }
}

/**
 * ヘッダレコードの生成
 * @param {Object} clientInfo - 振込依頼人情報
 * @param {number} dataCount - データ件数
 * @return {string} ヘッダレコード（120バイト）
 */
function createHeaderRecord(clientInfo, dataCount) {
  const fields = [];
  
  // 1. レコード種別（1桁）
  fields.push(ZENGIN_FORMAT.HEADER_TYPE);
  
  // 2. 種別コード（2桁）
  fields.push(padLeft(clientInfo.categoryCode, 2, '0'));
  
  // 3. コード区分（1桁）- 固定値'0'
  fields.push('0');
  
  // 4. 委託者コード（10桁）
  fields.push(padLeft(clientInfo.clientCode, 10, '0'));
  
  // 5. 委託者名（40桁）- 半角カナ
  fields.push(padRight(toHalfwidthKana(clientInfo.clientName), 40, ' '));
  
  // 6. 引落日（4桁）- MMDD形式
  const today = new Date();
  const withdrawalDate = String(today.getMonth() + 1).padStart(2, '0') + 
                        String(today.getDate()).padStart(2, '0');
  fields.push(withdrawalDate);
  
  // 7. 取引銀行番号（4桁）
  fields.push(padLeft(clientInfo.bankCode, 4, '0'));
  
  // 8. 取引銀行名（15桁）- 半角カナ
  fields.push(padRight(toHalfwidthKana(clientInfo.bankName), 15, ' '));
  
  // 9. 取引支店番号（3桁）
  fields.push(padLeft(clientInfo.branchCode, 3, '0'));
  
  // 10. 取引支店名（15桁）- 半角カナ
  fields.push(padRight(toHalfwidthKana(clientInfo.branchName), 15, ' '));
  
  // 11. 預金種目（1桁）
  fields.push(clientInfo.accountType);
  
  // 12. 口座番号（7桁）
  fields.push(padLeft(clientInfo.accountNumber, 7, '0'));
  
  // 13. ダミー（17桁）
  fields.push(padRight('', 17, ' '));
  
  const record = fields.join('');
  
  // 120バイトの確認
  if (record.length !== ZENGIN_FORMAT.RECORD_LENGTH) {
    Logger.log(`ヘッダレコード長エラー: ${record.length}バイト（期待値: ${ZENGIN_FORMAT.RECORD_LENGTH}バイト）`);
  }
  
  return record;
}

/**
 * データレコードの生成
 * @param {Object} transferData - 振込データ
 * @param {number} sequenceNumber - 連番
 * @param {string} nameOutputMode - 銀行名・支店名出力モード
 * @return {string} データレコード（120バイト）
 */
function createDataRecord(transferData, sequenceNumber, nameOutputMode) {
  const fields = [];
  
  // 1. レコード種別（1桁）
  fields.push(ZENGIN_FORMAT.DATA_TYPE);
  
  // 2. 銀行番号（4桁）
  fields.push(padLeft(transferData.bankCode, 4, '0'));
  
  // 3. 銀行名（15桁）- 標準モードではスペース埋め、名称出力モードでは実データ
  if (nameOutputMode === NAME_OUTPUT_MODES.OUTPUT_NAME.value && transferData.bankName) {
    fields.push(padRight(toHalfwidthKana(transferData.bankName), 15, ' '));
  } else {
    fields.push(padRight('', 15, ' '));
  }
  
  // 4. 支店番号（3桁）
  fields.push(padLeft(transferData.branchCode, 3, '0'));
  
  // 5. 支店名（15桁）- 標準モードではスペース埋め、名称出力モードでは実データ
  if (nameOutputMode === NAME_OUTPUT_MODES.OUTPUT_NAME.value && transferData.branchName) {
    fields.push(padRight(toHalfwidthKana(transferData.branchName), 15, ' '));
  } else {
    fields.push(padRight('', 15, ' '));
  }
  
  // 6. 手形交換所番号（4桁）- 未使用
  fields.push(padRight('', 4, ' '));
  
  // 7. 預金種目（1桁）
  fields.push(transferData.accountType);
  
  // 8. 口座番号（7桁）
  fields.push(padLeft(transferData.accountNumber, 7, '0'));
  
  // 9. 受取人名（30桁）- 半角カナ
  fields.push(padRight(toHalfwidthKana(transferData.recipientName), 30, ' '));
  
  // 10. 振込金額（10桁）
  fields.push(padLeft(String(Math.floor(transferData.amount)), 10, '0'));
  
  // 11. 新規コード（1桁）- 固定値'1'
  fields.push('1');
  
  // 12. 顧客番号（20桁）- EDI情報も統合
  const customerInfo = transferData.customerCode + (transferData.ediInfo || '');
  fields.push(padRight(customerInfo, 20, ' '));
  
  // 13. 振込指定区分（1桁）- 未使用
  fields.push(' ');
  
  // 14. 識別表示（1桁）
  fields.push(transferData.identification || ' ');
  
  // 15. ダミー（7桁）
  fields.push(padRight('', 7, ' '));
  
  const record = fields.join('');
  
  // 120バイトの確認
  if (record.length !== ZENGIN_FORMAT.RECORD_LENGTH) {
    Logger.log(`データレコード${sequenceNumber}長エラー: ${record.length}バイト（期待値: ${ZENGIN_FORMAT.RECORD_LENGTH}バイト）`);
  }
  
  return record;
}

/**
 * トレーラレコードの生成
 * @param {Array} transferDataList - 振込データリスト
 * @return {string} トレーラレコード（120バイト）
 */
function createTrailerRecord(transferDataList) {
  const fields = [];
  
  // 1. レコード種別（1桁）
  fields.push(ZENGIN_FORMAT.TRAILER_TYPE);
  
  // 2. 合計件数（6桁）
  fields.push(padLeft(String(transferDataList.length), 6, '0'));
  
  // 3. 合計金額（12桁）
  const totalAmount = transferDataList.reduce((sum, data) => sum + data.amount, 0);
  fields.push(padLeft(String(Math.floor(totalAmount)), 12, '0'));
  
  // 4. ダミー（101桁）
  fields.push(padRight('', 101, ' '));
  
  const record = fields.join('');
  
  // 120バイトの確認
  if (record.length !== ZENGIN_FORMAT.RECORD_LENGTH) {
    Logger.log(`トレーラレコード長エラー: ${record.length}バイト（期待値: ${ZENGIN_FORMAT.RECORD_LENGTH}バイト）`);
  }
  
  return record;
}

/**
 * エンドレコードの生成
 * @param {number} totalRecords - 総レコード数
 * @return {string} エンドレコード（120バイト）
 */
function createEndRecord(totalRecords) {
  const fields = [];
  
  // 1. レコード種別（1桁）
  fields.push(ZENGIN_FORMAT.END_TYPE);
  
  // 2. ダミー（119桁）
  fields.push(padRight('', 119, ' '));
  
  const record = fields.join('');
  
  // 120バイトの確認
  if (record.length !== ZENGIN_FORMAT.RECORD_LENGTH) {
    Logger.log(`エンドレコード長エラー: ${record.length}バイト（期待値: ${ZENGIN_FORMAT.RECORD_LENGTH}バイト）`);
  }
  
  return record;
}

/**
 * ファイル名の生成
 * @param {Object} clientInfo - 振込依頼人情報
 * @return {string} ファイル名
 */
function generateFileName(clientInfo) {
  const today = new Date();
  const dateStr = today.getFullYear().toString().slice(-2) + 
                 String(today.getMonth() + 1).padStart(2, '0') + 
                 String(today.getDate()).padStart(2, '0');
  
  const timeStr = String(today.getHours()).padStart(2, '0') + 
                 String(today.getMinutes()).padStart(2, '0');
  
  const extension = clientInfo.fileExtension || '.dat';
  
  return `FB${clientInfo.clientCode}_${dateStr}_${timeStr}${extension}`;
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
 * 全角カナを半角カナに変換
 * @param {string} str - 変換対象文字列
 * @return {string} 半角カナ文字列
 */
function toHalfwidthKana(str) {
  if (!str) return '';
  
  // 全角→半角カナの変換マップ
  const kanaMap = {
    'ア': 'ｱ', 'イ': 'ｲ', 'ウ': 'ｳ', 'エ': 'ｴ', 'オ': 'ｵ',
    'カ': 'ｶ', 'キ': 'ｷ', 'ク': 'ｸ', 'ケ': 'ｹ', 'コ': 'ｺ',
    'サ': 'ｻ', 'シ': 'ｼ', 'ス': 'ｽ', 'セ': 'ｾ', 'ソ': 'ｿ',
    'タ': 'ﾀ', 'チ': 'ﾁ', 'ツ': 'ﾂ', 'テ': 'ﾃ', 'ト': 'ﾄ',
    'ナ': 'ﾅ', 'ニ': 'ﾆ', 'ヌ': 'ﾇ', 'ネ': 'ﾈ', 'ノ': 'ﾉ',
    'ハ': 'ﾊ', 'ヒ': 'ﾋ', 'フ': 'ﾌ', 'ヘ': 'ﾍ', 'ホ': 'ﾎ',
    'マ': 'ﾏ', 'ミ': 'ﾐ', 'ム': 'ﾑ', 'メ': 'ﾒ', 'モ': 'ﾓ',
    'ヤ': 'ﾔ', 'ユ': 'ﾕ', 'ヨ': 'ﾖ',
    'ラ': 'ﾗ', 'リ': 'ﾘ', 'ル': 'ﾙ', 'レ': 'ﾚ', 'ロ': 'ﾛ',
    'ワ': 'ﾜ', 'ヲ': 'ｦ', 'ン': 'ﾝ',
    'ガ': 'ｶﾞ', 'ギ': 'ｷﾞ', 'グ': 'ｸﾞ', 'ゲ': 'ｹﾞ', 'ゴ': 'ｺﾞ',
    'ザ': 'ｻﾞ', 'ジ': 'ｼﾞ', 'ズ': 'ｽﾞ', 'ゼ': 'ｾﾞ', 'ゾ': 'ｿﾞ',
    'ダ': 'ﾀﾞ', 'ヂ': 'ﾁﾞ', 'ヅ': 'ﾂﾞ', 'デ': 'ﾃﾞ', 'ド': 'ﾄﾞ',
    'バ': 'ﾊﾞ', 'ビ': 'ﾋﾞ', 'ブ': 'ﾌﾞ', 'ベ': 'ﾍﾞ', 'ボ': 'ﾎﾞ',
    'パ': 'ﾊﾟ', 'ピ': 'ﾋﾟ', 'プ': 'ﾌﾟ', 'ペ': 'ﾍﾟ', 'ポ': 'ﾎﾟ',
    'ャ': 'ｬ', 'ュ': 'ｭ', 'ョ': 'ｮ', 'ッ': 'ｯ',
    'ァ': 'ｧ', 'ィ': 'ｨ', 'ゥ': 'ｩ', 'ェ': 'ｪ', 'ォ': 'ｫ',
    'ー': 'ｰ', '　': ' ', '・': '･'
  };
  
  let result = '';
  for (let i = 0; i < str.length; i++) {
    const char = str[i];
    result += kanaMap[char] || char;
  }
  
  return result;
}

/**
 * 数値コードの正規化（０埋め処理）
 * @param {any} value - 元の値
 * @param {number} length - 目標桁数
 * @return {string} ０埋めされた文字列
 */
function normalizeNumericCode(value, length) {
  const stringValue = String(value || '').trim();
  if (!stringValue) return '';
  
  // 数値の場合のみ０埋め処理を行う
  if (/^\d+$/.test(stringValue)) {
    return stringValue.padStart(length, '0');
  }
  
  return stringValue;
}

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
 * 全銀協フォーマットファイルのダウンロード
 * （実際の実装ではGASの制限により、DriveAPIでファイル作成後にダウンロードURLを提供）
 * @param {string} fileId - ファイルID
 * @return {string} ダウンロード用メッセージ
 */
function downloadZenginFile(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    const downloadUrl = `https://drive.google.com/file/d/${fileId}/view`;
    
    return `ファイルが作成されました。\n` +
           `ファイル名: ${file.getName()}\n` +
           `ダウンロードURL: ${downloadUrl}\n\n` +
           `※ 上記URLからファイルをダウンロードしてください。`;
  } catch (error) {
    Logger.log('ダウンロード準備エラー: ' + error.toString());
    return 'ファイルのダウンロード準備に失敗しました: ' + error.message;
  }
}

/**
 * Shift_JISエンコーディングでのBlob生成
 * @param {string} content - ファイル内容
 * @param {string} fileName - ファイル名
 * @return {Blob} Shift_JISエンコーディングのBlob
 */
function createShiftJISBlob(content, fileName) {
  try {
    // GAS制限への対応: 手動でShift_JISバイト変換
    const shiftJISBytes = convertToShiftJISBytes(content);
    
    // バイト配列からBlobを生成（MIMEタイプは指定しない）
    const blob = Utilities.newBlob(shiftJISBytes, 'application/octet-stream', fileName);
    
    Logger.log(`Shift_JISファイル生成: ${fileName}, サイズ: ${shiftJISBytes.length}バイト`);
    return blob;
    
  } catch (error) {
    Logger.log('Shift_JISBlob生成エラー: ' + error.toString());
    
    // フォールバック1: charset指定でのBlob生成
    try {
      return Utilities.newBlob(content, 'text/plain; charset=shift_jis', fileName);
    } catch (fallbackError) {
      // フォールバック2: 通常のBlob生成
      Logger.log('フォールバック実行: ' + fallbackError.toString());
      return Utilities.newBlob(content, 'text/plain', fileName);
    }
  }
}

/**
 * 文字列をShift_JISバイト配列に変換
 * @param {string} str - 変換対象文字列
 * @return {number[]} Shift_JISバイト配列
 */
function convertToShiftJISBytes(str) {
  const bytes = [];
  
  for (let i = 0; i < str.length; i++) {
    const char = str.charAt(i);
    const code = char.charCodeAt(0);
    
    // ASCII文字 (0x00-0x7F)
    if (code <= 0x7F) {
      bytes.push(code);
    }
    // 半角カナ (0xFF61-0xFF9F) → Shift_JIS (0xA1-0xDF)
    else if (code >= 0xFF61 && code <= 0xFF9F) {
      bytes.push(code - 0xFF61 + 0xA1);
    }
    // 全角スペース → 半角スペース２つ
    else if (code === 0x3000) {
      bytes.push(0x20, 0x20);
      continue;
    }
    // その他のよく使用される文字の直接マッピング
    else {
      const shiftJISBytes = getShiftJISMapping(char);
      if (shiftJISBytes) {
        shiftJISBytes.forEach(b => bytes.push(b));
      } else {
        Logger.log(`警告: Shift_JIS非対応文字 '${char}' (U+${code.toString(16).toUpperCase()}) を '?' に置換`);
        bytes.push(0x3F); // '?' のASCIIコード
      }
    }
  }
  
  return bytes;
}

/**
 * 文字のShift_JISマッピングを取得
 * @param {string} char - 文字
 * @return {number[]|null} Shift_JISバイト配列、または null
 */
function getShiftJISMapping(char) {
  // よく使用される文字のShift_JISマッピング
  const mappings = {
    // 数字関連
    '０': [0x82, 0x4F], '１': [0x82, 0x50], '２': [0x82, 0x51], '３': [0x82, 0x52], '４': [0x82, 0x53],
    '５': [0x82, 0x54], '６': [0x82, 0x55], '７': [0x82, 0x56], '８': [0x82, 0x57], '９': [0x82, 0x58],
    
    // アルファベット大文字
    'Ａ': [0x82, 0x60], 'Ｂ': [0x82, 0x61], 'Ｃ': [0x82, 0x62], 'Ｄ': [0x82, 0x63], 'Ｅ': [0x82, 0x64],
    'Ｆ': [0x82, 0x65], 'Ｇ': [0x82, 0x66], 'Ｈ': [0x82, 0x67], 'Ｉ': [0x82, 0x68], 'Ｊ': [0x82, 0x69],
    'Ｋ': [0x82, 0x6A], 'Ｌ': [0x82, 0x6B], 'Ｍ': [0x82, 0x6C], 'Ｎ': [0x82, 0x6D], 'Ｏ': [0x82, 0x6E],
    'Ｐ': [0x82, 0x6F], 'Ｑ': [0x82, 0x70], 'Ｒ': [0x82, 0x71], 'Ｓ': [0x82, 0x72], 'Ｔ': [0x82, 0x73],
    'Ｕ': [0x82, 0x74], 'Ｖ': [0x82, 0x75], 'Ｗ': [0x82, 0x76], 'Ｘ': [0x82, 0x77], 'Ｙ': [0x82, 0x78], 'Ｚ': [0x82, 0x79],
    
    // 記号
    '（': [0x81, 0x69], '）': [0x81, 0x6A], '－': [0x81, 0x5C], '＿': [0x81, 0x51],
    '．': [0x81, 0x44], '，': [0x81, 0x43], '：': [0x81, 0x46], '；': [0x81, 0x47]
  };
  
  return mappings[char] || null;
}

/**
 * 文字エンコーディングの検証
 * @param {string} content - 検証対象文字列
 * @return {Object} { isValidSJIS: boolean, message: string }
 */
function validateShiftJISCompatibility(content) {
  try {
    // Shift_JISで表現できない文字の検出
    const problematicChars = [];
    
    for (let i = 0; i < content.length; i++) {
      const char = content[i];
      const code = char.charCodeAt(0);
      
      // 基本的なShift_JIS範囲外の文字をチェック
      if (code > 0xFF7F && !isShiftJISCompatible(char)) {
        problematicChars.push(`${char}(U+${code.toString(16).toUpperCase()})`);
      }
    }
    
    if (problematicChars.length > 0) {
      return {
        isValidSJIS: false,
        message: `Shift_JIS非対応文字が含まれています: ${problematicChars.slice(0, 10).join(', ')}${problematicChars.length > 10 ? '...' : ''}`
      };
    }
    
    return {
      isValidSJIS: true,
      message: 'Shift_JIS互換性OK'
    };
    
  } catch (error) {
    return {
      isValidSJIS: false,
      message: '文字エンコーディング検証エラー: ' + error.message
    };
  }
}

/**
 * 文字がShift_JIS互換かチェック
 * @param {string} char - 文字
 * @return {boolean} Shift_JIS互換性
 */
function isShiftJISCompatible(char) {
  const code = char.charCodeAt(0);
  
  // ASCII範囲
  if (code <= 0x7F) return true;
  
  // 半角カナ範囲
  if (code >= 0xFF61 && code <= 0xFF9F) return true;
  
  // 一般的な日本語文字（簡易チェック）
  if (code >= 0x3040 && code <= 0x309F) return true; // ひらがな
  if (code >= 0x30A0 && code <= 0x30FF) return true; // カタカナ
  if (code >= 0x4E00 && code <= 0x9FAF) return true; // 漢字（基本範囲）
  
  return false;
}

/**
 * Shift_JIS変換のテスト用関数
 * @param {string} testString - テスト文字列
 * @return {Object} テスト結果
 */
function testShiftJISConversion(testString) {
  if (!testString) {
    testString = 'ﾃｽﾄｷｷﾞｮｳ ABC123 ･';
  }
  
  try {
    const bytes = convertToShiftJISBytes(testString);
    const hexString = bytes.map(b => '0x' + b.toString(16).toUpperCase().padStart(2, '0')).join(' ');
    
    Logger.log(`入力文字列: "${testString}"`);
    Logger.log(`変換後バイト数: ${bytes.length}`);
    Logger.log(`16進表示: ${hexString}`);
    
    return {
      success: true,
      input: testString,
      byteCount: bytes.length,
      hexBytes: hexString,
      bytes: bytes
    };
    
  } catch (error) {
    Logger.log('テスト変換エラー: ' + error.toString());
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 全銀協フォーマット生成のデバッグ版
 * @return {Object} デバッグ情報を含む生成結果
 */
function createZenginFormatFileDebug() {
  Logger.log('=== 全銀協フォーマット生成デバッグ開始 ===');
  
  const result = createZenginFormatFile();
  
  if (result.success) {
    Logger.log(`ファイル生成成功: ${result.fileName}`);
    Logger.log(`ファイルID: ${result.fileId}`);
    
    // 生成されたファイルの内容をチェック
    try {
      const file = DriveApp.getFileById(result.fileId);
      const blob = file.getBlob();
      const bytes = blob.getBytes();
      
      Logger.log(`実際のファイルサイズ: ${bytes.length}バイト`);
      Logger.log(`先頭20バイトの16進: ${bytes.slice(0, 20).map(b => '0x' + (b & 0xFF).toString(16).toUpperCase().padStart(2, '0')).join(' ')}`);
      
      result.debugInfo = {
        actualFileSize: bytes.length,
        firstBytesHex: bytes.slice(0, 20).map(b => '0x' + (b & 0xFF).toString(16).toUpperCase().padStart(2, '0')).join(' ')
      };
      
    } catch (debugError) {
      Logger.log('デバッグ情報取得エラー: ' + debugError.toString());
    }
  }
  
  Logger.log('=== 全銀協フォーマット生成デバッグ終了 ===');
  return result;
}

/**
 * データレコードフォーマットのテスト（開発用）
 * 標準モードと名称出力モードの違いを確認
 */
function testDataRecordFormat() {
  const testData = {
    bankCode: '0001',
    bankName: 'ﾐｽﾞﾎ',
    branchCode: '001',
    branchName: 'ﾎﾝﾃﾝ',
    accountType: '1',
    accountNumber: '1234567',
    recipientName: 'ﾀﾅｶ ﾀﾛｳ',
    amount: 100000,
    customerCode: 'EMP001',
    identification: 'Y',
    ediInfo: ''
  };
  
  Logger.log('=== データレコードフォーマットテスト ===');
  
  // 標準モード
  const standardRecord = createDataRecord(testData, 1, NAME_OUTPUT_MODES.STANDARD.value);
  Logger.log('標準モード（銀行名・支店名スペース埋め）:');
  Logger.log(standardRecord);
  Logger.log(`レコード長: ${standardRecord.length}バイト`);
  Logger.log('');
  
  // 名称出力モード
  const nameOutputRecord = createDataRecord(testData, 1, NAME_OUTPUT_MODES.OUTPUT_NAME.value);
  Logger.log('名称出力モード（銀行名・支店名出力）:');
  Logger.log(nameOutputRecord);
  Logger.log(`レコード長: ${nameOutputRecord.length}バイト`);
  Logger.log('');
  
  // フィールド位置の詳細
  Logger.log('フィールド位置（バイト単位）:');
  Logger.log('1: レコード種別（1バイト）');
  Logger.log('2-5: 銀行番号（4バイト）');
  Logger.log('6-20: 銀行名（15バイト）← 標準/名称出力の違い');
  Logger.log('21-23: 支店番号（3バイト）');
  Logger.log('24-38: 支店名（15バイト）← 標準/名称出力の違い');
  Logger.log('39-42: 手形交換所番号（4バイト）- 未使用');
  Logger.log('43: 預金種目（1バイト）');
  Logger.log('44-50: 口座番号（7バイト）');
  Logger.log('51-80: 受取人名（30バイト）');
  Logger.log('81-90: 振込金額（10バイト）');
  Logger.log('91: 新規コード（1バイト）');
  Logger.log('92-111: 顧客番号（20バイト）');
  Logger.log('112: 振込指定区分（1バイト）- 未使用');
  Logger.log('113: 識別表示（1バイト）');
  Logger.log('114-120: ダミー（7バイト）');
  
  return {
    standardRecord,
    nameOutputRecord,
    testData
  };
} 