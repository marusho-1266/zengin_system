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
      
      // 詳細な警告をユーザーに表示
      const detailWarning = `文字エンコーディングの問題が検出されました。\n\n` +
        encodingCheck.message + `\n\n` +
        `これらの文字はShift_JISで正しく表現できないため、` +
        `金融機関での処理時にエラーとなる可能性があります。\n\n` +
        `推奨対応:\n` +
        `1. 問題のある文字を半角カナまたは英数字に置換してください\n` +
        `2. 受取人名や銀行名に特殊文字が含まれていないか確認してください\n` +
        `3. それでも解決しない場合は、システム管理者にご相談ください`;
      
      SpreadsheetApp.getUi().alert('文字エンコーディング警告', detailWarning, SpreadsheetApp.getUi().ButtonSet.OK);
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
      clientCode: normalizeNumericCode(sheet.getRange(CLIENT_INFO_CELLS.CLIENT_CODE).getValue(), 10),
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
          amount: Math.floor(Number(amount)),
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
    
    Logger.log(`Shift_JISファイル生成成功: ${fileName}, サイズ: ${shiftJISBytes.length}バイト`);
    return blob;
    
  } catch (error) {
    Logger.log('Shift_JISBlob生成エラー: ' + error.toString());
    
    // フォールバック: UTF-8でのBlob生成（警告付き）
    const warningMessage = `
【警告】Shift_JISエンコーディングに失敗しました。
UTF-8形式でファイルを生成します。
金融機関によってはこのファイルが受け付けられない可能性があります。

推奨対応:
1. 受取人名や銀行名に特殊文字が含まれていないか確認してください
2. 全角文字を半角カナに変換してください
3. 記号は全銀協フォーマット対応文字のみを使用してください

エラー詳細: ${error.message}
`;
    
    SpreadsheetApp.getUi().alert('文字エンコーディング警告', warningMessage, SpreadsheetApp.getUi().ButtonSet.OK);
    
    // UTF-8でのBlob生成
    Logger.log('フォールバック実行: UTF-8形式でファイル生成');
    return Utilities.newBlob(content, 'text/plain; charset=utf-8', fileName);
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
  // 全銀協フォーマットで使用される文字のShift_JISマッピング
  const mappings = {
    // === 半角カナ文字（濁点・半濁点を含む） ===
    // 基本カナ
    'ｱ': [0xB1], 'ｲ': [0xB2], 'ｳ': [0xB3], 'ｴ': [0xB4], 'ｵ': [0xB5],
    'ｶ': [0xB6], 'ｷ': [0xB7], 'ｸ': [0xB8], 'ｹ': [0xB9], 'ｺ': [0xBA],
    'ｻ': [0xBB], 'ｼ': [0xBC], 'ｽ': [0xBD], 'ｾ': [0xBE], 'ｿ': [0xBF],
    'ﾀ': [0xC0], 'ﾁ': [0xC1], 'ﾂ': [0xC2], 'ﾃ': [0xC3], 'ﾄ': [0xC4],
    'ﾅ': [0xC5], 'ﾆ': [0xC6], 'ﾇ': [0xC7], 'ﾈ': [0xC8], 'ﾉ': [0xC9],
    'ﾊ': [0xCA], 'ﾋ': [0xCB], 'ﾌ': [0xCC], 'ﾍ': [0xCD], 'ﾎ': [0xCE],
    'ﾏ': [0xCF], 'ﾐ': [0xD0], 'ﾑ': [0xD1], 'ﾒ': [0xD2], 'ﾓ': [0xD3],
    'ﾔ': [0xD4], 'ﾕ': [0xD5], 'ﾖ': [0xD6],
    'ﾗ': [0xD7], 'ﾘ': [0xD8], 'ﾙ': [0xD9], 'ﾚ': [0xDA], 'ﾛ': [0xDB],
    'ﾜ': [0xDC], 'ｦ': [0xA6], 'ﾝ': [0xDD],
    
    // 小文字
    'ｧ': [0xA7], 'ｨ': [0xA8], 'ｩ': [0xA9], 'ｪ': [0xAA], 'ｫ': [0xAB],
    'ｬ': [0xAC], 'ｭ': [0xAD], 'ｮ': [0xAE], 'ｯ': [0xAF],
    
    // 濁点・半濁点（単独）
    'ﾞ': [0xDE], 'ﾟ': [0xDF],
    
    // 濁音（組み合わせ）
    'ｶﾞ': [0xB6, 0xDE], 'ｷﾞ': [0xB7, 0xDE], 'ｸﾞ': [0xB8, 0xDE], 'ｹﾞ': [0xB9, 0xDE], 'ｺﾞ': [0xBA, 0xDE],
    'ｻﾞ': [0xBB, 0xDE], 'ｼﾞ': [0xBC, 0xDE], 'ｽﾞ': [0xBD, 0xDE], 'ｾﾞ': [0xBE, 0xDE], 'ｿﾞ': [0xBF, 0xDE],
    'ﾀﾞ': [0xC0, 0xDE], 'ﾁﾞ': [0xC1, 0xDE], 'ﾂﾞ': [0xC2, 0xDE], 'ﾃﾞ': [0xC3, 0xDE], 'ﾄﾞ': [0xC4, 0xDE],
    'ﾊﾞ': [0xCA, 0xDE], 'ﾋﾞ': [0xCB, 0xDE], 'ﾌﾞ': [0xCC, 0xDE], 'ﾍﾞ': [0xCD, 0xDE], 'ﾎﾞ': [0xCE, 0xDE],
    
    // 半濁音（組み合わせ）
    'ﾊﾟ': [0xCA, 0xDF], 'ﾋﾟ': [0xCB, 0xDF], 'ﾌﾟ': [0xCC, 0xDF], 'ﾍﾟ': [0xCD, 0xDF], 'ﾎﾟ': [0xCE, 0xDF],
    
    // === 記号類 ===
    'ｰ': [0xB0],    // 長音
    '･': [0xA5],    // 中点
    '｡': [0xA1],    // 句点
    '､': [0xA4],    // 読点
    '｢': [0xA2],    // 開き鍵括弧
    '｣': [0xA3],    // 閉じ鍵括弧
    
    // === 全角文字のマッピング（必要な場合） ===
    // 全角数字
    '０': [0x82, 0x4F], '１': [0x82, 0x50], '２': [0x82, 0x51], '３': [0x82, 0x52], '４': [0x82, 0x53],
    '５': [0x82, 0x54], '６': [0x82, 0x55], '７': [0x82, 0x56], '８': [0x82, 0x57], '９': [0x82, 0x58],
    
    // 全角アルファベット大文字
    'Ａ': [0x82, 0x60], 'Ｂ': [0x82, 0x61], 'Ｃ': [0x82, 0x62], 'Ｄ': [0x82, 0x63], 'Ｅ': [0x82, 0x64],
    'Ｆ': [0x82, 0x65], 'Ｇ': [0x82, 0x66], 'Ｈ': [0x82, 0x67], 'Ｉ': [0x82, 0x68], 'Ｊ': [0x82, 0x69],
    'Ｋ': [0x82, 0x6A], 'Ｌ': [0x82, 0x6B], 'Ｍ': [0x82, 0x6C], 'Ｎ': [0x82, 0x6D], 'Ｏ': [0x82, 0x6E],
    'Ｐ': [0x82, 0x6F], 'Ｑ': [0x82, 0x70], 'Ｒ': [0x82, 0x71], 'Ｓ': [0x82, 0x72], 'Ｔ': [0x82, 0x73],
    'Ｕ': [0x82, 0x74], 'Ｖ': [0x82, 0x75], 'Ｗ': [0x82, 0x76], 'Ｘ': [0x82, 0x77], 'Ｙ': [0x82, 0x78], 
    'Ｚ': [0x82, 0x79],
    
    // 全角記号
    '（': [0x81, 0x69], '）': [0x81, 0x6A], '－': [0x81, 0x5C], '＿': [0x81, 0x51],
    '．': [0x81, 0x44], '，': [0x81, 0x43], '：': [0x81, 0x46], '；': [0x81, 0x47],
    '／': [0x81, 0x5E], '＼': [0x81, 0x5F], '＠': [0x81, 0x97], '＃': [0x81, 0x94],
    '＄': [0x81, 0x90], '％': [0x81, 0x93], '＆': [0x81, 0x95], '＊': [0x81, 0x96],
    '＋': [0x81, 0x7B], '＝': [0x81, 0x81], '＜': [0x81, 0x83], '＞': [0x81, 0x84],
    '？': [0x81, 0x48], '！': [0x81, 0x49], '～': [0x81, 0x60], '￥': [0x81, 0x8F],
    
    // 全角スペース
    '　': [0x81, 0x40]
  };
  
  return mappings[char] || null;
}

/**
 * 文字エンコーディングの検証
 * @param {string} content - 検証対象文字列
 * @return {Object} { isValidSJIS: boolean, message: string, problematicChars: Array }
 */
function validateShiftJISCompatibility(content) {
  try {
    // Shift_JISで表現できない文字の検出
    const problematicChars = [];
    const charLocations = [];
    
    const lines = content.split(/\r?\n/);
    for (let lineNum = 0; lineNum < lines.length; lineNum++) {
      const line = lines[lineNum];
      for (let charPos = 0; charPos < line.length; charPos++) {
        const char = line[charPos];
        const code = char.charCodeAt(0);
        
        // ASCII範囲はOK
        if (code <= 0x7F) continue;
        
        // 半角カナ範囲はOK
        if (code >= 0xFF61 && code <= 0xFF9F) continue;
        
        // Shift_JISマッピングが存在するかチェック
        const shiftJISMapping = getShiftJISMapping(char);
        if (!shiftJISMapping) {
          problematicChars.push({
            char: char,
            code: code,
            line: lineNum + 1,
            position: charPos + 1,
            hex: 'U+' + code.toString(16).toUpperCase().padStart(4, '0')
          });
        }
      }
    }
    
    if (problematicChars.length > 0) {
      const detailMessage = problematicChars.slice(0, 5).map(info => 
        `'${info.char}' (${info.hex}) at 行${info.line}:列${info.position}`
      ).join(', ');
      
      return {
        isValidSJIS: false,
        message: `Shift_JIS非対応文字が${problematicChars.length}個見つかりました: ${detailMessage}${problematicChars.length > 5 ? '...' : ''}`,
        problematicChars: problematicChars
      };
    }
    
    return {
      isValidSJIS: true,
      message: 'Shift_JIS互換性OK',
      problematicChars: []
    };
    
  } catch (error) {
    return {
      isValidSJIS: false,
      message: '文字エンコーディング検証エラー: ' + error.message,
      problematicChars: []
    };
  }
}







 