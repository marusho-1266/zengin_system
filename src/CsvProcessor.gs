/**
 * CSV処理機能
 * 振込データと金融機関マスタのCSVインポート機能
 */

/**
 * 振込データCSVの取込処理（エラー行スキップ対応版）
 * @param {string} csvData - CSVデータ（文字列）
 * @param {string} importMode - 取込モード（'overwrite' or 'append'）
 * @return {string} 処理結果メッセージ
 */
function importTransferDataFromCsv(csvData, importMode) {
  try {
    logSystemActivity('importTransferDataFromCsv', `CSV取込処理開始 - モード: ${importMode}`, 'INFO');
    
    // CSVデータの解析
    const parsedData = parseCSV(csvData);
    logSystemActivity('importTransferDataFromCsv', `CSVデータ解析完了 - ${parsedData.length}行`, 'INFO');
    
    if (parsedData.length === 0) {
      logSystemActivity('importTransferDataFromCsv', 'CSVデータが空です', 'WARNING');
      throw new Error('CSVデータが空です。');
    }
    
    // ヘッダー行を除く
    const dataRows = parsedData.slice(1);
    if (dataRows.length === 0) {
      logSystemActivity('importTransferDataFromCsv', 'データ行がありません（ヘッダーのみ）', 'WARNING');
      throw new Error('データ行がありません。ヘッダー行のみのCSVファイルです。');
    }
    
    // 事前全体検証
    const preValidationResult = preValidateTransferDataRows(dataRows);
    
    // エラーがある場合はユーザーに選択肢を提示
    if (preValidationResult.errorRows.length > 0) {
      const continueWithErrors = showErrorConfirmationDialog(preValidationResult);
      if (!continueWithErrors) {
        logSystemActivity('importTransferDataFromCsv', 'ユーザーによる処理中断', 'INFO');
        throw new Error('エラーが検出されたため処理を中断しました。');
      }
    }
    
    // 振込データシートを取得
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEET_NAMES.TRANSFER_DATA);
    
    if (!sheet) {
      logSystemActivity('importTransferDataFromCsv', '振込データシートが見つかりません', 'ERROR');
      throw new Error('振込データシートが見つかりません。');
    }
    
    logSystemActivity('importTransferDataFromCsv', `振込データシート名: ${sheet.getName()}`, 'INFO');
    
    let startRow = 2; // ヘッダー行の次から開始
    
    if (importMode === 'overwrite') {
      // 既存データをクリア
      const lastRow = sheet.getLastRow();
      logSystemActivity('importTransferDataFromCsv', `最終行: ${lastRow}`, 'INFO');
      
      if (lastRow > 1) {
        sheet.getRange(2, 1, lastRow - 1, Object.keys(TRANSFER_DATA_COLUMNS).length).clear();
        logSystemActivity('importTransferDataFromCsv', `既存データクリア完了 - ${lastRow - 1}行削除`, 'INFO');
      }
    } else if (importMode === 'append') {
      // 追記モードの場合は最後の行の次から開始
      startRow = sheet.getLastRow() + 1;
      logSystemActivity('importTransferDataFromCsv', `追記モード - 開始行: ${startRow}`, 'INFO');
    }
    
    // データの書き込み（エラー行スキップ対応）
    const processedData = [];
    let successCount = 0;
    let skipCount = 0;
    
    for (let i = 0; i < dataRows.length; i++) {
      const rowNum = i + 2; // ヘッダー行を考慮
      try {
        // 行の検証
        const rowErrors = validateCsvRow(dataRows[i], rowNum);
        if (rowErrors.length > 0) {
          skipCount++;
          logSystemActivity('importTransferDataFromCsv', `行${rowNum}をスキップ: ${rowErrors[0]}`, 'WARNING');
          continue;
        }
        
        // 正常行の処理
        const processedRow = processTransferDataRow(dataRows[i]);
        processedData.push(processedRow);
        successCount++;
      } catch (error) {
        skipCount++;
        logSystemActivity('importTransferDataFromCsv', `行${rowNum}処理エラーによりスキップ: ${error.message}`, 'WARNING');
      }
    }
    
    // 正常データの書き込み
    if (processedData.length > 0) {
      sheet.getRange(startRow, 1, processedData.length, processedData[0].length).setValues(processedData);
      logSystemActivity('importTransferDataFromCsv', `データ書き込み完了 - 成功: ${successCount}件, スキップ: ${skipCount}件`, 'INFO');
    }
    
    // 自動補完実行（オプション）
    try {
      logSystemActivity('importTransferDataFromCsv', '自動補完処理開始', 'INFO');
      const autoCompleteResult = bulkAutoComplete();
      logSystemActivity('importTransferDataFromCsv', `自動補完完了 - 銀行名: ${autoCompleteResult.bankNameCompletions}件, 支店名: ${autoCompleteResult.branchNameCompletions}件`, 'INFO');
    } catch (autoCompleteError) {
      logSystemActivity('importTransferDataFromCsv', `自動補完エラー: ${autoCompleteError.message}`, 'ERROR');
      Logger.log('自動補完エラー: ' + autoCompleteError.toString());
      // 自動補完エラーは無視して継続
    }
    
    const resultMessage = `CSV取込完了\n` +
                         `モード: ${importMode === 'overwrite' ? '上書き' : '追記'}\n` +
                         `総件数: ${dataRows.length}件\n` +
                         `成功: ${successCount}件\n` +
                         `スキップ: ${skipCount}件\n` +
                         `開始行: ${startRow}行`;
    
    logSystemActivity('importTransferDataFromCsv', resultMessage.replace(/\n/g, ', '), 'INFO');
    Logger.log(resultMessage);
    return resultMessage;
    
  } catch (error) {
    logSystemActivity('importTransferDataFromCsv', `エラー: ${error.message}`, 'ERROR');
    Logger.log('CSV取込エラー: ' + error.toString());
    throw error;
  }
}

/**
 * 金融機関マスタCSVの取込処理
 * @param {string} csvData - CSVデータ（文字列）
 * @param {boolean} duplicateCheck - 重複チェックを行うかどうか
 * @return {string} 処理結果メッセージ
 */
function importBankMasterFromCsv(csvData, duplicateCheck = true) {
  try {
    if (!csvData || csvData.trim() === '') {
      throw new Error('CSVデータが空です。');
    }
    
    // CSVデータの解析
    const rows = parseCSV(csvData);
    if (rows.length === 0) {
      throw new Error('有効なデータが見つかりません。');
    }
    
    // ヘッダー行の確認（オプション）
    const expectedHeaders = Object.values(BANK_MASTER_HEADERS);
    let dataStartIndex = 0;
    
    // 最初の行がヘッダーかチェック
    if (isHeaderRow(rows[0], expectedHeaders)) {
      dataStartIndex = 1;
    }
    
    const dataRows = rows.slice(dataStartIndex);
    if (dataRows.length === 0) {
      throw new Error('データ行が見つかりません。');
    }
    
    // データ検証
    const validationResult = validateBankMasterCsvData(dataRows);
    if (!validationResult.isValid) {
      throw new Error('データ検証エラー:\n' + validationResult.errors.slice(0, 10).join('\n') + 
                     (validationResult.errors.length > 10 ? '\n...他' + (validationResult.errors.length - 10) + '件' : ''));
    }
    
    // 金融機関マスタシートの取得
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.BANK_MASTER);
    if (!sheet) {
      throw new Error('金融機関マスタシートが見つかりません。');
    }
    
    // 重複チェック処理
    let newDataCount = 0;
    let updateCount = 0;
    let skipCount = 0;
    
    if (duplicateCheck) {
      const existingData = getBankMasterData();
      const existingMap = new Map();
      
      // 既存データのマップを作成
      existingData.forEach((row, index) => {
        const key = `${row[BANK_MASTER_COLUMNS.BANK_CODE - 1]}-${row[BANK_MASTER_COLUMNS.BRANCH_CODE - 1]}`;
        existingMap.set(key, { row, index: index + 2 }); // +2はヘッダー行とインデックス調整
      });
      
      // 新しいデータの処理
      const processedRows = [];
      
      for (const csvRow of dataRows) {
        const processedRow = processBankMasterDataRow(csvRow);
        const key = `${processedRow[BANK_MASTER_COLUMNS.BANK_CODE - 1]}-${processedRow[BANK_MASTER_COLUMNS.BRANCH_CODE - 1]}`;
        
        if (existingMap.has(key)) {
          // 既存データの更新
          const existing = existingMap.get(key);
          sheet.getRange(existing.index, 1, 1, processedRow.length).setValues([processedRow]);
          updateCount++;
        } else {
          // 新規データとして追加
          processedRows.push(processedRow);
          newDataCount++;
        }
      }
      
      // 新規データの追加
      if (processedRows.length > 0) {
        const startRow = sheet.getLastRow() + 1;
        sheet.getRange(startRow, 1, processedRows.length, processedRows[0].length).setValues(processedRows);
      }
      
    } else {
      // 重複チェックなしの場合は単純に追加
      const processedData = dataRows.map(row => processBankMasterDataRow(row));
      const startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, processedData.length, processedData[0].length).setValues(processedData);
      newDataCount = processedData.length;
    }
    
    // キャッシュをクリア
    CacheService.getScriptCache().remove(CACHE_CONFIG.BANK_MASTER_KEY);
    
    const resultMessage = `金融機関マスタ取込完了\n` +
                         `新規追加: ${newDataCount}件\n` +
                         `更新: ${updateCount}件\n` +
                         `スキップ: ${skipCount}件\n` +
                         `重複チェック: ${duplicateCheck ? '有効' : '無効'}`;
    
    Logger.log(resultMessage);
    return resultMessage;
    
  } catch (error) {
    Logger.log('金融機関マスタCSV取込エラー: ' + error.toString());
    throw error;
  }
}

/**
 * CSVデータの解析
 * @param {string} csvData - CSVデータ
 * @return {Array[]} 解析されたデータの2次元配列
 */
function parseCSV(csvData) {
  try {
    const lines = csvData.split(/\r?\n/);
    const result = [];
    
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i].trim();
      if (line === '') continue; // 空行をスキップ
      
      // 簡易CSV解析（カンマ区切り、ダブルクォート対応）
      const row = [];
      let current = '';
      let inQuotes = false;
      
      for (let j = 0; j < line.length; j++) {
        const char = line[j];
        
        if (char === '"' && !inQuotes) {
          inQuotes = true;
        } else if (char === '"' && inQuotes) {
          if (j + 1 < line.length && line[j + 1] === '"') {
            // エスケープされたクォート
            current += '"';
            j++; // 次の文字をスキップ
          } else {
            inQuotes = false;
          }
        } else if (char === ',' && !inQuotes) {
          row.push(current.trim());
          current = '';
        } else {
          current += char;
        }
      }
      
      // 最後のフィールドを追加
      row.push(current.trim());
      
      if (row.some(cell => cell !== '')) {
        result.push(row);
      }
    }
    
    return result;
  } catch (error) {
    Logger.log('CSV解析エラー: ' + error.toString());
    throw new Error('CSVデータの解析に失敗しました: ' + error.message);
  }
}

/**
 * ヘッダー行かどうかを判定
 * @param {Array} row - 行データ
 * @param {Array} expectedHeaders - 期待されるヘッダー
 * @return {boolean} ヘッダー行かどうか
 */
function isHeaderRow(row, expectedHeaders) {
  if (!row || row.length === 0) return false;
  
  // 少なくとも半分以上のヘッダーが一致すればヘッダー行と判定
  let matchCount = 0;
  const threshold = Math.ceil(expectedHeaders.length / 2);
  
  for (let i = 0; i < Math.min(row.length, expectedHeaders.length); i++) {
    if (row[i] && expectedHeaders[i] && 
        row[i].toString().includes(expectedHeaders[i]) || 
        expectedHeaders[i].includes(row[i].toString())) {
      matchCount++;
    }
  }
  
  return matchCount >= threshold;
}

/**
 * 振込データCSVの検証（仕様書準拠の8項目フォーマット）
 * @param {Array[]} dataRows - データ行の配列
 * @return {Object} 検証結果
 */
function validateCsvData(dataRows) {
  const errors = [];
  
  // 最大件数チェック
  if (dataRows.length > VALIDATION_RULES.MAX_RECORDS) {
    errors.push(`処理可能件数(${VALIDATION_RULES.MAX_RECORDS}件)を超えています。データ件数: ${dataRows.length}件`);
  }
  
  // 各行の検証
  for (let i = 0; i < dataRows.length; i++) {
    const row = dataRows[i];
    const rowNum = i + 1;
    
    // 最小必要列数チェック（仕様書準拠の8項目）
    if (row.length < 6) { // 銀行コード〜振込金額まで最低6列（必須項目）
      errors.push(`行${rowNum}: 必要な列数が不足しています。現在: ${row.length}列, 最低: 6列`);
      continue;
    }
    
    // 必須項目チェック（仕様書準拠フォーマット）
    const bankCode = (row[0] || '').toString().trim();         // 銀行コード
    const branchCode = (row[1] || '').toString().trim();       // 支店コード
    const accountType = (row[2] || '').toString().trim();      // 預金種目
    const accountNumber = (row[3] || '').toString().trim();    // 口座番号
    const recipientName = (row[4] || '').toString().trim();    // 受取人名
    const amount = (row[5] || '').toString().trim();           // 振込金額
    
    if (!bankCode) errors.push(`行${rowNum}: 銀行コードが入力されていません。`);
    if (!branchCode) errors.push(`行${rowNum}: 支店コードが入力されていません。`);
    if (!accountType) errors.push(`行${rowNum}: 預金種目が入力されていません。`);
    if (!accountNumber) errors.push(`行${rowNum}: 口座番号が入力されていません。`);
    if (!recipientName) errors.push(`行${rowNum}: 受取人名が入力されていません。`);
    if (!amount) errors.push(`行${rowNum}: 振込金額が入力されていません。`);
    
    // 形式チェック
    if (bankCode && !/^\d{4}$/.test(bankCode)) {
      errors.push(`行${rowNum}: 銀行コードは4桁の数字で入力してください。`);
    }
    if (branchCode && !/^\d{3}$/.test(branchCode)) {
      errors.push(`行${rowNum}: 支店コードは3桁の数字で入力してください。`);
    }
    if (accountType && !Object.values(ACCOUNT_TYPES).some(type => type.code === accountType)) {
      errors.push(`行${rowNum}: 預金種目は1(普通), 2(当座)のいずれかを入力してください。（4:貯蓄は全銀協対象外）`);
    }
    if (amount && (isNaN(parseFloat(amount)) || parseFloat(amount) <= 0)) {
      errors.push(`行${rowNum}: 振込金額は正の数値で入力してください。`);
    }
          if (recipientName && !/^[A-Z0-9ｱ-ﾝﾞﾟｧ-ｯ ・ー().\-/]+$/.test(recipientName)) {
      errors.push(`行${rowNum}: 受取人名は全銀協フォーマット対応文字（半角カナ・英数字・記号、カンマ除く）で入力してください。`);
    }
    
    // オプション項目の検証（顧客コード、識別表示）
    const customerCode = (row[6] || '').toString().trim();
    const identification = (row[7] || '').toString().trim();
    
    if (customerCode && customerCode.length > VALIDATION_RULES.CUSTOMER_CODE_MAX_LENGTH) {
      errors.push(`行${rowNum}: 顧客コードは${VALIDATION_RULES.CUSTOMER_CODE_MAX_LENGTH}文字以内で入力してください。`);
    }
    if (identification && !/^[YB]?$/.test(identification)) {
      errors.push(`行${rowNum}: 識別表示はY(給与)、B(賞与)、または空白で入力してください。`);
    }
  }
  
  return { isValid: errors.length === 0, errors };
}

/**
 * 金融機関マスタCSVの検証（5項目フォーマット対応）
 * @param {Array[]} dataRows - データ行の配列
 * @return {Object} 検証結果
 */
function validateBankMasterCsvData(dataRows) {
  const errors = [];
  
  // 各行の検証
  for (let i = 0; i < dataRows.length; i++) {
    const row = dataRows[i];
    const rowNum = i + 1;
    
    // 最小必要列数チェック（銀行コード、銀行名、支店コード、支店名の最低4列）
    if (row.length < 4) {
      errors.push(`行${rowNum}: 必要な列数が不足しています。現在: ${row.length}列, 最低: 4列`);
      continue;
    }
    
    // 必須項目チェック
    const bankCode = (row[0] || '').toString().trim();
    const bankName = (row[1] || '').toString().trim();
    const branchCode = (row[2] || '').toString().trim();
    const branchName = (row[3] || '').toString().trim();
    const status = (row[4] || '').toString().trim(); // 5列目が状態（オプション）
    
    if (!bankCode) errors.push(`行${rowNum}: 銀行コードが入力されていません。`);
    if (!bankName) errors.push(`行${rowNum}: 銀行名が入力されていません。`);
    if (!branchCode) errors.push(`行${rowNum}: 支店コードが入力されていません。`);
    if (!branchName) errors.push(`行${rowNum}: 支店名が入力されていません。`);
    
    // 形式チェック
    if (bankCode && !/^\d{4}$/.test(bankCode)) {
      errors.push(`行${rowNum}: 銀行コードは4桁の数字で入力してください。`);
    }
    if (branchCode && !/^\d{3}$/.test(branchCode)) {
      errors.push(`行${rowNum}: 支店コードは3桁の数字で入力してください。`);
    }
    
    // 状態の検証（CSVに5列目があり、値が設定されている場合のみ）
    if (row.length > 4 && status) {
      if (!STATUS_OPTIONS.includes(status)) {
        errors.push(`行${rowNum}: 状態は「有効」または「無効」で入力してください。現在値: ${status}`);
      }
    }
  }
  
  return { isValid: errors.length === 0, errors };
}

/**
 * 振込データ行の処理（仕様書準拠の8項目フォーマット）
 * @param {Array} csvRow - CSV行データ
 * @return {Array} 処理されたシート用データ
 */
function processTransferDataRow(csvRow) {
  const result = new Array(Object.keys(TRANSFER_DATA_COLUMNS).length).fill('');
  
  // CSVデータをシート列に対応付け（仕様書準拠の8項目フォーマット）
  result[TRANSFER_DATA_COLUMNS.BANK_CODE - 1] = (csvRow[0] || '').toString().trim();    // 銀行コード
  // 銀行名は空白（自動補完で設定）
  result[TRANSFER_DATA_COLUMNS.BANK_NAME - 1] = '';
  result[TRANSFER_DATA_COLUMNS.BRANCH_CODE - 1] = (csvRow[1] || '').toString().trim();  // 支店コード
  // 支店名は空白（自動補完で設定）
  result[TRANSFER_DATA_COLUMNS.BRANCH_NAME - 1] = '';
  result[TRANSFER_DATA_COLUMNS.ACCOUNT_TYPE - 1] = (csvRow[2] || '').toString().trim(); // 預金種目
  result[TRANSFER_DATA_COLUMNS.ACCOUNT_NUMBER - 1] = (csvRow[3] || '').toString().trim(); // 口座番号
  result[TRANSFER_DATA_COLUMNS.RECIPIENT_NAME - 1] = (csvRow[4] || '').toString().trim(); // 受取人名
  
  // 振込金額は数値に変換
  const amount = (csvRow[5] || '').toString().trim();
  result[TRANSFER_DATA_COLUMNS.AMOUNT - 1] = amount ? parseFloat(amount) : '';
  
  // オプション項目
  result[TRANSFER_DATA_COLUMNS.CUSTOMER_CODE - 1] = (csvRow[6] || '').toString().trim(); // 顧客コード
  result[TRANSFER_DATA_COLUMNS.IDENTIFICATION - 1] = (csvRow[7] || '').toString().trim(); // 識別表示
  // EDI情報は空白（仕様書では対象外）
  result[TRANSFER_DATA_COLUMNS.EDI_INFO - 1] = '';
  
  return result;
}

/**
 * 金融機関マスタ行の処理
 * @param {Array} csvRow - CSV行データ
 * @return {Array} 処理されたシート用データ
 */
function processBankMasterDataRow(csvRow) {
  const result = new Array(Object.keys(BANK_MASTER_COLUMNS).length).fill('');
  
  // CSVデータをシート列に対応付け
  result[BANK_MASTER_COLUMNS.BANK_CODE - 1] = (csvRow[0] || '').toString().trim();
  result[BANK_MASTER_COLUMNS.BANK_NAME - 1] = (csvRow[1] || '').toString().trim();
  result[BANK_MASTER_COLUMNS.BRANCH_CODE - 1] = (csvRow[2] || '').toString().trim();
  result[BANK_MASTER_COLUMNS.BRANCH_NAME - 1] = (csvRow[3] || '').toString().trim();
  
  // 更新日の設定（常に現在日時を設定）
  result[BANK_MASTER_COLUMNS.UPDATE_DATE - 1] = new Date();
  
  // 状態の設定（CSVの5列目があれば使用、なければ'有効'）
  if (csvRow.length > 4 && csvRow[4] && csvRow[4].toString().trim()) {
    result[BANK_MASTER_COLUMNS.STATUS - 1] = csvRow[4].toString().trim();
  } else {
    result[BANK_MASTER_COLUMNS.STATUS - 1] = '有効';
  }
  
  return result;
}

/**
 * CSVデータのデバッグ用関数
 * @param {string} csvData - CSVデータ
 */
function debugCsvData(csvData) {
  try {
    Logger.log('=== CSV解析デバッグ開始 ===');
    const parsedData = parseCSV(csvData);
    
    Logger.log(`解析行数: ${parsedData.length}`);
    
    for (let i = 0; i < Math.min(parsedData.length, 5); i++) {
      Logger.log(`行${i + 1}: [${parsedData[i].map((cell, idx) => `[${idx}]"${cell}"`).join(', ')}]`);
      
      if (i > 0) { // ヘッダー行以外をデバッグ
        const processed = processBankMasterDataRow(parsedData[i]);
        Logger.log(`処理後: [${processed.map((cell, idx) => `[${idx}]"${cell}"`).join(', ')}]`);
      }
    }
    
    Logger.log('=== CSV解析デバッグ終了 ===');
  } catch (error) {
    Logger.log('デバッグエラー: ' + error.toString());
  }
}

/**
 * テスト用金融機関マスタCSV取込
 */
function testBankMasterCsvImport() {
  try {
    // テストCSVデータ
    const testCsvData = `銀行コード,銀行名,支店コード,支店名,更新日,状態
0001,ﾐｽﾞﾎ,001,ﾄｳｷｮｳｴｲｷﾞｮｳﾌﾞ,2025-01-16,有効
0005,ﾐﾂﾋﾞｼUFJ,002,ﾏﾙﾉｳﾁ,2025-01-16,有効`;
    
    Logger.log('テスト用CSV取込開始');
    debugCsvData(testCsvData);
    
    const result = importBankMasterFromCsv(testCsvData, false);
    Logger.log('取込結果: ' + result);
    
  } catch (error) {
    Logger.log('テストエラー: ' + error.toString());
  }
}

/**
 * 受取人名の検証テスト
 */
function testRecipientNameValidation() {
  try {
    Logger.log('=== 受取人名検証テスト開始 ===');
    
    const testNames = [
      'ﾔﾏﾀﾞ ﾀﾛｳ',
      'ｻﾄｳ ﾊﾅｺ',
      'ｽｽﾞｷ ｼﾞﾛｳ',
      'ﾀﾅｶ ﾐｻｷ',
      'ｱﾍﾞ ﾊﾙﾄ',
      'ﾔﾏｶﾞﾀ ﾕｳﾀ',
      'ﾌｼﾞﾜﾗ ﾀｸﾐ',
      'ﾏﾂﾓﾄ ﾋﾛｼ'
    ];
    
    const regex = /^[ｱ-ﾝｧ-ｯﾞﾟｰ ]+$/;
    
    testNames.forEach((name, index) => {
      const isValid = regex.test(name);
      Logger.log(`テスト${index + 1}: "${name}" -> ${isValid ? 'OK' : 'NG'}`);
      
      // 文字コードも確認
      const charCodes = [];
      for (let i = 0; i < name.length; i++) {
        charCodes.push(`${name[i]}(${name.charCodeAt(i)})`);
      }
      Logger.log(`  文字コード: ${charCodes.join(', ')}`);
    });
    
    Logger.log('=== 受取人名検証テスト終了 ===');
    return { success: true };
    
  } catch (error) {
    Logger.log('受取人名検証テストエラー: ' + error.toString());
    return { success: false, error: error.message };
  }
}

/**
 * 振込データCSVのテスト取込
 */
function testTransferDataCsvImport() {
  try {
    Logger.log('=== 振込データCSVテスト開始 ===');
    
    // テストCSVデータ（最初の3行のみ）
    const testCsvData = `銀行コード,銀行名,支店コード,支店名,預金種目,口座番号,受取人名,振込金額,顧客コード,識別表示,EDI情報
0001,みずほ銀行,001,東京営業部,1,1234567,ﾔﾏﾀﾞ ﾀﾛｳ,250000,C001,Y,SAL202506
0005,三菱UFJ銀行,002,丸の内支店,1,2345678,ｻﾄｳ ﾊﾅｺ,180000,C002,Y,SAL202506`;
    
    Logger.log('解析開始...');
    const parsedData = parseCSV(testCsvData);
    Logger.log(`解析行数: ${parsedData.length}`);
    
    // ヘッダーを除いたデータ行を検証
    const dataRows = parsedData.slice(1);
    Logger.log('検証開始...');
    const validationResult = validateCsvData(dataRows);
    
    Logger.log(`検証結果: ${validationResult.isValid ? 'OK' : 'NG'}`);
    if (!validationResult.isValid) {
      Logger.log('エラー内容:');
      validationResult.errors.forEach(error => Logger.log(`  - ${error}`));
    }
    
    Logger.log('=== 振込データCSVテスト終了 ===');
    return { success: true, isValid: validationResult.isValid, errors: validationResult.errors };
    
  } catch (error) {
    Logger.log('振込データCSVテストエラー: ' + error.toString());
    return { success: false, error: error.message };
  }
}

/**
 * 事前検証：全行をチェックしてエラー行を特定
 * @param {Array} dataRows - データ行の配列
 * @return {Object} 検証結果 { validRows: Array, errorRows: Array }
 */
function preValidateTransferDataRows(dataRows) {
  const validRows = [];
  const errorRows = [];
  
  for (let i = 0; i < dataRows.length; i++) {
    const rowNum = i + 2; // ヘッダー行を考慮
    const errors = validateCsvRow(dataRows[i], rowNum);
    
    if (errors.length > 0) {
      errorRows.push({
        rowNum,
        data: dataRows[i],
        errors
      });
    } else {
      validRows.push({
        rowNum,
        data: dataRows[i]
      });
    }
  }
  
  return {
    validRows,
    errorRows,
    totalRows: dataRows.length,
    validCount: validRows.length,
    errorCount: errorRows.length
  };
}

/**
 * エラー確認ダイアログの表示
 * @param {Object} validationResult - 事前検証結果
 * @return {boolean} 処理を継続するかどうか
 */
function showErrorConfirmationDialog(validationResult) {
  const errorSummary = validationResult.errorRows.slice(0, 10).map(error => 
    `行${error.rowNum}: ${error.errors[0]}`
  ).join('\n');
  
  const additionalErrors = validationResult.errorCount > 10 ? 
    `\n...他${validationResult.errorCount - 10}件のエラー` : '';
  
  const message = `CSV取り込み時にエラーが検出されました：\n\n` +
    `総件数: ${validationResult.totalRows}件\n` +
    `正常: ${validationResult.validCount}件\n` +
    `エラー: ${validationResult.errorCount}件\n\n` +
    `【主なエラー内容】\n${errorSummary}${additionalErrors}\n\n` +
    `エラー行をスキップして正常なデータのみ取り込みますか？\n` +
    `「OK」: エラー行をスキップして継続\n` +
    `「キャンセル」: 処理を中断`;
  
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'CSV取り込みエラー確認',
    message,
    ui.ButtonSet.OK_CANCEL
  );
  
  return response === ui.Button.OK;
}

/**
 * CSV行の個別検証
 * @param {Array} row - CSV行データ
 * @param {number} rowNum - 行番号
 * @return {Array} エラーメッセージ配列
 */
function validateCsvRow(row, rowNum) {
  const errors = [];
  
  try {
    // 必要な項目数チェック
    if (!row || row.length < 8) {
      errors.push(`項目数が不足しています（必要：8項目、実際：${row ? row.length : 0}項目）`);
      return errors;
    }
    
    const [bankCode, branchCode, accountType, accountNumber, recipientName, amount, customerCode, identification] = row;
    
    // 銀行コードチェック
    if (!bankCode || !/^\d{1,4}$/.test(String(bankCode))) {
      errors.push('銀行コードが正しくありません（1-4桁の数字）');
    }
    
    // 支店コードチェック
    if (!branchCode || !/^\d{1,3}$/.test(String(branchCode))) {
      errors.push('支店コードが正しくありません（1-3桁の数字）');
    }
    
    // 預金種目チェック（重要：貯蓄(4)を検出）
    const accountTypeStr = String(accountType || '').trim();
    if (!accountTypeStr) {
      errors.push('預金種目が入力されていません');
    } else if (accountTypeStr === '4') {
      errors.push('預金種目「4:貯蓄」は全銀協フォーマットでは対象外です（1:普通、2:当座のみ対応）');
    } else if (!['1', '2'].includes(accountTypeStr)) {
      errors.push('預金種目は1(普通)、2(当座)のいずれかを入力してください');
    }
    
    // 口座番号チェック
    if (!accountNumber || !/^\d{1,7}$/.test(String(accountNumber))) {
      errors.push('口座番号が正しくありません（1-7桁の数字）');
    }
    
    // 受取人名チェック
    if (!recipientName || String(recipientName).trim() === '') {
      errors.push('受取人名が入力されていません');
    } else if (!/^[A-Z0-9ｱ-ﾝﾞﾟｧ-ｯ ・ー().\-/]+$/.test(String(recipientName))) {
      errors.push('受取人名は全銀協フォーマット対応文字（半角カナ・英数字・記号、カンマ除く）で入力してください');
    }
    
    // 振込金額チェック
    if (!amount || isNaN(parseFloat(amount)) || parseFloat(amount) <= 0) {
      errors.push('振込金額は正の数値で入力してください');
    }
    
  } catch (error) {
    errors.push(`データ処理エラー: ${error.message}`);
  }
  
  return errors;
} 