/**
 * 自動補完機能
 * 銀行コード・支店コード入力時の自動名称補完
 */

/**
 * セル編集時のイベントハンドラ（リアルタイム自動補完）
 * @param {Object} e - 編集イベント
 */
function onEdit(e) {
  try {
    const sheet = e.source.getActiveSheet();
    const sheetName = sheet.getName();
    
    // 振込データシートでの自動補完
    if (sheetName === SHEET_NAMES.TRANSFER_DATA) {
      handleTransferDataAutoComplete(e);
    }
    // 振込依頼人情報シートでの自動補完
    else if (sheetName === SHEET_NAMES.CLIENT_INFO) {
      handleClientInfoAutoComplete(e);
    }
  } catch (error) {
    Logger.log('onEdit エラー: ' + error.toString());
    // エラーが発生してもユーザーの編集を妨げないようにする
  }
}

/**
 * 振込データシートでの自動補完処理
 * @param {Object} e - 編集イベント
 */
function handleTransferDataAutoComplete(e) {
  const range = e.range;
  const row = range.getRow();
  const col = range.getColumn();
  
  // ヘッダー行はスキップ
  if (row <= 1) return;
  
  const sheet = e.source.getActiveSheet();
  
  // 銀行コード入力時の自動補完
  if (col === TRANSFER_DATA_COLUMNS.BANK_CODE) {
    const bankCode = normalizeCode(range.getValue(), VALIDATION_RULES.BANK_CODE_LENGTH);
    if (bankCode && bankCode.length === VALIDATION_RULES.BANK_CODE_LENGTH) {
      autoCompleteBankName(sheet, row, bankCode);
    }
  }
  
  // 支店コード入力時の自動補完
  if (col === TRANSFER_DATA_COLUMNS.BRANCH_CODE) {
    const bankCode = normalizeCode(sheet.getRange(row, TRANSFER_DATA_COLUMNS.BANK_CODE).getValue(), VALIDATION_RULES.BANK_CODE_LENGTH);
    const branchCode = normalizeCode(range.getValue(), VALIDATION_RULES.BRANCH_CODE_LENGTH);
    
    if (bankCode && branchCode && 
        bankCode.length === VALIDATION_RULES.BANK_CODE_LENGTH && 
        branchCode.length === VALIDATION_RULES.BRANCH_CODE_LENGTH) {
      autoCompleteBranchName(sheet, row, bankCode, branchCode);
    }
  }
}

/**
 * 振込依頼人情報シートでの自動補完処理
 * @param {Object} e - 編集イベント
 */
function handleClientInfoAutoComplete(e) {
  const range = e.range;
  const cell = range.getA1Notation();
  const sheet = e.source.getActiveSheet();
  
  // 銀行コード入力時の自動補完
  if (cell === CLIENT_INFO_CELLS.BANK_CODE) {
    const bankCode = normalizeCode(range.getValue(), VALIDATION_RULES.BANK_CODE_LENGTH);
    if (bankCode && bankCode.length === VALIDATION_RULES.BANK_CODE_LENGTH) {
      const bankName = findBankName(bankCode);
      if (bankName) {
        sheet.getRange(CLIENT_INFO_CELLS.BANK_NAME).setValue(bankName);
      }
    }
  }
  
  // 支店コード入力時の自動補完
  if (cell === CLIENT_INFO_CELLS.BRANCH_CODE) {
    const bankCode = normalizeCode(sheet.getRange(CLIENT_INFO_CELLS.BANK_CODE).getValue(), VALIDATION_RULES.BANK_CODE_LENGTH);
    const branchCode = normalizeCode(range.getValue(), VALIDATION_RULES.BRANCH_CODE_LENGTH);
    
    if (bankCode && branchCode && 
        bankCode.length === VALIDATION_RULES.BANK_CODE_LENGTH && 
        branchCode.length === VALIDATION_RULES.BRANCH_CODE_LENGTH) {
      const branchName = findBranchName(bankCode, branchCode);
      if (branchName) {
        sheet.getRange(CLIENT_INFO_CELLS.BRANCH_NAME).setValue(branchName);
      }
    }
  }
}

/**
 * 銀行名の自動補完
 * @param {Sheet} sheet - 対象シート
 * @param {number} row - 行番号
 * @param {string} bankCode - 銀行コード
 */
function autoCompleteBankName(sheet, row, bankCode) {
  const bankName = findBankName(bankCode);
  if (bankName) {
    sheet.getRange(row, TRANSFER_DATA_COLUMNS.BANK_NAME).setValue(bankName);
    Logger.log(`銀行名自動補完: 行${row}, ${maskBankCode(bankCode)} -> ${bankName}`);
  }
}

/**
 * 支店名の自動補完
 * @param {Sheet} sheet - 対象シート
 * @param {number} row - 行番号
 * @param {string} bankCode - 銀行コード
 * @param {string} branchCode - 支店コード
 */
function autoCompleteBranchName(sheet, row, bankCode, branchCode) {
  const branchName = findBranchName(bankCode, branchCode);
  if (branchName) {
    sheet.getRange(row, TRANSFER_DATA_COLUMNS.BRANCH_NAME).setValue(branchName);
    Logger.log(`支店名自動補完: 行${row}, ${maskBankCode(bankCode)}-${maskBankCode(branchCode)} -> ${branchName}`);
  }
}

/**
 * 銀行名を検索
 * @param {string} bankCode - 銀行コード
 * @return {string|null} 銀行名またはnull
 */
function findBankName(bankCode) {
  try {
    const masterData = getBankMasterData();
    const bankRow = masterData.find(row => 
      row[BANK_MASTER_COLUMNS.BANK_CODE - 1] == bankCode && 
      row[BANK_MASTER_COLUMNS.STATUS - 1] === '有効'
    );
    
    return bankRow ? bankRow[BANK_MASTER_COLUMNS.BANK_NAME - 1] : null;
  } catch (error) {
    Logger.log('銀行名検索エラー: ' + error.toString());
    return null;
  }
}

/**
 * 支店名を検索
 * @param {string} bankCode - 銀行コード
 * @param {string} branchCode - 支店コード
 * @return {string|null} 支店名またはnull
 */
function findBranchName(bankCode, branchCode) {
  try {
    const masterData = getBankMasterData();
    const branchRow = masterData.find(row => 
      row[BANK_MASTER_COLUMNS.BANK_CODE - 1] == bankCode && 
      row[BANK_MASTER_COLUMNS.BRANCH_CODE - 1] == branchCode && 
      row[BANK_MASTER_COLUMNS.STATUS - 1] === '有効'
    );
    
    return branchRow ? branchRow[BANK_MASTER_COLUMNS.BRANCH_NAME - 1] : null;
  } catch (error) {
    Logger.log('支店名検索エラー: ' + error.toString());
    return null;
  }
}

/**
 * 金融機関マスタデータの取得（キャッシュ機能付き）
 * @return {Array[]} 金融機関マスタデータの2次元配列
 */
function getBankMasterData() {
  try {
    // キャッシュチェック
    const cache = CacheService.getScriptCache();
    const cacheKey = CACHE_CONFIG.BANK_MASTER_KEY + '_v2';
    const cached = cache.get(cacheKey);
    
    if (cached) {
      logSystemActivity('getBankMasterData', 'キャッシュからデータ取得', 'INFO');
      Logger.log('金融機関マスタデータ: キャッシュヒット');
      return JSON.parse(cached);
    }
    
    logSystemActivity('getBankMasterData', 'キャッシュミス - スプレッドシートから取得', 'INFO');
    Logger.log('金融機関マスタデータ: キャッシュミス');
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEET_NAMES.BANK_MASTER);
    
    if (!sheet) {
      logSystemActivity('getBankMasterData', '金融機関マスタシートが見つかりません', 'ERROR');
      Logger.log('金融機関マスタシートが見つかりません');
      return [];
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      logSystemActivity('getBankMasterData', '金融機関マスタデータがありません', 'WARNING');
      Logger.log('金融機関マスタデータがありません');
      return [];
    }
    
    // ヘッダー行を除いてデータを取得
    const data = sheet.getRange(2, 1, lastRow - 1, Object.keys(BANK_MASTER_COLUMNS).length).getValues();
    
    // データの正規化処理
    const normalizedData = data.map(row => {
      const normalizedRow = [...row]; // 配列をコピー
      
      // 銀行コードの正規化
      if (normalizedRow[BANK_MASTER_COLUMNS.BANK_CODE - 1] !== null && normalizedRow[BANK_MASTER_COLUMNS.BANK_CODE - 1] !== '') {
        normalizedRow[BANK_MASTER_COLUMNS.BANK_CODE - 1] = String(normalizedRow[BANK_MASTER_COLUMNS.BANK_CODE - 1]).padStart(VALIDATION_RULES.BANK_CODE_LENGTH, '0');
      }
      
      // 支店コードの正規化
      if (normalizedRow[BANK_MASTER_COLUMNS.BRANCH_CODE - 1] !== null && normalizedRow[BANK_MASTER_COLUMNS.BRANCH_CODE - 1] !== '') {
        normalizedRow[BANK_MASTER_COLUMNS.BRANCH_CODE - 1] = String(normalizedRow[BANK_MASTER_COLUMNS.BRANCH_CODE - 1]).padStart(VALIDATION_RULES.BRANCH_CODE_LENGTH, '0');
      }
      
      // 詳細なログは削減（性能改善）
      // Logger.log(`マスタデータ正規化: 銀行 ${maskBankCode(row[BANK_MASTER_COLUMNS.BANK_CODE - 1])} -> ${maskBankCode(normalizedRow[BANK_MASTER_COLUMNS.BANK_CODE - 1])}, 支店 ${maskBankCode(row[BANK_MASTER_COLUMNS.BRANCH_CODE - 1])} -> ${maskBankCode(normalizedRow[BANK_MASTER_COLUMNS.BRANCH_CODE - 1])}`);
      
      return normalizedRow;
    });
    
    // 有効なデータのみをフィルタリング
    const filteredData = normalizedData.filter(row => {
      return row[BANK_MASTER_COLUMNS.STATUS - 1] === '有効' && 
             row[BANK_MASTER_COLUMNS.BANK_CODE - 1] && 
             row[BANK_MASTER_COLUMNS.BANK_NAME - 1];
    });
    
    // キャッシュに保存（5分間）
    try {
      cache.put(cacheKey, JSON.stringify(filteredData), 300); // 300秒 = 5分
      logSystemActivity('getBankMasterData', 'データをキャッシュに保存', 'INFO');
    } catch (cacheError) {
      // キャッシュエラーは無視（データは正常に取得できている）
      Logger.log('キャッシュ保存エラー（無視）: ' + cacheError.toString());
    }
    
    logSystemActivity('getBankMasterData', `金融機関マスタデータ取得完了: ${filteredData.length}件`, 'INFO');
    Logger.log(`金融機関マスタデータ取得: ${filteredData.length}件`);
    
    return filteredData;
  } catch (error) {
    logSystemActivity('getBankMasterData', `エラー: ${error.message}`, 'ERROR');
    Logger.log('金融機関マスタデータ取得エラー: ' + error.toString());
    return [];
  }
}

/**
 * 手動一括補完機能（メイン処理）
 * @return {Object} 補完結果レポート
 */
function bulkAutoComplete() {
  const startTime = Date.now();
  
  try {
    logInfo('一括補完開始');
    
    // 1. 初期化とデータ準備
    const initResult = initializeBulkAutoComplete();
    if (!initResult.success) {
      return initResult.result;
    }
    
    const { sheet, masterData, values, dataRange } = initResult.data;
    
    // 2. 一括補完処理実行
    const completionResult = executeBulkCompletion(values, masterData);
    
    // 3. 結果をスプレッドシートに反映
    const updateResult = updateSheetWithCompletions(dataRange, values, completionResult.hasUpdates);
    
    // 4. 結果レポート生成
    const processingTime = Date.now() - startTime;
    const report = generateCompletionReport(values.length, completionResult, processingTime);
    
    logInfo(`一括補完完了: 銀行名${report.bankNameCompletions}件, 支店名${report.branchNameCompletions}件, 失敗${report.failures}件, 処理時間${processingTime}ms`);
    
    return report;
    
  } catch (error) {
    logError('一括補完エラー: ' + error.toString());
    throw error;
  }
}

/**
 * 一括補完の初期化処理
 * @return {Object} 初期化結果
 */
function initializeBulkAutoComplete() {
  try {
    // キャッシュをクリアして最新データを取得
    const cache = CacheService.getScriptCache();
    cache.remove(CACHE_CONFIG.BANK_MASTER_KEY);
    logDebug('キャッシュクリア完了');
    
    // シート取得
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.TRANSFER_DATA);
    if (!sheet) {
      throw new Error('振込データシートが見つかりません');
    }
    logDebug(`振込データシート名: ${sheet.getName()}`);
    
    // データ範囲チェック
    const lastRow = sheet.getLastRow();
    logDebug(`最終行: ${lastRow}`);
    if (lastRow <= 1) {
      return {
        success: true,
        result: {
          totalRows: 0,
          bankNameCompletions: 0,
          branchNameCompletions: 0,
          failures: 0,
          processingTime: 0
        }
      };
    }
    
    // マスタデータ取得
    const masterData = getBankMasterData();
    logInfo(`マスタデータ件数: ${masterData.length}`);
    if (masterData.length === 0) {
      throw new Error('金融機関マスタデータが存在しません');
    }
    
    // データ範囲取得
    const dataRange = sheet.getRange(2, 1, lastRow - 1, Object.keys(TRANSFER_DATA_COLUMNS).length);
    const values = dataRange.getValues();
    
    return {
      success: true,
      data: { sheet, masterData, values, dataRange }
    };
    
  } catch (error) {
    return {
      success: false,
      result: {
        totalRows: 0,
        bankNameCompletions: 0,
        branchNameCompletions: 0,
        failures: 0,
        processingTime: Date.now() - startTime,
        error: error.message
      }
    };
  }
}

/**
 * 一括補完処理の実行
 * @param {Array[]} values - スプレッドシートのデータ
 * @param {Array[]} masterData - 金融機関マスタデータ
 * @return {Object} 補完結果
 */
function executeBulkCompletion(values, masterData) {
  let bankNameCompletions = 0;
  let branchNameCompletions = 0;
  let failures = 0;
  let hasUpdates = false;
  
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const rowNum = i + 2;
    
    // 空行スキップ
    if (isEmptyRow(row)) {
      logDebug(`行${rowNum}: 空行のためスキップ`);
      continue;
    }
    
    // 行ごとの補完処理
    const rowResult = processRowCompletion(row, rowNum, masterData, i, values);
    
    bankNameCompletions += rowResult.bankNameCompletion ? 1 : 0;
    branchNameCompletions += rowResult.branchNameCompletion ? 1 : 0;
    failures += rowResult.failures;
    
    if (rowResult.hasUpdate) {
      hasUpdates = true;
    }
  }
  
  return {
    bankNameCompletions,
    branchNameCompletions,
    failures,
    hasUpdates
  };
}

/**
 * 行ごとの補完処理
 * @param {Array} row - 行データ
 * @param {number} rowNum - 行番号
 * @param {Array[]} masterData - 金融機関マスタデータ
 * @param {number} index - 配列インデックス
 * @param {Array[]} values - 全データ配列（更新用）
 * @return {Object} 行の補完結果
 */
function processRowCompletion(row, rowNum, masterData, index, values) {
  const rawBankCode = row[TRANSFER_DATA_COLUMNS.BANK_CODE - 1];
  const rawBranchCode = row[TRANSFER_DATA_COLUMNS.BRANCH_CODE - 1];
  const currentBankName = String(row[TRANSFER_DATA_COLUMNS.BANK_NAME - 1] || '').trim();
  const currentBranchName = String(row[TRANSFER_DATA_COLUMNS.BRANCH_NAME - 1] || '').trim();
  
  // 数値を適切な桁数に0埋めして文字列に変換
  const bankCode = normalizeCode(rawBankCode, VALIDATION_RULES.BANK_CODE_LENGTH);
  const branchCode = normalizeCode(rawBranchCode, VALIDATION_RULES.BRANCH_CODE_LENGTH);
  
  logDebug(`行${rowNum}: 銀行=${bankCode}, 支店=${branchCode}`);
  
  let bankNameCompletion = false;
  let branchNameCompletion = false;
  let failures = 0;
  let hasUpdate = false;
  
  // 銀行名補完
  if (bankCode && bankCode.length === VALIDATION_RULES.BANK_CODE_LENGTH && !currentBankName) {
    const bankName = findBankNameFromCache(masterData, bankCode);
    if (bankName) {
      logDebug(`行${rowNum}: 銀行名補完成功 - ${bankCode} -> ${bankName}`);
      values[index][TRANSFER_DATA_COLUMNS.BANK_NAME - 1] = bankName;
      bankNameCompletion = true;
      hasUpdate = true;
    } else {
      logDebug(`行${rowNum}: 銀行名補完失敗 - 銀行コード ${bankCode} が見つかりません`);
      failures++;
    }
  }
  
  // 支店名補完
  if (bankCode && branchCode && 
      bankCode.length === VALIDATION_RULES.BANK_CODE_LENGTH && 
      branchCode.length === VALIDATION_RULES.BRANCH_CODE_LENGTH && 
      !currentBranchName) {
    const branchName = findBranchNameFromCache(masterData, bankCode, branchCode);
    if (branchName) {
      logDebug(`行${rowNum}: 支店名補完成功 - ${bankCode}-${branchCode} -> ${branchName}`);
      values[index][TRANSFER_DATA_COLUMNS.BRANCH_NAME - 1] = branchName;
      branchNameCompletion = true;
      hasUpdate = true;
    } else {
      logDebug(`行${rowNum}: 支店名補完失敗 - ${bankCode}-${branchCode} が見つかりません`);
      failures++;
    }
  }
  
  return {
    bankNameCompletion,
    branchNameCompletion,
    failures,
    hasUpdate
  };
}

/**
 * スプレッドシートへの更新反映
 * @param {Range} dataRange - データ範囲
 * @param {Array[]} values - 更新データ
 * @param {boolean} hasUpdates - 更新があるかどうか
 * @return {boolean} 更新成功フラグ
 */
function updateSheetWithCompletions(dataRange, values, hasUpdates) {
  if (hasUpdates) {
    dataRange.setValues(values);
    logInfo('一括更新完了（バッチ処理）');
    return true;
  }
  return false;
}

/**
 * 補完結果レポートの生成
 * @param {number} totalRows - 総行数
 * @param {Object} completionResult - 補完結果
 * @param {number} processingTime - 処理時間
 * @return {Object} レポート
 */
function generateCompletionReport(totalRows, completionResult, processingTime) {
  return {
    totalRows,
    bankNameCompletions: completionResult.bankNameCompletions,
    branchNameCompletions: completionResult.branchNameCompletions,
    failures: completionResult.failures,
    processingTime
  };
}

/**
 * キャッシュデータから銀行名を検索
 * @param {Array[]} masterData - マスタデータ
 * @param {string} bankCode - 銀行コード
 * @return {string|null} 銀行名またはnull
 */
function findBankNameFromCache(masterData, bankCode) {
  logDebug(`銀行名検索: ${bankCode}`);
  
  const bankRow = masterData.find(row => {
    const rowBankCode = String(row[BANK_MASTER_COLUMNS.BANK_CODE - 1] || '').trim();
    const rowStatus = String(row[BANK_MASTER_COLUMNS.STATUS - 1] || '').trim();
    
    return rowBankCode == bankCode && rowStatus === '有効';
  });
  
  const result = bankRow ? bankRow[BANK_MASTER_COLUMNS.BANK_NAME - 1] : null;
  if (result) {
    logDebug(`銀行名検索結果: ${bankCode} -> ${result}`);
  }
  return result;
}

/**
 * キャッシュデータから支店名を検索
 * @param {Array[]} masterData - マスタデータ
 * @param {string} bankCode - 銀行コード
 * @param {string} branchCode - 支店コード
 * @return {string|null} 支店名またはnull
 */
function findBranchNameFromCache(masterData, bankCode, branchCode) {
  logDebug(`支店名検索: ${bankCode}-${branchCode}`);
  
  const branchRow = masterData.find(row => {
    const rowBankCode = String(row[BANK_MASTER_COLUMNS.BANK_CODE - 1] || '').trim();
    const rowBranchCode = String(row[BANK_MASTER_COLUMNS.BRANCH_CODE - 1] || '').trim();
    const rowStatus = String(row[BANK_MASTER_COLUMNS.STATUS - 1] || '').trim();
    
    return rowBankCode == bankCode && rowBranchCode == branchCode && rowStatus === '有効';
  });
  
  const result = branchRow ? branchRow[BANK_MASTER_COLUMNS.BRANCH_NAME - 1] : null;
  if (result) {
    logDebug(`支店名検索結果: ${bankCode}-${branchCode} -> ${result}`);
  }
  return result;
}

/**
 * マスタデータ整備機能
 * @return {Object} 整備結果レポート
 */
function cleanupMasterData() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.BANK_MASTER);
    if (!sheet) {
      throw new Error('金融機関マスタシートが見つかりません');
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return { duplicatesRemoved: 0, dataFixed: 0, invalidData: 0 };
    }
    
    const dataRange = sheet.getRange(2, 1, lastRow - 1, Object.keys(BANK_MASTER_COLUMNS).length);
    const values = dataRange.getValues();
    
    let duplicatesRemoved = 0;
    let dataFixed = 0;
    let invalidData = 0;
    const invalidDataDetails = [];
    
    const seenKeys = new Set();
    const validRows = [];
    
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const rowNum = i + 2; // 実際の行番号（ヘッダー行を除く）
      
      // 空行スキップ
      if (isEmptyRow(row)) continue;
      
      const bankCode = String(row[BANK_MASTER_COLUMNS.BANK_CODE - 1] || '').trim();
      const branchCode = String(row[BANK_MASTER_COLUMNS.BRANCH_CODE - 1] || '').trim();
      let bankName = String(row[BANK_MASTER_COLUMNS.BANK_NAME - 1] || '').trim();
      let branchName = String(row[BANK_MASTER_COLUMNS.BRANCH_NAME - 1] || '').trim();
      let status = String(row[BANK_MASTER_COLUMNS.STATUS - 1] || '').trim();
      
      // データ検証（詳細な理由を記録）
      const invalidReasons = [];
      
      if (!bankCode) {
        invalidReasons.push('銀行コードが空白');
      } else if (bankCode.length > 4) {
        invalidReasons.push(`銀行コードが桁数オーバー(${bankCode.length}桁)`);
      } else if (!/^\d+$/.test(bankCode)) {
        invalidReasons.push('銀行コードに数字以外の文字');
      }
      
      if (!branchCode) {
        invalidReasons.push('支店コードが空白');
      } else if (branchCode.length > 3) {
        invalidReasons.push(`支店コードが桁数オーバー(${branchCode.length}桁)`);
      } else if (!/^\d+$/.test(branchCode)) {
        invalidReasons.push('支店コードに数字以外の文字');
      }
      
      if (invalidReasons.length > 0) {
        invalidData++;
        invalidDataDetails.push({
          row: rowNum,
          reasons: invalidReasons
        });
        
        // 無効データも保持するが、重複チェックはスキップ
        let updateDate = row[BANK_MASTER_COLUMNS.UPDATE_DATE - 1];
        if (!updateDate || !(updateDate instanceof Date)) {
          updateDate = new Date();
        }
        if (!status || (status !== '有効' && status !== '無効')) {
          status = '有効';
        }
        validRows.push([
          bankCode,
          bankName,
          branchCode,
          branchName,
          updateDate,
          status
        ]);
        continue;
      }
      
      // 重複チェック
      const key = `${bankCode}-${branchCode}`;
      if (seenKeys.has(key)) {
        duplicatesRemoved++;
        continue;
      }
      seenKeys.add(key);
      
      // データ修正
      let dataModified = false;
      
      // 状態の正規化
      if (!status || (status !== '有効' && status !== '無効')) {
        status = '有効';
        dataModified = true;
      }
      
      // 更新日設定
      let updateDate = row[BANK_MASTER_COLUMNS.UPDATE_DATE - 1];
      if (!updateDate || !(updateDate instanceof Date)) {
        updateDate = new Date();
        dataModified = true;
      }
      
      if (dataModified) {
        dataFixed++;
      }
      
      // 修正されたデータを保存
      validRows.push([
        bankCode,
        bankName,
        branchCode,
        branchName,
        updateDate,
        status
      ]);
    }
    
    // データをクリアして再設定
    if (validRows.length > 0) {
      // データ部分をクリア
      if (lastRow > 1) {
        sheet.getRange(2, 1, lastRow - 1, Object.keys(BANK_MASTER_COLUMNS).length).clear();
      }
      
      // 有効なデータを再設定
      sheet.getRange(2, 1, validRows.length, Object.keys(BANK_MASTER_COLUMNS).length).setValues(validRows);
    }
    
    // キャッシュをクリア
    CacheService.getScriptCache().remove(CACHE_CONFIG.BANK_MASTER_KEY);
    
    Logger.log(`マスタデータ整備完了: 重複削除${duplicatesRemoved}件, データ修正${dataFixed}件, 無効データ${invalidData}件`);
    if (invalidDataDetails.length > 0) {
      invalidDataDetails.forEach(detail => {
        Logger.log(`無効データ行${detail.row}: ${detail.reasons.join(', ')}`);
      });
    }
    
    return { duplicatesRemoved, dataFixed, invalidData, invalidDataDetails };
  } catch (error) {
    Logger.log('マスタデータ整備エラー: ' + error.toString());
    throw error;
  }
}

/**
 * キャッシュの強制更新
 */
function refreshMasterDataCache() {
  try {
    CacheService.getScriptCache().remove(CACHE_CONFIG.BANK_MASTER_KEY);
    getBankMasterData(); // 新しいデータでキャッシュを再構築
    Logger.log('マスタデータキャッシュを更新しました');
  } catch (error) {
    Logger.log('キャッシュ更新エラー: ' + error.toString());
    throw error;
  }
}

 