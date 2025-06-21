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
    Logger.log(`銀行名自動補完: 行${row}, ${bankCode} -> ${bankName}`);
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
    Logger.log(`支店名自動補完: 行${row}, ${bankCode}-${branchCode} -> ${branchName}`);
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
      
      Logger.log(`マスタデータ正規化: 銀行 ${row[BANK_MASTER_COLUMNS.BANK_CODE - 1]} -> ${normalizedRow[BANK_MASTER_COLUMNS.BANK_CODE - 1]}, 支店 ${row[BANK_MASTER_COLUMNS.BRANCH_CODE - 1]} -> ${normalizedRow[BANK_MASTER_COLUMNS.BRANCH_CODE - 1]}`);
      
      return normalizedRow;
    });
    
    // 有効なデータのみをフィルタリング
    const filteredData = normalizedData.filter(row => {
      return row[BANK_MASTER_COLUMNS.STATUS - 1] === '有効' && 
             row[BANK_MASTER_COLUMNS.BANK_CODE - 1] && 
             row[BANK_MASTER_COLUMNS.BANK_NAME - 1];
    });
    
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
 * 手動一括補完機能
 * @return {Object} 補完結果レポート
 */
function bulkAutoComplete() {
  const startTime = Date.now();
  let bankNameCompletions = 0;
  let branchNameCompletions = 0;
  let failures = 0;
  
  try {
    Logger.log('=== 一括補完デバッグ開始 ===');
    
    // キャッシュをクリアして最新データを取得
    const cache = CacheService.getScriptCache();
    cache.remove(CACHE_CONFIG.BANK_MASTER_KEY);
    Logger.log('キャッシュクリア完了');
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.TRANSFER_DATA);
    if (!sheet) {
      throw new Error('振込データシートが見つかりません');
    }
    Logger.log(`振込データシート名: ${sheet.getName()}`);
    
    const lastRow = sheet.getLastRow();
    Logger.log(`最終行: ${lastRow}`);
    if (lastRow <= 1) {
      return {
        totalRows: 0,
        bankNameCompletions: 0,
        branchNameCompletions: 0,
        failures: 0,
        processingTime: Date.now() - startTime
      };
    }
    
    // マスタデータを事前に取得
    const masterData = getBankMasterData();
    Logger.log(`マスタデータ件数: ${masterData.length}`);
    if (masterData.length === 0) {
      throw new Error('金融機関マスタデータが存在しません');
    }
    
    // マスタデータの最初の数件をログ出力
    for (let i = 0; i < Math.min(3, masterData.length); i++) {
      Logger.log(`マスタデータ[${i}]: ${JSON.stringify(masterData[i])}`);
    }
    
    // 一括処理用のデータ準備
    const dataRange = sheet.getRange(2, 1, lastRow - 1, Object.keys(TRANSFER_DATA_COLUMNS).length);
    const values = dataRange.getValues();
    const updates = [];
    
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const rowNum = i + 2;
      
      // 空行スキップ
      if (isEmptyRow(row)) {
        Logger.log(`行${rowNum}: 空行のためスキップ`);
        continue;
      }
      
      const rawBankCode = row[TRANSFER_DATA_COLUMNS.BANK_CODE - 1];
      const rawBranchCode = row[TRANSFER_DATA_COLUMNS.BRANCH_CODE - 1];
      const currentBankName = String(row[TRANSFER_DATA_COLUMNS.BANK_NAME - 1] || '').trim();
      const currentBranchName = String(row[TRANSFER_DATA_COLUMNS.BRANCH_NAME - 1] || '').trim();
      
      // 数値を適切な桁数に0埋めして文字列に変換
      const bankCode = normalizeCode(rawBankCode, VALIDATION_RULES.BANK_CODE_LENGTH);
      const branchCode = normalizeCode(rawBranchCode, VALIDATION_RULES.BRANCH_CODE_LENGTH);
      
      Logger.log(`行${rowNum}: 生銀行コード="${rawBankCode}", 正規化後="${bankCode}", 生支店コード="${rawBranchCode}", 正規化後="${branchCode}"`);
      Logger.log(`行${rowNum}: 現在の銀行名="${currentBankName}", 現在の支店名="${currentBranchName}"`);
      
      // 銀行コードがある場合の銀行名補完
      if (bankCode && bankCode.length === VALIDATION_RULES.BANK_CODE_LENGTH && !currentBankName) {
        Logger.log(`行${rowNum}: 銀行名補完を試行 - 銀行コード: ${bankCode}`);
        const bankName = findBankNameFromCache(masterData, bankCode);
        if (bankName) {
          Logger.log(`行${rowNum}: 銀行名補完成功 - ${bankCode} -> ${bankName}`);
          updates.push({
            row: rowNum,
            col: TRANSFER_DATA_COLUMNS.BANK_NAME,
            value: bankName
          });
          bankNameCompletions++;
        } else {
          Logger.log(`行${rowNum}: 銀行名補完失敗 - 銀行コード ${bankCode} が見つかりません`);
          failures++;
        }
      } else {
        Logger.log(`行${rowNum}: 銀行名補完条件不一致 - コード長=${bankCode.length}, 現在名="${currentBankName}"`);
      }
      
      // 支店コードがある場合の支店名補完
      if (bankCode && branchCode && 
          bankCode.length === VALIDATION_RULES.BANK_CODE_LENGTH && 
          branchCode.length === VALIDATION_RULES.BRANCH_CODE_LENGTH && 
          !currentBranchName) {
        Logger.log(`行${rowNum}: 支店名補完を試行 - 銀行コード: ${bankCode}, 支店コード: ${branchCode}`);
        const branchName = findBranchNameFromCache(masterData, bankCode, branchCode);
        if (branchName) {
          Logger.log(`行${rowNum}: 支店名補完成功 - ${bankCode}-${branchCode} -> ${branchName}`);
          updates.push({
            row: rowNum,
            col: TRANSFER_DATA_COLUMNS.BRANCH_NAME,
            value: branchName
          });
          branchNameCompletions++;
        } else {
          Logger.log(`行${rowNum}: 支店名補完失敗 - ${bankCode}-${branchCode} が見つかりません`);
          failures++;
        }
      } else {
        Logger.log(`行${rowNum}: 支店名補完条件不一致 - 銀行コード長=${bankCode.length}, 支店コード長=${branchCode.length}, 現在名="${currentBranchName}"`);
      }
    }
    
    // 一括更新実行
    if (updates.length > 0) {
      updates.forEach(update => {
        sheet.getRange(update.row, update.col).setValue(update.value);
      });
    }
    
    const processingTime = Date.now() - startTime;
    Logger.log(`一括補完完了: 銀行名${bankNameCompletions}件, 支店名${branchNameCompletions}件, 失敗${failures}件, 処理時間${processingTime}ms`);
    
    return {
      totalRows: lastRow - 1,
      bankNameCompletions,
      branchNameCompletions,
      failures,
      processingTime
    };
  } catch (error) {
    Logger.log('一括補完エラー: ' + error.toString());
    throw error;
  }
}

/**
 * キャッシュデータから銀行名を検索
 * @param {Array[]} masterData - マスタデータ
 * @param {string} bankCode - 銀行コード
 * @return {string|null} 銀行名またはnull
 */
function findBankNameFromCache(masterData, bankCode) {
  Logger.log(`銀行名検索: ${bankCode} をマスタデータから検索中...`);
  
  const bankRow = masterData.find(row => {
    const rowBankCode = String(row[BANK_MASTER_COLUMNS.BANK_CODE - 1] || '').trim();
    const rowStatus = String(row[BANK_MASTER_COLUMNS.STATUS - 1] || '').trim();
    
    Logger.log(`  比較中: マスタ銀行コード="${rowBankCode}", 状態="${rowStatus}"`);
    
    return rowBankCode == bankCode && rowStatus === '有効';
  });
  
  const result = bankRow ? bankRow[BANK_MASTER_COLUMNS.BANK_NAME - 1] : null;
  Logger.log(`銀行名検索結果: ${bankCode} -> ${result}`);
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
  Logger.log(`支店名検索: ${bankCode}-${branchCode} をマスタデータから検索中...`);
  
  const branchRow = masterData.find(row => {
    const rowBankCode = String(row[BANK_MASTER_COLUMNS.BANK_CODE - 1] || '').trim();
    const rowBranchCode = String(row[BANK_MASTER_COLUMNS.BRANCH_CODE - 1] || '').trim();
    const rowStatus = String(row[BANK_MASTER_COLUMNS.STATUS - 1] || '').trim();
    
    Logger.log(`  比較中: マスタ銀行コード="${rowBankCode}", マスタ支店コード="${rowBranchCode}", 状態="${rowStatus}"`);
    
    return rowBankCode == bankCode && rowBranchCode == branchCode && rowStatus === '有効';
  });
  
  const result = branchRow ? branchRow[BANK_MASTER_COLUMNS.BRANCH_NAME - 1] : null;
  Logger.log(`支店名検索結果: ${bankCode}-${branchCode} -> ${result}`);
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