/**
 * データ検証機能
 * 振込依頼人情報と振込データの検証処理
 */

/**
 * 振込依頼人情報の検証
 * @return {Object} 検証結果 { isValid: boolean, errors: string[] }
 */
function validateClientInfo() {
  const errors = [];
  
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.CLIENT_INFO);
    if (!sheet) {
      errors.push('振込依頼人情報シートが見つかりません。');
      return { isValid: false, errors };
    }
    
    // 各項目の検証
    const validations = [
      { 
        cell: CLIENT_INFO_CELLS.CLIENT_CODE, 
        label: CLIENT_INFO_LABELS.CLIENT_CODE, 
        maxLength: VALIDATION_RULES.CLIENT_CODE_MAX_LENGTH,
        required: true,
        type: 'alphanumeric'
      },
      { 
        cell: CLIENT_INFO_CELLS.CLIENT_NAME, 
        label: CLIENT_INFO_LABELS.CLIENT_NAME, 
        maxLength: VALIDATION_RULES.CLIENT_NAME_MAX_LENGTH,
        required: true,
        type: 'zenginFormat'
      },
      { 
        cell: CLIENT_INFO_CELLS.BANK_CODE, 
        label: CLIENT_INFO_LABELS.BANK_CODE, 
        exactLength: VALIDATION_RULES.BANK_CODE_LENGTH,
        required: true,
        type: 'number'
      },
      { 
        cell: CLIENT_INFO_CELLS.BANK_NAME, 
        label: CLIENT_INFO_LABELS.BANK_NAME, 
        maxLength: VALIDATION_RULES.BANK_NAME_MAX_LENGTH,
        required: true,
        type: 'zenginFormat'
      },
      { 
        cell: CLIENT_INFO_CELLS.BRANCH_CODE, 
        label: CLIENT_INFO_LABELS.BRANCH_CODE, 
        exactLength: VALIDATION_RULES.BRANCH_CODE_LENGTH,
        required: true,
        type: 'number'
      },
      { 
        cell: CLIENT_INFO_CELLS.BRANCH_NAME, 
        label: CLIENT_INFO_LABELS.BRANCH_NAME, 
        maxLength: VALIDATION_RULES.BRANCH_NAME_MAX_LENGTH,
        required: true,
        type: 'zenginFormat'
      },
      { 
        cell: CLIENT_INFO_CELLS.ACCOUNT_TYPE, 
        label: CLIENT_INFO_LABELS.ACCOUNT_TYPE, 
        required: true,
        type: 'select',
        validValues: Object.values(ACCOUNT_TYPES).map(type => type.code)
      },
      { 
        cell: CLIENT_INFO_CELLS.ACCOUNT_NUMBER, 
        label: CLIENT_INFO_LABELS.ACCOUNT_NUMBER, 
        maxLength: VALIDATION_RULES.ACCOUNT_NUMBER_MAX_LENGTH,
        required: true,
        type: 'number'
      },
      { 
        cell: CLIENT_INFO_CELLS.CATEGORY_CODE, 
        label: CLIENT_INFO_LABELS.CATEGORY_CODE, 
        required: true,
        type: 'select',
        validValues: Object.values(CATEGORY_CODES).map(code => code.code)
      },
      { 
        cell: CLIENT_INFO_CELLS.FILE_EXTENSION, 
        label: CLIENT_INFO_LABELS.FILE_EXTENSION, 
        required: true,
        type: 'select',
        validValues: FILE_EXTENSIONS
      }
    ];
    
    for (const validation of validations) {
      const value = sheet.getRange(validation.cell).getValue();
      const fieldErrors = validateField(value, validation);
      if (fieldErrors.length > 0) {
        errors.push(...fieldErrors);
      }
    }
    
    // 金融機関コードの存在チェック（オプション）
    // 注意: 金融機関マスタが不完全な場合はこのチェックをコメントアウト
    /*
    const bankCode = sheet.getRange(CLIENT_INFO_CELLS.BANK_CODE).getValue();
    const branchCode = sheet.getRange(CLIENT_INFO_CELLS.BRANCH_CODE).getValue();
    if (bankCode && branchCode) {
      const bankMasterValid = validateBankMasterExists(bankCode, branchCode);
      if (!bankMasterValid.isValid) {
        errors.push(...bankMasterValid.errors);
      }
    }
    */
    
  } catch (error) {
    Logger.log('振込依頼人情報検証エラー: ' + error.toString());
    errors.push('振込依頼人情報の検証中にエラーが発生しました: ' + error.message);
  }
  
  return { isValid: errors.length === 0, errors };
}

/**
 * 振込データの検証
 * @return {Object} 検証結果 { isValid: boolean, errors: string[] }
 */
function validateTransferData() {
  const errors = [];
  
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.TRANSFER_DATA);
    if (!sheet) {
      errors.push('振込データシートが見つかりません。');
      return { isValid: false, errors };
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      errors.push('振込データが入力されていません。');
      return { isValid: false, errors };
    }
    
    // データ件数チェック
    const dataRowCount = lastRow - 1;
    if (dataRowCount > VALIDATION_RULES.MAX_RECORDS) {
      errors.push(`処理可能件数(${VALIDATION_RULES.MAX_RECORDS}件)を超えています。現在: ${dataRowCount}件`);
    }
    
    // 全データ行の検証
    const dataRange = sheet.getRange(2, 1, dataRowCount, Object.keys(TRANSFER_DATA_COLUMNS).length);
    const values = dataRange.getValues();
    
    const duplicateCheckMap = new Map();
    
    for (let i = 0; i < values.length; i++) {
      const rowNum = i + 2;
      const row = values[i];
      
      // 空行スキップ
      if (isEmptyRow(row)) continue;
      
      const rowErrors = validateTransferDataRow(row, rowNum);
      if (rowErrors.length > 0) {
        errors.push(...rowErrors);
      }
      
      // 重複チェック用キー作成
      const bankCode = row[TRANSFER_DATA_COLUMNS.BANK_CODE - 1];
      const branchCode = row[TRANSFER_DATA_COLUMNS.BRANCH_CODE - 1];
      const accountNumber = row[TRANSFER_DATA_COLUMNS.ACCOUNT_NUMBER - 1];
      
      if (bankCode && branchCode && accountNumber) {
        const duplicateKey = `${bankCode}-${branchCode}-${accountNumber}`;
        if (duplicateCheckMap.has(duplicateKey)) {
          errors.push(`行${rowNum}: 重複する口座が存在します (初回出現: 行${duplicateCheckMap.get(duplicateKey)})`);
        } else {
          duplicateCheckMap.set(duplicateKey, rowNum);
        }
      }
    }
    
  } catch (error) {
    Logger.log('振込データ検証エラー: ' + error.toString());
    errors.push('振込データの検証中にエラーが発生しました: ' + error.message);
  }
  
  return { isValid: errors.length === 0, errors };
}

/**
 * 個別フィールドの検証
 * @param {any} value - 検証対象の値
 * @param {Object} validation - 検証ルール
 * @return {string[]} エラーメッセージの配列
 */
function validateField(value, validation) {
  const errors = [];
  let stringValue = String(value || '').trim();
  
  // 数値項目の０埋め正規化処理
  if (validation.type === 'number' && validation.exactLength && stringValue) {
    // 数値の場合、指定桁数まで左０埋めを行う
    if (/^\d+$/.test(stringValue)) {
      stringValue = stringValue.padStart(validation.exactLength, '0');
    }
  }
  
  // 必須チェック
  if (validation.required && !stringValue) {
    errors.push(`${validation.label}: 必須項目が入力されていません。`);
    return errors;
  }
  
  if (!stringValue) return errors; // 空の場合はこれ以上チェックしない
  
  // 文字数チェック
  if (validation.maxLength && stringValue.length > validation.maxLength) {
    errors.push(`${validation.label}: 文字数制限(${validation.maxLength}文字)を超えています。現在: ${stringValue.length}文字`);
  }
  
  if (validation.exactLength && stringValue.length !== validation.exactLength) {
    errors.push(`${validation.label}: ${validation.exactLength}桁で入力してください。現在: ${stringValue.length}桁`);
  }
  
  // データ型チェック
  switch (validation.type) {
    case 'number':
      if (!/^\d+$/.test(stringValue)) {
        errors.push(`${validation.label}: 数字のみ入力してください。`);
      }
      break;
    
    case 'kana':
      if (!/^[ｱ-ﾝﾞﾟｧ-ｯ ・ー]+$/.test(stringValue)) {
        errors.push(`${validation.label}: 半角カナのみ入力してください。`);
      }
      break;
    
    case 'zenginFormat':
      if (!isValidZenginFormat(stringValue)) {
        errors.push(`${validation.label}: 全銀協フォーマット対応文字（半角カナ・英数字・記号、カンマ除く）で入力してください。`);
      }
      break;
    
    case 'alphanumeric':
      if (!/^[A-Za-z0-9]+$/.test(stringValue)) {
        errors.push(`${validation.label}: 英数字のみ入力してください。`);
      }
      break;
    
    case 'select':
      if (validation.validValues && !validation.validValues.includes(stringValue)) {
        // プルダウン値が「1:普通」形式の場合、「1」部分のみを抽出して再チェック
        const extractedValue = stringValue.includes(':') ? stringValue.split(':')[0] : stringValue;
        if (!validation.validValues.includes(extractedValue)) {
          errors.push(`${validation.label}: 有効な値を選択してください。`);
        }
      }
      break;
  }
  
  return errors;
}

/**
 * 振込データ行の検証
 * @param {Array} row - 行データ
 * @param {number} rowNum - 行番号
 * @return {string[]} エラーメッセージの配列
 */
function validateTransferDataRow(row, rowNum) {
  const errors = [];
  
  const validations = [
    { 
      value: row[TRANSFER_DATA_COLUMNS.BANK_CODE - 1], 
      label: '銀行コード', 
      exactLength: VALIDATION_RULES.BANK_CODE_LENGTH,
      required: true,
      type: 'number'
    },
    { 
      value: row[TRANSFER_DATA_COLUMNS.BRANCH_CODE - 1], 
      label: '支店コード', 
      exactLength: VALIDATION_RULES.BRANCH_CODE_LENGTH,
      required: true,
      type: 'number'
    },
    { 
      value: row[TRANSFER_DATA_COLUMNS.ACCOUNT_TYPE - 1], 
      label: '預金種目', 
      required: true,
      type: 'select',
      validValues: Object.values(ACCOUNT_TYPES).map(type => type.code)
    },
    { 
      value: row[TRANSFER_DATA_COLUMNS.ACCOUNT_NUMBER - 1], 
      label: '口座番号', 
      maxLength: VALIDATION_RULES.ACCOUNT_NUMBER_MAX_LENGTH,
      required: true,
      type: 'number'
    },
    { 
      value: row[TRANSFER_DATA_COLUMNS.RECIPIENT_NAME - 1], 
      label: '受取人名', 
      maxLength: VALIDATION_RULES.RECIPIENT_NAME_MAX_LENGTH,
      required: true,
      type: 'zenginFormat'
    },
    { 
      value: row[TRANSFER_DATA_COLUMNS.AMOUNT - 1], 
      label: '振込金額', 
      required: true,
      type: 'amount'
    },
    { 
      value: row[TRANSFER_DATA_COLUMNS.CUSTOMER_CODE - 1], 
      label: '顧客コード', 
      maxLength: VALIDATION_RULES.CUSTOMER_CODE_MAX_LENGTH,
      required: false,
      type: 'alphanumeric'
    },
    { 
      value: row[TRANSFER_DATA_COLUMNS.IDENTIFICATION - 1], 
      label: '識別表示', 
      required: false,
      type: 'select',
      validValues: Object.values(IDENTIFICATION_CODES)
    },
    { 
      value: row[TRANSFER_DATA_COLUMNS.EDI_INFO - 1], 
      label: 'EDI情報', 
      maxLength: VALIDATION_RULES.EDI_INFO_MAX_LENGTH,
      required: false,
      type: 'any'
    }
  ];
  
  for (const validation of validations) {
    const fieldErrors = validateField(validation.value, validation);
    if (fieldErrors.length > 0) {
      errors.push(...fieldErrors.map(error => `行${rowNum}: ${error}`));
    }
  }
  
  // 振込金額の特別チェック
  const amount = row[TRANSFER_DATA_COLUMNS.AMOUNT - 1];
  if (amount) {
    if (typeof amount !== 'number' || amount <= 0) {
      errors.push(`行${rowNum}: 振込金額は正の数値で入力してください。`);
    } else if (amount > VALIDATION_RULES.MAX_AMOUNT) {
      errors.push(`行${rowNum}: 振込金額が上限(${VALIDATION_RULES.MAX_AMOUNT.toLocaleString()}円)を超えています。`);
    }
  }
  
  // 金融機関コードの存在チェック（オプション）
  // 注意: 金融機関マスタが不完全な場合はこのチェックをコメントアウト
  /*
  const bankCode = row[TRANSFER_DATA_COLUMNS.BANK_CODE - 1];
  const branchCode = row[TRANSFER_DATA_COLUMNS.BRANCH_CODE - 1];
  if (bankCode && branchCode) {
    const bankMasterValid = validateBankMasterExists(bankCode, branchCode);
    if (!bankMasterValid.isValid) {
      errors.push(...bankMasterValid.errors.map(error => `行${rowNum}: ${error}`));
    }
  }
  */
  
  return errors;
}

/**
 * 金融機関マスタの存在チェック
 * @param {string} bankCode - 銀行コード
 * @param {string} branchCode - 支店コード
 * @return {Object} 検証結果
 */
function validateBankMasterExists(bankCode, branchCode) {
  try {
    const masterData = getMasterData();
    const exists = masterData.some(row => 
      row[BANK_MASTER_COLUMNS.BANK_CODE - 1] == bankCode && 
      row[BANK_MASTER_COLUMNS.BRANCH_CODE - 1] == branchCode && 
      row[BANK_MASTER_COLUMNS.STATUS - 1] === '有効'
    );
    
    if (!exists) {
      return { 
        isValid: false, 
        errors: [`金融機関コード(${bankCode}-${branchCode})が金融機関マスタに存在しないか、無効です。`] 
      };
    }
    
    return { isValid: true, errors: [] };
  } catch (error) {
    Logger.log('金融機関マスタ存在チェックエラー: ' + error.toString());
    return { 
      isValid: false, 
      errors: ['金融機関マスタの確認中にエラーが発生しました。'] 
    };
  }
}

/**
 * 行が空かどうかチェック
 * @param {Array} row - 行データ
 * @return {boolean} 空行かどうか
 */
function isEmptyRow(row) {
  return row.every(cell => !cell || String(cell).trim() === '');
}

/**
 * 半角カナ文字列の検証
 * @param {string} str - 検証対象文字列
 * @return {boolean} 有効かどうか
 */
function isValidKana(str) {
  // 半角カナ文字とスペースのみ許可
  return /^[ｱ-ﾝ ]+$/.test(str);
}

/**
 * 全銀協フォーマット対応文字列の検証
 * @param {string} str - 検証対象文字列
 * @return {boolean} 有効かどうか
 */
function isValidZenginFormat(str) {
  // 全銀協フォーマット対応文字: 半角カナ、英数字、記号（カンマ除く）、スペース
  // 使用可能文字: A-Z, 0-9, ｱ-ﾝ, 濁点, 半濁点, 長音, 中点, 各種記号（カンマ除く）
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