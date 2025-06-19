/**
 * スプレッドシート構成設定
 * 振込依頼人情報、振込データ、金融機関マスタの3シートを設定
 */

/**
 * 全シートの初期設定を実行
 */
function setupAllSheets() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    setupClientInfoSheet(spreadsheet);
    setupTransferDataSheet(spreadsheet);
    setupBankMasterSheet(spreadsheet);
    
    SpreadsheetApp.getUi().alert(
      'シート設定完了',
      '全てのシートの設定が完了しました。',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (error) {
    Logger.log('シート設定エラー: ' + error.toString());
    SpreadsheetApp.getUi().alert(
      'エラー',
      'シート設定中にエラーが発生しました: ' + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * 振込依頼人情報シートの設定
 * @param {Spreadsheet} spreadsheet - 対象スプレッドシート
 */
function setupClientInfoSheet(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(SHEET_NAMES.CLIENT_INFO);
  
  // シートが存在しない場合は作成
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAMES.CLIENT_INFO);
  }
  
  // シートをクリア
  sheet.clear();
  
  // タイトル行の設定
  sheet.getRange('A1').setValue('振込依頼人情報登録').setFontWeight('bold').setFontSize(14);
  
  // ラベルと入力フィールドの設定
  const labels = Object.values(CLIENT_INFO_LABELS);
  const cells = Object.keys(CLIENT_INFO_CELLS);
  
  for (let i = 0; i < labels.length; i++) {
    const rowNum = i + 2;
    const labelCell = 'A' + rowNum;
    const inputCell = CLIENT_INFO_CELLS[cells[i]];
    
    // ラベル設定
    sheet.getRange(labelCell).setValue(labels[i]).setFontWeight('bold');
    
    // 入力フィールドの背景色設定
    sheet.getRange(inputCell).setBackground('#f0f8ff');
  }
  
  // プルダウン設定
  setupClientInfoValidation(sheet);
  
  // 列幅調整
  sheet.setColumnWidth(1, 180); // A列（ラベル）
  sheet.setColumnWidth(2, 200); // B列（入力）
}

/**
 * 振込依頼人情報シートの入力検証設定
 * @param {Sheet} sheet - 対象シート
 */
function setupClientInfoValidation(sheet) {
  // 預金種目のプルダウン
  const accountTypeValues = Object.values(ACCOUNT_TYPES).map(type => type.code + ':' + type.name);
  const accountTypeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(accountTypeValues)
    .setAllowInvalid(false)
    .setHelpText('預金種目を選択してください')
    .build();
  sheet.getRange(CLIENT_INFO_CELLS.ACCOUNT_TYPE).setDataValidation(accountTypeRule);
  
  // 種別コードのプルダウン
  const categoryCodeValues = Object.values(CATEGORY_CODES).map(code => code.code + ':' + code.name);
  const categoryCodeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(categoryCodeValues)
    .setAllowInvalid(false)
    .setHelpText('種別コードを選択してください')
    .build();
  sheet.getRange(CLIENT_INFO_CELLS.CATEGORY_CODE).setDataValidation(categoryCodeRule);
  
  // 出力ファイル拡張子のプルダウン
  const fileExtRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(FILE_EXTENSIONS)
    .setAllowInvalid(false)
    .setHelpText('出力ファイル拡張子を選択してください')
    .build();
  sheet.getRange(CLIENT_INFO_CELLS.FILE_EXTENSION).setDataValidation(fileExtRule);
  
  // 数値項目の入力検証と書式設定
  const numberRule = SpreadsheetApp.newDataValidation()
    .requireNumberGreaterThan(0)
    .setAllowInvalid(false)
    .build();
  
  // 銀行コード: 4桁の数値書式
  const bankCodeCell = sheet.getRange(CLIENT_INFO_CELLS.BANK_CODE);
  bankCodeCell.setDataValidation(numberRule);
  bankCodeCell.setNumberFormat('0000');
  
  // 支店コード: 3桁の数値書式
  const branchCodeCell = sheet.getRange(CLIENT_INFO_CELLS.BRANCH_CODE);
  branchCodeCell.setDataValidation(numberRule);
  branchCodeCell.setNumberFormat('000');
  
  // 口座番号: 7桁の数値書式
  const accountNumberCell = sheet.getRange(CLIENT_INFO_CELLS.ACCOUNT_NUMBER);
  accountNumberCell.setDataValidation(numberRule);
  accountNumberCell.setNumberFormat('0000000');
}

/**
 * 振込データシートの設定
 * @param {Spreadsheet} spreadsheet - 対象スプレッドシート
 */
function setupTransferDataSheet(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(SHEET_NAMES.TRANSFER_DATA);
  
  // シートが存在しない場合は作成
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAMES.TRANSFER_DATA);
  }
  
  // シートをクリア
  sheet.clear();
  
  // ヘッダー行の設定
  const headers = Object.values(TRANSFER_DATA_HEADERS);
  for (let i = 0; i < headers.length; i++) {
    sheet.getRange(1, i + 1).setValue(headers[i]);
  }
  
  // ヘッダー行のスタイル設定
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold')
    .setBackground('#4CAF50')
    .setFontColor('white')
    .setHorizontalAlignment('center');
  
  // 列幅調整
  sheet.setColumnWidth(TRANSFER_DATA_COLUMNS.BANK_CODE, 80);      // 銀行コード
  sheet.setColumnWidth(TRANSFER_DATA_COLUMNS.BANK_NAME, 120);     // 銀行名
  sheet.setColumnWidth(TRANSFER_DATA_COLUMNS.BRANCH_CODE, 80);    // 支店コード
  sheet.setColumnWidth(TRANSFER_DATA_COLUMNS.BRANCH_NAME, 120);   // 支店名
  sheet.setColumnWidth(TRANSFER_DATA_COLUMNS.ACCOUNT_TYPE, 80);   // 預金種目
  sheet.setColumnWidth(TRANSFER_DATA_COLUMNS.ACCOUNT_NUMBER, 100); // 口座番号
  sheet.setColumnWidth(TRANSFER_DATA_COLUMNS.RECIPIENT_NAME, 150); // 受取人名
  sheet.setColumnWidth(TRANSFER_DATA_COLUMNS.AMOUNT, 100);        // 振込金額
  sheet.setColumnWidth(TRANSFER_DATA_COLUMNS.CUSTOMER_CODE, 100); // 顧客コード
  sheet.setColumnWidth(TRANSFER_DATA_COLUMNS.IDENTIFICATION, 80); // 識別表示
  sheet.setColumnWidth(TRANSFER_DATA_COLUMNS.EDI_INFO, 120);      // EDI情報
  
  // データ行の入力検証設定
  setupTransferDataValidation(sheet);
  
  // 行の固定（ヘッダー行）
  sheet.setFrozenRows(1);
}

/**
 * 振込データシートの入力検証設定
 * @param {Sheet} sheet - 対象シート
 */
function setupTransferDataValidation(sheet) {
  // 預金種目のプルダウン（データ行用）
  const accountTypeValues = Object.values(ACCOUNT_TYPES).map(type => type.code);
  const accountTypeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(accountTypeValues)
    .setAllowInvalid(false)
            .setHelpText('預金種目: 1=普通, 2=当座（4=貯蓄は全銀協対象外）')
    .build();
  
  // 識別表示のプルダウン
  const identificationValues = Object.values(IDENTIFICATION_CODES);
  const identificationRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(identificationValues)
    .setAllowInvalid(false)
    .setHelpText('識別表示: Y=給与, B=賞与')
    .build();
  
  // 大量行に対する検証設定（2行目から1000行目まで）
  const dataRange = sheet.getRange(2, 1, 999, Object.keys(TRANSFER_DATA_COLUMNS).length);
  
  // 預金種目列の検証設定
  sheet.getRange(2, TRANSFER_DATA_COLUMNS.ACCOUNT_TYPE, 999, 1).setDataValidation(accountTypeRule);
  
  // 識別表示列の検証設定
  sheet.getRange(2, TRANSFER_DATA_COLUMNS.IDENTIFICATION, 999, 1).setDataValidation(identificationRule);
  
  // 振込金額の数値検証
  const amountRule = SpreadsheetApp.newDataValidation()
    .requireNumberGreaterThan(0)
    .setAllowInvalid(false)
    .setHelpText('正の数値を入力してください')
    .build();
  sheet.getRange(2, TRANSFER_DATA_COLUMNS.AMOUNT, 999, 1).setDataValidation(amountRule);
  
  // 数値コード項目の書式設定
  // 銀行コード: 4桁の数値書式
  sheet.getRange(2, TRANSFER_DATA_COLUMNS.BANK_CODE, 999, 1).setNumberFormat('0000');
  
  // 支店コード: 3桁の数値書式
  sheet.getRange(2, TRANSFER_DATA_COLUMNS.BRANCH_CODE, 999, 1).setNumberFormat('000');
  
  // 口座番号: 7桁の数値書式
  sheet.getRange(2, TRANSFER_DATA_COLUMNS.ACCOUNT_NUMBER, 999, 1).setNumberFormat('0000000');
}

/**
 * 金融機関マスタシートの設定
 * @param {Spreadsheet} spreadsheet - 対象スプレッドシート
 */
function setupBankMasterSheet(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(SHEET_NAMES.BANK_MASTER);
  
  // シートが存在しない場合は作成
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAMES.BANK_MASTER);
  }
  
  // シートをクリア
  sheet.clear();
  
  // ヘッダー行の設定
  const headers = Object.values(BANK_MASTER_HEADERS);
  for (let i = 0; i < headers.length; i++) {
    sheet.getRange(1, i + 1).setValue(headers[i]);
  }
  
  // ヘッダー行のスタイル設定
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold')
    .setBackground('#2196F3')
    .setFontColor('white')
    .setHorizontalAlignment('center');
  
  // 列幅調整
  sheet.setColumnWidth(BANK_MASTER_COLUMNS.BANK_CODE, 80);      // 銀行コード
  sheet.setColumnWidth(BANK_MASTER_COLUMNS.BANK_NAME, 150);     // 銀行名
  sheet.setColumnWidth(BANK_MASTER_COLUMNS.BRANCH_CODE, 80);    // 支店コード
  sheet.setColumnWidth(BANK_MASTER_COLUMNS.BRANCH_NAME, 150);   // 支店名
  sheet.setColumnWidth(BANK_MASTER_COLUMNS.UPDATE_DATE, 100);   // 更新日
  sheet.setColumnWidth(BANK_MASTER_COLUMNS.STATUS, 80);         // 状態
  
  // データ行の入力検証設定
  setupBankMasterValidation(sheet);
  
  // 行の固定（ヘッダー行）
  sheet.setFrozenRows(1);
}

/**
 * 金融機関マスタシートの入力検証設定
 * @param {Sheet} sheet - 対象シート
 */
function setupBankMasterValidation(sheet) {
  // 状態のプルダウン（データ行用）
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(STATUS_OPTIONS)
    .setAllowInvalid(false)
    .setHelpText('状態を選択してください')
    .build();
  
  // 大量行に対する検証設定（2行目から10000行目まで）
  sheet.getRange(2, BANK_MASTER_COLUMNS.STATUS, 9999, 1).setDataValidation(statusRule);
  
  // 更新日の日付検証
  const dateRule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(false)
    .setHelpText('有効な日付を入力してください')
    .build();
  sheet.getRange(2, BANK_MASTER_COLUMNS.UPDATE_DATE, 9999, 1).setDataValidation(dateRule);
  
  // 数値コード項目の書式設定
  // 銀行コード: 4桁の数値書式
  sheet.getRange(2, BANK_MASTER_COLUMNS.BANK_CODE, 9999, 1).setNumberFormat('0000');
  
  // 支店コード: 3桁の数値書式
  sheet.getRange(2, BANK_MASTER_COLUMNS.BRANCH_CODE, 9999, 1).setNumberFormat('000');
}

/**
 * 指定シートのヘッダー行を取得
 * @param {string} sheetName - シート名
 * @return {Array} ヘッダー行の配列
 */
function getSheetHeaders(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return [];
  
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return headerRow.filter(cell => cell !== '');
} 