/**
 * カスタムメニューとイベントハンドラの実装
 * スプレッドシート開く時に自動でメニューを追加
 */

/**
 * スプレッドシート開く時に実行される関数
 */
function onOpen() {
  createCustomMenu();
}

/**
 * カスタムメニューの作成
 */
function createCustomMenu() {
  const ui = SpreadsheetApp.getUi();
  
  // メインメニューの作成
  const menu = ui.createMenu('全銀協システム');
  
  // 基本機能メニュー
  menu.addItem('振込用CSV取込処理', 'showCsvImportDialog');
  menu.addItem('振込データ作成処理', 'generateZenginFile');
  menu.addSeparator();
  menu.addItem('データ検証', 'validateAllData');
  
  // デバッグ・テスト機能
  const debugSubmenu = ui.createMenu('デバッグ・テスト');
  debugSubmenu.addItem('Shift_JIS変換テスト', 'runShiftJISTest');
  debugSubmenu.addItem('ファイル生成デバッグ', 'runZenginFileDebug');
  debugSubmenu.addItem('自動補完テスト', 'runAutoCompleteTest');
  debugSubmenu.addItem('金融機関マスタCSV取込テスト', 'runBankMasterCsvTest');
  debugSubmenu.addItem('受取人名検証テスト', 'runRecipientNameTest');
  debugSubmenu.addItem('振込データCSV取込テスト', 'runTransferDataCsvTest');
  
  menu.addSubMenu(debugSubmenu);
  
  // 金融機関マスタ管理サブメニュー
  const bankMasterSubmenu = ui.createMenu('金融機関マスタ管理');
  bankMasterSubmenu.addItem('金融機関データ一括取込', 'showBankMasterImportDialog');
  bankMasterSubmenu.addItem('銀行・支店名自動補完', 'executeAutoComplete');
  bankMasterSubmenu.addItem('マスタデータ整備', 'cleanupBankMasterData');
  
  menu.addSubMenu(bankMasterSubmenu);
  
  menu.addSeparator();
  menu.addItem('システム設定', 'showSystemSettings');
  
  // メニューを追加
  menu.addToUi();
}

/**
 * 振込用CSV取込処理ダイアログの表示
 */
function showCsvImportDialog() {
  try {
    const htmlTemplate = HtmlService.createTemplate(`
      <div style="font-family: Arial, sans-serif; padding: 20px;">
        <h3>振込用CSV取込処理</h3>
        <div style="margin-bottom: 15px;">
          <label for="fileInput">CSVファイルを選択:</label><br>
          <input type="file" id="fileInput" accept=".csv" style="margin-top: 5px;">
        </div>
        <div style="margin-bottom: 15px;">
          <label>
            <input type="radio" name="importMode" value="overwrite" checked>
            既存データを上書き
          </label><br>
          <label>
            <input type="radio" name="importMode" value="append">
            既存データに追記
          </label>
        </div>
        <div>
          <button onclick="handleCsvImport()" style="background-color: #4CAF50; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; margin-right: 10px;">
            取込実行
          </button>
          <button onclick="google.script.host.close()" style="background-color: #f44336; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer;">
            キャンセル
          </button>
        </div>
      </div>
      
      <script>
        function handleCsvImport() {
          const fileInput = document.getElementById('fileInput');
          const importMode = document.querySelector('input[name="importMode"]:checked').value;
          
          if (!fileInput.files[0]) {
            alert('ファイルを選択してください。');
            return;
          }
          
          const file = fileInput.files[0];
          const reader = new FileReader();
          
          reader.onload = function(e) {
            const csvData = e.target.result;
            google.script.run
              .withSuccessHandler(onImportSuccess)
              .withFailureHandler(onImportFailure)
              .importTransferDataFromCsv(csvData, importMode);
          };
          
          reader.readAsText(file, 'UTF-8');
        }
        
        function onImportSuccess(result) {
          alert('CSV取込が完了しました。\\n' + result);
          google.script.host.close();
        }
        
        function onImportFailure(error) {
          alert('エラーが発生しました: ' + error.message);
        }
      </script>
    `);
    
    const htmlOutput = htmlTemplate.evaluate()
      .setWidth(UI_CONFIG.DIALOG_WIDTH)
      .setHeight(UI_CONFIG.DIALOG_HEIGHT);
    
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, '振込用CSV取込処理');
  } catch (error) {
    Logger.log('CSV取込ダイアログエラー: ' + error.toString());
    SpreadsheetApp.getUi().alert('エラー', 'ダイアログの表示に失敗しました: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 振込データ作成処理の実行
 */
function generateZenginFile() {
  try {
    logSystemActivity('generateZenginFile', '振込データ作成処理開始', 'INFO');
    
    // データ検証を先に実行
    const validationResult = validateAllData();
    
    if (!validationResult.isValid) {
      logSystemActivity('generateZenginFile', `データ検証エラー: ${validationResult.errors.length}件`, 'ERROR');
      SpreadsheetApp.getUi().alert(
        'データ検証エラー',
        'データにエラーがあります:\n' + validationResult.errors.join('\n'),
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    logSystemActivity('generateZenginFile', 'データ検証OK - ファイル生成開始', 'INFO');
    
    // 全銀協フォーマットファイルの生成
    const result = createZenginFormatFile();
    
    if (result.success) {
      logSystemActivity('generateZenginFile', `ファイル生成成功 - ${result.fileName}, ${result.recordCount}件`, 'INFO');
      SpreadsheetApp.getUi().alert(
        'ファイル生成完了',
        '全銀協フォーマットファイルが生成されました。\n\n' +
        'ファイル名: ' + result.fileName + '\n' +
        '処理件数: ' + result.recordCount + '件\n' +
        'ダウンロードしてください。',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } else {
      logSystemActivity('generateZenginFile', `ファイル生成失敗: ${result.error}`, 'ERROR');
      SpreadsheetApp.getUi().alert(
        'エラー',
        'ファイル生成に失敗しました: ' + result.error,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  } catch (error) {
    logSystemActivity('generateZenginFile', `例外エラー: ${error.message}`, 'ERROR');
    Logger.log('振込データ作成エラー: ' + error.toString());
    SpreadsheetApp.getUi().alert(
      'エラー',
      '振込データ作成中にエラーが発生しました: ' + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * データ検証の実行
 */
function validateAllData() {
  try {
    logSystemActivity('validateAllData', 'データ検証開始', 'INFO');
    
    const clientValidation = validateClientInfo();
    const transferValidation = validateTransferData();
    
    const errors = [];
    
    if (!clientValidation.isValid) {
      errors.push('【振込依頼人情報】');
      errors.push(...clientValidation.errors);
    }
    
    if (!transferValidation.isValid) {
      errors.push('【振込データ】');
      errors.push(...transferValidation.errors);
    }
    
    const isValid = errors.length === 0;
    
    if (isValid) {
      logSystemActivity('validateAllData', 'データ検証完了 - エラーなし', 'INFO');
      SpreadsheetApp.getUi().alert(
        'データ検証結果',
        '全てのデータが正常です。',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } else {
      logSystemActivity('validateAllData', `データ検証完了 - エラー${errors.length}件`, 'WARNING');
      SpreadsheetApp.getUi().alert(
        'データ検証結果',
        'エラーが見つかりました:\n\n' + errors.join('\n'),
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    
    return { isValid, errors };
  } catch (error) {
    logSystemActivity('validateAllData', `例外エラー: ${error.message}`, 'ERROR');
    Logger.log('データ検証エラー: ' + error.toString());
    return { isValid: false, errors: ['データ検証中にエラーが発生しました: ' + error.message] };
  }
}

/**
 * 金融機関データ一括取込ダイアログの表示
 */
function showBankMasterImportDialog() {
  try {
    const htmlTemplate = HtmlService.createTemplate(`
      <div style="font-family: Arial, sans-serif; padding: 20px;">
        <h3>金融機関データ一括取込</h3>
        <div style="margin-bottom: 15px;">
          <p>対応フォーマット: 銀行コード,銀行名,支店コード,支店名,状態</p>
          <label for="fileInput">CSVファイルを選択:</label><br>
          <input type="file" id="fileInput" accept=".csv" style="margin-top: 5px;">
        </div>
        <div style="margin-bottom: 15px;">
          <label>
            <input type="checkbox" id="duplicateCheck" checked>
            重複データをチェックする
          </label>
        </div>
        <div>
          <button onclick="handleBankMasterImport()" style="background-color: #2196F3; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; margin-right: 10px;">
            取込実行
          </button>
          <button onclick="google.script.host.close()" style="background-color: #f44336; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer;">
            キャンセル
          </button>
        </div>
      </div>
      
      <script>
        function handleBankMasterImport() {
          const fileInput = document.getElementById('fileInput');
          const duplicateCheck = document.getElementById('duplicateCheck').checked;
          
          if (!fileInput.files[0]) {
            alert('ファイルを選択してください。');
            return;
          }
          
          const file = fileInput.files[0];
          const reader = new FileReader();
          
          reader.onload = function(e) {
            const csvData = e.target.result;
            google.script.run
              .withSuccessHandler(onImportSuccess)
              .withFailureHandler(onImportFailure)
              .importBankMasterFromCsv(csvData, duplicateCheck);
          };
          
          reader.readAsText(file, 'UTF-8');
        }
        
        function onImportSuccess(result) {
          alert('金融機関データ取込が完了しました。\\n' + result);
          google.script.host.close();
        }
        
        function onImportFailure(error) {
          alert('エラーが発生しました: ' + error.message);
        }
      </script>
    `);
    
    const htmlOutput = htmlTemplate.evaluate()
      .setWidth(UI_CONFIG.DIALOG_WIDTH + 100)
      .setHeight(UI_CONFIG.DIALOG_HEIGHT);
    
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, '金融機関データ一括取込');
  } catch (error) {
    Logger.log('金融機関取込ダイアログエラー: ' + error.toString());
    SpreadsheetApp.getUi().alert('エラー', 'ダイアログの表示に失敗しました: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 銀行・支店名自動補完の実行
 */
function executeAutoComplete() {
  try {
    logSystemActivity('executeAutoComplete', '自動補完処理開始', 'INFO');
    
    const result = bulkAutoComplete();
    
    logSystemActivity('executeAutoComplete', `自動補完完了 - 銀行名: ${result.bankNameCompletions}件, 支店名: ${result.branchNameCompletions}件, 失敗: ${result.failures}件`, 'INFO');
    
    SpreadsheetApp.getUi().alert(
      '自動補完完了',
      `補完処理が完了しました。\n\n` +
      `検査対象行数: ${result.totalRows}行\n` +
      `銀行名補完件数: ${result.bankNameCompletions}件\n` +
      `支店名補完件数: ${result.branchNameCompletions}件\n` +
      `補完失敗件数: ${result.failures}件\n` +
      `処理時間: ${result.processingTime}ms`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (error) {
    logSystemActivity('executeAutoComplete', `エラー: ${error.message}`, 'ERROR');
    Logger.log('自動補完エラー: ' + error.toString());
    SpreadsheetApp.getUi().alert(
      'エラー',
      '自動補完中にエラーが発生しました: ' + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * マスタデータ整備の実行
 */
function cleanupBankMasterData() {
  try {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'マスタデータ整備',
      '重複データの削除とデータ整合性チェックを実行します。\n続行しますか？',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      const result = cleanupMasterData();
      
      ui.alert(
        '整備完了',
        `マスタデータ整備が完了しました。\n\n` +
        `重複削除件数: ${result.duplicatesRemoved}件\n` +
        `データ修正件数: ${result.dataFixed}件\n` +
        `無効データ件数: ${result.invalidData}件`,
        ui.ButtonSet.OK
      );
    }
  } catch (error) {
    Logger.log('マスタデータ整備エラー: ' + error.toString());
    SpreadsheetApp.getUi().alert(
      'エラー',
      'マスタデータ整備中にエラーが発生しました: ' + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * システム設定の表示
 */
function showSystemSettings() {
  try {
    const ui = SpreadsheetApp.getUi();
    const menu = ui.createMenu('システム設定');
    
    menu.addItem('シート構成の初期化', 'setupAllSheets');
    menu.addItem('キャッシュクリア', 'clearAllCache');
    menu.addItem('ログ表示', 'showSystemLogs');
    
    // 設定メニューをダイアログで表示
    const htmlTemplate = HtmlService.createTemplate(`
      <div style="font-family: Arial, sans-serif; padding: 20px;">
        <h3>システム設定</h3>
        <div style="margin-bottom: 15px;">
          <button onclick="runFunction('setupAllSheets')" style="width: 100%; padding: 10px; margin-bottom: 10px; background-color: #4CAF50; color: white; border: none; border-radius: 4px; cursor: pointer;">
            シート構成の初期化
          </button>
          <button onclick="runFunction('clearAllCache')" style="width: 100%; padding: 10px; margin-bottom: 10px; background-color: #FF9800; color: white; border: none; border-radius: 4px; cursor: pointer;">
            キャッシュクリア
          </button>
          <button onclick="runFunction('showSystemLogs')" style="width: 100%; padding: 10px; margin-bottom: 10px; background-color: #2196F3; color: white; border: none; border-radius: 4px; cursor: pointer;">
            ログ表示
          </button>
        </div>
        <div>
          <button onclick="google.script.host.close()" style="width: 100%; padding: 10px; background-color: #f44336; color: white; border: none; border-radius: 4px; cursor: pointer;">
            閉じる
          </button>
        </div>
      </div>
      
      <script>
        function runFunction(functionName) {
          google.script.run
            .withSuccessHandler(onSuccess)
            .withFailureHandler(onFailure)
            [functionName]();
        }
        
        function onSuccess(result) {
          if (result) {
            alert(result);
          }
        }
        
        function onFailure(error) {
          alert('エラーが発生しました: ' + error.message);
        }
      </script>
    `);
    
    const htmlOutput = htmlTemplate.evaluate()
      .setWidth(UI_CONFIG.DIALOG_WIDTH)
      .setHeight(UI_CONFIG.DIALOG_HEIGHT);
    
    ui.showModalDialog(htmlOutput, 'システム設定');
  } catch (error) {
    Logger.log('システム設定エラー: ' + error.toString());
    SpreadsheetApp.getUi().alert('エラー', 'システム設定の表示に失敗しました: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 全キャッシュのクリア
 */
function clearAllCache() {
  try {
    const cache = CacheService.getScriptCache();
    cache.removeAll(['bankMasterData', 'lastCacheUpdate']);
    return 'キャッシュをクリアしました。';
  } catch (error) {
    Logger.log('キャッシュクリアエラー: ' + error.toString());
    throw new Error('キャッシュクリアに失敗しました: ' + error.message);
  }
}

/**
 * システムログの表示
 */
function showSystemLogs() {
  try {
    // 従来のLogger.log()で記録されたログを取得
    const legacyLogs = Logger.getLog();
    
    // 最新の実行履歴からログ情報を構築
    const executionLogs = getRecentExecutionLogs();
    
    // 両方のログを結合
    let combinedLogs = '';
    
    if (executionLogs && executionLogs.trim()) {
      combinedLogs += '【最新実行ログ】\n';
      combinedLogs += executionLogs;
      combinedLogs += '\n\n';
    }
    
    if (legacyLogs && legacyLogs.trim()) {
      combinedLogs += '【詳細ログ（Logger.log）】\n';
      combinedLogs += legacyLogs.split('\n').slice(-50).join('\n'); // 最新50件
    }
    
    const ui = SpreadsheetApp.getUi();
    
    if (!combinedLogs.trim()) {
      ui.alert(
        'システムログ',
        'ログエントリがありません。\n\n※ログが表示されない場合は、\n1. 何らかの処理を実行してから再度確認してください\n2. GASの実行履歴で詳細ログを確認できます',
        ui.ButtonSet.OK
      );
    } else {
      // 長すぎる場合は最新部分のみ表示
      if (combinedLogs.length > 8000) {
        combinedLogs = '...(ログが長いため最新部分のみ表示)\n\n' + combinedLogs.slice(-8000);
      }
      
      ui.alert(
        'システムログ',
        combinedLogs,
        ui.ButtonSet.OK
      );
    }
    
    return 'ログを表示しました。';
  } catch (error) {
    console.log('ログ表示エラー: ' + error.toString());
    Logger.log('ログ表示エラー: ' + error.toString());
    throw new Error('ログ表示に失敗しました: ' + error.message);
  }
}

/**
 * 最新の実行履歴からログ情報を取得
 * @return {string} 実行ログ情報
 */
function getRecentExecutionLogs() {
  try {
    // PropertiesServiceを使用してログ情報を保存・取得
    const properties = PropertiesService.getScriptProperties();
    const storedLogs = properties.getProperty('systemExecutionLogs');
    
    if (storedLogs) {
      const logs = JSON.parse(storedLogs);
      const recentLogs = logs.slice(-20); // 最新20件
      
      return recentLogs.map(log => {
        const timestamp = new Date(log.timestamp).toLocaleString('ja-JP');
        return `${timestamp} - ${log.functionName}: ${log.message}`;
      }).join('\n');
    }
    
    return '';
  } catch (error) {
    console.log('実行ログ取得エラー: ' + error.toString());
    return '実行ログの取得に失敗しました: ' + error.message;
  }
}

/**
 * システムログを記録する関数
 * @param {string} functionName - 実行した関数名
 * @param {string} message - ログメッセージ
 * @param {string} level - ログレベル（INFO/WARNING/ERROR）
 */
function logSystemActivity(functionName, message, level = 'INFO') {
  try {
    // 従来のLoggerとconsole.logの両方に出力
    const logMessage = `[${level}] ${functionName}: ${message}`;
    Logger.log(logMessage);
    console.log(logMessage);
    
    // PropertiesServiceにも保存
    const properties = PropertiesService.getScriptProperties();
    let storedLogs = [];
    
    try {
      const existing = properties.getProperty('systemExecutionLogs');
      if (existing) {
        storedLogs = JSON.parse(existing);
      }
    } catch (parseError) {
      console.log('既存ログ解析エラー: ' + parseError.toString());
    }
    
    // 新しいログエントリを追加
    storedLogs.push({
      timestamp: new Date().toISOString(),
      functionName: functionName,
      message: message,
      level: level
    });
    
    // 最新100件のみ保持
    if (storedLogs.length > 100) {
      storedLogs = storedLogs.slice(-100);
    }
    
    properties.setProperty('systemExecutionLogs', JSON.stringify(storedLogs));
    
  } catch (error) {
    console.log('ログ記録エラー: ' + error.toString());
  }
}

/**
 * Shift_JIS変換テストの実行
 */
function runShiftJISTest() {
  try {
    const testString = 'ﾃｽﾄｷｷﾞｮｳ ABC123 ･';
    const result = testShiftJISConversion(testString);
    
    if (result.success) {
      SpreadsheetApp.getUi().alert(
        'Shift_JIS変換テスト結果',
        `入力文字列: "${result.input}"\n` +
        `変換後バイト数: ${result.byteCount}バイト\n` +
        `16進表示: ${result.hexBytes}\n\n` +
        '詳細はログをご確認ください。',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } else {
      SpreadsheetApp.getUi().alert(
        'テストエラー',
        'Shift_JIS変換テストに失敗しました: ' + result.error,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  } catch (error) {
    Logger.log('Shift_JISテストエラー: ' + error.toString());
    SpreadsheetApp.getUi().alert(
      'エラー',
      'テスト実行中にエラーが発生しました: ' + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * 全銀協ファイル生成デバッグの実行
 */
function runZenginFileDebug() {
  try {
    const result = createZenginFormatFileDebug();
    
    if (result.success) {
      let message = `ファイル生成デバッグ完了\n\n` +
                   `ファイル名: ${result.fileName}\n` +
                   `処理件数: ${result.recordCount}件\n` +
                   `ファイルID: ${result.fileId}`;
      
      if (result.debugInfo) {
        message += `\n\n【デバッグ情報】\n` +
                  `実際のファイルサイズ: ${result.debugInfo.actualFileSize}バイト\n` +
                  `先頭20バイト: ${result.debugInfo.firstBytesHex}`;
      }
      
      SpreadsheetApp.getUi().alert(
        'デバッグ実行完了',
        message + '\n\n詳細はログをご確認ください。',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } else {
      SpreadsheetApp.getUi().alert(
        'デバッグエラー',
        'デバッグ実行に失敗しました: ' + result.error,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  } catch (error) {
    Logger.log('デバッグ実行エラー: ' + error.toString());
    SpreadsheetApp.getUi().alert(
      'エラー',
      'デバッグ実行中にエラーが発生しました: ' + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * 自動補完テストの実行
 */
function runAutoCompleteTest() {
  try {
    const result = testAutoComplete();
    
    if (result.success) {
      let message = `自動補完テスト完了\n\n` +
                   `振込データシート発見: ${result.transferSheetFound ? 'はい' : 'いいえ'}\n` +
                   `金融機関マスタシート発見: ${result.masterSheetFound ? 'はい' : 'いいえ'}\n` +
                   `マスタデータ件数: ${result.masterDataCount}件\n` +
                   `テスト銀行名検索(0001): ${result.testBankName || '見つかりません'}\n` +
                   `テスト支店名検索(0001-021): ${result.testBranchName || '見つかりません'}`;
      
      SpreadsheetApp.getUi().alert(
        '自動補完テスト結果',
        message + '\n\n詳細はログをご確認ください。',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } else {
      SpreadsheetApp.getUi().alert(
        'テストエラー',
        '自動補完テストに失敗しました: ' + result.error,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  } catch (error) {
    Logger.log('自動補完テストエラー: ' + error.toString());
    SpreadsheetApp.getUi().alert(
      'エラー',
      'テスト実行中にエラーが発生しました: ' + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * 金融機関マスタCSV取込テストの実行
 */
function runBankMasterCsvTest() {
  try {
    Logger.log('=== 金融機関マスタCSVテスト開始 ===');
    
    const result = testBankMasterCsvImport();
    
    SpreadsheetApp.getUi().alert(
      '金融機関マスタCSVテスト結果',
      'テストが完了しました。\n\n詳細はログをご確認ください。',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    Logger.log('=== 金融機関マスタCSVテスト終了 ===');
    
  } catch (error) {
    Logger.log('金融機関マスタCSVテストエラー: ' + error.toString());
    SpreadsheetApp.getUi().alert(
      'テストエラー',
      '金融機関マスタCSVテストに失敗しました: ' + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * 受取人名検証テストの実行
 */
function runRecipientNameTest() {
  try {
    Logger.log('=== 受取人名検証テスト開始 ===');
    
    const result = testRecipientNameValidation();
    
    SpreadsheetApp.getUi().alert(
      '受取人名検証テスト結果',
      'テストが完了しました。\n\n詳細はログをご確認ください。',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    Logger.log('=== 受取人名検証テスト終了 ===');
    
  } catch (error) {
    Logger.log('受取人名検証テストエラー: ' + error.toString());
    SpreadsheetApp.getUi().alert(
      'テストエラー',
      '受取人名検証テストに失敗しました: ' + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * 振込データCSV取込テストの実行
 */
function runTransferDataCsvTest() {
  try {
    Logger.log('=== 振込データCSV取込テスト開始 ===');
    
    const result = testTransferDataCsvImport();
    
    let message = 'テストが完了しました。\n\n';
    if (result.success) {
      message += `検証結果: ${result.isValid ? 'OK' : 'NG'}\n`;
      if (!result.isValid && result.errors) {
        message += `エラー数: ${result.errors.length}件\n`;
      }
    } else {
      message += `エラー: ${result.error}\n`;
    }
    message += '\n詳細はログをご確認ください。';
    
    SpreadsheetApp.getUi().alert(
      '振込データCSV取込テスト結果',
      message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    Logger.log('=== 振込データCSV取込テスト終了 ===');
    
  } catch (error) {
    Logger.log('振込データCSV取込テストエラー: ' + error.toString());
    SpreadsheetApp.getUi().alert(
      'テストエラー',
      '振込データCSV取込テストに失敗しました: ' + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
} 