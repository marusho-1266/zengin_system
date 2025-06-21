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
  const startTime = new Date().getTime();
  try {
    logSystemActivityEnhanced('generateZenginFile', '振込データ作成処理開始', 'INFO', 'ファイル生成', { startTime: startTime });
    
    // データ検証を先に実行
    const validationResult = validateAllData();
    
    if (!validationResult.isValid) {
      const endTime = new Date().getTime();
      logSystemActivityEnhanced('generateZenginFile', `データ検証エラー: ${validationResult.errors.length}件`, 'ERROR', 'ファイル生成', { 
        processingTime: Math.round(endTime - startTime),
        errorCount: validationResult.errors.length 
      });
      SpreadsheetApp.getUi().alert(
        'データ検証エラー',
        'データにエラーがあります:\n' + validationResult.errors.join('\n'),
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    logSystemActivityEnhanced('generateZenginFile', 'データ検証OK - ファイル生成開始', 'INFO', 'ファイル生成');
    
    // 全銀協フォーマットファイルの生成
    const result = createZenginFormatFile();
    
    const endTime = new Date().getTime();
    const processingTime = Math.round(endTime - startTime);
    
    if (result.success) {
      logSystemActivityEnhanced('generateZenginFile', `ファイル生成成功 - ${result.fileName}, ${result.recordCount}件`, 'INFO', 'ファイル生成', {
        processingTime: processingTime,
        recordCount: result.recordCount,
        fileName: result.fileName
      });
      SpreadsheetApp.getUi().alert(
        'ファイル生成完了',
        '全銀協フォーマットファイルが生成されました。\n\n' +
        'ファイル名: ' + result.fileName + '\n' +
        '処理件数: ' + result.recordCount + '件\n' +
        'ダウンロードしてください。',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } else {
      logSystemActivityEnhanced('generateZenginFile', `ファイル生成失敗: ${result.error}`, 'ERROR', 'ファイル生成', {
        processingTime: processingTime,
        error: result.error
      });
      SpreadsheetApp.getUi().alert(
        'エラー',
        'ファイル生成に失敗しました: ' + result.error,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  } catch (error) {
    const endTime = new Date().getTime();
    logSystemActivityEnhanced('generateZenginFile', `例外エラー: ${error.message}`, 'ERROR', 'ファイル生成', {
      processingTime: Math.round(endTime - startTime),
      error: error.message
    });
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
  const startTime = new Date().getTime();
  try {
    logSystemActivityEnhanced('validateAllData', 'データ検証開始', 'INFO', 'データ検証', { startTime: startTime });
    
    const clientValidation = validateClientInfo();
    const transferValidation = validateTransferData();
    
    const errors = [];
    const warnings = [];
    
    if (!clientValidation.isValid) {
      errors.push('【振込依頼人情報】');
      errors.push(...clientValidation.errors);
    }
    
    if (!transferValidation.isValid) {
      errors.push('【振込データ】');
      errors.push(...transferValidation.errors);
    }
    
    // 警告の収集
    if (transferValidation.warnings && transferValidation.warnings.length > 0) {
      warnings.push('【振込データ - 警告】');
      warnings.push(...transferValidation.warnings);
    }
    
    const isValid = errors.length === 0;
    const endTime = new Date().getTime();
    const processingTime = Math.round(endTime - startTime);
    
    if (isValid && warnings.length === 0) {
      logSystemActivityEnhanced('validateAllData', 'データ検証完了 - エラー・警告なし', 'INFO', 'データ検証', {
        processingTime: processingTime,
        errorCount: 0,
        warningCount: 0
      });
      SpreadsheetApp.getUi().alert(
        'データ検証結果',
        '全てのデータが正常です。',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } else if (isValid && warnings.length > 0) {
      logSystemActivityEnhanced('validateAllData', `データ検証完了 - 警告${warnings.length}件`, 'WARNING', 'データ検証', {
        processingTime: processingTime,
        errorCount: 0,
        warningCount: warnings.length
      });
      SpreadsheetApp.getUi().alert(
        'データ検証結果',
        'エラーはありませんが、以下の警告があります:\n\n' + warnings.join('\n') + 
        '\n\n※ 警告は推奨事項です。処理は続行可能です。',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } else {
      logSystemActivityEnhanced('validateAllData', `データ検証完了 - エラー${errors.length}件, 警告${warnings.length}件`, 'ERROR', 'データ検証', {
        processingTime: processingTime,
        errorCount: errors.length,
        warningCount: warnings.length
      });
      let message = 'エラーが見つかりました:\n\n' + errors.join('\n');
      if (warnings.length > 0) {
        message += '\n\n【警告事項】\n' + warnings.join('\n');
      }
      SpreadsheetApp.getUi().alert(
        'データ検証結果',
        message,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    
    return { isValid, errors, warnings };
  } catch (error) {
    const endTime = new Date().getTime();
    logSystemActivityEnhanced('validateAllData', `例外エラー: ${error.message}`, 'ERROR', 'データ検証', {
      processingTime: Math.round(endTime - startTime),
      error: error.message
    });
    Logger.log('データ検証エラー: ' + error.toString());
    return { isValid: false, errors: ['データ検証中にエラーが発生しました: ' + error.message], warnings: [] };
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
      
      let message = `マスタデータ整備が完了しました。\n\n` +
        `重複削除件数: ${result.duplicatesRemoved}件\n` +
        `データ修正件数: ${result.dataFixed}件\n` +
        `無効データ件数: ${result.invalidData}件`;
      
      if (result.invalidDataDetails && result.invalidDataDetails.length > 0) {
        message += `\n\n無効データ詳細:`;
        result.invalidDataDetails.forEach(detail => {
          message += `\n行${detail.row}: ${detail.reasons.join(', ')}`;
        });
      }
      
      ui.alert('整備完了', message, ui.ButtonSet.OK);
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
          <button onclick="runFunction('showLogFilterDialog')" style="width: 100%; padding: 10px; margin-bottom: 10px; background-color: #9C27B0; color: white; border: none; border-radius: 4px; cursor: pointer;">
            ログフィルタ表示
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
 * システムログの表示（拡張HTML形式）
 */
function showSystemLogs() {
  try {
    logSystemActivityEnhanced('showSystemLogs', 'システムログ表示開始', 'INFO', 'システム');
    
    // PropertiesServiceからログデータを取得
    const properties = PropertiesService.getScriptProperties();
    const storedLogs = properties.getProperty('systemExecutionLogs');
    
    if (!storedLogs) {
      SpreadsheetApp.getUi().alert(
        'システムログ', 
        'ログエントリがありません。\n\n※ログが表示されない場合は、\n1. 何らかの処理を実行してから再度確認してください\n2. GASの実行履歴で詳細ログを確認できます', 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return 'ログエントリがありませんでした。';
    }
    
    let logs = JSON.parse(storedLogs);
    
    // 最新50件のログを取得
    const filteredLogs = logs.slice(-50);
    
    // デフォルトフィルタ設定（全レベル・全カテゴリ）
    const defaultFilters = {
      levels: ['INFO', 'WARNING', 'ERROR'],
      categories: ['CSV取込', 'データ検証', 'ファイル生成', '自動補完', 'マスタ管理', 'システム'],
      displayCount: 50,
      keyword: '',
      displayMode: 'enhanced'
    };
    
    // HTML形式でログを表示
    const htmlContent = createLogHtmlContent(filteredLogs, defaultFilters);
    const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
      .setWidth(900)
      .setHeight(600)
      .setTitle('システムログ - 最新50件');
      
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'システムログ - 最新50件');
    
    logSystemActivityEnhanced('showSystemLogs', `システムログ表示完了: ${filteredLogs.length}件表示`, 'INFO', 'システム');
    
    return 'ログを表示しました。';
  } catch (error) {
    logSystemActivityEnhanced('showSystemLogs', `ログ表示エラー: ${error.message}`, 'ERROR', 'システム');
    Logger.log('ログ表示エラー: ' + error.toString());
    throw new Error('ログ表示に失敗しました: ' + error.message);
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
 * 拡張されたシステムログを記録する関数（カテゴリ・パフォーマンス情報付き）
 * @param {string} functionName - 実行した関数名
 * @param {string} message - ログメッセージ
 * @param {string} level - ログレベル（INFO/WARNING/ERROR）
 * @param {string} category - 機能分類
 * @param {Object} details - 詳細情報（処理時間、件数等）
 */
function logSystemActivityEnhanced(functionName, message, level = 'INFO', category = 'システム', details = {}) {
  try {
    // 従来のLoggerとconsole.logの両方に出力
    const logMessage = `[${level}][${category}] ${functionName}: ${message}`;
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
    
    // 新しいログエントリを追加（拡張形式）
    storedLogs.push({
      timestamp: new Date().toISOString(),
      functionName: functionName,
      message: message,
      level: level,
      category: category,
      details: details
    });
    
    // 最新100件のみ保持
    if (storedLogs.length > 100) {
      storedLogs = storedLogs.slice(-100);
    }
    
    properties.setProperty('systemExecutionLogs', JSON.stringify(storedLogs));
    
  } catch (error) {
    console.log('拡張ログ記録エラー: ' + error.toString());
  }
}

/**
 * ログフィルタダイアログの表示
 */
function showLogFilterDialog() {
  try {
    const htmlTemplate = HtmlService.createTemplate(`
      <div style="font-family: Arial, sans-serif; padding: 20px;">
        <h3>ログフィルタ設定</h3>
        
        <!-- ログレベルフィルタ -->
        <div style="margin-bottom: 15px;">
          <label style="font-weight: bold;">ログレベル:</label><br>
          <label><input type="checkbox" id="filterINFO" checked> INFO</label><br>
          <label><input type="checkbox" id="filterWARNING" checked> WARNING</label><br>
          <label><input type="checkbox" id="filterERROR" checked> ERROR</label>
        </div>
        
        <!-- 機能分類フィルタ -->
        <div style="margin-bottom: 15px;">
          <label style="font-weight: bold;">機能分類:</label><br>
          <label><input type="checkbox" id="catCSV取込" checked> CSV取込</label><br>
          <label><input type="checkbox" id="catデータ検証" checked> データ検証</label><br>
          <label><input type="checkbox" id="catファイル生成" checked> ファイル生成</label><br>
          <label><input type="checkbox" id="cat自動補完" checked> 自動補完</label><br>
          <label><input type="checkbox" id="catマスタ管理" checked> マスタ管理</label><br>
          <label><input type="checkbox" id="catシステム" checked> システム</label>
        </div>
        
        <!-- 表示件数 -->
        <div style="margin-bottom: 15px;">
          <label style="font-weight: bold;">表示件数:</label><br>
          <select id="displayCount">
            <option value="20">最新20件</option>
            <option value="50" selected>最新50件</option>
            <option value="100">最新100件</option>
          </select>
        </div>
        
        <!-- キーワード検索 -->
        <div style="margin-bottom: 15px;">
          <label style="font-weight: bold;">キーワード検索:</label><br>
          <input type="text" id="searchKeyword" placeholder="検索キーワードを入力" style="width: 100%; padding: 5px;">
        </div>
        
        <div>
          <button onclick="applyLogFilter()" style="width: 48%; padding: 10px; background-color: #4CAF50; color: white; border: none; border-radius: 4px; cursor: pointer; margin-right: 4%;">
            フィルタ適用
          </button>
          <button onclick="google.script.host.close()" style="width: 48%; padding: 10px; background-color: #f44336; color: white; border: none; border-radius: 4px; cursor: pointer;">
            キャンセル
          </button>
        </div>
      </div>
      
      <script>
        function applyLogFilter() {
          // フィルタ条件を収集
          const filters = {
            levels: [],
            categories: [],
            displayCount: parseInt(document.getElementById('displayCount').value),
            keyword: document.getElementById('searchKeyword').value.trim(),
            displayMode: 'enhanced' // 常にHTML形式を使用
          };
          
          // ログレベルフィルタ
          if (document.getElementById('filterINFO').checked) filters.levels.push('INFO');
          if (document.getElementById('filterWARNING').checked) filters.levels.push('WARNING');
          if (document.getElementById('filterERROR').checked) filters.levels.push('ERROR');
          
          // 機能分類フィルタ
          if (document.getElementById('catCSV取込').checked) filters.categories.push('CSV取込');
          if (document.getElementById('catデータ検証').checked) filters.categories.push('データ検証');
          if (document.getElementById('catファイル生成').checked) filters.categories.push('ファイル生成');
          if (document.getElementById('cat自動補完').checked) filters.categories.push('自動補完');
          if (document.getElementById('catマスタ管理').checked) filters.categories.push('マスタ管理');
          if (document.getElementById('catシステム').checked) filters.categories.push('システム');
          
          // 常にHTML形式のログ表示を使用
          google.script.run
            .withSuccessHandler(onFilterSuccess)
            .withFailureHandler(onFilterFailure)
            .showFilteredLogs(filters);
        }
        
        function onFilterSuccess(result) {
          google.script.host.close();
          // 結果は別ダイアログで表示される
        }
        
        function onFilterFailure(error) {
          alert('フィルタ適用エラー: ' + error.message);
        }
      </script>
    `);
    
    const htmlOutput = htmlTemplate.evaluate()
      .setWidth(400)
      .setHeight(500);
    
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'ログフィルタ設定');
  } catch (error) {
    Logger.log('ログフィルタダイアログエラー: ' + error.toString());
    SpreadsheetApp.getUi().alert('エラー', 'ログフィルタダイアログの表示に失敗しました: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}



/**
 * ログ表示（HTML形式・色分け・構造化表示）
 * @param {Object} filters フィルタ条件
 */
function showFilteredLogs(filters) {
  try {
    logSystemActivityEnhanced('showFilteredLogs', `ログ表示開始: レベル[${filters.levels.join(',')}] カテゴリ[${filters.categories.join(',')}] 件数[${filters.displayCount}]`, 'INFO', 'システム');
    
    // PropertiesServiceからログデータを取得
    const properties = PropertiesService.getScriptProperties();
    const storedLogs = properties.getProperty('systemExecutionLogs');
    
    if (!storedLogs) {
      SpreadsheetApp.getUi().alert(
        'フィルタリング結果', 
        'ログエントリがありません。', 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    let logs = JSON.parse(storedLogs);
    
    // フィルタリング処理
    let filteredLogs = logs.filter(log => {
      // ログレベルフィルタ
      if (filters.levels.length > 0 && !filters.levels.includes(log.level)) {
        return false;
      }
      
      // 機能分類フィルタ（既存ログとの互換性を考慮）
      if (filters.categories.length > 0) {
        const logCategory = log.category || 'システム'; // デフォルト値
        if (!filters.categories.includes(logCategory)) {
          return false;
        }
      }
      
      // キーワード検索
      if (filters.keyword && filters.keyword.length > 0) {
        const searchText = `${log.functionName} ${log.message}`.toLowerCase();
        if (!searchText.includes(filters.keyword.toLowerCase())) {
          return false;
        }
      }
      
      return true;
    });
    
    // 表示件数制限
    filteredLogs = filteredLogs.slice(-filters.displayCount);
    
    // 結果表示
    if (filteredLogs.length === 0) {
      SpreadsheetApp.getUi().alert(
        'フィルタリング結果', 
        'フィルタ条件に一致するログが見つかりませんでした。', 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // HTML形式でログを表示
    const htmlContent = createLogHtmlContent(filteredLogs, filters);
    const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
      .setWidth(900)
      .setHeight(600)
      .setTitle('システムログ - フィルタ表示');
      
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'システムログ - フィルタ表示');
    
    logSystemActivityEnhanced('showFilteredLogs', `ログ表示完了: ${filteredLogs.length}件表示`, 'INFO', 'システム');
    
  } catch (error) {
    logSystemActivityEnhanced('showFilteredLogs', `ログ表示エラー: ${error.message}`, 'ERROR', 'システム');
    Logger.log('ログ表示エラー: ' + error.toString());
    throw new Error('ログ表示に失敗しました: ' + error.message);
  }
}

/**
 * ログ表示用のHTMLコンテンツを生成
 * @param {Array} logs ログデータ配列
 * @param {Object} filters フィルタ条件
 * @returns {string} HTMLコンテンツ
 */
function createLogHtmlContent(logs, filters) {
  const filterInfo = `レベル: ${filters.levels.length > 0 ? filters.levels.join(', ') : '全て'} | ` +
                    `カテゴリ: ${filters.categories.length > 0 ? filters.categories.join(', ') : '全て'} | ` +
                    `件数: ${filters.displayCount}件` +
                    (filters.keyword ? ` | キーワード: "${filters.keyword}"` : '');
  
  let html = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <title>システムログ - 拡張表示</title>
  <style>
    body {
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      margin: 10px;
      background-color: #f8f9fa;
    }
    .header {
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      color: white;
      padding: 15px;
      border-radius: 8px;
      margin-bottom: 20px;
      box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    .header h2 {
      margin: 0 0 5px 0;
      font-size: 18px;
    }
    .filter-info {
      font-size: 12px;
      opacity: 0.9;
    }
    .log-container {
      max-height: 450px;
      overflow-y: auto;
      border: 1px solid #ddd;
      border-radius: 8px;
      background: white;
    }
    .log-table {
      width: 100%;
      border-collapse: collapse;
      font-size: 12px;
    }
    .log-table th {
      background: #f1f3f4;
      padding: 10px 8px;
      text-align: left;
      font-weight: 600;
      border-bottom: 2px solid #ddd;
      position: sticky;
      top: 0;
      z-index: 1;
    }
    .log-table td {
      padding: 8px;
      border-bottom: 1px solid #eee;
      vertical-align: top;
    }
    .log-row:hover {
      background-color: #f8f9fa;
    }
    .level-ERROR {
      background-color: #ffebee;
      border-left: 4px solid #f44336;
    }
    .level-WARNING {
      background-color: #fff3e0;
      border-left: 4px solid #ff9800;
    }
    .level-INFO {
      background-color: #e8f5e8;
      border-left: 4px solid #4caf50;
    }
    .level-badge {
      display: inline-block;
      padding: 3px 8px;
      border-radius: 12px;
      font-size: 10px;
      font-weight: bold;
      color: white;
      text-align: center;
      min-width: 45px;
    }
    .level-ERROR .level-badge {
      background-color: #f44336;
    }
    .level-WARNING .level-badge {
      background-color: #ff9800;
    }
    .level-INFO .level-badge {
      background-color: #4caf50;
    }
    .category-badge {
      display: inline-block;
      padding: 2px 6px;
      border-radius: 8px;
      font-size: 9px;
      background-color: #e0e0e0;
      color: #424242;
      margin-left: 5px;
    }
    .timestamp {
      color: #666;
      font-size: 11px;
    }
    .function-name {
      font-weight: 600;
      color: #1976d2;
    }
    .message {
      line-height: 1.4;
    }
    .performance-info {
      color: #666;
      font-size: 10px;
      font-style: italic;
      margin-top: 3px;
    }
    .summary {
      text-align: center;
      padding: 10px;
      background: #e3f2fd;
      margin-bottom: 15px;
      border-radius: 6px;
      font-weight: 600;
      color: #1565c0;
    }
  </style>
</head>
<body>
  <div class="header">
    <h2>📊 システムログ - 拡張表示</h2>
    <div class="filter-info">🔍 ${filterInfo}</div>
  </div>
  
  <div class="summary">
    合計 ${logs.length} 件のログエントリを表示中
  </div>
  
  <div class="log-container">
    <table class="log-table">
      <thead>
        <tr>
          <th style="width: 12%">日時</th>
          <th style="width: 8%">レベル</th>
          <th style="width: 12%">カテゴリ</th>
          <th style="width: 15%">関数名</th>
          <th style="width: 53%">メッセージ・詳細</th>
        </tr>
      </thead>
      <tbody>
  `;
  
  // ログエントリをHTML行として追加
  logs.forEach(log => {
    const timestamp = new Date(log.timestamp).toLocaleString('ja-JP', {
      month: '2-digit',
      day: '2-digit',
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit'
    });
    const category = log.category || 'システム';
    const level = log.level || 'INFO';
    
    // パフォーマンス情報の表示
    let performanceInfo = '';
    if (log.details) {
      const perfItems = [];
      if (log.details.processingTime) perfItems.push(`処理時間: ${log.details.processingTime}ms`);
      if (log.details.recordCount) perfItems.push(`件数: ${log.details.recordCount}`);
      if (log.details.fileName) perfItems.push(`ファイル: ${log.details.fileName}`);
      if (log.details.errorCount) perfItems.push(`エラー: ${log.details.errorCount}`);
      if (log.details.warningCount) perfItems.push(`警告: ${log.details.warningCount}`);
      
      if (perfItems.length > 0) {
        performanceInfo = `<div class="performance-info">📈 ${perfItems.join(' | ')}</div>`;
      }
    }
    
    html += `
        <tr class="log-row level-${level}">
          <td class="timestamp">${timestamp}</td>
          <td><span class="level-badge">${level}</span></td>
          <td>${category}<span class="category-badge">${category}</span></td>
          <td class="function-name">${log.functionName}</td>
          <td>
            <div class="message">${escapeHtml(log.message)}</div>
            ${performanceInfo}
          </td>
        </tr>
    `;
  });
  
  html += `
      </tbody>
    </table>
  </div>
</body>
</html>
  `;
  
  return html;
}

/**
 * HTMLエスケープ処理
 * @param {string} text エスケープする文字列
 * @returns {string} エスケープ後の文字列
 */
function escapeHtml(text) {
  if (typeof text !== 'string') return '';
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#x27;');
}













