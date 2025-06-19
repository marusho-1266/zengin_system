/**
 * 全銀協フォーマット対応データ作成システム 定数定義
 * スプレッドシートのセル位置、項目名、設定値等を一元管理
 */

// ===== シート名定数 =====
const SHEET_NAMES = {
  CLIENT_INFO: '振込依頼人情報',
  TRANSFER_DATA: '振込データ',
  BANK_MASTER: '金融機関マスタ'
};

// ===== 振込依頼人情報シート定数 =====
const CLIENT_INFO_CELLS = {
  CLIENT_CODE: 'B2',
  CLIENT_NAME: 'B3',
  BANK_CODE: 'B4',
  BANK_NAME: 'B5',
  BRANCH_CODE: 'B6',
  BRANCH_NAME: 'B7',
  ACCOUNT_TYPE: 'B8',
  ACCOUNT_NUMBER: 'B9',
  CATEGORY_CODE: 'B10',
  FILE_EXTENSION: 'B11'
};

const CLIENT_INFO_LABELS = {
  CLIENT_CODE: '委託者コード',
  CLIENT_NAME: '委託者名',
  BANK_CODE: '取引銀行コード',
  BANK_NAME: '取引銀行名',
  BRANCH_CODE: '取引支店コード',
  BRANCH_NAME: '取引支店名',
  ACCOUNT_TYPE: '預金種目',
  ACCOUNT_NUMBER: '口座番号',
  CATEGORY_CODE: '種別コード',
  FILE_EXTENSION: '出力ファイル拡張子'
};

// ===== 振込データシート定数 =====
const TRANSFER_DATA_COLUMNS = {
  BANK_CODE: 1,        // A列
  BANK_NAME: 2,        // B列
  BRANCH_CODE: 3,      // C列
  BRANCH_NAME: 4,      // D列
  ACCOUNT_TYPE: 5,     // E列
  ACCOUNT_NUMBER: 6,   // F列
  RECIPIENT_NAME: 7,   // G列
  AMOUNT: 8,           // H列
  CUSTOMER_CODE: 9,    // I列
  IDENTIFICATION: 10,  // J列
  EDI_INFO: 11        // K列
};

const TRANSFER_DATA_HEADERS = {
  BANK_CODE: '銀行コード',
  BANK_NAME: '銀行名',
  BRANCH_CODE: '支店コード',
  BRANCH_NAME: '支店名',
  ACCOUNT_TYPE: '預金種目',
  ACCOUNT_NUMBER: '口座番号',
  RECIPIENT_NAME: '受取人名',
  AMOUNT: '振込金額',
  CUSTOMER_CODE: '顧客コード',
  IDENTIFICATION: '識別表示',
  EDI_INFO: 'EDI情報'
};

// ===== 金融機関マスタシート定数 =====
const BANK_MASTER_COLUMNS = {
  BANK_CODE: 1,       // A列
  BANK_NAME: 2,       // B列
  BRANCH_CODE: 3,     // C列
  BRANCH_NAME: 4,     // D列
  UPDATE_DATE: 5,     // E列
  STATUS: 6           // F列
};

const BANK_MASTER_HEADERS = {
  BANK_CODE: '銀行コード',
  BANK_NAME: '銀行名',
  BRANCH_CODE: '支店コード',
  BRANCH_NAME: '支店名',
  UPDATE_DATE: '更新日',
  STATUS: '状態'
};

// ===== 選択肢定数 =====
const ACCOUNT_TYPES = {
  ORDINARY: { code: '1', name: '普通' },
  CHECKING: { code: '2', name: '当座' }
  // 注意：貯蓄(4)は全銀協フォーマットでは対象外のため削除
};

const CATEGORY_CODES = {
  SALARY: { code: '11', name: '給与' },
  BONUS: { code: '12', name: '賞与' }
  // 注意：配当金等(31)は全銀協フォーマットでは非対応のため削除
};

const FILE_EXTENSIONS = ['.dat', '.txt', '.fb'];

const STATUS_OPTIONS = ['有効', '無効'];

const IDENTIFICATION_CODES = {
  SALARY: 'Y',
  BONUS: 'B'
};

// ===== データ検証定数 =====
const VALIDATION_RULES = {
  BANK_CODE_LENGTH: 4,
  BRANCH_CODE_LENGTH: 3,
  ACCOUNT_NUMBER_MAX_LENGTH: 7,
  CLIENT_CODE_MAX_LENGTH: 10,
  CLIENT_NAME_MAX_LENGTH: 40,  // 英数カナ（全銀協準拠）
  BANK_NAME_MAX_LENGTH: 15,    // 英数カナ（全銀協準拠）
  BRANCH_NAME_MAX_LENGTH: 15,  // 英数カナ（全銀協準拠）
  RECIPIENT_NAME_MAX_LENGTH: 30, // 英数カナ（全銀協準拠）
  CUSTOMER_CODE_MAX_LENGTH: 10,
  EDI_INFO_MAX_LENGTH: 20,
  MAX_AMOUNT: 99999999,        // 最大振込金額
  MAX_RECORDS: 1000            // 最大処理件数
};

// ===== 全銀協フォーマット定数 =====
const ZENGIN_FORMAT = {
  RECORD_LENGTH: 120,          // レコード長（バイト）
  ENCODING: 'JIS',             // 文字コード（全銀協標準）
  HEADER_TYPE: '1',           // ヘッダレコード種別
  DATA_TYPE: '2',             // データレコード種別
  TRAILER_TYPE: '8',          // トレーラレコード種別
  END_TYPE: '9'               // エンドレコード種別
};

// ===== メッセージ定数 =====
const MESSAGES = {
  SUCCESS: {
    DATA_IMPORTED: 'データの取込が完了しました。',
    FILE_GENERATED: '全銀協フォーマットファイルの生成が完了しました。',
    AUTO_COMPLETE: '自動補完が完了しました。'
  },
  ERROR: {
    INVALID_FORMAT: 'データ形式が正しくありません。',
    BANK_NOT_FOUND: '金融機関コードが見つかりません。',
    REQUIRED_FIELD: '必須項目が入力されていません。',
    DUPLICATE_DATA: '重複するデータが見つかりました。',
    FILE_TOO_LARGE: 'ファイルサイズが制限を超えています。'
  },
  WARNING: {
    DATA_WILL_BE_OVERWRITTEN: '既存のデータが上書きされます。続行しますか？',
    INVALID_CHARACTERS: '使用できない文字が含まれています。'
  }
};

// ===== キャッシュ定数 =====
const CACHE_CONFIG = {
  EXPIRATION_TIME: 300000,     // 5分 (ミリ秒)
  BANK_MASTER_KEY: 'bankMasterData'
};

// ===== UI設定定数 =====
const UI_CONFIG = {
  DIALOG_WIDTH: 400,
  DIALOG_HEIGHT: 300,
  MAX_DISPLAY_ROWS: 1000
}; 