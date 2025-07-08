/**
 * アプリケーション全体で使用する定数を定義
 */

// ステータス定数
const STATUS = {
  PENDING: '承認待ち',
  APPROVED: '承認済み',
  REJECTED: '拒否済み',
  UPDATED: '更新済み',
  CANCELLED: 'キャンセル済み'
};

// スプレッドシートの列番号（0始まり）
const COLUMNS = {
  TIMESTAMP: 0,        // A: タイムスタンプ
  APPLICANT_EMAIL: 1,  // B: 申請者メール
  EVENT_TITLE: 2,      // C: イベントタイトル
  START_TIME: 3,       // D: 開始日時
  END_TIME: 4,         // E: 終了日時
  LOCATION: 5,         // F: 場所
  DESCRIPTION: 6,      // G: 説明
  STATUS: 7,           // H: ステータス
  APPROVER_EMAIL: 8,   // I: 承認者メール
  PROCESSED_AT: 9,     // J: 処理日時
  REJECTION_REASON: 10,// K: 否認理由
  EVENT_ID: 11,        // L: カレンダーイベントID
  EDIT_TOKEN: 12,      // M: 編集用トークン
  PARENT_ROW: 13       // N: 親レコードの行番号（更新元）
};

// メール設定
const EMAIL = {
  APPROVAL_REQUEST_SUBJECT: '【承認依頼】カレンダー登録申請',
  APPROVAL_NOTICE_SUBJECT: '【承認完了】カレンダー登録が承認されました',
  REJECTION_NOTICE_SUBJECT: '【差し戻し】カレンダー登録申請について',
  EDIT_LINK_EXPIRY_DAYS: 30  // 編集リンクの有効期限（日）
};

// スクリプトプロパティのキー
const PROPERTY_KEYS = {
  APPROVER_EMAIL: 'approverEmail',
  CALENDAR_ID: 'calendarId',
  SCRIPT_URL: 'scriptUrl'
};

// アクションタイプ
const ACTIONS = {
  APPROVE: 'approve',
  REJECT: 'reject',
  EDIT: 'edit',
  CANCEL: 'cancel'
};

// エラーメッセージ
const ERROR_MESSAGES = {
  INVALID_REQUEST: '無効なリクエストです。',
  ALREADY_PROCESSED: 'このリクエストは既に処理済みです。',
  LINK_EXPIRED: 'このリンクは有効期限が切れています。',
  NOT_AUTHORIZED: 'この操作を実行する権限がありません。',
  EVENT_NOT_FOUND: '指定されたイベントが見つかりません。'
};