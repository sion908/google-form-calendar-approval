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

// フォーム項目のタイトル
const FORM_ITEMS = {
  FULL_NAME: '氏名(フルネーム)',
  AFFILIATION: '所属（企業・大学・NPOなど）',
  EMAIL: 'メールアドレス',
  EVENT_DATE: 'イベント開催の希望日',
  START_TIME: 'イベント開始時刻',
  END_TIME: 'イベント終了時刻',
  EVENT_TITLE: 'イベント名',
  EVENT_DETAILS: 'イベント内容（目的・詳細）',
  PARTICIPANT_COUNT: 'イベント参加予定人数（だいたいで構いません）',
  NOTES: '備考（例：事前準備したいこと）',
  FLYER: 'チラシ（既に用意できている場合は共有お願いします）'
};

// スプレッドシートの列番号（0始まり）
const COLUMNS = {
  TIMESTAMP: 0,           // A: タイムスタンプ
  FULL_NAME: 1,           // B: 氏名
  AFFILIATION: 2,         // C: 所属
  EMAIL: 3,               // D: メールアドレス
  EVENT_DATE: 4,          // E: イベント開催日
  START_TIME: 5,          // F: 開始時刻
  END_TIME: 6,            // G: 終了時刻
  EVENT_TITLE: 7,         // H: イベント名
  EVENT_DETAILS: 8,       // I: イベント内容
  PARTICIPANT_COUNT: 9,   // J: 参加予定人数
  NOTES: 10,              // K: 備考
  FLYER_URL: 11,          // L: チラシURL
  STATUS: 12,             // M: ステータス
  APPROVER_EMAIL: 13,     // N: 承認者メール
  PROCESSED_AT: 14,       // O: 処理日時
  REJECTION_REASON: 15,   // P: 否認理由
  EVENT_ID: 16,           // Q: カレンダーイベントID
  EDIT_TOKEN: 17,         // R: 編集用トークン
  PARENT_ROW: 18          // S: 親レコードの行番号
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
  SCRIPT_URL: 'scriptUrl',
  EMAIL_SUBJECT_PREFIX: 'emailSubjectPrefix',
  ADMIN_EMAIL: 'adminEmail'
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

// ファイルURLからファイルIDを抽出するヘルパー関数
function getFileIdFromUrl(url) {
  if (!url) return null;
  const match = url.match(/[\w-]{25,}(?!.*[\w-]{25,})/);
  return match ? match[0] : null;
}