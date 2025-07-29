/**
 * @OnlyCurrentDoc
 */

// グローバル変数
let _activeSpreadsheet = null;

/**
 * アクティブなスプレッドシートを取得する
 * @return {GoogleAppsScript.Spreadsheet.Spreadsheet} スプレッドシートオブジェクト
 */
function getActiveSpreadsheet() {
  if (!_activeSpreadsheet) {
    _activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  }
  return _activeSpreadsheet;
}

/**
 * WebアプリケーションとしてGETリクエストが来た際に実行される関数
 * @param {Object} e - イベントオブジェクト
 * @return {GoogleAppsScript.HTML.HtmlOutput} HTMLレスポンス
 */
function doGet(e) {
  try {
    const action = e.parameter.action;
    const row = parseInt(e.parameter.row, 10);
    const token = e.parameter.token;
    
    // アクションに応じた処理を実行
    switch (action) {
      case ACTIONS.APPROVE:
        return handleApproveAction(row);
        
      case ACTIONS.REJECT:
        return handleRejectAction(row);
        
      case ACTIONS.EDIT:
        return handleEditAction(row, token);
        
      case ACTIONS.CANCEL:
        return handleCancelAction(row, token);
        
      default:
        return createErrorResponse(ERROR_MESSAGES.INVALID_REQUEST, 400);
    }
  } catch (error) {
    console.error('doGetでエラーが発生しました:', error);
    return createErrorResponse(error.message || '内部エラーが発生しました。', 500);
  }
}

/**
 * WebアプリケーションとしてPOSTリクエストが来た際に実行される関数
 * @param {Object} e - イベントオブジェクト
 * @return {GoogleAppsScript.HTML.HtmlOutput} HTMLレスポンス
 */
function doPost(e) {
  try {
    const params = e.parameter;
    const action = params.action;
    const row = parseInt(params.row, 10);
    const token = params.token;
    
    // アクションに応じた処理を実行
    switch (action) {
      case ACTIONS.REJECT:
        const rejectionReason = params.rejectionReason || '';
        return handleRejectSubmit(row, rejectionReason);
        
      case ACTIONS.EDIT:
        return handleEditSubmit(row, params, token);
        
      default:
        return createErrorResponse(ERROR_MESSAGES.INVALID_REQUEST, 400);
    }
  } catch (error) {
    console.error('doPostでエラーが発生しました:', error);
    return createErrorResponse(error.message || '内部エラーが発生しました。', 500);
  }
}

/**
 * 承認アクションを処理する
 * @param {number} row - 行番号
 * @return {GoogleAppsScript.HTML.HtmlOutput} HTMLレスポンス
 */
function handleApproveAction(row) {
  try {
    if (!row || isNaN(row)) {
      return createErrorResponse('無効なリクエストパラメータです。', 400);
    }
    
    const result = approveEvent(row);
    
    if (result.success) {
      const htmlOutput = HtmlService.createHtmlOutput(`
        <!DOCTYPE html>
        <html>
          <head>
            <base target="_top">
            <meta charset="UTF-8">
            <title>予定を承認しました</title>
            <style>
              body { 
                font-family: 'Arial', 'Meiryo', sans-serif; 
                line-height: 1.6; 
                max-width: 600px; 
                margin: 0 auto; 
                padding: 20px; 
              }
              .success { 
                color: #155724; 
                background-color: #d4edda; 
                border: 1px solid #c3e6cb; 
                padding: 20px; 
                border-radius: 4px; 
                margin: 20px 0;
              }
              .btn { 
                display: inline-block; 
                padding: 10px 20px; 
                background-color: #28a745; 
                color: white; 
                text-decoration: none; 
                border-radius: 4px; 
                margin-top: 15px;
                font-weight: bold;
              }
              .btn:hover { 
                background-color: #218838;
                text-decoration: none;
              }
              h2 { 
                margin-top: 0; 
                color: #155724;
              }
            </style>
          </head>
          <body>
            <div class="success">
              <h2>✅ 予定を承認しました</h2>
              <p>${result.message}</p>
              ${result.eventUrl ? `<p><a href="${result.eventUrl}" target="_blank" class="btn">カレンダーで確認する</a></p>` : ''}
            </div>
            <p><a href="${ScriptApp.getService().getUrl()}">トップに戻る</a></p>
          </body>
        </html>
      `);
      
      return htmlOutput;
    } else {
      return createErrorResponse(result.message, 500);
    }
  } catch (error) {
    console.error('承認処理中にエラーが発生しました:', error);
    return createErrorResponse('承認処理中にエラーが発生しました。', 500);
  }
}

/**
 * 拒否アクションを処理する
 * @param {number} row - 行番号
 * @return {GoogleAppsScript.HTML.HtmlOutput} HTMLレスポンス
 */
function handleRejectAction(row) {
  if (!row || isNaN(row)) {
    return createErrorResponse('無効なリクエストパラメータです。', 400);
  }
  
  // スプレッドシートからイベント情報を取得
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getRange(row, 1, 1, COLUMNS.PARENT_ROW + 2).getValues()[0];
  
  // テンプレートに渡すデータを準備
  const template = HtmlService.createTemplateFromFile('rejection_form_template');
  template.eventTitle = data[COLUMNS.EVENT_TITLE] || '（タイトルなし）';
  template.eventDate = formatDateTime(data[COLUMNS.START_TIME]) + ' 〜 ' + formatDateTime(data[COLUMNS.END_TIME]);
  template.applicantEmail = data[COLUMNS.APPLICANT_EMAIL] || '（メールアドレスなし）';
  template.row = row; // 行番号もテンプレートに渡す
  
  // 拒否理由入力フォームを表示
  return template
    .evaluate()
    .setTitle('予定を拒否 - 理由を入力')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * 拒否フォームの送信を処理する
 * @param {number} row - 行番号
 * @param {string} rejectionReason - 拒否理由
 * @return {GoogleAppsScript.HTML.HtmlOutput} HTMLレスポンス
 */
function handleRejectSubmit(row, rejectionReason) {
  if (!row || isNaN(row) || !rejectionReason) {
    return createErrorResponse('無効なリクエストパラメータです。', 400);
  }
  
  const result = rejectEvent(row, rejectionReason);
  
  if (result.success) {
    return HtmlService.createHtmlOutput(`
      <!DOCTYPE html>
      <html>
        <head>
          <base target="_top">
          <meta charset="UTF-8">
          <title>予定を拒否しました</title>
          <style>
            body { font-family: Arial, sans-serif; line-height: 1.6; max-width: 600px; margin: 0 auto; padding: 20px; }
            .info { color: #0c5460; background-color: #d1ecf1; border: 1px solid #bee5eb; padding: 15px; border-radius: 4px; }
          </style>
        </head>
        <body>
          <div class="info">
            <h2>⛔ 予定を拒否しました</h2>
            <p>申請者に通知が送信されました。</p>
          </div>
        </body>
      </html>
    `);
  } else {
    return createErrorResponse(result.message, 500);
  }
}

/**
 * 編集アクションを処理する
 * @param {number} row - 行番号
 * @param {string} token - 編集トークン
 * @return {GoogleAppsScript.HTML.HtmlOutput} HTMLレスポンス
 */
function handleEditAction(row, token) {
  // TODO: 編集フォームを表示する実装
  return createErrorResponse('編集機能は現在準備中です。', 501);
}

/**
 * 編集フォームの送信を処理する
 * @param {number} row - 行番号
 * @param {Object} params - フォームパラメータ
 * @param {string} token - 編集トークン
 * @return {GoogleAppsScript.HTML.HtmlOutput} HTMLレスポンス
 */
function handleEditSubmit(row, params, token) {
  // TODO: 編集内容を処理する実装
  return createErrorResponse('編集機能は現在準備中です。', 501);
}

/**
 * キャンセルアクションを処理する
 * @param {number} row - 行番号
 * @param {string} token - 編集トークン
 * @return {GoogleAppsScript.HTML.HtmlOutput} HTMLレスポンス
 */
function handleCancelAction(row, token) {
  // TODO: キャンセル処理を実装
  return createErrorResponse('キャンセル機能は現在準備中です。', 501);
}

/**
 * エラーレスポンスを作成する
 * @param {string} message - エラーメッセージ
 * @param {number} statusCode - HTTPステータスコード
 * @return {GoogleAppsScript.HTML.HtmlOutput} HTMLレスポンス
 */
function createErrorResponse(message, statusCode = 500) {
  const htmlOutput = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <meta charset="UTF-8">
        <title>エラー</title>
        <style>
          body { 
            font-family: 'Arial', 'Meiryo', sans-serif; 
            line-height: 1.6; 
            margin: 0; 
            padding: 20px; 
            max-width: 800px;
            margin: 0 auto;
          }
          .error { 
            color: #721c24; 
            background-color: #f8d7da; 
            border: 1px solid #f5c6cb; 
            padding: 20px; 
            border-radius: 4px; 
            margin: 20px 0;
          }
          h2 { margin-top: 0; }
        </style>
      </head>
      <body>
        <div class="error">
          <h2>⚠️ エラーが発生しました</h2>
          <p>${message}</p>
        </div>
        <p><a href="${ScriptApp.getService().getUrl()}">トップに戻る</a></p>
      </body>
    </html>
  `);
  
  // ステータスコードを設定
  if (statusCode) {
    const response = {}
    response[1] = statusCode;
    response[0] = htmlOutput;
    return response;
  }
  
  return htmlOutput;
}

/**
 * スクリプトのプロパティを設定する（管理者用）
 * @param {Object} properties - 設定するプロパティ
 * @return {Object} 処理結果
 */
function setScriptProperties(properties) {
  try {
    // 管理者のみ実行可能
    if (!Session.getEffectiveUser().getEmail().endsWith('@your-domain.com')) {
      throw new Error('この操作を実行する権限がありません。');
    }
    
    setProperties(properties);
    return { success: true, message: 'プロパティを更新しました。' };
  } catch (error) {
    console.error('プロパティの設定中にエラーが発生しました:', error);
    return { success: false, message: `プロパティの設定に失敗しました: ${error.message}` };
  }
}

/**
 * スクリプトのプロパティを取得する（管理者用）
 * @return {Object} プロパティ一覧
 */
function getScriptProperties() {
  try {
    // 管理者のみ実行可能
    if (!Session.getEffectiveUser().getEmail().endsWith('@your-domain.com')) {
      throw new Error('この操作を実行する権限がありません。');
    }
    
    const properties = PropertiesService.getScriptProperties().getProperties();
    return { success: true, data: properties };
  } catch (error) {
    console.error('プロパティの取得中にエラーが発生しました:', error);
    return { success: false, message: `プロパティの取得に失敗しました: ${error.message}` };
  }
}
