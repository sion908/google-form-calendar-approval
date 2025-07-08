/**
 * 承認フローを管理する関数
 */

/**
 * フォーム送信時に実行される関数
 * @param {GoogleAppsScript.Events.SheetsOnFormSubmit} e - フォーム送信イベント
 */
function onFormSubmit(e) {
  try {
    // フォーム回答を取得
    const formResponse = e.response;
    if (!formResponse) {
      throw new Error('フォームの回答を取得できませんでした。');
    }

    // スプレッドシートとアクティブシートを取得
    const sheet = e.range.getSheet();
    const row = e.range.getRow();
    
    // フォーム回答をオブジェクトに変換
    const formData = {
      timestamp: new Date(),
      applicantEmail: formResponse.getRespondentEmail() || 'guest@example.com',
      title: formResponse.getItemResponses().find(r => r.getItem().getTitle() === 'イベント名')?.getResponse() || '（無題）',
      startTime: formResponse.getItemResponses().find(r => r.getItem().getTitle() === '開始日時')?.getResponse(),
      endTime: formResponse.getItemResponses().find(r => r.getItem().getTitle() === '終了日時')?.getResponse(),
      location: formResponse.getItemResponses().find(r => r.getItem().getTitle() === '場所')?.getResponse() || '',
      description: formResponse.getItemResponses().find(r => r.getItem().getTitle() === '説明')?.getResponse() || ''
    };

    // スプレッドシートにステータスを記録
    sheet.getRange(row, COLUMNS.STATUS + 1).setValue(STATUS.PENDING);
    sheet.getRange(row, COLUMNS.APPLICANT_EMAIL + 1).setValue(formData.applicantEmail);
    
    // 編集用トークンを生成して保存
    const editToken = Utilities.getUuid();
    sheet.getRange(row, COLUMNS.EDIT_TOKEN + 1).setValue(editToken);
    
    // 承認依頼メールを送信
    const result = sendApprovalRequestEmail(formData, row);
    
    if (!result.success) {
      throw new Error(result.message);
    }
    
    console.log(`承認依頼を送信しました。行: ${row}, 申請者: ${formData.applicantEmail}`);
  } catch (error) {
    console.error('フォーム送信処理でエラーが発生しました:', error);
    // エラーをスプレッドシートに記録
    if (e && e.range) {
      const sheet = e.range.getSheet();
      const row = e.range.getRow();
      sheet.getRange(row, COLUMNS.STATUS + 1).setValue(`エラー: ${error.message}`);
    }
  }
}

/**
 * イベントを承認する
 * @param {number} row - スプレッドシートの行番号
 * @return {Object} 処理結果
 */
function approveEvent(row) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data = sheet.getRange(row, 1, 1, COLUMNS.PARENT_ROW + 2).getValues()[0];
    
    // 既に処理済みか確認
    if (data[COLUMNS.STATUS] !== STATUS.PENDING) {
      return { 
        success: false, 
        message: ERROR_MESSAGES.ALREADY_PROCESSED 
      };
    }
    
    // イベントデータを準備
    const eventData = {
      title: data[COLUMNS.EVENT_TITLE],
      startTime: new Date(data[COLUMNS.START_TIME]),
      endTime: new Date(data[COLUMNS.END_TIME]),
      location: data[COLUMNS.LOCATION],
      description: data[COLUMNS.DESCRIPTION],
      organizerEmail: data[COLUMNS.APPLICANT_EMAIL]
    };
    
    // カレンダーにイベントを作成
    const calendarResult = createCalendarEvent(eventData, [data[COLUMNS.APPLICANT_EMAIL]]);
    
    if (!calendarResult.success) {
      throw new Error(calendarResult.message);
    }
    
    // スプレッドシートを更新
    sheet.getRange(row, COLUMNS.STATUS + 1).setValue(STATUS.APPROVED);
    sheet.getRange(row, COLUMNS.PROCESSED_AT + 1).setValue(new Date());
    sheet.getRange(row, COLUMNS.EVENT_ID + 1).setValue(calendarResult.eventId);
    
    // 承認完了メールを送信
    const emailResult = sendApprovalNoticeEmail(
      eventData, 
      calendarResult.eventUrl
    );
    
    if (emailResult.success && emailResult.editToken) {
      sheet.getRange(row, COLUMNS.EDIT_TOKEN + 1).setValue(emailResult.editToken);
    }
    
    return { 
      success: true, 
      message: 'イベントを承認し、カレンダーに登録しました。',
      eventUrl: calendarResult.eventUrl
    };
  } catch (error) {
    console.error('イベントの承認中にエラーが発生しました:', error);
    return { 
      success: false, 
      message: `イベントの承認に失敗しました: ${error.message}` 
    };
  }
}

/**
 * イベントを拒否する
 * @param {number} row - スプレッドシートの行番号
 * @param {string} rejectionReason - 拒否理由
 * @return {Object} 処理結果
 */
function rejectEvent(row, rejectionReason = '') {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data = sheet.getRange(row, 1, 1, COLUMNS.PARENT_ROW + 2).getValues()[0];
    
    // 既に処理済みか確認
    if (data[COLUMNS.STATUS] !== STATUS.PENDING) {
      return { 
        success: false, 
        message: ERROR_MESSAGES.ALREADY_PROCESSED 
      };
    }
    
    // 編集用トークンを生成
    const editToken = Utilities.getUuid();
    const editUrl = `${getScriptUrl()}?action=${ACTIONS.EDIT}&token=${editToken}`;
    
    // スプレッドシートを更新
    sheet.getRange(row, COLUMNS.STATUS + 1).setValue(STATUS.REJECTED);
    sheet.getRange(row, COLUMNS.PROCESSED_AT + 1).setValue(new Date());
    sheet.getRange(row, COLUMNS.REJECTION_REASON + 1).setValue(rejectionReason);
    sheet.getRange(row, COLUMNS.EDIT_TOKEN + 1).setValue(editToken);
    
    // イベントデータを準備
    const eventData = {
      title: data[COLUMNS.EVENT_TITLE],
      startTime: new Date(data[COLUMNS.START_TIME]),
      endTime: new Date(data[COLUMNS.END_TIME]),
      location: data[COLUMNS.LOCATION],
      description: data[COLUMNS.DESCRIPTION],
      applicantEmail: data[COLUMNS.APPLICANT_EMAIL]
    };
    
    // 拒否通知メールを送信
    sendRejectionNoticeEmail(eventData, rejectionReason, editUrl);
    
    return { 
      success: true, 
      message: 'イベントを拒否しました。',
      editUrl: editUrl
    };
  } catch (error) {
    console.error('イベントの拒否中にエラーが発生しました:', error);
    return { 
      success: false, 
      message: `イベントの拒否に失敗しました: ${error.message}` 
    };
  }
}

/**
 * イベントをキャンセルする
 * @param {number} row - スプレッドシートの行番号
 * @return {Object} 処理結果
 */
function cancelEvent(row) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data = sheet.getRange(row, 1, 1, COLUMNS.PARENT_ROW + 2).getValues()[0];
    
    // 既にキャンセル済みか確認
    if (data[COLUMNS.STATUS] === STATUS.CANCELLED) {
      return { 
        success: false, 
        message: 'このイベントは既にキャンセル済みです。' 
      };
    }
    
    // カレンダーからイベントを削除
    if (data[COLUMNS.EVENT_ID]) {
      const deleteResult = deleteCalendarEvent(data[COLUMNS.EVENT_ID]);
      if (!deleteResult.success) {
        console.warn(deleteResult.message);
      }
    }
    
    // ステータスを更新
    sheet.getRange(row, COLUMNS.STATUS + 1).setValue(STATUS.CANCELLED);
    sheet.getRange(row, COLUMNS.PROCESSED_AT + 1).setValue(new Date());
    
    return { 
      success: true, 
      message: 'イベントをキャンセルしました。' 
    };
  } catch (error) {
    console.error('イベントのキャンセル中にエラーが発生しました:', error);
    return { 
      success: false, 
      message: `イベントのキャンセルに失敗しました: ${error.message}` 
    };
  }
}