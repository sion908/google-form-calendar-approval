/**
 * メール送信に関するユーティリティ関数
 */

/**
 * ファイルをパブリックに共有可能なURLに変換する
 * @param {string} fileUrl - ファイルのURL
 * @return {string} 共有可能なURL
 */
function getPublicFileUrl(fileUrl) {
  try {
    if (!fileUrl) return '';
    
    // Google ドライブのファイルIDを抽出
    const fileId = fileUrl.match(/[\w-]{25,}/);
    if (!fileId) return fileUrl; // ファイルIDが取得できない場合は元のURLを返す
    
    const file = DriveApp.getFileById(fileId[0]);
    
    // すでに共有設定があるか確認
    const access = file.getSharingAccess();
    if (access !== DriveApp.Access.ANYONE_WITH_LINK && access !== DriveApp.Access.ANYONE) {
      // 共有設定を「リンクを知っている全員」に変更
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    }
    
    return file.getUrl();
  } catch (error) {
    console.error('ファイルの共有設定中にエラーが発生しました:', error);
    return fileUrl; // エラーが発生した場合は元のURLを返す
  }
}

/**
 * 承認者のメールアドレスを取得する
 * @return {string} 承認者のメールアドレス
 */
function getApproverEmail() {
  return PropertiesService.getScriptProperties().getProperty(PROPERTY_KEYS.APPROVER_EMAIL);
}

/**
 * スクリプトの公開URLを取得する
 * @return {string} スクリプトの公開URL
 */
function getScriptUrl() {
  return PropertiesService.getScriptProperties().getProperty(PROPERTY_KEYS.SCRIPT_URL);
}

/**
 * 承認依頼メールを送信する
 * @param {Object} eventData - イベントデータ
 * @param {string} eventData.title - イベント名
 * @param {Date} eventData.startTime - 開始日時
 * @param {Date} eventData.endTime - 終了日時
 * @param {string} eventData.eventDetails - イベント内容
 * @param {string} eventData.applicantName - 申請者名
 * @param {string} eventData.applicantEmail - 申請者のメールアドレス
 * @param {string} eventData.affiliation - 所属
 * @param {number} eventData.participantCount - 参加予定人数
 * @param {string} eventData.notes - 備考
 * @param {string} eventData.flyerUrl - チラシURL
 * @param {number} row - スプレッドシートの行番号
 * @return {Object} 送信結果
 */
function sendApprovalRequestEmail(eventData, row) {
  try {
    const approverEmail = getApproverEmail();
    if (!approverEmail) {
      throw new Error('承認者のメールアドレスが設定されていません。');
    }

    const scriptUrl = getScriptUrl();
    const approvalUrl = `${scriptUrl}?action=${ACTIONS.APPROVE}&row=${row}`;
    const rejectionUrl = `${scriptUrl}?action=${ACTIONS.REJECT}&row=${row}`;
    
    const subject = EMAIL.APPROVAL_REQUEST_SUBJECT;
    
    // チラシの表示（URLがある場合）
    const flyerText = eventData.flyerUrl 
      ? `【チラシ】
${eventData.flyerUrl}

` 
      : '';
    
    const body = `
以下の予定登録の承認をお願いします。

【イベント名】
${eventData.title}

【日時】
${formatDateTime(eventData.startTime)} 〜 ${formatDateTime(eventData.endTime)}

【イベント内容】
${eventData.eventDetails || '特になし'}

【申請者情報】
・氏名: ${eventData.applicantName}
・メール: ${eventData.applicantEmail}
・所属: ${eventData.affiliation || '未設定'}
・参加予定人数: ${eventData.participantCount || '未設定'}名

【備考】
${eventData.notes || '特になし'}

${flyerText}
以下のリンクから承認または拒否を行ってください。

✅ 承認する: ${approvalUrl}
❌ 拒否する: ${rejectionUrl}

※このリンクは30日間有効です。
`;

    GmailApp.sendEmail(approverEmail, subject, body, {
      name: 'カレンダー承認システム',
      replyTo: eventData.applicantEmail
    });

    return { success: true, message: '承認依頼メールを送信しました。' };
  } catch (error) {
    console.error('承認依頼メールの送信中にエラーが発生しました:', error);
    return { success: false, message: `メール送信エラー: ${error.message}` };
  }
}

/**
 * 承認完了メールを送信する
 * @param {Object} eventData - イベントデータ
 * @param {string} calendarEventUrl - カレンダーイベントのURL
 * @return {Object} 送信結果
 */
function sendApprovalNoticeEmail(eventData, calendarEventUrl) {
  try {
    const editToken = Utilities.getUuid();
    const editUrl = `${getScriptUrl()}?action=${ACTIONS.EDIT}&token=${editToken}`;
    
    const subject = EMAIL.APPROVAL_NOTICE_SUBJECT;
    const body = `
以下の予定が承認され、カレンダーに登録されました。

【イベント名】
${eventData.title}

【日時】
${formatDateTime(eventData.startTime)} 〜 ${formatDateTime(eventData.endTime)}

【場所】
${eventData.location || '未設定'}

【説明】
${eventData.description || '特になし'}

【カレンダーで確認】
${calendarEventUrl}

【予定を編集する】
${editUrl}

※このリンクは30日間有効です。
`;

    GmailApp.sendEmail(eventData.applicantEmail, subject, body, {
      name: 'カレンダー承認システム',
      replyTo: getApproverEmail()
    });

    return { 
      success: true, 
      message: '承認完了メールを送信しました。',
      editToken: editToken
    };
  } catch (error) {
    console.error('承認完了メールの送信中にエラーが発生しました:', error);
    return { success: false, message: `メール送信エラー: ${error.message}` };
  }
}

/**
 * 拒否通知メールを送信する
 * @param {Object} eventData - イベントデータ
 * @param {string} rejectionReason - 拒否理由
 * @param {string} editUrl - 編集用URL
 * @return {Object} 送信結果
 */
function sendRejectionNoticeEmail(eventData, rejectionReason, editUrl = '') {
  try {
    const subject = EMAIL.REJECTION_NOTICE_SUBJECT;
    let body = `
以下の予定登録申請は承認されませんでした。

【イベント名】
${eventData.title}

【日時】
${formatDateTime(eventData.startTime)} 〜 ${formatDateTime(eventData.endTime)}

【場所】
${eventData.location || '未設定'}

【説明】
${eventData.description || '特になし'}

【拒否理由】
${rejectionReason || '理由は指定されていません。'}
`;

    if (editUrl) {
      body += `
【申請を修正する】
${editUrl}

※このリンクは30日間有効です。
`;
    }

    GmailApp.sendEmail(eventData.applicantEmail, subject, body, {
      name: 'カレンダー承認システム',
      replyTo: getApproverEmail()
    });

    return { success: true, message: '拒否通知メールを送信しました。' };
  } catch (error) {
    console.error('拒否通知メールの送信中にエラーが発生しました:', error);
    return { success: false, message: `メール送信エラー: ${error.message}` };
  }
}

/**
 * 日時をフォーマットする
 * @param {Date} date - フォーマットする日時
 * @return {string} フォーマットされた日時文字列
 */
function formatDateTime(date) {
  if (!(date instanceof Date)) {
    date = new Date(date);
  }
  
  // 日本時間に変換
  const options = {
    timeZone: 'Asia/Tokyo',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    hour12: false
  };
  
  return date.toLocaleString('ja-JP', options);
}