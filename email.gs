/**
 * メール送信に関するユーティリティ関数
 */

/**
 * 承認依頼メールを送信する
 * @param {Object} eventData - イベントデータ
 * @param {string} eventData.applicantEmail - 申請者のメールアドレス
 * @param {string} eventData.title - イベントタイトル
 * @param {Date} eventData.startTime - 開始日時
 * @param {Date} eventData.endTime - 終了日時
 * @param {string} eventData.location - 場所
 * @param {string} eventData.description - 説明
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
    const body = `
以下の予定登録の承認をお願いします。

【イベント名】
${eventData.title}

【日時】
${formatDateTime(eventData.startTime)} 〜 ${formatDateTime(eventData.endTime)}

【場所】
${eventData.location || '未設定'}

【説明】
${eventData.description || '特になし'}

【申請者】
${eventData.applicantEmail}

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