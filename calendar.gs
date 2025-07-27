/**
 * Googleカレンダー操作に関するユーティリティ関数
 */

/**
 * カレンダーにイベントを作成する
 * @param {Object} eventData - イベントデータ
 * @param {string} eventData.title - イベントタイトル
 * @param {Date} eventData.startTime - 開始日時
 * @param {Date} eventData.endTime - 終了日時
 * @param {string} eventData.location - 場所
 * @param {string} eventData.description - 説明
 * @param {string} eventData.organizerEmail - 主催者メール
 * @param {Array<string>} [attendees=[]] - 参加者のメールアドレス配列
 * @return {Object} 作成されたイベントの情報
 */
function createCalendarEvent(eventData, attendees = []) {
  try {
    const calendarId = getCalendarId();
    const calendar = CalendarApp.getCalendarById(calendarId);
    
    if (!calendar) {
      throw new Error('カレンダーが見つかりません。カレンダーIDを確認してください。');
    }

    // イベントを作成
    const event = calendar.createEvent(
      eventData.title,
      new Date(eventData.startTime),
      new Date(eventData.endTime),
      {
        description: eventData.description || '',
        location: eventData.location || '',
        guests: attendees.join(','),
        sendInvites: true
      }
    );

    // 主催者を設定
    if (eventData.organizerEmail) {
      try {
        event.setGuestStatus(eventData.organizerEmail, CalendarApp.GuestStatus.OWNER);
      } catch (e) {
        console.warn('主催者の設定に失敗しました:', e.message);
      }
    }

    // イベントURLを構築
    const eventId = event.getId().replace('@google.com', '');
    const eventUrl = `https://calendar.google.com/calendar/event?eid=${encodeURIComponent(event.getId())}&ctz=Asia/Tokyo`;
    
    return {
      success: true,
      eventId: eventId,
      eventUrl: eventUrl,
      calendarName: calendar.getName(),
      message: 'カレンダーにイベントを作成しました。'
    };
  } catch (error) {
    console.error('カレンダーイベントの作成中にエラーが発生しました:', error);
    return {
      success: false,
      message: `カレンダーイベントの作成に失敗しました: ${error.message}`
    };
  }
}

/**
 * 既存のカレンダーイベントを更新する
 * @param {string} eventId - 更新するイベントID
 * @param {Object} eventData - 更新するイベントデータ
 * @param {string} [eventData.title] - イベントタイトル
 * @param {Date} [eventData.startTime] - 開始日時
 * @param {Date} [eventData.endTime] - 終了日時
 * @param {string} [eventData.location] - 場所
 * @param {string} [eventData.description] - 説明
 * @return {Object} 更新結果
 */
function updateCalendarEvent(eventId, eventData) {
  try {
    const calendarId = getCalendarId();
    const calendar = CalendarApp.getCalendarById(calendarId);
    
    if (!calendar) {
      throw new Error('カレンダーが見つかりません。カレンダーIDを確認してください。');
    }

    // イベントを取得
    const event = calendar.getEventById(eventId);
    if (!event) {
      throw new Error('更新するイベントが見つかりません。');
    }

    // イベントを更新
    if (eventData.title) event.setTitle(eventData.title);
    if (eventData.startTime) event.setTime(
      new Date(eventData.startTime),
      eventData.endTime ? new Date(eventData.endTime) : event.getEndTime()
    );
    if (eventData.location !== undefined) event.setLocation(eventData.location || '');
    if (eventData.description !== undefined) event.setDescription(eventData.description || '');

    // イベントURLを構築
    const eventId = event.getId().replace('@google.com', '');
    const eventUrl = `https://calendar.google.com/calendar/event?eid=${encodeURIComponent(event.getId())}&ctz=Asia/Tokyo`;
    
    return {
      success: true,
      eventId: eventId,
      eventUrl: eventUrl,
      message: 'カレンダーイベントを更新しました。'
    };
  } catch (error) {
    console.error('カレンダーイベントの更新中にエラーが発生しました:', error);
    return {
      success: false,
      message: `カレンダーイベントの更新に失敗しました: ${error.message}`
    };
  }
}

/**
 * カレンダーイベントを削除する
 * @param {string} eventId - 削除するイベントID
 * @return {Object} 削除結果
 */
function deleteCalendarEvent(eventId) {
  try {
    const calendarId = getCalendarId();
    const calendar = CalendarApp.getCalendarById(calendarId);
    
    if (!calendar) {
      throw new Error('カレンダーが見つかりません。カレンダーIDを確認してください。');
    }

    // イベントを取得して削除
    const event = calendar.getEventById(eventId);
    if (!event) {
      throw new Error('削除するイベントが見つかりません。');
    }

    event.deleteEvent();

    return {
      success: true,
      message: 'カレンダーイベントを削除しました。'
    };
  } catch (error) {
    console.error('カレンダーイベントの削除中にエラーが発生しました:', error);
    return {
      success: false,
      message: `カレンダーイベントの削除に失敗しました: ${error.message}`
    };
  }
}

/**
 * 指定された日時のイベントが存在するか確認する
 * @param {Date} startTime - 開始日時
 * @param {Date} endTime - 終了日時
 * @param {string} [excludeEventId] - チェックから除外するイベントID
 * @return {boolean} イベントが存在するかどうか
 */
function hasEventInTimeRange(startTime, endTime, excludeEventId = '') {
  try {
    const calendarId = getCalendarId();
    const calendar = CalendarApp.getCalendarById(calendarId);
    
    if (!calendar) {
      throw new Error('カレンダーが見つかりません。カレンダーIDを確認してください。');
    }

    const events = calendar.getEvents(new Date(startTime), new Date(endTime));
    
    // 指定された時間帯に他のイベントがあるか確認
    for (const event of events) {
      const eventId = event.getId().replace('@google.com', '');
      // 除外するイベントIDでない場合、かつ時間が重なっているか確認
      if (eventId !== excludeEventId && 
          event.getStartTime() < new Date(endTime) && 
          event.getEndTime() > new Date(startTime)) {
        return true;
      }
    }
    
    return false;
  } catch (error) {
    console.error('イベントの重複チェック中にエラーが発生しました:', error);
    throw error; // 呼び出し元で適切に処理する
  }
}