/**
 * calendar.gs の関数を使用するサンプルコード集
 */

/**
 * カレンダーイベントを作成するサンプル
 */
function exampleCreateEvent() {
  // 現在の日時を取得
  const now = new Date();
  const startTime = new Date(now.getTime() + 60 * 60 * 1000); // 1時間後
  const endTime = new Date(startTime.getTime() + 60 * 60 * 1000); // 開始から1時間後
  
  // イベントデータを準備
  const eventData = {
    title: 'サンプルミーティング',
    startTime: startTime,
    endTime: endTime,
    location: '会議室A',
    description: 'これはサンプルのミーティングです。',
    organizerEmail: Session.getActiveUser().getEmail()
  };
  
  // 参加者（任意）
  const attendees = ['user1@example.com', 'user2@example.com'];
  
  // イベントを作成
  const result = createCalendarEvent(eventData, attendees);
  
  if (result.success) {
    console.log('イベントが作成されました:', result.eventUrl);
    console.log('イベントID:', result.eventId);
  } else {
    console.error('エラー:', result.message);
  }
  
  return result;
}

/**
 * カレンダーイベントを更新するサンプル
 * @param {string} eventId - 更新するイベントID
 */
function exampleUpdateEvent(eventId) {
  // 現在の日時を取得
  const now = new Date();
  const newStartTime = new Date(now.getTime() + 48 * 60 * 60 * 1000); // 2日後
  const newEndTime = new Date(newStartTime.getTime() + 60 * 60 * 1000); // 開始から1時間後
  
  // 更新するイベントデータ
  const updateData = {
    title: '更新されたミーティング',
    startTime: newStartTime,
    endTime: newEndTime,
    location: '会議室B',
    description: 'このミーティングは更新されました。'
  };
  
  // イベントを更新
  const result = updateCalendarEvent(eventId, updateData);
  
  if (result.success) {
    console.log('イベントが更新されました:', result.eventUrl);
  } else {
    console.error('エラー:', result.message);
  }
  
  return result;
}

/**
 * カレンダーイベントを削除するサンプル
 * @param {string} eventId - 削除するイベントID
 */
function exampleDeleteEvent(eventId) {
  const result = deleteCalendarEvent(eventId);
  
  if (result.success) {
    console.log('イベントが削除されました');
  } else {
    console.error('エラー:', result.message);
  }
  
  return result;
}

/**
 * 指定された時間帯にイベントが存在するか確認するサンプル
 */
function exampleCheckEventConflict() {
  // チェックしたい時間帯を設定
  const checkStartTime = new Date('2025-07-28T14:00:00+09:00');
  const checkEndTime = new Date('2025-07-28T15:00:00+09:00');
  
  try {
    const hasConflict = hasEventInTimeRange(checkStartTime, checkEndTime);
    
    if (hasConflict) {
      console.log('指定された時間帯には既にイベントが存在します。');
    } else {
      console.log('指定された時間帯にはイベントがありません。');
    }
    
    return hasConflict;
  } catch (error) {
    console.error('エラーが発生しました:', error.message);
    throw error;
  }
}

/**
 * すべてのサンプルを実行する
 */
function runAllExamples() {
  console.log('=== カレンダー機能のサンプルを開始します ===');
  
  // 1. イベント作成のサンプルを実行
  console.log('\n1. イベント作成のサンプル');
  const createResult = exampleCreateEvent();
  
  if (!createResult.success) {
    console.error('イベントの作成に失敗しました。以降のサンプルをスキップします。');
    return;
  }
  
  // 2. イベント更新のサンプルを実行
  console.log('\n2. イベント更新のサンプル');
  exampleUpdateEvent(createResult.eventId);
  
  // 3. イベント重複チェックのサンプルを実行
  console.log('\n3. イベント重複チェックのサンプル');
  exampleCheckEventConflict();
  
  // 4. イベント削除のサンプルを実行（コメントアウトしています）
  // console.log('\n4. イベント削除のサンプル');
  // exampleDeleteEvent(createResult.eventId);
  
  console.log('\n=== サンプルが完了しました ===');
}

// スクリプトエディタのメニューにカスタムメニューを追加
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('カレンダーサンプル')
    .addItem('サンプルを実行', 'runAllExamples')
    .addToUi();
}
