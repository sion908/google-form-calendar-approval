/**
 * ユーティリティ関数を提供する
 */

/**
 * 日時をフォーマットする
 * @param {Date|string} date - フォーマットする日時
 * @return {string} フォーマットされた日時文字列
 */
function formatDateTime(date) {
  if (!date) return '';
  
  if (!(date instanceof Date)) {
    date = new Date(date);
  }
  
  // 有効な日付かチェック
  if (isNaN(date.getTime())) {
    return '日時が不正です';
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

/**
 * 日付と時刻を組み合わせてDateオブジェクトを作成する
 * @param {Date|string} date - 日付
 * @param {Date|string} time - 時刻
 * @return {Date} 結合されたDateオブジェクト
 */
function combineDateTime(date, time) {
  if (!date || !time) return null;
  
  const dateObj = new Date(date);
  const timeObj = new Date(time);
  
  if (isNaN(dateObj.getTime()) || isNaN(timeObj.getTime())) {
    return null;
  }
  
  dateObj.setHours(timeObj.getHours(), timeObj.getMinutes(), 0, 0);
  return dateObj;
}
