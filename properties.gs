/**
 * スクリプトプロパティを管理するためのユーティリティ関数
 */

/**
 * スクリプトプロパティを初期化する
 * 初回セットアップ時に実行する
 */
function initializeProperties() {
  const properties = PropertiesService.getScriptProperties();
  
  // 必要なプロパティが設定されていない場合、デフォルト値を設定
  const defaultProperties = {
    [PROPERTY_KEYS.APPROVER_EMAIL]: '',  // 承認者のメールアドレス
    [PROPERTY_KEYS.CALENDAR_ID]: 'primary',  // デフォルトはプライマリカレンダー
    [PROPERTY_KEYS.SCRIPT_URL]: ScriptApp.getService().getUrl() || ''
  };
  
  properties.setProperties(defaultProperties);
  return defaultProperties;
}

/**
 * スクリプトプロパティを取得する
 * @param {string} key - プロパティのキー
 * @param {*} defaultValue - プロパティが存在しない場合のデフォルト値
 * @return {*} プロパティの値
 */
function getProperty(key, defaultValue = null) {
  const properties = PropertiesService.getScriptProperties();
  return properties.getProperty(key) || defaultValue;
}

/**
 * スクリプトプロパティを設定する
 * @param {string} key - プロパティのキー
 * @param {string} value - 設定する値
 */
function setProperty(key, value) {
  const properties = PropertiesService.getScriptProperties();
  properties.setProperty(key, value);
}

/**
 * 承認者のメールアドレスを取得する
 * @return {string} 承認者のメールアドレス
 */
function getApproverEmail() {
  return getProperty(PROPERTY_KEYS.APPROVER_EMAIL);
}

/**
 * カレンダーIDを取得する
 * @return {string} カレンダーID
 */
function getCalendarId() {
  return getProperty(PROPERTY_KEYS.CALENDAR_ID);
}

/**
 * スクリプトのURLを取得する
 * @return {string} スクリプトのURL
 */
function getScriptUrl() {
  return getProperty(PROPERTY_KEYS.SCRIPT_URL) || ScriptApp.getService().getUrl();
}

/**
 * プロパティを一括で設定する
 * @param {Object} properties - 設定するプロパティのキーと値のオブジェクト
 */
function setProperties(properties) {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperties(properties);
}

/**
 * すべてのプロパティを削除する（デバッグ用）
 */
function deleteAllProperties() {
  const properties = PropertiesService.getScriptProperties();
  properties.deleteAllProperties();
}