/**
 * メニューを追加
 */
function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('カレンダー承認システム')
        .addItem('承認者メールを設定', 'setApproverEmail')
        .addToUi();

    // トリガーが設定されているか確認
    checkTrigger(true);
}

/**
 * 承認者メールアドレスを設定
 */
function setApproverEmail() {
    const ui = SpreadsheetApp.getUi();
    const properties = PropertiesService.getScriptProperties();
    const currentEmail = properties.getProperty(PROPERTY_KEYS.APPROVER_EMAIL) || '';

    const result = ui.prompt(
        '承認者メールアドレスの設定',
        '承認を行うメールアドレスを入力してください:',
        ui.ButtonSet.OK_CANCEL
    );

    if (result.getSelectedButton() === ui.Button.OK) {
        const email = result.getResponseText().trim();

        // メールアドレスのバリデーション
        if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
            ui.alert('エラー', '無効なメールアドレスです。', ui.ButtonSet.OK);
            return;
        }

        // プロパティを更新
        properties.setProperty(PROPERTY_KEYS.APPROVER_EMAIL, email);
        ui.alert('承認者メールアドレスを更新しました。');
    }
}

/**
 * トリガーを確認
 * @param {boolean} silent サイレントモード（アラートを表示しない）
 */
function checkTrigger(silent = false) {
    const triggers = ScriptApp.getProjectTriggers();
    const hasFormSubmitTrigger = triggers.some(
        trigger => trigger.getHandlerFunction() === 'onFormSubmit'
    );

    if (!hasFormSubmitTrigger && !silent) {
        const ui = SpreadsheetApp.getUi();
        ui.alert(
            'トリガーが設定されていません',
            'フォーム送信トリガーが設定されていません。\n' +
            '「リソース」>「プロジェクトのトリガー」から手動で設定してください。',
            ui.ButtonSet.OK
        );
    }

    return hasFormSubmitTrigger;
}


/**
 * メインのエントリーポイント
 * フォーム送信時にトリガーされる
 * @param {GoogleAppsScript.Events.SheetsOnFormSubmit} e - フォーム送信イベント
 */
function createTrigger() {
    // 既存のトリガーを削除
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
        if (trigger.getHandlerFunction() === 'onFormSubmit') {
            ScriptApp.deleteTrigger(trigger);
        }
    });

    // 新しいトリガーを作成
    ScriptApp.newTrigger('onFormSubmit')
        .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
        .onFormSubmit()
        .create();

    Logger.log('フォーム送信トリガーを設定しました');
}
