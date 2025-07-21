/**
 * デバッグ情報を記録する
 * @param {string} message - デバッグメッセージ
 * @param {Object} [data] - 追加データ
 */
function sendDisco(message = "uni") {
    //取得したWebhookURLを追加
    const WEBHOOK_URL = "dicroedwebhookURL";

    const payload = {
        username: "花火",
        content: message,
    };

    UrlFetchApp.fetch(WEBHOOK_URL, {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload),
    });
}
