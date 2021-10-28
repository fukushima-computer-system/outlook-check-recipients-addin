import Common = require('../common');

Office.onReady(() => {
    const g = Common.getGlobal() as any;
    // the add-in command functions need to be available in global scope
    g.inspect = inspect;
});

async function inspect(event: any) {
    // メールアイテムの情報はローカルにはない。すべて EWS リクエストとして取得する必要がある。
    // https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/querystring-querystringtype

    const isOnline = Office.context.platform == Office.PlatformType.OfficeOnline;
    let finalStatus = false;
    let userCancelled = false;

    try {
        if ("message" === Office.context.mailbox.item.itemType) {
            window.localStorage.clear();
            window.localStorage.setItem("time_activated", new Date().toISOString());

            await Common.promiseSaveAsDraft();
            await Common.promiseDisplayMessage('メッセージを下書き保存しました。');
            await Common.promiseDisplayMessage('アイテムから情報を取得しました。');

            // 個人用設定から復元
            Common.setLocalStorageFromRoamingSetting(Common.SettingKeyDelaySendKey, "delay_send_0");
            Common.setLocalStorageFromRoamingSetting("width_recipient_name", "120");
            Common.setLocalStorageFromRoamingSetting("width_recipient_address", "220");
            Common.setLocalStorageFromRoamingSetting("width_attachment_name", "120");
            Common.setLocalStorageFromRoamingSetting("width_attachment_size", "70");
            Common.setLocalStorageFromRoamingSetting(Common.SettingKeyAccessibilitySize, Common.FontSizeChoice[1]);

            // 宛先確認ダイアログを開く
            if (!await Common.promiseOpenDialog(isOnline)) {
                userCancelled = true;
                throw new Error("キャンセルしました。");
            }
            if (window.localStorage.getItem("timeout") === "timeout") {
                throw new Error("処理がタイムアウトしました。");
            }

            // 遅延送信 (Online のみ)
            if (isOnline) {
                const delaySendKey = window.localStorage.getItem(Common.SettingKeyDelaySendKey);
                Office.context.roamingSettings.set(Common.SettingKeyDelaySendKey, delaySendKey);
                const delaySendMinutesFloatValue = Math.floor(parseFloat(delaySendKey.substring(delaySendKey.lastIndexOf("_") + 1)));
                if (delaySendMinutesFloatValue > 0 && delaySendMinutesFloatValue < 30) {
                    const iid = await Common.promiseGetItemId();
                    const getItemXmlRes = await Common.promiseGetItem(iid);
                    const customMsgInfo: Common.ICustomMessageItemInfo = Common.extractCustomMessageItemInfo(getItemXmlRes);
                    await Common.promiseSetDefferedDeliveryTime(customMsgInfo, delaySendMinutesFloatValue, isOnline);
                }
                else if (delaySendMinutesFloatValue >= 30) {
                    throw new Error("遅延送信情報が不正です。");
                }
            }

            // 表の列幅を保存
            Common.setRoamingSettingFromLocalStorage("width_recipient_name");
            Common.setRoamingSettingFromLocalStorage("width_recipient_address");
            Common.setRoamingSettingFromLocalStorage("width_attachment_name");
            Common.setRoamingSettingFromLocalStorage("width_attachment_size");

            // ダイアログ側の処理では保存できない。
            await Common.promiseSaveRoamingSetting();

            finalStatus = true;
        }
        else {
            // Outlook のメッセージアイテムではない場合。
            finalStatus = true;
        }
    } catch (err) {
        try {
            console.error(err);
            if (!userCancelled) {
                await Common.promiseDisplayMessage('アドインの処理エラーにより送信が中止されました。詳細：' + err.message);
            }
            else {
                await Common.promiseDisplayMessage('確認ダイアログにより送信はキャンセルされました。');
            }
        } catch (err2) {
            console.error(err2);
        }
    }
    finally {
        try {
            window.localStorage.clear();
        }
        catch (err) {
            console.error(err);
        }
        event.completed({ allowEvent: finalStatus });
    }
}
