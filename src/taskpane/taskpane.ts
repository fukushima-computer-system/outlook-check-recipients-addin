import Common = require('../common');

const setSampleFontSize = function (radioButtonId) {
  Common.FontSizeChoice.forEach((v, idx) => {
    if (radioButtonId.endsWith(v)) {
      document.getElementById("size-sample").style.fontSize = Common.BasicFontSizePt[idx] + "pt";
    }
  });
}

Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("save").onclick = save;
    document.getElementsByName("accessibility-size").forEach(radioElem => {
      radioElem.onclick = (_) => {
        setSampleFontSize(radioElem.id);
      }
    });
    // アクセシビリティの設定
    const sizeSettingValue = Office.context.roamingSettings.get(Common.SettingKeyAccessibilitySize);
    if (typeof sizeSettingValue === "undefined") {
      (<HTMLInputElement>document.getElementById("accessibility-size-small")).checked = true;
    }
    else {
      const radioButtonId = "accessibility-size-" + sizeSettingValue;
      (<HTMLInputElement>document.getElementById(radioButtonId)).checked = true;
      setSampleFontSize(radioButtonId);
    }
    (<HTMLButtonElement>document.getElementById("save")).disabled = false;
  }
});


//保存ボタンクリック時
async function save() {
  const domMsg = document.getElementById("save-message");
  try {
    domMsg.classList.remove("is-error");
    domMsg.textContent = "保存しています ...";

    const sizeSettingValue = Common.FontSizeChoice.filter(v => {
      return (<HTMLInputElement>document.getElementById("accessibility-size-" + v)).checked;
    })[0];
    Office.context.roamingSettings.set(Common.SettingKeyAccessibilitySize, sizeSettingValue);
    await Common.promiseSaveRoamingSetting();
    domMsg.textContent = "保存しました。";
  } catch (err) {
    domMsg.classList.add("is-error");
    domMsg.textContent = "保存できませんでした。" + err.message;
    console.log("[save] 下書きテキスト保存に失敗しました。err.message : " + err.message);
  }
}
