export const NotificationKeyInfo = 'OnFcsMailAddin';
export const NotificationKeyError = 'OnFcsMailAddinError';
export const DialogTimeoutWarning = 269 * 1000;
export const DialogTimeout = 299 * 1000;
export const MaximumRecipients = 100;

export const SettingKeyDelaySendKey = "mail_delay_send_key"; // 変更すると互換性を失う
export const SettingKeyAccessibilitySize = "accessibility_size"; // 変更すると互換性を失う
export const FontSizeChoice = ["small", "medium", "large"]; // 変更すると互換性を失う

export const BasicFontSizePt = [12, 14, 16];
export const DialogMessageCancel = "<NO>";

export enum MailPropertyType { None, From, To, Cc, Bcc, Subject, Body };

export interface ICustomAttachment {
  name: string;
  data: Office.AttachmentContent;
  size: number;
  id: string;
}

export interface ICheckListItem {
  identity: string;
  key: number;
  recipientType: string;
  name: string;
  message: string;
  size: string;
  address: string;
  isAttachment: boolean;
  isPrivateDL: boolean;
  isExternal: boolean;
  disabled: boolean;
  checked: boolean;
}

export interface ICustomRecipient {
  name: string;
  type: MailPropertyType;
  isPrivateDL: boolean;
  emailAddress: string;
  routingType: string;
  mailboxType: string;
  itemId: string;
  isExternal: boolean;
}

export interface ICustomMessageItemInfo {
  id: string,
  internetMessageId: string,
  changeKey: string,
  debug: string,
}

/* global global, Office, self, window */
export function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
      ? window
      : typeof global !== "undefined"
        ? global
        : undefined;
}

function getDateStringISO8601(dt: Date) {
  console.log('[START] getDateStringISO8601');
  let y = ("0000" + dt.getFullYear()).slice(-4);
  let m = ("00" + (dt.getMonth() + 1)).slice(-2);
  let d = ("00" + dt.getDate()).slice(-2);
  let h = ("00" + dt.getHours()).slice(-2);
  let M = ("00" + dt.getMinutes()).slice(-2);
  let s = ("00" + dt.getSeconds()).slice(-2);
  let result = y + '-' + m + '-' + d + 'T' + h + ':' + M + ':' + s;
  console.log('[END] getDateStringISO8601');
  return result;
}

function promiseReplaceNotificationMessage(key: string, type: Office.MailboxEnums.ItemNotificationMessageType, msg: string): Promise<void> {
  console.log('[START] promiseReplaceNotificationMessage');
  let jsonlike = {
    message: msg.substr(0, 120),
    type: type,
  };
  if (type == Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage) {
    jsonlike['icon'] = "Icon.80x80";
    jsonlike['persistent'] = true;
  }
  console.log('[END] promiseReplaceNotificationMessage');
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.notificationMessages.replaceAsync(key, jsonlike, null,
      (asRes: Office.AsyncResult<void>) => {
        console.log('[promiseReplaceNotificationMessage] 完了しました。status : ' + asRes.status);
        if (asRes.status == Office.AsyncResultStatus.Failed) {
          reject(asRes.error);
        }
        else {
          resolve();
        }
      });
  });
}

function promiseGetTextBody(itemId: string): Promise<string> {
  const reqXml = `<?xml version="1.0" encoding="utf-8"?>
  <soap:Envelope xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance"
                 xmlns:xsd="https://www.w3.org/2001/XMLSchema"
                 xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" 
                 xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" 
                 xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
    <soap:Header>
      <t:RequestServerVersion Version="Exchange2013" soap:mustUnderstand="0" />
    </soap:Header>
    <soap:Body>
      <m:GetItem>
        <m:ItemShape>
          <t:BaseShape>IdOnly</t:BaseShape>
          <t:AdditionalProperties>
            <t:FieldURI FieldURI="item:TextBody" />
          </t:AdditionalProperties>
        </m:ItemShape>
        <m:ItemIds>
          <t:ItemId Id="${itemId}" />
        </m:ItemIds>
      </m:GetItem>
    </soap:Body>
  </soap:Envelope>`;
  return new Promise((resolve, reject) => {
    console.log("[promiseGetTextBody] START/END");
    Office.context.mailbox.makeEwsRequestAsync(reqXml, (asRes => {
      if (asRes.status == Office.AsyncResultStatus.Failed) {
        reject(asRes.error);
      }
      else {
        try {
          const domparser = new DOMParser();
          const htmlDom = domparser.parseFromString(asRes.value, "text/xml");
          const textBody = htmlDom.getElementsByTagNameNS("http://schemas.microsoft.com/exchange/services/2006/types", "TextBody")[0];
          resolve(textBody.innerHTML);
        } catch (err) {
          console.error(err);
          reject(err);
        }
      }
    }));
  });
}

function promiseGetProperty(isOnline: boolean, thisDomain: string, propType: MailPropertyType): Promise<string | Office.EmailAddressDetails | ICustomRecipient[]> {
  console.log('[START/END] promiseGetProperty');
  return new Promise(async (resolve, reject) => {
    if (propType == MailPropertyType.Body) {
      try {
        if (isOnline) {
          let textBody = await promiseGetTextBody(await promiseGetItemId());
          // オンラインの場合は、改行まで正確にとれる
          resolve(unescapeForXml(textBody));
        }
        else {
          Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, null, (asResHtml: Office.AsyncResult<string>) => {
            console.log('[promiseGetProperty] 完了しました。(Html) status : ' + asResHtml.status);
            if (asResHtml.status == Office.AsyncResultStatus.Failed) {
              reject(asResHtml.error);
            }
            else {
              resolve(asResHtml.value);
            }
          });
        }
      }
      catch (err) {
        reject(err);
      }
    }
    else {
      let asyncGetter = null;
      if (propType == MailPropertyType.From) {
        asyncGetter = Office.context.mailbox.item.from; // single
      }
      else if (propType == MailPropertyType.To) {
        asyncGetter = Office.context.mailbox.item.to; // array
      }
      else if (propType == MailPropertyType.Cc) {
        asyncGetter = Office.context.mailbox.item.cc; // array
      }
      else if (propType == MailPropertyType.Bcc) {
        asyncGetter = Office.context.mailbox.item.bcc; // array
      }
      else if (propType == MailPropertyType.Subject) {
        asyncGetter = Office.context.mailbox.item.subject; // string
      }
      asyncGetter.getAsync(null,
        (asRes: Office.AsyncResult<string | Office.EmailAddressDetails | Office.EmailAddressDetails[]>) => {
          console.log('[promiseGetProperty] 完了しました。status : ' + asRes.status);
          if (asRes.status == Office.AsyncResultStatus.Failed) {
            reject(asRes.error);
          }
          else {
            if (propType == MailPropertyType.To || propType == MailPropertyType.Cc || propType == MailPropertyType.Bcc) {
              let ret: ICustomRecipient[] = [];
              for (let det of (asRes.value as Office.EmailAddressDetails[])) {
                ret.push({
                  name: det.displayName,
                  isPrivateDL: !det.emailAddress,
                  type: propType,
                  emailAddress: det.emailAddress || "",
                  routingType: "",
                  mailboxType: "",
                  itemId: "",
                  isExternal: det.emailAddress && thisDomain !== getDomainFromEmailAddress(det.emailAddress)
                });
              }
              resolve(ret);
            }
            else if (propType == MailPropertyType.Subject) {
              resolve(asRes.value as string);
            }
            else if (propType == MailPropertyType.From) {
              resolve(asRes.value as Office.EmailAddressDetails);
            }
          }
        });
    }
  });
}

function unescapeForXml(x: string) {
  x = x.replace(/&amp;/g, "&");
  x = x.replace(/&lt;/g, "<");
  x = x.replace(/&gt;/g, ">");
  x = x.replace(/&quot;/g, "\"");
  x = x.replace(/&apos;/g, "'");
  return x;
}

export function promiseGetItem(itemId: string): Promise<string> {
  // よくわからんが、XML のスキーマが https だったり http だったりして公式通りにやっても通らないので
  // 公式のサンプルと少し違う
  let reqXml = `<?xml version="1.0" encoding="utf-8"?>
  <soap:Envelope xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance"
                 xmlns:xsd="https://www.w3.org/2001/XMLSchema"
                 xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" 
                 xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" 
                 xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
    <soap:Header>
      <t:RequestServerVersion Version="Exchange2013" soap:mustUnderstand="0" />
    </soap:Header>
    <soap:Body>
      <m:GetItem>
        <m:ItemShape>
          <t:BaseShape>AllProperties</t:BaseShape>
          <t:AdditionalProperties>
            <t:ExtendedFieldURI PropertyTag="16367" PropertyType="SystemTime" />
          </t:AdditionalProperties>
        </m:ItemShape>
        <m:ItemIds>
          <t:ItemId Id="${itemId}" />
        </m:ItemIds>
      </m:GetItem>
    </soap:Body>
  </soap:Envelope>`;

  return new Promise((resolve, reject) => {
    Office.context.mailbox.makeEwsRequestAsync(reqXml,
      async (asRes: Office.AsyncResult<string>) => {
        if (asRes.status == Office.AsyncResultStatus.Failed) {
          reject(asRes.error);
        }
        else {
          resolve(asRes.value);
        }
      });
  });
}

export async function promiseDisplayMessage(msg: string) {
  await promiseReplaceNotificationMessage(
    NotificationKeyInfo,
    Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    '誤送信防止アドイン : ' + msg);
}

export async function promiseDisplayErrorMessage(msg: string) {
  await promiseReplaceNotificationMessage(
    NotificationKeyError,
    Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
    '誤送信防止アドイン : ' + msg);
}

function promiseGetAttachments(): Promise<Office.AttachmentDetailsCompose[]> {
  console.log('[START/END] promiseGetAttachments');
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.getAttachmentsAsync(null,
      (asRes: Office.AsyncResult<Office.AttachmentDetailsCompose[]>) => {
        console.log('[promiseGetAttachments] 完了しました。status : ' + asRes.status);
        if (asRes.status == Office.AsyncResultStatus.Failed) {
          reject(asRes.error);
        }
        else {
          resolve(asRes.value);
        }
      });
  });
}

export async function promiseOpenDialog(isOnline: boolean): Promise<boolean> {
  // ダイアログ側の処理では取得できない。
  const subject = await promiseGetProperty(isOnline, null, MailPropertyType.Subject) as string;
  const body = await promiseGetProperty(isOnline, null, MailPropertyType.Body) as string;
  const thisDomain = getDomainFromEmailAddress(Office.context.mailbox.userProfile.emailAddress);
  const itemTo = await promiseGetProperty(isOnline, thisDomain, MailPropertyType.To) as ICustomRecipient[];
  const itemCc = await promiseGetProperty(isOnline, thisDomain, MailPropertyType.Cc) as ICustomRecipient[];
  const itemBcc = await promiseGetProperty(isOnline, thisDomain, MailPropertyType.Bcc) as ICustomRecipient[];
  const recipients = itemTo.concat(itemCc).concat(itemBcc);
  window.localStorage.setItem("mail_subject", subject);
  window.localStorage.setItem("mail_body", body);
  window.localStorage.setItem("mail_attachments", JSON.stringify(await promiseGetAttachments()));
  window.localStorage.setItem("mail_recipients", JSON.stringify(recipients));

  // To, Cc, Bcc の上限はバイト数で決まる。2017 年頃の情報だと 32000 バイト。
  setTimeout(async function () {
    let testItemIds = await promiseGetSentItems(recipients);
    if (testItemIds.length > 0) {
      window.localStorage.setItem("mail_recipient_pattern_exists", "yes");
    }
    else {
      window.localStorage.setItem("mail_recipient_pattern_exists", "no");
    }
  });

  console.log('[START/END] promiseOpenDialog');
  return new Promise((resolve, reject) => {
    Office.context.ui.displayDialogAsync(window.location.origin + "/dialog.html",
      {
        height: 70, // percent
        width: 70, // percent
        displayInIframe: true,
      },
      function (asRes: Office.AsyncResult<Office.Dialog>) {
        if (asRes.status == Office.AsyncResultStatus.Failed) {
          // In addition to general system errors, there are 3 specific errors for 
          // displayDialogAsync that you can handle individually.
          // 12004 : Domain is not trusted.
          // 12005 : HTTPS is required.
          // 12007 : Already opened.
          // 12009 : Dialog box is ignored by user
          console.error('[promiseOpenDialog] ダイアログを開けませんでした。asRes.error.code : ' + asRes.error.code);
          reject(asRes.error);
        }
        else {
          let dialog = asRes.value;
          dialog.addEventHandler(Office.EventType.DialogMessageReceived,
            async function (arg: any) {
              console.log('[promiseOpenDialog] メッセージを受信しました。arg.message : ' + arg.message);
              try {
                if (arg.message.indexOf("<ENCRYPT>") >= 0) {
                  await promiseDisplayMessage("処理が完了しました。");
                }
                dialog.close();
                if (arg.message.indexOf("<YES>") >= 0) {
                  resolve(true);
                }
                else {
                  resolve(false);
                }
              } catch (err) {
                console.error(err);
                reject(err);
              }
            }
          );
          dialog.addEventHandler(Office.EventType.DialogEventReceived,
            function (arg: any) {
              console.log('[promiseOpenDialog] イベントを受信しました。arg.error : ' + arg.error);
              // Events are sent by the platform in response to user actions or errors.
              let msg = "";
              switch (arg.error) {
                case 12002:
                  msg = "Cannot load URL, no such page or bad URL syntax.";
                  break;
                case 12003:
                  msg = "HTTPS is required.";
                  break;
                case 12006:
                  // The dialog has been closed (pressed X). Treat as a cancel.
                  resolve(false);
                  break;
                default:
                  msg = "Unknown error.";
                  break;
              }
              reject({ code: arg.error, message: msg, name: "DialogEventReceived" });
            }
          );
        }
      });
  });
}

export const getDomainFromEmailAddress = function (addr: string) {
  return addr.substr(1 + addr.lastIndexOf("@"));
}

export function promiseSaveAsDraft(): Promise<string> {
  console.log('[START/END] promiseSaveAsDraft');
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.saveAsync(null,
      (asRes: Office.AsyncResult<string>) => {
        console.log('[promiseSaveAsDraft] 完了しました。status : ' + asRes.status);
        if (asRes.status == Office.AsyncResultStatus.Failed) {
          reject(asRes.error);
        }
        else {
          resolve(asRes.value);
        }
      }
    );
  });
}

export function promiseGetItemId(): Promise<string> {
  console.log('[START/END] promiseGetItemId');
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.getItemIdAsync(null,
      (asRes: Office.AsyncResult<string>) => {
        console.log('[promiseGetItemId] 完了しました。status : ' + asRes.status);
        if (asRes.status == Office.AsyncResultStatus.Failed) {
          reject(asRes.error);
        }
        else {
          resolve(asRes.value);
        }
      });
  });
}

export function extractCustomMessageItemInfo(xmlRes: string): ICustomMessageItemInfo {
  let domparser = new DOMParser();
  let xml = domparser.parseFromString(xmlRes, "text/xml");
  let itemIdElem = xml.getElementsByTagNameNS("http://schemas.microsoft.com/exchange/services/2006/types", "ItemId")[0];
  let itemId = itemIdElem.getAttribute("Id");
  let changeKey = itemIdElem.getAttribute("ChangeKey");
  let imIdElem = xml.getElementsByTagNameNS("http://schemas.microsoft.com/exchange/services/2006/types", "InternetMessageId");
  let imId = "";
  if (imIdElem.length > 0) {
    imId = imIdElem[0].innerHTML;
    imId = imId.substring(4, imId.length - 4); // よくわからないが <> の実態参照で囲まれている。
  }
  let ret: ICustomMessageItemInfo = { id: itemId, changeKey: changeKey, debug: xmlRes, internetMessageId: imId };
  return ret;
}

export function promiseSetDefferedDeliveryTime(
  info: ICustomMessageItemInfo,
  sendAfterInMinute: number,
  isOnline: boolean
): Promise<boolean> {
  // JST->GMT
  let defferedDate = new Date((new Date()).getTime() + sendAfterInMinute * 60000 - 9 * (3600000));
  let defferedTimestamp = getDateStringISO8601(defferedDate);
  let reqXml = `<?xml version="1.0" encoding="utf-8"?>
  <soap:Envelope xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance" 
               xmlns:xsd="https://www.w3.org/2001/XMLSchema" 
               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" 
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
    <soap:Header>
      <t:RequestServerVersion Version="Exchange2013" soap:mustUnderstand="0" />
    </soap:Header>
    <soap:Body>
      <UpdateItem MessageDisposition="${isOnline ? "SaveOnly" : "SendOnly"}" ConflictResolution="AutoResolve" 
                  xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">
        <ItemChanges>
          <t:ItemChange>
            <t:ItemId Id="${info.id}" ChangeKey="${info.changeKey}" />
            <t:Updates>
              <t:SetItemField>
                <t:ExtendedFieldURI PropertyTag="16367" PropertyType="SystemTime" />
                <t:Message>
                  <t:ExtendedProperty>
                    <t:ExtendedFieldURI PropertyTag="16367"
                                        PropertyType="SystemTime" />
                    <t:Value>${defferedTimestamp}</t:Value>
                  </t:ExtendedProperty>
                </t:Message>
              </t:SetItemField>
            </t:Updates>
          </t:ItemChange>
        </ItemChanges>
      </UpdateItem>
    </soap:Body>
  </soap:Envelope>`;
  console.log('[START/END] promiseSetDefferedDeliveryTime');
  return new Promise((resolve, reject) => {
    Office.context.mailbox.makeEwsRequestAsync(reqXml,
      (asRes: Office.AsyncResult<string>) => {
        console.log('[promiseSetDefferedDeliveryTime] 完了しました。status : ' + asRes.status);
        if (asRes.status == Office.AsyncResultStatus.Failed) {
          reject(asRes.error);
        }
        else {
          try {
            let domparser = new DOMParser();
            let xml = domparser.parseFromString(asRes.value, "text/xml");
            let responseElem = xml.getElementsByTagNameNS("http://schemas.microsoft.com/exchange/services/2006/messages", "UpdateItemResponseMessage")[0];
            let result = responseElem.getAttribute("ResponseClass");
            if (result === "Success") {
              resolve(true);
            }
            else {
              throw new Error("promiseSetDefferedDeliveryTime failed.");
            }
          } catch (err) {
            err.message = "[promiseSetDefferedDeliveryTime] " + err.message;
            reject(err);
          }
        }
      });
  });
}

function promiseGetSentItems(recipients: ICustomRecipient[]): Promise<string[]> {
  // https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/querystring-querystringtype
  // Traversal=Shallow : Returns only Id
  // AdditionalProperties : 取得する情報や、検索条件で必要なフィールド等も指定する。
  // IndexedPageItemView : ビューの状態（件数とオフセット）
  const concatAsQueryString = function (recipients: ICustomRecipient[]) {
    let conds = [];
    for (let recipient of recipients) {
      let ea = recipient.emailAddress;
      let et = "";
      switch (recipient.type) {
        case MailPropertyType.To: et = "to"; break;
        case MailPropertyType.Cc: et = "cc"; break;
        case MailPropertyType.Bcc: et = "bcc"; break;
      }
      if (ea && ea.indexOf("@") >= 0) {
        conds.push(et + ":" + ea);
      }
    }
    return conds.join(" AND ");
  }

  const queryString = concatAsQueryString(recipients);
  const reqXml = `<?xml version="1.0" encoding="utf-8"?>
  <soap:Envelope xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance" 
               xmlns:xsd="https://www.w3.org/2001/XMLSchema" 
               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" 
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
    <soap:Header>
      <t:RequestServerVersion Version="Exchange2013" />
    </soap:Header>
    <soap:Body>
      <m:FindItem Traversal="Shallow">
        <m:ItemShape>
          <t:BaseShape>IdOnly</t:BaseShape>
          <t:AdditionalProperties>
            <t:FieldURI FieldURI="message:ToRecipients" />
            <t:FieldURI FieldURI="message:CcRecipients" />
            <t:FieldURI FieldURI="message:BccRecipients" />
          </t:AdditionalProperties>
        </m:ItemShape>
        <m:IndexedPageItemView MaxEntriesReturned="1" Offset="0" BasePoint="Beginning" />
        <m:ParentFolderIds>
          <t:DistinguishedFolderId Id="sentitems"/>
        </m:ParentFolderIds>
        <m:QueryString>${queryString}</m:QueryString>
      </m:FindItem>
    </soap:Body>
  </soap:Envelope>`;
  return new Promise((resolve, reject) => {
    Office.context.mailbox.makeEwsRequestAsync(reqXml,
      (asRes: Office.AsyncResult<string>) => {
        console.log('[promiseGetSentItems] 完了しました。status : ' + asRes.status);
        if (asRes.status == Office.AsyncResultStatus.Failed) {
          reject(asRes.error);
        }
        else {
          let domparser = new DOMParser();
          let xml = domparser.parseFromString(asRes.value, "text/xml");
          let responseElem = xml.getElementsByTagNameNS("http://schemas.microsoft.com/exchange/services/2006/messages", "FindItemResponseMessage")[0];
          let result = responseElem.getAttribute("ResponseClass");
          let ret = [];
          if (result === "Success") {
            let elems = xml.getElementsByTagNameNS("http://schemas.microsoft.com/exchange/services/2006/types", "ItemId");
            for (let elem_idx = 0; elem_idx < elems.length; ++elem_idx) {
              let id = elems[elem_idx].getAttribute("Id");
              ret.push(id);
            }
          }
          resolve(ret);
        }
      });
  });
}

export function promiseSaveRoamingSetting(): Promise<void> {
  console.log('[START/END] promiseSaveRoamingSetting');
  return new Promise((resolve, reject) => {
    Office.context.roamingSettings.saveAsync(function (result) {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        reject(result.error);
      } else {
        console.log('[promiseSaveRoamingSetting] 完了しました。status : ' + result.status);
        resolve();
      }
    });
  });
}

export function setLocalStorageFromRoamingSetting(key: string, initial: string) {
  let val = Office.context.roamingSettings.get(key) || initial;
  window.localStorage.setItem(key, val);
}

export function setRoamingSettingFromLocalStorage(key: string) {
  Office.context.roamingSettings.set(key, window.localStorage.getItem(key));
}
