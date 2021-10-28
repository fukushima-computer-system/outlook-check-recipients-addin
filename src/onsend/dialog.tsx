// React Section
import * as React from "react";
import { useEffect } from "react";
import * as ReactDOM from "react-dom";

// Fluent UI Section
import "office-ui-fabric-react/dist/css/fabric.min.css";
import { useId, useBoolean } from '@uifabric/react-hooks';
import { initializeIcons } from "office-ui-fabric-react/lib/Icons";

import {
    Callout,
    Checkbox,
    DefaultButton,
    Dialog,
    DialogFooter,
    DialogType,
    DetailsRow,
    Dropdown,
    DropdownMenuItemType,
    Icon,
    IDetailsRowStyles,
    IDropdownOption,
    IDropdownStyles,
    Label,
    MessageBar,
    MessageBarType,
    Pivot,
    PivotItem,
    PrimaryButton,
    ProgressIndicator,
    Stack,
    Text,
} from 'office-ui-fabric-react';

import {
    DetailsList,
    DetailsListLayoutMode,
    IColumn,
    SelectionMode
} from 'office-ui-fabric-react/lib/DetailsList';

import Common = require('../common');

enum CheckStage { None, First, Second };

// 大・中・小の設定値
const sizeDict = (() => {
    let val = window.localStorage.getItem(Common.SettingKeyAccessibilitySize);
    let idx = Common.FontSizeChoice.findIndex(v => v == val);
    let baseSize = Common.BasicFontSizePt[idx];
    return {
        message: (baseSize - 1) + "pt",
        label: (baseSize - 1) + "pt",
        detailsListCell: baseSize + "pt",
        detailsListColumnCellName: (baseSize - 1) + "pt",
        dropdownOptionText: (baseSize) + "pt",
        subject: baseSize + "pt",
        body: baseSize + "pt",
        checkBoxMarginTop: "2px",
        checkBox: ([20, 24, 28])[idx] + "px",
    };
})();

const MyMessageBar = (msg: JSX.Element, barType: MessageBarType) => (
    <MessageBar
        messageBarType={barType}
        style={{ fontSize: sizeDict["message"] }}>{msg}</MessageBar>
);

const MyLabel = (msg: JSX.Element) => (
    <Label style={{ paddingTop: "8px", paddingBottom: "8px", fontSize: sizeDict["label"] }}>{msg}</Label>
);

type CheckListProps = {
    items: Common.ICheckListItem[];
    columnParams: CreateColumnParams;
}

const CheckList: React.FunctionComponent<CheckListProps> = ({ items, columnParams }) => {
    // ピボット内で再描画する際に、列幅が反映されないため、レンダするたびに列も再生成する。
    const columns = createDetailsListColumns(columnParams);
    return (
        <DetailsList
            items={items}
            columns={columns}
            setKey="set"
            styles={{
                root: {
                    height: "100%",
                    overflow: "auto",
                },
            }}
            onShouldVirtualize={() => false}
            onColumnResize={(column, newWidth) => {
                if (column.isResizable) {
                    window.localStorage.setItem("width_" + column.key, "" + newWidth);
                }
            }}
            selectionMode={SelectionMode.none}
            onRenderRow={props => {
                const customStyles: Partial<IDetailsRowStyles> = { cell: { fontSize: sizeDict["detailsListCell"] } };
                if (props) {
                    // Every other row renders with a different background color
                    // customStyles.root = { backgroundColor: (!props.item.is_attachment&&props.item.warn) ? "rgb(253, 231, 233) !important" : null};
                    return <DetailsRow {...props} styles={customStyles} />;
                }
                return null;
            }}
            layoutMode={DetailsListLayoutMode.fixedColumns}
        />
    );
}

const convertRecipientToListItem = function (idx: number, data: Common.ICustomRecipient, isDoubleCheckList: boolean): Common.ICheckListItem {
    const getRecipientTypeString = function (type: Common.MailPropertyType): string {
        let ret: string = "";
        switch (type) {
            case Common.MailPropertyType.To:
                ret = "To";
                break;
            case Common.MailPropertyType.Cc:
                ret = "Cc";
                break;
            case Common.MailPropertyType.Bcc:
                ret = "Bcc";
                break;
        }
        return ret;
    }

    const getMessage = function (data: Common.ICustomRecipient): string {
        let ret = "";
        ret += data.isPrivateDL ? "プライベート配布リストは、展開しなければ送信できません。一度ダイアログを閉じて、当該配布リストの左側の「＋」マークをダブルクリックすると、展開することができます。" : "";
        return ret;
    }

    return {
        key: idx,
        recipientType: getRecipientTypeString(data.type),
        name: data.name,
        message: getMessage(data),
        size: "",
        identity: "fcs_recipient_" + (isDoubleCheckList ? "second_" : "first_") + idx,
        address: data.emailAddress,
        isAttachment: false,
        isPrivateDL: data.isPrivateDL,
        isExternal: data.isExternal,
        checked: isDoubleCheckList ? false : (!data.isPrivateDL && !data.isExternal),
        disabled: data.isPrivateDL || !data.isExternal
    }
}

type DescriptionCalloutProps = {
    title: string
    description: string,
}

const DetailsListCallout: React.FunctionComponent<DescriptionCalloutProps> = ({
    title,
    description,
}) => {
    const [isCalloutVisible, { toggle: toggleIsCalloutVisible }] = useBoolean(false);

    const buttonId: string = useId('callout-button');
    const labelId: string = useId('callout-label');
    const descriptionId: string = useId('callout-description');
    return (
        <>
            <DefaultButton
                styles={{
                    textContainer: {
                        fontSize: sizeDict["label"]
                    }
                }}
                style={{ height: "auto", border: "0", background: "none" }}
                iconProps={{ iconName: "IncidentTriangle" }}
                id={buttonId}
                onClick={toggleIsCalloutVisible}
                text={isCalloutVisible ? '詳細を非表示' : '詳細を表示'} />
            {isCalloutVisible && (
                <Callout
                    styles={{ root: { width: "480px" } }}
                    ariaLabelledBy={labelId}
                    ariaDescribedBy={descriptionId}
                    role="alertdialog"
                    gapSpace={0}
                    target={`#${buttonId}`}
                    onDismiss={toggleIsCalloutVisible}
                    setInitialFocus
                >
                    <Stack>
                        <Text id={labelId} style={{ fontSize: "17pt", margin: "16px", marginBottom: "0" }}>
                            {title}
                        </Text>
                        <Text id={descriptionId} style={{ fontSize: sizeDict["label"], margin: "16px" }}>
                            {description}
                        </Text>
                    </Stack>
                </Callout>
            )}
        </>
    );
};

type CheckListCheckboxProps = {
    className: string,
    item: Common.ICheckListItem,
    changedEventHandlerCreator: (isChecked: boolean, toggler: () => void) =>
        (ev: React.FormEvent<HTMLElement | HTMLInputElement>, checked: boolean) =>
            void
}

const CheckListCheckbox: React.FunctionComponent<CheckListCheckboxProps> = ({
    className,
    item,
    changedEventHandlerCreator: creator
}) => {
    const [isChecked, { toggle: toggler }] = useBoolean(false);
    return <Checkbox
        className={className}
        styles={{
            checkbox: {
                width: sizeDict["checkBox"],
                height: sizeDict["checkBox"]
            }
        }}
        id={item.identity}
        checked={item.checked}
        disabled={item.disabled}
        onChange={creator(isChecked, toggler)}
    />;
}

const createDetailsListColumns = function (
    params: CreateColumnParams
): IColumn[] {
    const width_recipient_name = parseInt(window.localStorage.getItem("width_recipient_name"));
    const width_recipient_address = parseInt(window.localStorage.getItem("width_recipient_address"));
    const width_attachment_name = parseInt(window.localStorage.getItem("width_attachment_name"));
    const width_attachment_size = parseInt(window.localStorage.getItem("width_attachment_size"));
    const className = params.className;
    const checkboxChangedEventHandlerCreator = params.handlerCreator;
    const isAttachment = params.isAttachment;

    let ret: IColumn[] = [
        {
            key: 'check',
            name: '',
            fieldName: 'check',
            minWidth: 40,
            maxWidth: 40,
            isResizable: false,
            onRender: (item: Common.ICheckListItem) => {
                return <CheckListCheckbox className={className} changedEventHandlerCreator={checkboxChangedEventHandlerCreator} item={item} />;
            }
        },
    ];
    ret = ret.concat(
        isAttachment ?
            [
                {
                    key: 'attachment_name',
                    name: '添付ファイル名',
                    fieldName: 'name',
                    minWidth: 40,
                    styles: { cellName: { fontSize: sizeDict["detailsListColumnCellName"] } },
                    maxWidth: width_attachment_name,
                    isResizable: true,
                    onRender: (item: Common.ICheckListItem) => {
                        return item.name;
                    }
                },
                {
                    key: 'attachment_size',
                    name: 'サイズ',
                    fieldName: 'size',
                    minWidth: 40,
                    styles: { cellName: { fontSize: sizeDict["detailsListColumnCellName"] } },
                    maxWidth: width_attachment_size,
                    isResizable: true,
                    onRender: (item: Common.ICheckListItem) => {
                        return item.size;
                    }
                },
            ]
            :
            [
                {
                    key: 'type',
                    name: '種別',
                    fieldName: 'recipientType',
                    minWidth: 40,
                    styles: { cellName: { fontSize: sizeDict["detailsListColumnCellName"] } },
                    isResizable: false,
                    onRender: (item: Common.ICheckListItem) => {
                        return item.recipientType;
                    }
                },
                {
                    key: 'recipient_name',
                    name: '名前',
                    fieldName: 'name',
                    minWidth: 40,
                    styles: { cellName: { fontSize: sizeDict["detailsListColumnCellName"] } },
                    maxWidth: width_recipient_name,
                    isResizable: true,
                    onRender: (item: Common.ICheckListItem) => {
                        if (item.isExternal) {
                            return <span style={{ fontWeight: 'bold', color: '#ca4414' }}>{item.name}</span>;
                        }
                        else {
                            return item.name;
                        }
                    }
                },
                {
                    key: 'recipient_address',
                    name: 'メールアドレス',
                    fieldName: 'address',
                    minWidth: 40,
                    styles: { cellName: { fontSize: sizeDict["detailsListColumnCellName"] } },
                    maxWidth: width_recipient_address,
                    isResizable: true,
                    onRender: (item: Common.ICheckListItem) => {
                        if (item.isExternal) {
                            return <span style={{ fontWeight: 'bold', color: '#ca4414' }}>{item.address}</span>;
                        }
                        else if (item.isPrivateDL) {
                            return <DetailsListCallout title="プライベート配布リスト" description={item.message} />;
                        }
                        else {
                            return <span>{item.address}</span>;
                        }
                    }
                },
            ]);
    return ret;
}

type CheckModalProps = {
    attachmentList: Common.ICheckListItem[],
    recipientList: Common.ICheckListItem[],
    doubleCheckRecipientList: Common.ICheckListItem[],
    subject: string,
    body: string,
    addinActivatedDate: Date,
}

type TimeoutMessageBarProps = {
    addinActivatedDate: Date
};

const TimeoutMessageBar: React.FunctionComponent<TimeoutMessageBarProps> = ({
    addinActivatedDate
}) => {
    const [currentDate, setCurrentDate] = React.useState<Date>();
    const isDialogTimeout = (timeoutMs) => {
        return currentDate && (currentDate.getTime() - addinActivatedDate.getTime() > timeoutMs);
    }
    const isOnline = Office.context.platform == Office.PlatformType.OfficeOnline;
    useEffect(() => {
        const dateUpdateTimer = setInterval(() => {
            setCurrentDate(new Date());
            if (isDialogTimeout(Common.DialogTimeout)) {
                window.localStorage.setItem("timeout", "timeout");
            }
        }, 10000);
        return () => {
            clearInterval(dateUpdateTimer);
        }
    });

    return <>
        {
            isDialogTimeout(Common.DialogTimeout) ?
                MyMessageBar(<span>アドインは無効化されました。メッセージは下書きに保存されています。
                    {isOnline ? "ページを再読み込みして" : "アドインのウィンドウを閉じて"}ください。</span>, MessageBarType.severeWarning) :
                (isDialogTimeout(Common.DialogTimeoutWarning) ?
                    MyMessageBar(<span>アドインはまもなく Outlook によって無効化されます。ダイアログを一度閉じて、開き直すことを推奨します。</span>, MessageBarType.warning)
                    : null)
        }
    </>;
}

type CreateColumnParams = {
    className: string;
    isAttachment: boolean;
    handlerCreator: (_: boolean, __: () => void) => (_: React.FormEvent<HTMLElement | HTMLInputElement>, __: boolean) => void;
}

const CheckModal: React.FunctionComponent<CheckModalProps> = ({
    attachmentList,
    recipientList,
    doubleCheckRecipientList,
    subject,
    body,
    addinActivatedDate
}) => {
    // Get Basic info.
    const isOnline = Office.context.platform == Office.PlatformType.OfficeOnline;
    const [isDoubleCheckModalHidden, { setTrue: hideDoubleCheckModal, setFalse: showDoubleCheckModal }] = useBoolean(true);
    const [isYesDisabled, setYesDisabled] = React.useState<boolean>(true);
    const [isDoubleCheckYesDisabled, setDoubleCheckYesDisabled] = React.useState<boolean>(true);
    const [selectedItem, setSelectedItem] = React.useState<IDropdownOption>();
    const [recipientPatternAlreadyExists, setRecipientPatternAlreadyExists] = React.useState<string>();

    useEffect(() => {
        updateYesButtonDisabled(CheckStage.First);
        const patternPeekTimer = setInterval(() => {
            let val = window.localStorage.getItem("mail_recipient_pattern_exists");
            if (val) {
                setRecipientPatternAlreadyExists(val);
                clearInterval(patternPeekTimer);
            }
        }, 500);
        return () => {
            clearInterval(patternPeekTimer);
        }
    });

    const dropdownStyles: Partial<IDropdownStyles> = {
        label: { fontSize: sizeDict["label"] },
        title: { fontSize: sizeDict["dropdownOptionText"] },
        dropdownOptionText: { fontSize: sizeDict["dropdownOptionText"] },
        dropdown: { width: 300 },
    };
    const sendImmediatelyKey = 'delay_send_0';
    const options: IDropdownOption[] = [
        { key: sendImmediatelyKey, text: 'すぐに送信する', data: 0 },
        { key: 'delay_send_div', text: '-', itemType: DropdownMenuItemType.Divider },
        { key: 'delay_send_1', text: '1 分後に送信する', data: 1 },
        { key: 'delay_send_3', text: '3 分後に送信する', data: 3 },
        { key: 'delay_send_5', text: '5 分後に送信する', data: 5 },
    ];
    const isDelaySendEnabled: boolean = isOnline;
    const differentDomainCount = Object.keys(recipientList.reduce((prev, cur) => {
        if (cur.isExternal) {
            prev[Common.getDomainFromEmailAddress(cur.address)] = true;
        }
        return prev;
    }, {})).length;

    const updateYesButtonDisabled = (whichCheck: CheckStage) => {
        const totalCheckCount = whichCheck === CheckStage.First ? recipientList.length + attachmentList.length : doubleCheckRecipientList.length;
        const checks = whichCheck == CheckStage.First ? attachmentList.concat(recipientList).filter(v => v.checked) : doubleCheckRecipientList.filter(v => v.checked);
        if (checks.length == totalCheckCount) {
            if (whichCheck === CheckStage.First) {
                setYesDisabled(false);
            }
            else {
                setDoubleCheckYesDisabled(false);
            }
        }
        else {
            if (whichCheck === CheckStage.First) {
                setYesDisabled(true);
            }
            else {
                setDoubleCheckYesDisabled(true);
            }
        }
    }

    const checkboxChangedEventHandlerCreator = (_: boolean, toggler: () => void) =>
        (ev: React.FormEvent<HTMLElement | HTMLInputElement>, checked: boolean) => {
            const candClassList = ev.currentTarget.parentElement.classList;
            const whichCheck: CheckStage = candClassList.contains("fcsCheck") ? CheckStage.First :
                candClassList.contains("fcsDoubleCheck") ? CheckStage.Second :
                    CheckStage.None;
            if (whichCheck !== CheckStage.None) {
                const identity = ev.currentTarget.id;
                attachmentList.concat(recipientList).concat(doubleCheckRecipientList).forEach(v => {
                    if (v.identity == identity) {
                        v.checked = checked;
                        toggler();
                    }
                });
                updateYesButtonDisabled(whichCheck);
            }
        };

    const delaySendKey = window.localStorage.getItem(Common.SettingKeyDelaySendKey);

    const onClose = async function (isYes: boolean) {
        let msg = isYes ? "<YES>" : Common.DialogMessageCancel;
        if (isYes && differentDomainCount > 0 && attachmentList.length > 0) {
            msg += "<ENCRYPT>";
        }
        if (isYes) {
            await Common.promiseSaveRoamingSetting();
        }
        Office.context.ui.messageParent(msg);
    }

    const debugInformationForDesktop = window.localStorage.getItem('debug_desktop');

    return (
        <Stack
            style={{
                height: "100%",
                display: "flex",
                flexFlow: "column"
            }}
        >
            <Dialog
                hidden={isDoubleCheckModalHidden}
                isBlocking={false}
                isDarkOverlay={false}
                minWidth="85%"
                maxWidth="85%"
                styles={{
                    root: {
                        display: "flex",
                        flexFlow: "column",
                    },
                }}

                dialogContentProps={{
                    type: DialogType.normal,
                    title: "外部ドメインの確認",
                }}
            >
                <Label style={{ fontSize: sizeDict["label"] }}>複数の異なるドメインが含まれています。間違いありませんか ?</Label>
                <Stack style={{ width: "100%", height: "320px", display: "flex", flexFlow: "column", overflow: "hidden" }}>
                    <CheckList items={doubleCheckRecipientList} columnParams={{
                        className: "fcsDoubleCheck",
                        isAttachment: false,
                        handlerCreator: checkboxChangedEventHandlerCreator
                    }} />
                </Stack>
                <DialogFooter>
                    <PrimaryButton
                        text="送信"
                        styles={{
                            textContainer: {
                                fontSize: sizeDict["label"]
                            }
                        }}
                        onClick={async () => {
                            hideDoubleCheckModal();
                            await onClose(true);
                        }}
                        disabled={isDoubleCheckYesDisabled} />
                    <DefaultButton
                        text="キャンセル"
                        styles={{
                            textContainer: {
                                fontSize: sizeDict["label"]
                            }
                        }}
                        onClick={() => {
                            hideDoubleCheckModal();
                        }}
                        disabled={false} />
                </DialogFooter>
            </Dialog>
            <Stack>
                <TimeoutMessageBar addinActivatedDate={addinActivatedDate} />
                {
                    !recipientPatternAlreadyExists ?
                        MyMessageBar(<><Icon iconName="Robot" /><ProgressIndicator description="受信者の組み合わせを送信済みアイテムと照合しています。結果を待たずに送信することもできます。"></ProgressIndicator></>, MessageBarType.info) :
                        (recipientPatternAlreadyExists === "yes" ?
                            MyMessageBar(<><Icon iconName="Robot" /><span>このメールの受信者の組み合わせは、過去に送信したことがあります。</span></>, MessageBarType.success) :
                            MyMessageBar(<><Icon iconName="Robot" /><span>このメールの受信者の組み合わせは、過去に送信したことがありません。</span></>, MessageBarType.severeWarning))
                }
                {differentDomainCount > 1 ? MyMessageBar(<span>受信者に送信者と異なるドメインが<strong> 2 種類以上</strong>含まれています。</span>, MessageBarType.severeWarning) : null}
                {differentDomainCount == 1 ? MyMessageBar(<span>受信者に送信者と異なるドメインが 1 種類だけ含まれています。</span>, MessageBarType.warning) : null}
                {differentDomainCount == 0 ? MyMessageBar(<span>受信者と送信者は全て同じドメインです。</span>, MessageBarType.info) : null}
                {(isDelaySendEnabled && (selectedItem ? selectedItem.key : delaySendKey) !== sendImmediatelyKey) ?
                    MyMessageBar(<span>このメールは遅延送信されます。</span>, MessageBarType.info) :
                    MyMessageBar(<span>このメールはすぐに送信されます。</span>, MessageBarType.info)}
                {(debugInformationForDesktop && debugInformationForDesktop.length) > 0 ? <div>{debugInformationForDesktop}</div> : null}
            </Stack>
            <Pivot aria-label="mail_details" style={{
                flex: "1",
                borderBottom: "1px solid #ccc",
                borderTop: "1px solid #ccc",
                overflow: "hidden",
                display: "flex",
                flexFlow: "column",
            }} styles={{
                itemContainer: {
                    display: "flex",
                    flexFlow: "column",
                    flex: "1",
                    overflow: "hidden",
                },
                text: {
                    fontSize: sizeDict["label"]
                }
            }}>
                <PivotItem headerText="受信者の一覧"
                    style={{
                        height: "100%",
                        paddingLeft: "16px",
                        paddingRight: "16px"
                    }}
                    itemIcon="People">
                    <Stack horizontal style={{
                        height: "100%",
                    }}>
                        <Stack style={{
                            width: attachmentList.length > 0 ? "60%" : "100%",
                            display: "flex",
                            flexFlow: "column",
                            height: "100%",
                        }}>
                            <CheckList items={recipientList} columnParams={{
                                className: "fcsCheck",
                                isAttachment: false,
                                handlerCreator: checkboxChangedEventHandlerCreator
                            }} />
                        </Stack>
                        {attachmentList.length > 0 ?
                            <Stack style={{
                                width: "40%",
                                display: "flex",
                                flexFlow: "column",
                                height: "100%",
                            }}>
                                <Stack style={{ flex: 1 }}>
                                    <CheckList items={attachmentList} columnParams={{
                                        className: "fcsCheck",
                                        isAttachment: true,
                                        handlerCreator: checkboxChangedEventHandlerCreator
                                    }} />
                                </Stack>
                            </Stack> : null
                        }
                    </Stack>
                </PivotItem>
                <PivotItem headerText="件名と本文"
                    style={{
                        display: "flex",
                        height: "100%",
                        flexFlow: "column",
                        overflow: "hidden",
                        paddingTop: "16px", // DetailsListと合わせる
                        paddingBottom: "16px",
                        paddingRight: "10%",
                        paddingLeft: "10%",
                    }}
                    itemIcon="EditMail">
                    <Stack style={{
                        paddingBottom: "4px"
                    }}>
                        {MyLabel(<span>件名</span>)}
                    </Stack>
                    <Stack style={{
                        fontSize: sizeDict["subject"],
                        border: "1px solid #ccc",
                        borderRadius: "10px",
                        padding: "8px"
                    }}>
                        {subject}
                    </Stack>
                    <Stack style={{
                        paddingBottom: "4px"
                    }}>
                        {
                            MyLabel(
                                <div>
                                    <span>本文</span>
                                </div>
                            )
                        }
                    </Stack>
                    <Stack style={{
                        flex: "1",
                        overflowY: "auto",
                    }}>
                        {<textarea value={body} readOnly={true} style={{
                            resize: "none",
                            userSelect: "none",
                            fontFamily: "inherit",
                            height: "100%",
                            border: "1px solid #ccc",
                            borderRadius: "10px",
                            padding: "8px",
                            outline: "none",
                            fontSize: sizeDict["body"]
                        }} />}
                    </Stack>
                </PivotItem>
                <PivotItem headerText="その他" itemIcon="Settings" style={{ padding: "16px" }}>
                    <Stack>
                        <Dropdown
                            disabled={!isDelaySendEnabled}
                            options={options}
                            label="遅延送信オプション"
                            styles={dropdownStyles}
                            onChange={(_, item) => {
                                setSelectedItem(item);
                                window.localStorage.setItem(Common.SettingKeyDelaySendKey, item.key.toString());
                            }}
                            selectedKey={selectedItem ? selectedItem.key : delaySendKey}
                        />
                    </Stack>
                </PivotItem>
            </Pivot>
            <Stack horizontal verticalAlign="center" horizontalAlign="end" style={{
                paddingTop: "16px",
                paddingBottom: "16px",
            }}>
                <PrimaryButton
                    styles={{
                        textContainer: {
                            fontSize: sizeDict["label"]
                        }
                    }}
                    style={{
                        marginRight: "16px"
                    }}
                    text={differentDomainCount > 1 ? "確認" : "送信"}
                    onClick={async () => {
                        if (differentDomainCount > 1) {
                            doubleCheckRecipientList.forEach(v => {
                                v.checked = false;
                            });
                            setDoubleCheckYesDisabled(true);
                            showDoubleCheckModal();
                        }
                        else {
                            await onClose(true);
                        }
                    }}
                    disabled={isYesDisabled} />
                <DefaultButton
                    styles={{
                        textContainer: {
                            fontSize: sizeDict["label"]
                        }
                    }}
                    style={{
                        marginRight: "16px"
                    }}
                    text="キャンセル"
                    onClick={() => {
                        onClose(false);
                    }}
                    disabled={false} />
            </Stack>
        </Stack>
    );
}

const render = async () => {
    try {
        initializeIcons();
        const dateTime = new Date(window.localStorage.getItem("time_activated"));
        const itemAttachments: Common.ICustomAttachment[] = JSON.parse(window.localStorage.getItem("mail_attachments"));
        const itemSubject = window.localStorage.getItem("mail_subject");
        const itemBody = window.localStorage.getItem("mail_body");
        const recipientList: Common.ICustomRecipient[] = JSON.parse(window.localStorage.getItem("mail_recipients"));
        const attachmentList = itemAttachments.map((att) => {
            let catt: Common.ICustomAttachment = {
                id: att.id,
                name: att.name,
                data: null,
                size: att.size
            };
            return catt;
        });

        const recipientDataList = recipientList.map((val, idx) => {
            return convertRecipientToListItem(idx, val, false);
        });

        const doubleCheckRecipientDataList = recipientList.filter(val => {
            return val.isExternal;
        }).map((val, idx) => {
            return convertRecipientToListItem(idx, val, true);
        });

        const sizeExpression = (size) => {
            let ret = "";
            if (size >= 1024 * 1024) {
                ret = (size / (1024 * 1024)).toFixed(1) + "MB";
            }
            else {
                ret = (size / 1024).toFixed() + "KB";
            }
            return ret;
        };

        const attachmentDataList = attachmentList.map((att, idx) => {
            let cli: Common.ICheckListItem = {
                key: idx,
                recipientType: "",
                size: sizeExpression(att.size),
                name: att.name,
                message: "",
                address: '',
                identity: 'fcs_attachment_' + idx,
                isExternal: false,
                isPrivateDL: false,
                isAttachment: true,
                checked: false,
                disabled: false
            };
            return cli;
        });

        ReactDOM.render(
            <CheckModal
                recipientList={recipientDataList}
                attachmentList={attachmentDataList}
                doubleCheckRecipientList={doubleCheckRecipientDataList}
                body={itemBody}
                subject={itemSubject}
                addinActivatedDate={dateTime}
            />,
            document.getElementById("container")
        );
    }
    catch (err) {
        console.error(err);
    }
};

Office.onReady(async () => {
    await render();
});
