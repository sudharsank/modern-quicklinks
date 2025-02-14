import * as React from 'react';
import { useEffect, useContext, useReducer } from 'react';
import tileStyle from '../Tile/Tile.module.scss';
import * as strings from 'ModernQuickLinksWebPartStrings';
import { css } from 'office-ui-fabric-react/lib/Utilities';
import { IconPicker } from '@pnp/spfx-controls-react/lib/IconPicker';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { Stack, IStackStyles } from 'office-ui-fabric-react/lib/Stack';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import AppContext, { IAppContextProps } from '../../../../common/AppContext';
import { IKLists, LoaderType, MessageScope } from '../../../../common/Constants';
import MessageContainer from '../../../../common/components/Message';
import ContentLoader from '../../../../common/components/ContentLoader';
import { isValidURL } from '../../../../common/util';
import { useLaunchPadHelper } from '../../../../common/Helpers/useLaunchPadHelper';
import { FilePicker, IFilePickerResult } from '@pnp/spfx-controls-react/lib/FilePicker';
import { BaseWebPartContext } from '@microsoft/sp-webpart-base';
import { useCommonHelper } from '../../../../common/Helpers/useCommonHelper';
import { Image } from '@fluentui/react/lib/Image';
import { isNumber } from 'lodash';

export interface IAddUpdateTileProps {
    myTilesCallback?: () => void;
    onDismissPanel: () => void;
    item?: any;
    customTileCount: number;
    userTileList: string;
    context: BaseWebPartContext;
}

export interface IAddUpdateTileState {
    tileTitle: string;
    tileURL: string;
    tileDesc: string;
    tileSeq: number;
    newWindow: boolean;
    isIcon: boolean;
    iconName: string;
    iconImage: IFilePickerResult;
    message: string;
    messageScope: MessageScope;
    saving: boolean;
    isEdit: boolean;
    disableSave: boolean;
}

const stackTokens = { childrenGap: 10 };
const stackStyles: Partial<IStackStyles> = { root: { width: '100%' } };

const initialState: IAddUpdateTileState = {
    tileTitle: '',
    tileURL: '',
    tileDesc: '',
    tileSeq: undefined,
    newWindow: false,
    isIcon: true,
    iconName: '',
    iconImage: undefined,
    message: '',
    messageScope: MessageScope.Info,
    saving: false,
    isEdit: false,
    disableSave: false
};

const addUpdateReducer = (state: IAddUpdateTileState, action: any): IAddUpdateTileState => {
    switch (action.type) {
        case 'RESET':
            return initialState;
        case 'SET_TILE_TITLE':
            return { ...state, tileTitle: action.payload };
        case 'SET_TILE_URL':
            return { ...state, tileURL: action.payload };
        case 'SET_TILE_DESC':
            return { ...state, tileDesc: action.payload };
        case 'SET_TILE_SEQ':
            return { ...state, tileSeq: action.payload };
        case 'SET_NEW_WINDOW':
            return { ...state, newWindow: action.payload };
        case 'SET_IS_ICON':
            return { ...state, isIcon: action.payload };
        case 'SET_ICON_IMAGE':
            return { ...state, iconImage: action.payload };
        case 'SET_ICON_NAME':
            return { ...state, iconName: action.payload };
        case 'CLEAR_MSG':
            return { ...state, message: '', messageScope: MessageScope.Info };
        case 'SET_MSG':
            return { ...state, message: action.payload.message, messageScope: action.payload.messageScope };
        case 'SET_SAVING':
            return { ...state, saving: action.payload };
        case 'SET_EDIT':
            return { ...state, isEdit: action.payload };
        case 'SET_DISABLE_SAVE':
            return { ...state, disableSave: true };
        case 'SAVE_SUCCESS':
            return {
                ...state, message: action.payload.message, messageScope: action.payload.messageScope, saving: action.payload.saving,
                isEdit: action.payload.isEdit, disableSave: action.payload.disableSave
            };
        case 'EDIT_ITEM':
            return {
                ...state, isEdit: action.payload.isEdit, tileTitle: action.payload.tileTitle, tileURL: action.payload.tileURL,
                tileDesc: action.payload.tileDesc, newWindow: action.payload.newWindow, iconName: action.payload.iconName,
                isIcon: action.payload.isIcon, tileSeq: action.payload.tileSeq,
                iconImage: action.payload.iconImage?.serverRelativeUrl ? {
                    fileAbsoluteUrl: action.payload.iconImage.serverRelativeUrl,
                    fileName: action.payload.iconImage.fileName,
                    fileNameWithoutExtension: '',
                    downloadFileContent: undefined
                } : undefined
            };
    }
}

const AddUpdateTile: React.FC<IAddUpdateTileProps> = (props) => {
    const appContext: IAppContextProps = useContext(AppContext);
    const [state, dispatch] = useReducer(addUpdateReducer, initialState);
    const { manageTiles, getTilesCount } = useLaunchPadHelper(appContext.spService);
    const { returnImageFieldInfoForWrite } = useCommonHelper(appContext.spService);

    const _changeTileTitle = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => dispatch({ type: 'SET_TILE_TITLE', payload: newValue });
    const _changeTileURL = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => dispatch({ type: 'SET_TILE_URL', payload: newValue });
    const _changeTileDesc = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => dispatch({ type: 'SET_TILE_DESC', payload: newValue });
    const _changeTileSeq = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => dispatch({ type: 'SET_TILE_SEQ', payload: newValue });
    const _changeNewWindow = (event: React.MouseEvent<HTMLElement, MouseEvent>, checked?: boolean) => dispatch({ type: 'SET_NEW_WINDOW', payload: checked });
    const _changeIsIcon = (event: React.MouseEvent<HTMLElement, MouseEvent>, checked?: boolean) => dispatch({ type: 'SET_IS_ICON', payload: checked });
    const getIconImageInfo = (filePickerResult: IFilePickerResult[]) => {
        if (filePickerResult && filePickerResult.length > 0) {
            dispatch({ type: 'SET_ICON_IMAGE', payload: filePickerResult[0] });
        }
    };
    /**
     * Once saved clear the controls in the new tile form
     */
    const _clearControls = () => {
        dispatch({ type: 'RESET' });
    };
    /**
     * Save the tile to the My Launch pad list
     */
    const _saveTileLink = async () => {
        const { tileTitle, tileURL, tileSeq, isIcon, iconImage, tileDesc, iconName, isEdit, newWindow } = state;
        dispatch({ type: 'CLEAR_MSG' });
        if (!tileTitle || tileTitle.trim().length <= 0 || !tileURL && tileURL.trim().length <= 0) {
            dispatch({ type: 'SET_MSG', payload: { message: strings.Msg_Mandatory, messageScope: MessageScope.Failure } });
        } else {
            if (!isValidURL(tileURL.trim())) {
                dispatch({ type: 'SET_MSG', payload: { message: strings.Msg_InvalidURL, messageScope: MessageScope.Failure } });
            } else {
                if (!isIcon && !iconImage) {
                    dispatch({ type: 'SET_MSG', payload: { message: strings.Msg_ReqImage, messageScope: MessageScope.Failure } });
                } else {
                    dispatch({ type: 'SET_SAVING', payload: true });
                    let icoImage: string = undefined;
                    if (!isIcon && iconImage && iconImage.downloadFileContent)
                        icoImage = await returnImageFieldInfoForWrite(iconImage, `${IKLists.Asset}/ModernQuickLinks`, appContext.serRelativeUrl);
                    let totalTilesLength: number = 0;
                    if (!isEdit) {
                        totalTilesLength = await getTilesCount('', props.userTileList);
                        if (totalTilesLength < props.customTileCount) {
                            await manageTiles({
                                Title: tileTitle.trim(),
                                URL: { Url: tileURL.trim(), Description: tileDesc.trim() },
                                Sequence: isNumber(tileSeq) ? tileSeq : undefined,
                                IconName: iconName,
                                IconImage: icoImage,
                                NewWindow: newWindow,
                                IsActive: true
                            }, '', props.userTileList);
                        } else {
                            dispatch({ type: 'SET_MSG', payload: { message: strings.ItemsLimitAdd, messageScope: MessageScope.Info } });
                        }
                    } else {
                        if (props.item) {
                            if (isEdit && iconImage && iconImage.downloadFileContent) {
                                await manageTiles({
                                    Id: props.item.ID,
                                    Title: tileTitle.trim(),
                                    URL: { Url: tileURL.trim(), Description: tileDesc.trim() },
                                    Sequence: isNumber(tileSeq) ? tileSeq : undefined,
                                    IconName: iconName,
                                    IconImage: !isIcon ? icoImage : null,
                                    NewWindow: newWindow,
                                }, '', props.userTileList);
                            } else {
                                await manageTiles({
                                    Id: props.item.ID,
                                    Title: tileTitle.trim(),
                                    URL: { Url: tileURL.trim(), Description: tileDesc.trim() },
                                    Sequence: isNumber(tileSeq) ? tileSeq : undefined,
                                    IconName: iconName,
                                    IconImage: null,
                                    NewWindow: newWindow,
                                }, '', props.userTileList);
                            }
                        } else {
                            await manageTiles({
                                Title: tileTitle.trim(),
                                URL: { Url: tileURL.trim(), Description: tileDesc.trim() },
                                Sequence: isNumber(tileSeq) ? tileSeq : undefined,
                                IconName: iconName,
                                IconImage: icoImage,
                                NewWindow: newWindow,
                                IsActive: true
                            }, '', props.userTileList);
                        }
                    }
                    totalTilesLength = await getTilesCount('', props.userTileList);
                    dispatch({
                        type: 'SAVE_SUCCESS', payload: {
                            message: strings.Msg_SaveSuccess, messageScope: MessageScope.Success,
                            saving: false, isEdit: false, disableSave: totalTilesLength >= props.customTileCount ? true : false
                        }
                    });
                    _clearControls();
                    props.myTilesCallback();
                }
            }
        }
    };

    useEffect(() => {
        if (props.item) {
            console.log("Edit: ", props.item);
            if (props.item.IconImage) {
                let icoImage = JSON.parse(props.item.IconImage);
                dispatch({
                    type: 'EDIT_ITEM', payload: {
                        isEdit: true, tileTitle: props.item.Title, tileURL: props.item.URL.Url, isIcon: false,
                        tileDesc: props.item.URL.Description, newWindow: props.item.NewWindow, iconName: props.item.IconName,
                        tileSeq: props.item.Sequence,
                        iconImage: {
                            fileAbsoluteUrl: icoImage.serverRelativeUrl,
                            fileName: icoImage.fileName,
                            fileNameWithoutExtension: '',
                            downloadFileContent: undefined
                        }
                    }
                });
            } else {
                dispatch({
                    type: 'EDIT_ITEM', payload: {
                        isEdit: true, tileTitle: props.item.Title, tileURL: props.item.URL.Url, isIcon: true,
                        tileDesc: props.item.URL.Description, newWindow: props.item.NewWindow, iconName: props.item.IconName,
                        tileSeq: props.item.Sequence
                    }
                });
            }

            // setIsEdit(true);
            // settileTitle(props.item.Title);
            // settileURL(props.item.URL.Url);
            // settileDesc(props.item.URL.Description);
            // setnewWindow(props.item.NewWindow);
            // seticonName(props.item.IconName);
            // if (props.item.IconImage) {
            //     let icoImage = JSON.parse(props.item.IconImage);
            //     setIconImage({
            //         fileAbsoluteUrl: icoImage.serverRelativeUrl,
            //         fileName: icoImage.fileName,
            //         fileNameWithoutExtension: '',
            //         downloadFileContent: undefined
            //     });
            //     setIsIcon(false);
            // }
        }
    }, [props.item]);

    useEffect(() => {
        if (state.disableSave) {
            setTimeout(() => {
                dispatch({ type: 'SET_MSG', payload: { message: strings.ItemsLimitAdd, messageScope: MessageScope.Info } });
            }, 1000);
        }
    }, [state.disableSave]);

    return (
        <Stack tokens={stackTokens} styles={stackStyles}>
            <Stack>
                {state.message &&
                    <MessageContainer MessageScope={state.messageScope} Message={state.message} />
                }
            </Stack>
            <Stack className={tileStyle.divFormContainer}>
                <div className={tileStyle.divLabel}>
                    <Icon iconName={"TextField"} />
                    <label className={css(tileStyle.fieldTitle, tileStyle.fieldTitleReq)}>{strings.Fld_lbl_Title}</label>
                </div>
                <div className={tileStyle.divField}>
                    <TextField maxLength={20} placeholder={`${strings.Fld_lbl_Title}...`} disabled={state.saving}
                        description={strings.Allowed20Chars} value={state.tileTitle} onChange={_changeTileTitle} />
                </div>
            </Stack>
            <Stack className={tileStyle.divFormContainer}>
                <div className={tileStyle.divLabel}>
                    <Icon iconName={"Link"} />
                    <label className={css(tileStyle.fieldTitle, tileStyle.fieldTitleReq)}>{strings.Fld_lbl_Url}</label>
                </div>
                <div className={tileStyle.divField}>
                    <TextField placeholder={`${strings.Fld_lbl_Url}...`} disabled={state.saving}
                        value={state.tileURL} onChange={_changeTileURL} multiline={true} />
                </div>
            </Stack>
            <Stack className={tileStyle.divFormContainer}>
                <div className={tileStyle.divLabel}>
                    <Icon iconName={"AlignLeft"} />
                    <label className={css(tileStyle.fieldTitle)}>{strings.Fld_lbl_Desc}</label>
                </div>
                <div className={tileStyle.divField}>
                    <TextField maxLength={30} placeholder={`${strings.Fld_lbl_Desc}...`} disabled={state.saving}
                        description={strings.Allowed30Chars} value={state.tileDesc} onChange={_changeTileDesc} />
                </div>
            </Stack>
            <Stack className={tileStyle.divFormContainer}>
                <div className={tileStyle.divLabel}>
                    <Icon iconName={"AlignLeft"} />
                    <label className={css(tileStyle.fieldTitle)}>{strings.Fld_lbl_Seq}</label>
                </div>
                <div className={tileStyle.divField}>
                    <TextField maxLength={3} inputMode='numeric' placeholder={`${strings.Fld_lbl_Seq}...`} disabled={state.saving}
                        value={state.tileSeq?.toString()} onChange={_changeTileSeq} />
                </div>
            </Stack>
            <Stack className={tileStyle.divFormContainer}>
                <div className={tileStyle.divLabel}>
                    <Icon iconName={"OpenInNewTab"} />
                    <label className={css(tileStyle.fieldTitle)}>{strings.Fld_lbl_NewWin}</label>
                </div>
                <div className={tileStyle.divField}>
                    <Toggle onText="On" offText="Off" onChange={_changeNewWindow} checked={state.newWindow} disabled={state.saving} />
                </div>
            </Stack>
            <Stack className={tileStyle.divFormContainer}>
                <div className={tileStyle.divLabel}>
                    <Icon iconName={"Photo2"} />
                    <label className={css(tileStyle.fieldTitle)}>{strings.Fld_lbl_IcoImg}</label>
                </div>
                <div className={tileStyle.divField}>
                    <Toggle onText="Icon" offText="Image" onChange={_changeIsIcon} checked={state.isIcon} disabled={state.saving} />
                </div>
            </Stack>
            {state.isIcon ? (
                <Stack className={tileStyle.divFormContainer}>
                    <div className={tileStyle.divLabel}>
                        <Icon iconName={"AppIconDefault"} />
                        <label className={css(tileStyle.fieldTitle)}>{strings.Fld_lbl_Icon}</label>
                    </div>
                    <div className={tileStyle.divField} style={{ display: 'inline-flex' }}>
                        <span><Icon iconName={state.iconName ? state.iconName : "PageLink"} style={{ fontSize: '20px', marginRight: '10px' }} /></span>
                        <IconPicker buttonLabel={'Icon Gallery'} buttonClassName={tileStyle.galleryButton} disabled={state.saving}
                            onChange={(icon: string) => { dispatch({ type: 'SET_ICON_NAME', payload: icon }); }}
                            onSave={(icon: string) => { dispatch({ type: 'SET_ICON_NAME', payload: icon }); }} />
                    </div>
                </Stack>
            ) : (
                <Stack className={tileStyle.divFormContainer}>
                    <div className={tileStyle.divLabel}>
                        <Icon iconName={"AppIconDefault"} />
                        <label className={css(tileStyle.fieldTitle)}>{strings.Fld_lbl_Img}</label>
                    </div>
                    <div className={tileStyle.divField} style={{ display: 'inline-flex' }}>
                        {state.iconImage && state.iconImage.fileAbsoluteUrl &&
                            <span>
                                <Image src={state.iconImage.fileAbsoluteUrl} width={100} />
                            </span>
                        }
                        <FilePicker
                            accepts={[".gif", ".jpg", ".jpeg", ".bmp", ".png"]}
                            buttonIcon="FileImage"
                            buttonLabel='Select Image'
                            onSave={(filePickerResult: IFilePickerResult[]) => { getIconImageInfo(filePickerResult); console.log(filePickerResult); }}
                            context={props.context as any}
                        />
                    </div>
                </Stack>
            )}
            <div className={tileStyle.formActions}>
                <PrimaryButton onClick={_saveTileLink} disabled={state.saving || state.disableSave} primaryDisabled={state.saving || state.disableSave}>
                    <Icon iconName={"Save"} />&nbsp;{strings.Btn_Save}
                </PrimaryButton>
                <DefaultButton onClick={props.onDismissPanel} disabled={state.saving}>
                    <Icon iconName={"Blocked"} />&nbsp;{strings.Btn_Cancel}
                </DefaultButton>
                {state.saving &&
                    <div className={tileStyle.actionLoader}><ContentLoader loaderType={LoaderType.Spinner} spinSize={SpinnerSize.medium} /></div>
                }
            </div>
        </Stack>
    );
};

export default AddUpdateTile;