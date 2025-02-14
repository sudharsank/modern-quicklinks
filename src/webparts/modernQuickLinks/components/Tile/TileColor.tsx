import * as React from 'react';
import tileStyle from '../Tile/Tile.module.scss';
import { FC, useEffect, useReducer, useContext, useState } from 'react';
import { DesignTypes, fld_tileColors, LoaderType, MessageScope } from '../../../../common/Constants';
import MessageContainer from '../../../../common/components/Message';
import { IStackStyles, Stack } from '@fluentui/react/lib/Stack';
import { css, DefaultButton, Icon, IconButton, mergeStyleSets, PrimaryButton, SpinnerSize } from '@fluentui/react';
import strings from 'ModernQuickLinksWebPartStrings';
import ContentLoader from '../../../../common/components/ContentLoader';
import AppContext, { IAppContextProps } from '../../../../common/AppContext';
import { useLaunchPadHelper } from '../../../../common/Helpers';
import { ColorPicker, IColorPickerStyles, } from '@fluentui/react';

const stackTokens = { childrenGap: 10 };
const stackStyles: Partial<IStackStyles> = { root: { width: '100%' } };

export interface ITileColorProps {
    globalList: string;
    item: any;
    designType: DesignTypes;
    tileColors?: any | undefined;
    onDismissPanel: () => void;
    myTilesCallback?: () => void;
}

export interface ITileColorState {
    message: string;
    messageScope: string;
    saving: boolean;
    loading: boolean;
    disableControls: boolean;
    displayBGColorPicker: boolean;
    backgroundColor: string;
    displayFontColorPicker: boolean;
    fontColor: string;
    displayOverflowColorPicker: boolean;
    overflowBackgroundColor: string;
    displayOverflowFontColorPicker: boolean;
    overflowFontColor: string;
    displayActionIconColorPicker: boolean;
    actionIconColor: string;
    displayActionIconHoverColorPicker: boolean;
    actionIconHoverColor: string;
}

const initialState: ITileColorState = {
    message: undefined,
    messageScope: '',
    saving: false,
    loading: false,
    disableControls: false,
    displayBGColorPicker: false,
    backgroundColor: undefined,
    displayFontColorPicker: false,
    fontColor: undefined,
    displayOverflowColorPicker: false,
    overflowBackgroundColor: undefined,
    displayOverflowFontColorPicker: false,
    overflowFontColor: undefined,
    displayActionIconColorPicker: false,
    actionIconColor: undefined,
    displayActionIconHoverColorPicker: false,
    actionIconHoverColor: undefined
};

const tileColorReducer = (state: ITileColorState, action: any) => {
    switch (action.type.toUpperCase()) {
        case 'RESET':
            return initialState;
        case 'CLEAR_MSG':
            return { ...state, message: '', messageScope: MessageScope.Info };
        case 'SET_MSG':
            return { ...state, message: action.payload.message, messageScope: action.payload.messageScope };
        case 'SET_SAVING':
            return { ...state, saving: action.payload };
        case 'OPEN_CP':
            if (action.payload.toLowerCase() === 'bgcolor') return { ...state, displayBGColorPicker: true };
            else if (action.payload.toLowerCase() === 'fontcolor') return { ...state, displayFontColorPicker: true };
            else if (action.payload.toLowerCase() === 'overflowcolor') return { ...state, displayOverflowColorPicker: true };
            else if (action.payload.toLowerCase() === 'overflowfontcolor') return { ...state, displayOverflowFontColorPicker: true };
            else if (action.payload.toLowerCase() === 'actioniconcolor') return { ...state, displayActionIconColorPicker: true };
            else if (action.payload.toLowerCase() === 'actioniconhovercolor') return { ...state, displayActionIconHoverColorPicker: true };
            break;
        case 'CLOSE_CP':
            if (action.payload.toLowerCase() === 'bgcolor') return { ...state, displayBGColorPicker: false };
            else if (action.payload.toLowerCase() === 'fontcolor') return { ...state, displayFontColorPicker: false };
            else if (action.payload.toLowerCase() === 'overflowcolor') return { ...state, displayOverflowColorPicker: false };
            else if (action.payload.toLowerCase() === 'overflowfontcolor') return { ...state, displayOverflowFontColorPicker: false };
            else if (action.payload.toLowerCase() === 'actioniconcolor') return { ...state, displayActionIconColorPicker: false };
            else if (action.payload.toLowerCase() === 'actioniconhovercolor') return { ...state, displayActionIconHoverColorPicker: false };
            break;
        case 'SET_COLOR':
            if (action.payload.scope.toLowerCase() === 'bgcolor') return { ...state, backgroundColor: action.payload.color };
            else if (action.payload.scope.toLowerCase() === 'fontcolor') return { ...state, fontColor: action.payload.color };
            else if (action.payload.scope.toLowerCase() === 'overflowcolor') return { ...state, overflowBackgroundColor: action.payload.color };
            else if (action.payload.scope.toLowerCase() === 'overflowfontcolor') return { ...state, overflowFontColor: action.payload.color };
            else if (action.payload.scope.toLowerCase() === 'actioniconcolor') return { ...state, actionIconColor: action.payload.color };
            else if (action.payload.scope.toLowerCase() === 'actioniconhovercolor') return { ...state, actionIconHoverColor: action.payload.color };
            break;
        case 'SET_ALL_COLOR':
            return {
                ...state, backgroundColor: action.payload.backgroundColor, fontColor: action.payload.fontColor, overflowBackgroundColor: action.payload.overflowBackgroundColor,
                overflowFontColor: action.payload.overflowFontColor, actionIconColor: action.payload.actionIconColor, actionIconHoverColor: action.payload.actionIconHoverColor
            };
        case 'CLOSE_ALL_CP':
            return {
                ...state, displayBGColorPicker: false, displayFontColorPicker: false, displayOverflowColorPicker: false,
                displayOverflowFontColorPicker: false, displayActionIconColorPicker: false, displayActionIconHoverColorPicker: false
            };
        default:
            return state;
    }
};

const TileColor: FC<ITileColorProps> = (props) => {
    const appContext: IAppContextProps = useContext(AppContext);
    const [state, dispatch] = useReducer(tileColorReducer, initialState);
    const { checkForField, createField, manageTiles, getItemById } = useLaunchPadHelper(appContext.spService);
    const tileColorStyles = mergeStyleSets({
        swatch: {
            padding: '5px',
            background: '#fff',
            borderRadius: '1px',
            boxShadow: '0 0 0 1px rgba(0,0,0,.1)',
            display: 'inline-block',
            width: '95%',
            border: '1px solid'
        },
        color: {
            height: '14px',
            borderRadius: '2px'
        },
        pickerIcon: {
            cursor: 'pointer',
            verticalAlign: 'middle',
            fontSize: '20px',
            marginTop: '-2px',
            marginLeft: '-2px'
        }
    });
    const colorPickerStyles: Partial<IColorPickerStyles> = {
        panel: { padding: 12 },
        root: {
            maxWidth: 300,
            minWidth: 300,
        },
        colorRectangle: { height: 250 },
    };

    const _setColor = (scope: string, newColor: string) => dispatch({ type: 'set_color', payload: { scope: scope, color: newColor } });

    const _openColorPicker = (scope: string) => dispatch({ type: 'open_cp', payload: scope });

    const _closeColorPicker = (scope: string) => dispatch({ type: 'close_cp', payload: scope });

    const _getSavedTileColors = async () => {
        dispatch({ type: 'SET_SAVING', payload: true });
        dispatch({ type: 'clear_msg' });
        try {
            const tileColors = await getItemById(undefined, props.globalList, props.item?.ID.toString());
            if (tileColors && tileColors.TileColors) {
                const colors = JSON.parse(tileColors.TileColors);
                dispatch({ type: 'set_all_color', payload: colors });
            }
        } catch (err) {
            console.error(err);
        }
        dispatch({ type: 'SET_SAVING', payload: false });
    };

    const _saveTileColor = async () => {
        dispatch({ type: 'clear_msg' });
        dispatch({ type: 'close_all_cp' });
        try {
            const fldExists = await checkForField(undefined, props.globalList, fld_tileColors);
            console.log(fldExists);
            if (!fldExists) await createField(undefined, props.globalList, fld_tileColors, 'multiline', false, false);
            let isValid: boolean = true;
            switch (props.designType) {
                case DesignTypes.Tiles:
                case DesignTypes.Buttons:
                    if (!state.backgroundColor && !state.fontColor && !state.overflowBackgroundColor && !state.overflowFontColor && !state.actionIconColor && !state.actionIconHoverColor) {
                        isValid = false;
                    }
                    break;
                case DesignTypes.Compact:
                    if (!state.backgroundColor && !state.fontColor && !state.overflowBackgroundColor && !state.actionIconColor && !state.actionIconHoverColor) {
                        isValid = false;
                    }
                    break;
                case DesignTypes.Grid:
                    if (!state.backgroundColor && !state.fontColor && !state.overflowFontColor && !state.actionIconColor && !state.actionIconHoverColor) {
                        isValid = false;
                    }
                    break;
            }
            if (isValid) {
                dispatch({ type: 'SET_SAVING', payload: true });
                const finalColors = JSON.stringify({
                    backgroundColor: state.backgroundColor, fontColor: state.fontColor,
                    overflowBackgroundColor: state.overflowBackgroundColor, overflowFontColor: state.overflowFontColor,
                    actionIconColor: state.actionIconColor, actionIconHoverColor: state.actionIconHoverColor
                });
                await manageTiles({
                    Id: props.item.ID,
                    TileColors: finalColors
                }, undefined, props.globalList);
                props.myTilesCallback();
                dispatch({ type: 'SET_MSG', payload: { message: 'Successfully saved the colors', messageScope: MessageScope.Success } });
                dispatch({ type: 'SET_SAVING', payload: false });
            } else {
                dispatch({ type: 'SET_MSG', payload: { message: 'Choose atleast a color to save', messageScope: MessageScope.Failure } });
            }
        } catch (err) {
            console.error(err);
        }
    };

    useEffect(() => {
        if (props.item) {
            (async () => {
                await _getSavedTileColors();
            })();
        } else dispatch({ type: 'SET_MSG', payload: { message: 'No item found', messageScope: MessageScope.Failure } });
    }, [props.item]);

    return (
        <Stack tokens={stackTokens} styles={stackStyles}>
            <Stack>
                {state.message &&
                    <MessageContainer MessageScope={state.messageScope} Message={state.message} />
                }
            </Stack>
            {/* Background Color - Tiles,Buttons,Compact & Grid */}
            <Stack className={tileStyle.divFormContainer}>
                <div className={tileStyle.divLabel}>
                    <Icon iconName={"BucketColor"} />
                    <label className={css(tileStyle.fieldTitle)}>
                        {props.designType === DesignTypes.Compact || props.designType === DesignTypes.Buttons ? 'Border Color' : 'Background Color'}
                    </label>
                </div>
                <div className={tileStyle.divField}>
                    <Stack horizontal horizontalAlign='space-between' tokens={stackTokens}>
                        <Stack.Item grow>
                            <div className={tileColorStyles.swatch}>
                                <div className={tileColorStyles.color} style={{ backgroundColor: state.backgroundColor }} />
                            </div>
                            <div>
                                {state.displayBGColorPicker &&
                                    <>
                                        <ColorPicker color={state.backgroundColor} onChange={(e, color) => _setColor('bgcolor', color.str)} alphaType={'alpha'}
                                            showPreview={true} styles={colorPickerStyles} />
                                    </>
                                }
                            </div>
                        </Stack.Item>
                        <Stack.Item>
                            {!state.displayBGColorPicker ? (
                                <IconButton iconProps={{ iconName: "Color" }} onClick={() => _openColorPicker('bgcolor')} className={tileColorStyles.pickerIcon}
                                    disabled={state.saving || state.loading} />
                            ) : (
                                <IconButton iconProps={{ iconName: "ErrorBadge" }} onClick={() => _closeColorPicker('bgcolor')} className={tileColorStyles.pickerIcon}
                                    disabled={state.saving || state.loading} />
                            )}
                        </Stack.Item>
                    </Stack>
                </div>
            </Stack>
            {/* Font Color - Tiles,Buttons,Compact & Grid */}
            <Stack className={tileStyle.divFormContainer}>
                <div className={tileStyle.divLabel}>
                    <Icon iconName={"BucketColor"} />
                    <label className={css(tileStyle.fieldTitle)}>
                        {'Icon & Text Color'}
                    </label>
                </div>
                <div className={tileStyle.divField}>
                    <Stack horizontal horizontalAlign='space-between' tokens={stackTokens}>
                        <Stack.Item grow>
                            <div className={tileColorStyles.swatch}>
                                <div className={tileColorStyles.color} style={{ backgroundColor: state.fontColor }} />
                            </div>
                            <div>
                                {state.displayFontColorPicker &&
                                    <>
                                        <ColorPicker color={state.fontColor} onChange={(e, color) => _setColor('fontColor', color.str)} alphaType={'alpha'}
                                            showPreview={true} styles={colorPickerStyles} />
                                    </>
                                }
                            </div>
                        </Stack.Item>
                        <Stack.Item>
                            {!state.displayFontColorPicker ? (
                                <IconButton iconProps={{ iconName: "Color" }} onClick={() => _openColorPicker('fontColor')} className={tileColorStyles.pickerIcon}
                                    disabled={state.saving || state.loading} />
                            ) : (
                                <IconButton iconProps={{ iconName: "ErrorBadge" }} onClick={() => _closeColorPicker('fontColor')} className={tileColorStyles.pickerIcon}
                                    disabled={state.saving || state.loading} />
                            )}
                        </Stack.Item>
                    </Stack>
                </div>
            </Stack>
            {/* Font Color - Tiles,Buttons,Compact */}
            <Stack className={tileStyle.divFormContainer}>
                <div className={tileStyle.divLabel}>
                    <Icon iconName={"BucketColor"} />
                    <label className={css(tileStyle.fieldTitle)}>
                        {props.designType === DesignTypes.Compact ? 'Background Color' : props.designType === DesignTypes.Buttons ? 'Hover Background Color' : 'Overflow Background Color'}
                    </label>
                </div>
                <div className={tileStyle.divField}>
                    <Stack horizontal horizontalAlign='space-between' tokens={stackTokens}>
                        <Stack.Item grow>
                            <div className={tileColorStyles.swatch}>
                                <div className={tileColorStyles.color} style={{ backgroundColor: state.overflowBackgroundColor }} />
                            </div>
                            <div>
                                {state.displayOverflowColorPicker &&
                                    <>
                                        <ColorPicker color={state.overflowBackgroundColor} onChange={(e, color) => _setColor('overflowcolor', color.str)} alphaType={'alpha'}
                                            showPreview={true} styles={colorPickerStyles} />
                                    </>
                                }
                            </div>
                        </Stack.Item>
                        <Stack.Item>
                            {!state.displayOverflowColorPicker ? (
                                <IconButton iconProps={{ iconName: "Color" }} onClick={() => _openColorPicker('overflowcolor')} className={tileColorStyles.pickerIcon}
                                    disabled={state.saving || state.loading} />
                            ) : (
                                <IconButton iconProps={{ iconName: "ErrorBadge" }} onClick={() => _closeColorPicker('overflowcolor')} className={tileColorStyles.pickerIcon}
                                    disabled={state.saving || state.loading} />
                            )}
                        </Stack.Item>
                    </Stack>
                </div>
            </Stack>
            {/* Font Color - Tiles,Buttons,Grid */}
            <Stack className={tileStyle.divFormContainer}>
                <div className={tileStyle.divLabel}>
                    <Icon iconName={"BucketColor"} />
                    <label className={css(tileStyle.fieldTitle)}>
                        {props.designType === DesignTypes.Buttons ? 'Hover Icon & Text Color' : props.designType === DesignTypes.Grid ? 'Hover Color' : 'Overflow Font Color'}
                    </label>
                </div>
                <div className={tileStyle.divField}>
                    <Stack horizontal horizontalAlign='space-between' tokens={stackTokens}>
                        <Stack.Item grow>
                            <div className={tileColorStyles.swatch}>
                                <div className={tileColorStyles.color} style={{ backgroundColor: state.overflowFontColor }} />
                            </div>
                            <div>
                                {state.displayOverflowFontColorPicker &&
                                    <>
                                        <ColorPicker color={state.overflowFontColor} onChange={(e, color) => _setColor('overflowFontColor', color.str)} alphaType={'alpha'}
                                            showPreview={true} styles={colorPickerStyles} />
                                    </>
                                }
                            </div>
                        </Stack.Item>
                        <Stack.Item>
                            {!state.displayOverflowFontColorPicker ? (
                                <IconButton iconProps={{ iconName: "Color" }} onClick={() => _openColorPicker('overflowFontColor')} className={tileColorStyles.pickerIcon}
                                    disabled={state.saving || state.loading} />
                            ) : (
                                <IconButton iconProps={{ iconName: "ErrorBadge" }} onClick={() => _closeColorPicker('overflowFontColor')} className={tileColorStyles.pickerIcon}
                                    disabled={state.saving || state.loading} />
                            )}
                        </Stack.Item>
                    </Stack>
                </div>
            </Stack>
            {/* Action Icon Color - All */}
            <Stack className={tileStyle.divFormContainer}>
                <div className={tileStyle.divLabel}>
                    <Icon iconName={"BucketColor"} />
                    <label className={css(tileStyle.fieldTitle)}>
                        {'Action Icon Color'}
                    </label>
                </div>
                <div className={tileStyle.divField}>
                    <Stack horizontal horizontalAlign='space-between' tokens={stackTokens}>
                        <Stack.Item grow>
                            <div className={tileColorStyles.swatch}>
                                <div className={tileColorStyles.color} style={{ backgroundColor: state.actionIconColor }} />
                            </div>
                            <div>
                                {state.displayActionIconColorPicker &&
                                    <>
                                        <ColorPicker color={state.actionIconColor} onChange={(e, color) => _setColor('actionIconColor', color.str)} alphaType={'alpha'}
                                            showPreview={true} styles={colorPickerStyles} />
                                    </>
                                }
                            </div>
                        </Stack.Item>
                        <Stack.Item>
                            {!state.displayActionIconColorPicker ? (
                                <IconButton iconProps={{ iconName: "Color" }} onClick={() => _openColorPicker('actionIconColor')} className={tileColorStyles.pickerIcon}
                                    disabled={state.saving || state.loading} />
                            ) : (
                                <IconButton iconProps={{ iconName: "ErrorBadge" }} onClick={() => _closeColorPicker('actionIconColor')} className={tileColorStyles.pickerIcon}
                                    disabled={state.saving || state.loading} />
                            )}
                        </Stack.Item>
                    </Stack>
                </div>
            </Stack>
            {/* Action Icon Hover Color - All */}
            <Stack className={tileStyle.divFormContainer}>
                <div className={tileStyle.divLabel}>
                    <Icon iconName={"BucketColor"} />
                    <label className={css(tileStyle.fieldTitle)}>
                        {'Action Icon Hover Color'}
                    </label>
                </div>
                <div className={tileStyle.divField}>
                    <Stack horizontal horizontalAlign='space-between' tokens={stackTokens}>
                        <Stack.Item grow>
                            <div className={tileColorStyles.swatch}>
                                <div className={tileColorStyles.color} style={{ backgroundColor: state.actionIconHoverColor }} />
                            </div>
                            <div>
                                {state.displayActionIconHoverColorPicker &&
                                    <>
                                        <ColorPicker color={state.actionIconHoverColor} onChange={(e, color) => _setColor('actionIconHoverColor', color.str)} alphaType={'alpha'}
                                            showPreview={true} styles={colorPickerStyles} />
                                    </>
                                }
                            </div>
                        </Stack.Item>
                        <Stack.Item>
                            {!state.displayActionIconHoverColorPicker ? (
                                <IconButton iconProps={{ iconName: "Color" }} onClick={() => _openColorPicker('actionIconHoverColor')} className={tileColorStyles.pickerIcon}
                                    disabled={state.saving || state.loading} />
                            ) : (
                                <IconButton iconProps={{ iconName: "ErrorBadge" }} onClick={() => _closeColorPicker('actionIconHoverColor')} className={tileColorStyles.pickerIcon}
                                    disabled={state.saving || state.loading} />
                            )}
                        </Stack.Item>
                    </Stack>
                </div>
            </Stack>
            {/* <Stack verticalAlign='center' style={{ width: '100%' }}>
                <Stack.Item>
                    {props.designType === DesignTypes.Tiles &&
                        <div className={css(styles.tile)} style={tileStyle}>
                            {tileDesignActionButtons()}
                            <a href={props.item && props.item.URL ? props.item.URL.Url : null} onClick={props.onAddClick ? () => props.onAddClick() : null}
                                target={props.item && props.item.NewWindow ? '_blank' : ''}
                                title={props.item && props.item.Title ? props.item.Title : props.addTitle}
                                className={css(styles.linkTitle, customLinkStyles.linkTitle)} data-interception="off">
                                {tileIconElement()}
                                {tileTitleElement()}
                                {tileDesignDescElement()}
                            </a>
                        </div>
                    }
                </Stack.Item>
            </Stack> */}
            <div className={tileStyle.formActions}>
                <PrimaryButton onClick={_saveTileColor} disabled={state.saving || state.loading} primaryDisabled={state.saving || state.loading}>
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
}

export default TileColor;