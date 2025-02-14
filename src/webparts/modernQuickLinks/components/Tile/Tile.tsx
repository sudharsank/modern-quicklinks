import * as React from 'react';
import { useState, FC } from 'react';
import * as strings from 'ModernQuickLinksWebPartStrings';
import styles from './Tile.module.scss';
import { Icon } from '@fluentui/react/lib/Icon';
import { Image, ImageFit } from '@fluentui/react/lib/Image';
import { _getBoxStyleItemWidth, getTileWidth } from '../../../../common/util';
import Dialog, { DialogFooter, DialogType } from '@fluentui/react/lib/Dialog';
import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
import ContentLoader from '../../../../common/components/ContentLoader';
import { DesignTypes, ItemSize, LoaderType, PanelTypes } from '../../../../common/Constants';
import { SpinnerSize } from '@fluentui/react/lib/Spinner';
import { IStackItemStyles, Stack } from '@fluentui/react/lib/Stack';
import { ColorClassNames, css, HoverCard, IStackStyles } from 'office-ui-fabric-react';
import { GridLayout } from "@pnp/spfx-controls-react/lib/GridLayout";
import {
    DocumentCard, DocumentCardDetails, DocumentCardPreview, DocumentCardTitle,
    DocumentCardType, IDocumentCardDetailsStyles, IDocumentCardPreviewProps, IDocumentCardStyles
} from '@fluentui/react/lib/DocumentCard';
import { ISize } from '@fluentui/react/lib/Utilities';
import { FontSizes } from '@fluentui/react/lib/Styling';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { AppPanel } from '../../../../common/components/AppPanel';
import { mergeStyleSets } from '@fluentui/react/lib/Styling';
import { DisplayMode } from '@microsoft/sp-core-library';

export interface ITileProps {
    globalList?: string;
    items?: any[];
    item?: any;
    addTitle?: string;
    showAdd?: boolean;
    displayMode: DisplayMode;
    onAddClick?: () => void;
    deleteCallback: (itemid: number) => void;
    onEditClick?: (item: any) => void;
    onShowTileColor?: (item: any) => void;
    addDescription?: string;
    isCustom?: boolean;
    designType: DesignTypes;
    tileSize?: ItemSize;
    wpWidth: number;
    theme: IReadonlyTheme;
    isDarkTheme: boolean;
    enableTileInfo?: boolean;
    tileInfoPanelHeader?: string;
    tileColors?: any | undefined;
    useTileColors?: boolean;
}

const Tile: FC<ITileProps> = (props) => {
    const [saving, setSaving] = useState<boolean>(false);
    const [hideDialog, sethideDialog] = useState<boolean>(true);
    const [showInfoPanel, setInfoPanel] = useState<boolean>(false);
    const [delItem, setDelItem] = useState<any>(undefined);
    const [infoitem, setInfoItem] = useState<any>(undefined);
    const tileStyle: React.CSSProperties = {};
    tileStyle.width = getTileWidth(props.wpWidth, props.tileSize);
    const itemTileColors = props.useTileColors && props.item?.TileColors ? JSON.parse(props.item.TileColors) : undefined;
    const backgroundColor = itemTileColors ? (itemTileColors.backgroundColor ? itemTileColors.backgroundColor : props.tileColors?.backgroundColor) : props.tileColors?.backgroundColor;
    const fontColor = itemTileColors ? (itemTileColors.fontColor ? itemTileColors.fontColor : props.tileColors?.fontColor) : props.tileColors?.fontColor;
    const overflowBackgroundColor = itemTileColors ? (itemTileColors.overflowBackgroundColor ? itemTileColors.overflowBackgroundColor : props.tileColors?.overflowBackgroundColor) : props.tileColors?.overflowBackgroundColor;
    const overflowFontColor = itemTileColors ? (itemTileColors.overflowFontColor ? itemTileColors.overflowFontColor : props.tileColors?.overflowFontColor) : props.tileColors?.overflowFontColor;
    const actionIconColor = itemTileColors ? (itemTileColors.actionIconColor ? itemTileColors.actionIconColor : props.tileColors?.actionIconColor) : props.tileColors?.actionIconColor;
    const actionIconHoverColor = itemTileColors ? (itemTileColors.actionIconHoverColor ? itemTileColors.actionIconHoverColor : props.tileColors?.actionIconHoverColor) : props.tileColors?.actionIconHoverColor;

    const buttonItemStyle: IStackItemStyles = {
        root: {
            display: 'flex',
            height: 49,
            minWidth: 200,
            maxWidth: _getBoxStyleItemWidth(props.wpWidth, undefined),
            marginBottom: 10,
            border: '1px solid',
            borderColor: backgroundColor
        },
    };
    const compactItemStyle: IStackItemStyles = {
        root: {
            display: 'flex',
            height: 30,
            minWidth: 200,
            maxWidth: _getBoxStyleItemWidth(props.wpWidth, undefined),
            marginBottom: 10,
        },
    };

    const customLinkStyles = mergeStyleSets({
        linkTitle: {
            backgroundColor: backgroundColor,
            color: fontColor
        },
        overflowColors: {
            backgroundColor: overflowBackgroundColor,
            color: overflowFontColor
        },
        buttonLinkTitle: {
            selectors: {
                '&:hover': {
                    backgroundColor: overflowBackgroundColor,
                    color: overflowFontColor,
                    //borderbottomcolor: backgroundColor,
                }
            }
        },
        compactLinkTitle: {
            selectors: {
                '&:hover': {
                    border: '1px solid',
                    borderColor: backgroundColor,
                    backgroundColor: overflowBackgroundColor
                }
            }
        },
        docCard: {
            selectors: {
                '&::after': {
                    borderColor: `${overflowFontColor} !important`
                },
                '&:hover': {
                    backgroundColor: overflowFontColor,
                    borderColor: overflowFontColor
                },
                'div[class^="ms-DocumentCardPreview-iconContainer"]': {
                    backgroundColor: backgroundColor,
                    color: fontColor
                }
            }
        },
        tileIcon: {
            color: fontColor
        },
        tileTitle: {
            color: fontColor
        },
        actionIcon: {
            color: actionIconColor,
            selectors: {
                '&:hover': {
                    color: actionIconHoverColor
                }
            }
        }
    });

    const modalPropsStyles = { main: { minwidth: 400, maxWidth: 450 } };
    const dialogContentProps = {
        type: DialogType.close,
        title: 'Delete Link'
    };
    const modalProps = React.useMemo(
        () => ({
            isBlocking: true,
            styles: modalPropsStyles,
        }),
        [false],
    );
    const buttonActionButtonStyle: IStackStyles = {
        root: { height: '47px', paddingRight: '2px' }
    };
    const compactActionButtonStyle: IStackStyles = {
        root: {
            height: '30px',
            paddingRight: '3px',
            alignItems: 'center'
        }
    };
    const documentCardStyles: IDocumentCardStyles = {
        root: {
            border: '0px',
            width: '210px'
        }
    };
    const documentCardDetailsStyles: IDocumentCardDetailsStyles = {
        root: {
            backgroundColor: props.theme.semanticColors.inputBackground,
            border: '1px solid',
            borderColor: overflowFontColor, // props.isDarkTheme ? props.theme.semanticColors.primaryButtonBorder : props.theme.palette.themePrimary
            '&:hover': {
                borderColor: overflowFontColor,
            }
        }
    };
    const _dismissDelDialog = () => sethideDialog(true);

    const _deleteLink = async () => {
        setSaving(true);
        await props.deleteCallback(delItem?.ID);
        sethideDialog(true);
        setSaving(false);
    };

    const _deleteTile = (item?: any) => { setDelItem(props.item || item); sethideDialog(false); }

    const _editTile = async (item?: any) => { props.onEditClick(props.item || item); };

    const _showTitleInfo = async (item?: any) => { setInfoItem(item); setInfoPanel(true); };

    const _updateTileColor = async (item?: any) => { console.log('Update tile color', props.item?.ID); props.onShowTileColor(props.item || item); };

    const _onRenderGridItem = (item: any, finalSize?: ISize, isCompact?: boolean): JSX.Element => {
        let gridItemCardProps: IDocumentCardPreviewProps = undefined;
        if (item.Id === 0) {
            gridItemCardProps = {
                previewImages: [
                    {
                        previewIconProps: {
                            iconName: 'AddLink',
                            styles: {
                                root: {
                                    fontSize: FontSizes.mega
                                }
                            }
                        },
                        name: 'Add New',
                        linkProps: {
                            onClick: props.onAddClick,
                            href: null,
                            target: '_self'
                        },
                        height: 100
                    }
                ]
            }
        } else if (item.ImageUrl) {
            gridItemCardProps = {
                previewImages: [
                    {
                        name: item.Title ? item.Title : props.addTitle,
                        linkProps: {
                            href: item.URL ? item.URL.Url : null,
                            target: item && item.NewWindow ? '_blank' : ''
                        },
                        height: 100,
                        previewImageSrc: item.ImageUrl ? item.ImageUrl : '',
                        imageFit: ImageFit.centerCover
                    }
                ]
            }
        } else {
            gridItemCardProps = {
                previewImages: [
                    {
                        previewIconProps: {
                            iconName: item.IconName ? item.IconName : "PageLink",
                            styles: {
                                root: {
                                    fontSize: FontSizes.mega
                                }
                            }
                        },
                        name: item.Title,
                        linkProps: {
                            href: item.URL ? item.URL.Url : null,
                            target: item && item.NewWindow ? '_blank' : ''
                        },
                        height: 100
                    }
                ]
            }
        }
        return (
            <>
                <div data-is-focusable={true} aria-label={item.title} style={{ marginTop: '20px' }}>
                    {item.isUsers &&
                        <div className={styles.actionButtonsGrid}>
                            <div style={{ paddingRight: '3px' }}><a href={"#"} onClick={() => _deleteTile(item)}>
                                <Icon iconName={"Delete"} className={customLinkStyles.actionIcon} /></a>
                            </div>
                            <div><a href={"#"} onClick={() => _editTile(item)}>
                                <Icon iconName={"PageEdit"} className={customLinkStyles.actionIcon} /></a>
                            </div>
                        </div>
                    }
                    {!props.showAdd && !props.isCustom && props.enableTileInfo && item?.Description &&
                        <div className={styles.actionButtonsGrid}>
                            <div style={{ paddingRight: '3px' }}><a href={"#"} onClick={() => _showTitleInfo(item)}>
                                <Icon iconName={"Info12"} className={customLinkStyles.actionIcon} /></a>
                            </div>
                        </div>
                    }
                    {props.useTileColors && !props.isCustom && props.displayMode === DisplayMode.Edit &&
                        <div className={styles.actionButtonsGrid}>
                            <div style={{ paddingRight: '3px' }}><a href={"#"} onClick={_updateTileColor}>
                                <Icon iconName={"Color"} className={customLinkStyles.actionIcon} /></a>
                            </div>
                        </div>
                    }
                    <DocumentCard type={DocumentCardType.normal} styles={documentCardStyles} className={customLinkStyles.docCard}
                        onClickHref={item.URL ? item.URL.Url : null} onClickTarget={item && item.NewWindow ? '_blank' : ''}
                        onClick={item.Id === 0 ? props.onAddClick : undefined}>
                        <DocumentCardPreview {...gridItemCardProps} />
                        <DocumentCardDetails styles={documentCardDetailsStyles}>
                            <DocumentCardTitle title={item.Title ? item.Title : props.addTitle}
                                styles={{ root: { fontSize: '14px' } }} />
                        </DocumentCardDetails>
                    </DocumentCard>
                </div>
            </>
        );
    }
    const tileIconElement = (): JSX.Element => {
        return (
            <div className={css(styles.tileIcon, customLinkStyles.tileIcon)}>
                {props.showAdd ? (
                    <Icon iconName={"AddLink"} />
                ) : (
                    <>
                        {props.item.ImageUrl ? (
                            <Stack horizontalAlign='center'>
                                <Stack.Item align='center'>
                                    <Image src={props.item.ImageUrl} width={25} height={25} />
                                </Stack.Item>
                            </Stack>
                        ) : (
                            <Icon iconName={props.item.IconName ? props.item.IconName : "PageLink"} />
                        )}
                    </>
                )}
            </div>
        );
    };
    const tileTitleElement = (): JSX.Element => {
        return (
            <div className={css(styles.tileTitle, customLinkStyles.tileTitle)}>
                {props.item.Title ? props.item.Title : props.addTitle}
            </div>
        );
    };
    const actionButtonsElement = (): JSX.Element => {
        return (
            <>
                {!props.showAdd && !props.isCustom && props.enableTileInfo && props.item.FieldValuesAsText?.Description &&
                    <Stack.Item>
                        <a href={"#"} onClick={_showTitleInfo}>
                            <Icon iconName={"Info12"} className={customLinkStyles.actionIcon} /></a>
                    </Stack.Item>
                }
                {props.useTileColors && !props.isCustom && props.displayMode === DisplayMode.Edit &&
                    <Stack.Item>
                        <a href={"#"} onClick={_updateTileColor}>
                            <Icon iconName={"Color"} className={customLinkStyles.actionIcon} /></a>
                    </Stack.Item>
                }
                {props.isCustom &&
                    <>
                        <Stack.Item>
                            <a href={"#"} onClick={_editTile}>
                                <Icon iconName={"PageEdit"} className={customLinkStyles.actionIcon} /></a>
                        </Stack.Item>
                        <Stack.Item>
                            <a href={"#"} onClick={_deleteTile}>
                                <Icon iconName={"Delete"} className={customLinkStyles.actionIcon} /></a>
                        </Stack.Item>
                    </>
                }
            </>
        );
    };
    const tileDesignActionButtons = (): JSX.Element => {
        return (
            <>
                {props.isCustom &&
                    <div className={styles.actionButtons}>
                        <div style={{ paddingRight: '3px' }}><a href={"#"} onClick={_deleteTile}>
                            <Icon iconName={"Delete"} className={customLinkStyles.actionIcon} /></a>
                        </div>
                        <div><a href={"#"} onClick={_editTile}>
                            <Icon iconName={"PageEdit"} className={customLinkStyles.actionIcon} /></a>
                        </div>
                    </div>
                }
                {!props.showAdd && !props.isCustom && props.enableTileInfo && props.item.FieldValuesAsText?.Description &&
                    <div className={styles.actionButtons} style={{ left: props.useTileColors ? '8px' : 'unset' }}>
                        <div><a href={"#"} onClick={_showTitleInfo}>
                            <Icon iconName={"Info12"} className={customLinkStyles.actionIcon} /></a>
                        </div>
                    </div>
                }
                {props.useTileColors && !props.isCustom && props.displayMode === DisplayMode.Edit &&
                    <div className={styles.actionButtons}>
                        <div><a href={"#"} onClick={_updateTileColor}>
                            <Icon iconName={"Color"} className={customLinkStyles.actionIcon} /></a>
                        </div>
                    </div>
                }
            </>
        )
    };
    const tileDesignDescElement = (): JSX.Element => {
        return (
            <>
                {(props.item && props.item.URL && props.item.URL.Description) ? (
                    <div className={css(styles.overflow, customLinkStyles.overflowColors)}>
                        {props.item.URL.Description}
                    </div>
                ) : (
                    <>
                        {props.addDescription &&
                            <div className={css(styles.overflow, customLinkStyles.overflowColors)}>
                                {props.addDescription}
                            </div>
                        }
                    </>
                )}
            </>
        );
    };

    return (
        <>
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
            {props.designType === DesignTypes.Buttons &&
                <Stack.Item styles={buttonItemStyle} grow>
                    <a href={props.item.URL ? props.item.URL.Url : null} onClick={props.onAddClick ? () => props.onAddClick() : null}
                        target={props.item && props.item.NewWindow ? '_blank' : ''}
                        title={props.item.Title ? props.item.Title : props.addTitle}
                        data-interception="off" className={css(styles.buttonLinkTile, customLinkStyles.buttonLinkTitle)}>
                        {tileIconElement()}
                        <div className={styles.tileTitleContainer}>
                            {tileTitleElement()}
                        </div>
                        {!props.isCustom &&
                            <Stack verticalAlign='space-between' styles={buttonActionButtonStyle}>
                                {actionButtonsElement()}
                            </Stack>
                        }
                        {props.isCustom &&
                            <Stack verticalAlign='space-between' styles={buttonActionButtonStyle}>
                                {actionButtonsElement()}
                            </Stack>
                        }
                    </a>
                </Stack.Item>
            }
            {props.designType === DesignTypes.Compact &&
                <Stack.Item styles={compactItemStyle} grow>
                    <a href={props.item.URL ? props.item.URL.Url : null} onClick={props.onAddClick ? () => props.onAddClick() : null}
                        target={props.item && props.item.NewWindow ? '_blank' : ''}
                        title={props.item.Title ? props.item.Title : props.addTitle}
                        data-interception="off" className={css(styles.compactLink, customLinkStyles.compactLinkTitle)}>
                        {tileIconElement()}
                        <div className={styles.tileTitleContainer}>
                            {tileTitleElement()}
                        </div>
                        {!props.isCustom &&
                            <Stack verticalAlign='space-between' styles={buttonActionButtonStyle} style={{ height: '15px' }}>
                                {actionButtonsElement()}
                            </Stack>
                        }
                        {props.isCustom &&
                            <Stack tokens={{ childrenGap: 4 }} horizontal horizontalAlign='end' styles={compactActionButtonStyle}>
                                {actionButtonsElement()}
                            </Stack>
                        }
                    </a>
                </Stack.Item>
            }
            {props.designType === DesignTypes.Grid &&
                <>
                    <Stack horizontal horizontalAlign='space-evenly' verticalAlign='center' wrap>
                        {props.items && props.items.length > 0 && props.items.map((item: any, index: number) => {
                            return (
                                <>{_onRenderGridItem(item, undefined, false)}</>
                            )
                        })}
                    </Stack>
                    {/* <GridLayout
                        ariaLabel="Launch pad items."
                        items={props.items}
                        onRenderGridItem={(item: any, finalSize: ISize, isCompact: boolean) => _onRenderGridItem(item, finalSize, isCompact)}
                    /> */}
                </>
            }
            <Dialog
                hidden={hideDialog}
                onDismiss={_dismissDelDialog}
                dialogContentProps={dialogContentProps}
                modalProps={modalProps} >
                <div className={styles.dialogContent}>
                    <div>{strings.DeleteDialogDescription}</div>
                    <div><b>Note:</b> {strings.DeleteDialogNote}</div>
                </div>
                <DialogFooter>
                    <Stack tokens={{ childrenGap: 10 }} horizontal horizontalAlign='end'>
                        <Stack.Item>
                            {saving &&
                                <div style={{ marginTop: '-13px', display: 'inline-block' }}><ContentLoader loaderType={LoaderType.Spinner} spinSize={SpinnerSize.medium} /></div>
                            }
                        </Stack.Item>
                        <Stack.Item>
                            <PrimaryButton onClick={_deleteLink} text="Yes" disabled={saving} />
                        </Stack.Item>
                        <Stack.Item>
                            <DefaultButton onClick={_dismissDelDialog} text="No" disabled={saving} />
                        </Stack.Item>
                    </Stack>
                </DialogFooter>
            </Dialog>
            {showInfoPanel &&
                <AppPanel headerText={props.item?.Title ? `${props.item.Title}${props.tileInfoPanelHeader ? ' - ' + props.tileInfoPanelHeader : ''}` : props.tileInfoPanelHeader ? props.tileInfoPanelHeader : 'Information'}
                    panelType={PanelTypes.TileInfo} item={props.item ? props.item : infoitem}
                    dismissCallback={() => setInfoPanel(false)} tileColors={props.tileColors} />
            }
        </>
    );
};

export default Tile;