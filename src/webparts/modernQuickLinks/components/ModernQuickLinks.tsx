import * as React from 'react';
import { useEffect, FC, useReducer } from 'react';
import * as strings from 'ModernQuickLinksWebPartStrings';
import styles from './ModernQuickLinks.module.scss';
import commonStyles from '../../../common/components/common.module.scss';
import { IModernQuickLinksProps } from './IModernQuickLinksProps';
import AppContext from '../../../common/AppContext';
import { IModernQuickLinksState } from './IModernQuickLinksState';
import { IStackItemStyles } from 'office-ui-fabric-react/lib/Stack';
import { _getBoxStyleItemWidth, _isSearchPositionAvailable } from '../../../common/util';
import { merge, orderBy, sortBy, union } from 'lodash';
import { DesignTypes, fld_tileColors, LicenseMessage, LoaderType, LogSource, MessageScope, PanelTypes } from '../../../common/Constants';
import { useCommonHelper, useLaunchPadHelper, useLicenseHelper } from '../../../common/Helpers';
import { Stack } from '@fluentui/react/lib/Stack';
import ContentLoader from '../../../common/components/ContentLoader';
import { ConfigMissing } from '../../../common/components/ConfigMissing';
import { LicenseInfo } from '../../../common/components/License/LicenseInfo';
import { Icon } from '@fluentui/react/lib/Icon';
import Tile from './Tile/Tile';
import { AppPanel } from '../../../common/components/AppPanel';
import { ISearchBoxStyles, mergeStyles } from 'office-ui-fabric-react';
import MessageContainer from '../../../common/components/Message';
import moment from 'moment';
import { SearchBox } from '@fluentui/react';
import { IFieldInfo } from '@pnp/sp/fields';

const initialState: IModernQuickLinksState = {
    loading: true,
    isOpen: false,
    launchPadItems: [],
    searchItems: [],
    userTiles: [],
    showAddLink: true,
    addTitle: '',
    addDesc: '',
    reload: false,
    hideDialog: true,
    delItemID: 0,
    saving: false,
    showAdd: false,
    showEdit: false,
    showManage: false,
    showTileColor: false,
    item: undefined,
    editItem: null,
    showConfig: true,
    showConfigDialog: false,
    showSettings: true,
    showLicenseDialog: false,
    licenseMessage: undefined,
    licInfo: undefined
}

const modernLinkReducer = (state: IModernQuickLinksState, action: any): IModernQuickLinksState => {
    switch (action.type) {
        case 'SET_LIC_MSG':
            return { ...state, licenseMessage: action.payload };
        case 'INVALID_LIC':
            return { ...state, loading: false, showConfig: false, reload: false, licenseMessage: action.payload };
        case 'LICINFO':
            return { ...state, licInfo: action.payload };
        case 'LISTS_CONFIGURED':
            return { ...state, showConfig: false };
        case 'LISTS_NOT_CONFIGURED':
            return { ...state, showConfig: true, loading: false };
        case 'LOAD_LAUNCHPAD_ITEMS':
            return {
                ...state, launchPadItems: action.payload.finalLinks, userTiles: action.payload.userLinks, showAddLink: action.payload.showAddLink,
                reload: false, loading: false, searchItems: action.payload.finalLinks
            };
        case 'LICENSE_VALIDATED':
            return { ...state, loading: true, showLicenseDialog: false };
        case 'DISMISS_PANEL':
            return { ...state, showAdd: false, showEdit: false, showManage: false, showTileColor: false, item: null, editItem: null };
        case 'MANAGE_LINKS':
            return { ...state, isOpen: true, showManage: true };
        case 'MANAGE_LIC':
            return { ...state, showLicenseDialog: action.payload, showConfig: false };
        case 'CLOSE_LIC':
            return { ...state, showLicenseDialog: false };
        case 'MANAGE_LIC_FROM_CONFIG':
            return { ...state, showLicenseDialog: action.payload };
        case 'ADD_LINK':
            return { ...state, isOpen: true, showAdd: true };
        case 'EDIT_LINK':
            return { ...state, isOpen: true, showEdit: true, editItem: action.payload };
        case 'DELETE_LINK':
            return { ...state, delItemID: action.payload };
        case 'TILECOLOR':
            return { ...state, showTileColor: true, item: action.payload };
        case 'RELOAD':
            return { ...state, reload: true, showEdit: false };
        case 'SEARCH_ITEMS':
            return { ...state, searchItems: action.payload };
        case 'CLEAR_SEARCH':
            return { ...state, searchItems: state.launchPadItems };
    }
};

const ModernQuickLinks: FC<IModernQuickLinksProps> = (props) => {
    const { spService } = props;
    const { checkAndCreateOOBLists } = useCommonHelper(props.spService);
    const { checkForLicenseKey, getLicenseInfo } = useLicenseHelper(props.spService);
    const { getGlobalLinks, getUserLinks, delUserTile, checkAndCreateField } = useLaunchPadHelper(props.spService, props.wpInstanceId);
    const [newState, dispatch] = useReducer(modernLinkReducer, initialState);
    const isSearchPositionAvailable: boolean = _isSearchPositionAvailable(props.wpZoneWidthDynamic, props.wpZoneWidth);
    const searchBoxStyles: Partial<ISearchBoxStyles> = { root: { width: props.searchWidth ? props.searchWidth : 350 } };
    let tileColors: any = {};
    const getCustomColors = () => {
        tileColors = {
            backgroundColor: props.useThemeColors ? props.themeColors.primaryButtonBackground : props.backgroundColor,
            fontColor: props.useThemeColors ? props.themeColors.primaryButtonText : props.fontColor,
            overflowBackgroundColor: props.useThemeColors ? props.themeColors.overflowBg : props.overflowBackgroundColor,
            overflowFontColor: props.useThemeColors ? props.themeColors.overflowFontColor : props.overflowFontColor,
            actionIconColor: props.useThemeColors ? props.themeColors.actLink : props.actionIconColor,
            actionIconHoverColor: props.useThemeColors ? props.themeColors.actLinkHovered : props.actionIconHoverColor
        };
        if (props.designType === DesignTypes.Compact) {
            tileColors.fontColor = props.useThemeColors ? props.themeColors.bodyText : props.fontColor;
            tileColors.overflowBackgroundColor = props.useThemeColors ? props.themeColors.bodyBackground : props.overflowBackgroundColor;
        } else if (props.designType === DesignTypes.Buttons) {
            tileColors.fontColor = props.useThemeColors ? props.themeColors.bodyText : props.fontColor;
            tileColors.overflowBackgroundColor = props.useThemeColors ? props.themeColors.bodyBgHovered : props.overflowBackgroundColor;
            tileColors.overflowFontColor = props.useThemeColors ? props.themeColors.bodyText : props.overflowFontColor;
        } else if (props.designType === DesignTypes.Grid) {
            tileColors.backgroundColor = props.useThemeColors ? props.themeColors.cardPrevBg : props.backgroundColor;
            tileColors.fontColor = props.useThemeColors ? props.themeColors.cardPrevFontColor : props.fontColor;
            tileColors.overflowFontColor = props.useThemeColors ? props.themeColors.primaryButtonBackground : props.overflowFontColor;
        }
    };
    getCustomColors();
    const buttonItemStyle: IStackItemStyles = {
        root: {
            display: 'flex',
            height: 49,
            minWidth: 200,
            maxWidth: _getBoxStyleItemWidth(props.wpZoneWidthDynamic, props.wpZoneWidth),
            marginBottom: 10,
            //border: '1px solid',
            //borderColor: props.theme.palette.themePrimary
        },
    };
    const compactItemStyle: IStackItemStyles = {
        root: {
            display: 'flex',
            height: 30,
            minWidth: 200,
            maxWidth: _getBoxStyleItemWidth(props.wpZoneWidthDynamic, props.wpZoneWidth),
            marginBottom: 10,
        },
    };

    const _loadLaunchPadItems = async (updateCache?: boolean) => {
        if (newState.licenseMessage === LicenseMessage.Valid) {
            let finalLinks: any[] = [];
            let userLinks: any[] = [];
            const listFields: IFieldInfo[] = await spService.getFields('', props.globalList);
            let globalLinks: any[] = await getGlobalLinks(undefined, props.globalList, props.glSortBy ? props.glSortBy : 'ID',
                props.glSortOrder ? props.glSortOrder : false, updateCache, props.enableTileInfo, props.enableKeywordSearch, props.keywordSearchField,
                props.useTileColors, listFields, props.useGLFilter, props.enableGLPCache, props.useGLFilter ? props.listFilters : undefined);
            globalLinks = orderBy(globalLinks, [props.glSortBy], [props.glSortOrder ? 'asc' : 'desc']);
            if (props.enableCustomTiles && props.userList) userLinks = await getUserLinks(undefined, props.userList, props.ulSortBy ? props.ulSortBy : 'ID',
                props.ulSortOrder ? props.ulSortOrder : false, updateCache, props.enableULPCache);
            userLinks = orderBy(userLinks, [props.ulSortBy], [props.ulSortOrder ? 'asc' : 'desc']);
            userLinks.map(o => { return o.isUsers = true; });
            finalLinks = union(globalLinks, userLinks);
            finalLinks.map(o => {
                if (o.IconImage) {
                    let jsonImage = JSON.parse(o.IconImage);
                    return o.ImageUrl = jsonImage.serverRelativeUrl;
                } else return o.ImageUrl = undefined;
            });
            if (props.designType === DesignTypes.Grid && !(userLinks.length >= props.customTileCount || !props.enableCustomTiles)) {
                finalLinks.push({
                    Id: 0,
                    Title: strings.HDR_Add
                });
            }
            dispatch({
                type: 'LOAD_LAUNCHPAD_ITEMS', payload: {
                    finalLinks: finalLinks, userLinks: userLinks,
                    showAddLink: (userLinks.length >= props.customTileCount || !props.enableCustomTiles) ? false : true
                }
            });
        }
    };

    const _checkForProperties = () => {
        if (props.enableCustomTiles) {
            if (props.userList && props.globalList) dispatch({ type: 'LISTS_CONFIGURED' });
            else dispatch({ type: 'LISTS_NOT_CONFIGURED' });
        } else {
            if (props.globalList) dispatch({ type: 'LISTS_CONFIGURED' });
            else dispatch({ type: 'LISTS_NOT_CONFIGURED' });
        }
    };

    const _compLoad = async (): Promise<void> => {
        try {
            // Check for OOB lists
            checkAndCreateOOBLists();
            _checkForProperties();
        } catch (err) {
            spService.writeErrorLog(err, LogSource.LaunchPad);
        }
    };

    const _onLoad = async (): Promise<void> => {
        try {
            // Check for license Information
            const lic: LicenseMessage = await checkForLicenseKey();
            const licInfo: any = await getLicenseInfo();
            dispatch({ type: 'LICINFO', payload: licInfo });
            dispatch({ type: 'SET_LIC_MSG', payload: lic });
            //console.log(lic);
            if (lic === LicenseMessage.Valid) _compLoad();
            else dispatch({ type: 'INVALID_LIC', payload: lic });
        } catch (err) {
            dispatch({ type: 'SET_LIC_MSG', payload: LicenseMessage.ConfigError });
            spService.writeErrorLog(err, LogSource.LaunchPad);
        }
    };

    const _onLicenseValidated = async (): Promise<void> => {
        dispatch({ type: 'LICENSE_VALIDATED' });
        window.location.reload();
    };

    const _deleteLink = async (itemid: number) => {
        dispatch({ type: 'DELETE_LINK', payload: itemid });
        let delRes: boolean = await delUserTile('', props.userList, itemid.toString());
        if (delRes) dispatch({ type: 'DELETE_LINK', payload: undefined });
        _loadLaunchPadItems(true);
    };

    const _manageLinks = () => dispatch({ type: 'MANAGE_LINKS' });
    const _manageLic = () => dispatch({ type: 'MANAGE_LIC', payload: true });
    const _manageLicFromConfig = () => dispatch({ type: 'MANAGE_LIC_FROM_CONFIG', payload: true });
    const _onCloseLicDialog = () => {
        dispatch({ type: 'CLOSE_LIC' });
        _onLoad();
    }
    const _editLink = (item: any) => dispatch({ type: 'EDIT_LINK', payload: item });
    const _onDismissPanel = () => dispatch({ type: 'DISMISS_PANEL' });
    const _addNewLink = () => dispatch({ type: 'ADD_LINK' });
    const _showTileColor = (item: any) => dispatch({ type: 'TILECOLOR', payload: item });
    const _reloadTiles = () => dispatch({ type: 'RELOAD' });
    const _getTrialRemDays = () => {
        if (newState.licInfo && newState.licInfo.ed) {
            return moment(new Date(newState.licInfo.ed).toISOString()).diff(moment(new Date().toISOString()), 'days')
        }
        return 0;
    };
    const _onSearch = (e: any, searchText: string) => {
        if (searchText) {
            let searchItems = newState.launchPadItems.filter((item) => {
                return item.Title.toLowerCase().indexOf(searchText.toLowerCase()) !== -1 ||
                    (item.Description ? item.Description.toLowerCase().indexOf(searchText.toLowerCase()) !== -1 : false) ||
                    (item.URL ? item.URL.Description.toLowerCase().indexOf(searchText.toLowerCase()) !== -1 : false) ||
                    (item[props.keywordSearchField] ? item[props.keywordSearchField].toLowerCase().indexOf(searchText.toLowerCase()) !== -1 : false);
            });
            dispatch({ type: 'SEARCH_ITEMS', payload: searchItems });
        } else {
            dispatch({ type: 'CLEAR_SEARCH' });
        }
    };

    useEffect(() => {
        if (!newState.showConfig && newState.licenseMessage && newState.licenseMessage === LicenseMessage.Valid) _loadLaunchPadItems();
    }, [newState.showConfig]);

    useEffect(() => {
        if (newState.reload || (props.listFilters && props.listFilters.length > 0)) _loadLaunchPadItems(true);
    }, [newState.reload, props.listFilters, props.useGLFilter]);

    useEffect(() => dispatch({ type: 'RELOAD' }),
        [props.globalList, props.userList, props.glSortOrder, props.ulSortOrder, props.glSortBy, props.ulSortBy]);

    useEffect(() => {
        if (props.useTileColors) {
            (async () => {
                await checkAndCreateField(undefined, props.globalList, fld_tileColors, 'multiline', false, false);
            })();
        }
        _onLoad();
    }, [props]);

    return (
        <AppContext.Provider value={{ ...props }}>
            {newState.loading ? (
                <Stack verticalAlign={'center'}>
                    <ContentLoader loaderType={LoaderType.Spinner} loaderMsg={strings.Msg_Wait} />
                </Stack>
            ) : (

                <section className={`${styles.modernQuickLinks} ${props.hasTeamsContext ? styles.teams : ''}`}>
                    {newState.showConfig &&
                        <ConfigMissing displayMode={props.displayMode} propertyPane={props.propertyPane} onLicManage={_manageLicFromConfig} />
                    }
                    {(newState.licenseMessage !== LicenseMessage.Valid || newState.showLicenseDialog) &&
                        <LicenseInfo licMsg={newState.licenseMessage} onLicValidated={_onLicenseValidated} onCloseCallback={_onCloseLicDialog}
                            showLicenseForm={newState.showLicenseDialog} />
                    }
                    {newState.licenseMessage === LicenseMessage.Valid && !newState.showConfig &&
                        <div>
                            <div className={styles.modernQuickLinks}>
                                {props.showTitle && props.title &&
                                    <Stack tokens={{ childrenGap: 10 }} horizontal horizontalAlign='space-between' className={commonStyles.headerStack}>
                                        <Stack.Item align='start'>
                                            <div style={{ display: 'inline-flex', marginBottom: '5px' }}>
                                                <Icon iconName={props.titleIcon ? props.titleIcon : 'Tiles'} className={commonStyles.headerIcon}></Icon>
                                                <div className={commonStyles.headerDiv}>{props.title}</div>
                                            </div>
                                        </Stack.Item>
                                        {props.enableSearch && !props.enableSearchInContentArea && isSearchPositionAvailable &&
                                            <Stack.Item align='center' style={{ marginBottom: '15px' }}>
                                                <SearchBox placeholder={props.searchPlaceholder ? props.searchPlaceholder : 'Search...'} className={styles.searchBox} disableAnimation={true}
                                                    onChange={_onSearch} onClear={(e) => _onSearch(e, undefined)} styles={searchBoxStyles}
                                                    showIcon={true} />
                                            </Stack.Item>
                                        }
                                        <Stack.Item align='end'>
                                            <div className={styles.manageSection}>
                                                {props.enableManageTiles &&
                                                    <a className={styles.manageLink} href={"#"} onClick={_manageLinks} title={strings.Settings_Man_UTile}>
                                                        <Icon iconName='AddToShoppingList' className={commonStyles.manageListIcon} />
                                                    </a>
                                                }
                                                {props.isSiteAdmin &&
                                                    <a className={styles.manageLink} href={"#"} onClick={_manageLic} title={'Manage License'}>
                                                        <Icon iconName='AzureKeyVault' className={commonStyles.manageListIcon} />
                                                    </a>
                                                }
                                            </div>
                                        </Stack.Item>
                                    </Stack>
                                }
                                {props.enableSearch && props.enableSearchInContentArea &&
                                    <>
                                        <Stack horizontal horizontalAlign={props.searchPosition && isSearchPositionAvailable ? props.searchPosition : 'center'}>
                                            <Stack.Item style={{ marginBottom: '10px' }}>
                                                <SearchBox placeholder={props.searchPlaceholder ? props.searchPlaceholder : 'Search...'} className={styles.searchBox} disableAnimation={true}
                                                    onChange={_onSearch} onClear={(e) => _onSearch(e, undefined)} styles={searchBoxStyles}
                                                    showIcon={true} />
                                            </Stack.Item>
                                            {newState.showAddLink &&
                                                <Stack.Item style={{ marginTop: '5px', marginLeft: '10px' }}>
                                                    <a className={styles.manageLink} href={"#"} onClick={_addNewLink} title={'Add'}>
                                                        <Icon iconName='AddLink' className={commonStyles.addLinkIcon} style={{ fontSize: '20px' }} />
                                                    </a>
                                                </Stack.Item>
                                            }
                                        </Stack>
                                    </>
                                }
                                {!props.enableSearch && newState.showAddLink &&
                                    <Stack horizontal horizontalAlign={'end'}>
                                        <Stack.Item style={{ marginTop: '5px', marginLeft: '10px' }}>
                                            <a className={styles.manageLink} href={"#"} onClick={_addNewLink} title={'Add'}>
                                                <Icon iconName='AddLink' className={commonStyles.addLinkIcon} style={{ fontSize: '20px' }} />
                                            </a>
                                        </Stack.Item>
                                    </Stack>
                                }
                                {newState.searchItems ? (
                                    <div style={{
                                        height: !props.freeFlow && props.wpHeight ? `${props.wpHeight}px` : 'auto', overflowX: 'hidden',
                                        overflowY: !props.freeFlow && props.wpHeight ? 'auto' : 'hidden', paddingRight: !props.freeFlow && props.wpHeight ? '3px' : '0px'
                                    }}>
                                        {newState.licInfo && newState.licInfo.ed &&
                                            <div>
                                                <MessageContainer MessageScope={MessageScope.Info} Message={`Trial version. Trial ends in ${_getTrialRemDays()} day(s);`} />
                                            </div>
                                        }
                                        {props.designType === DesignTypes.Tiles &&
                                            <div className={styles.tilesList}>
                                                {newState.searchItems.map((tile, idx) =>
                                                    <Tile key={idx} item={tile} isCustom={tile.isUsers}
                                                        deleteCallback={_deleteLink} onEditClick={_editLink} wpWidth={props.wpZoneWidth || props.wpZoneWidthDynamic} theme={props.theme}
                                                        isDarkTheme={props.isDarkTheme} designType={props.designType} tileSize={props.tileSize} enableTileInfo={props.enableTileInfo}
                                                        tileInfoPanelHeader={props.tileInfoPanelHeader} tileColors={tileColors} useTileColors={props.useTileColors}
                                                        displayMode={props.displayMode} onShowTileColor={_showTileColor} globalList={props.globalList} />
                                                )}
                                                {newState.showAddLink &&
                                                    <Tile key={"customLinkKey"} item={{}} showAdd={true} tileSize={props.tileSize}
                                                        addTitle={"Add New"} addDescription={"Add New Link"} onAddClick={_addNewLink} deleteCallback={_reloadTiles}
                                                        wpWidth={props.wpZoneWidth || props.wpZoneWidthDynamic} theme={props.theme} isDarkTheme={props.isDarkTheme}
                                                        designType={props.designType} tileColors={tileColors} useTileColors={false} displayMode={props.displayMode} />}
                                            </div>
                                        }
                                        {(props.designType === DesignTypes.Buttons || props.designType === DesignTypes.Compact) &&
                                            <Stack tokens={{ childrenGap: 10 }} horizontal horizontalAlign='space-evenly' wrap>
                                                {newState.searchItems.map((tile, idx) =>
                                                    // <Stack.Item styles={props.designType === DesignTypes.Buttons ? buttonItemStyle : compactItemStyle} grow>
                                                    <Tile key={idx} item={tile} isCustom={tile.isUsers}
                                                        deleteCallback={_deleteLink} onEditClick={_editLink} designType={props.designType}
                                                        wpWidth={props.wpZoneWidth || props.wpZoneWidthDynamic} theme={props.theme} isDarkTheme={props.isDarkTheme}
                                                        enableTileInfo={props.enableTileInfo} tileInfoPanelHeader={props.tileInfoPanelHeader} tileColors={tileColors}
                                                        useTileColors={props.useTileColors} displayMode={props.displayMode} globalList={props.globalList}
                                                        onShowTileColor={_showTileColor} />
                                                    // </Stack.Item>
                                                )}
                                                {newState.showAddLink &&
                                                    // <Stack.Item styles={props.designType === DesignTypes.Buttons ? buttonItemStyle : compactItemStyle} grow>
                                                    <Tile key={"customLinkKey"} item={{}} showAdd={true} designType={props.designType}
                                                        addTitle={"Add New"} addDescription={"Add New Link"} onAddClick={_addNewLink} deleteCallback={_reloadTiles}
                                                        wpWidth={props.wpZoneWidth || props.wpZoneWidthDynamic} theme={props.theme} isDarkTheme={props.isDarkTheme}
                                                        tileColors={tileColors} useTileColors={false} displayMode={props.displayMode} />
                                                    // </Stack.Item>
                                                }
                                            </Stack>
                                        }
                                        {props.designType === DesignTypes.Grid &&
                                            <Tile items={newState.searchItems} deleteCallback={_deleteLink} onEditClick={_editLink} designType={props.designType}
                                                wpWidth={props.wpZoneWidth || props.wpZoneWidthDynamic} theme={props.theme} isDarkTheme={props.isDarkTheme}
                                                onAddClick={_addNewLink} enableTileInfo={props.enableTileInfo} tileInfoPanelHeader={props.tileInfoPanelHeader}
                                                tileColors={tileColors} useTileColors={props.useTileColors} displayMode={props.displayMode} globalList={props.globalList}
                                                onShowTileColor={_showTileColor} />
                                        }
                                    </div>
                                ) : (
                                    <></>
                                )}
                                {newState.showAdd &&
                                    <AppPanel headerText={strings.HDR_Add} panelType={PanelTypes.AddUpdateTiles} dismissCallback={_onDismissPanel} successCallback={_reloadTiles} customTileRowCount={props.customTileCount}
                                        userTileList={props.userList} />
                                }
                                {newState.showEdit && newState.editItem &&
                                    <AppPanel headerText={strings.HDR_Edit} panelType={PanelTypes.AddUpdateTiles} dismissCallback={_onDismissPanel} successCallback={_reloadTiles} item={newState.editItem} customTileRowCount={props.customTileCount}
                                        userTileList={props.userList} />
                                }
                                {newState.showManage &&
                                    <AppPanel headerText={strings.HDR_List} panelType={PanelTypes.ManageTiles} dismissCallback={_onDismissPanel} successCallback={_reloadTiles} items={newState.userTiles} customTileRowCount={props.customTileCount}
                                        userTileList={props.userList} deleteCallback={_deleteLink} />
                                }
                                {newState.showTileColor &&
                                    <AppPanel headerText={'Link Colors'} panelType={PanelTypes.TileColor} dismissCallback={_onDismissPanel} successCallback={_reloadTiles}
                                        item={newState.item} globalList={props.globalList} designType={props.designType} />
                                }
                            </div>
                        </div>
                    }
                    {/**
                     // ) : (
                    //     <LicenseInfo licMsg={newState.licenseMessage} onLicValidated={_onLicenseValidated} onCloseCallback={_onCloseLicDialog} />
                    // )}
                     */}
                </section>
            )}
        </AppContext.Provider>
    );
}

export default ModernQuickLinks;
