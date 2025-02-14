import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { DisplayMode } from '@microsoft/sp-core-library';
import { BaseWebPartContext, IPropertyPaneAccessor } from '@microsoft/sp-webpart-base';
import { DesignTypes, ItemSize } from '../../../common/Constants';
import { HttpClient } from '@microsoft/sp-http';
import SPService from '../../../common/Helpers/spService';
import { Alignment, IColor } from '@fluentui/react';


export interface IModernQuickLinksProps {
	// Global Properties
    spService: SPService;
    httpClient: HttpClient;
	tName: string;
	siteurl: string;
    serRelativeUrl: string;
    context?: BaseWebPartContext;
    isSiteAdmin: boolean;
	// Common properties
	theme: IReadonlyTheme;
	isDarkTheme: boolean;
	environmentMessage: string;
	hasTeamsContext: boolean;
	userDisplayName: string;
	displayMode: DisplayMode;
	propertyPane: IPropertyPaneAccessor;
	wpZoneWidth: number;
	wpZoneWidthDynamic: number;
	wpInstanceId: string;
	designType: DesignTypes;
	tileSize: ItemSize;
	// Webpart properties
	showTitle: boolean;
    title: string;
    titleIcon: string;
	customTileCount: number;
	enableCustomTiles: boolean;
	enableManageTiles: boolean;
	globalList: string;
	userList: string;
	enableGLPCache: boolean;
	enableULPCache: boolean;
    freeFlow: boolean;
    wpHeight: number;
    glSortOrder: boolean;
    ulSortOrder: boolean;
    glSortBy: string;
    ulSortBy: string;
    enableTileInfo: boolean;
    tileInfoPanelHeader: string;
    // Search properties
    enableSearch: boolean;
    searchPlaceholder: string;
    enableSearchInContentArea: boolean;
    searchWidth: number;
    searchPosition: Alignment;
    enableKeywordSearch: boolean;
    keywordSearchField: string;
    // Color Settings
    useThemeColors: boolean;
    themeColors: any;
    backgroundColor: IColor;
    fontColor: IColor;
    overflowBackgroundColor: IColor;
    overflowFontColor: IColor;
    actionIconColor: IColor;
    actionIconHoverColor: IColor;
    useTileColors: boolean;
    listFilters: any[];
    useGLFilter: boolean;
}