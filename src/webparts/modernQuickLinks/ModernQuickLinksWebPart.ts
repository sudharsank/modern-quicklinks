import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
    IPropertyPaneConfiguration, PropertyPaneChoiceGroup, PropertyPaneLink, PropertyPaneTextField,
    PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import * as strings from 'ModernQuickLinksWebPartStrings';
import ModernQuickLinks from './components/ModernQuickLinks';
import { IModernQuickLinksProps } from './components/IModernQuickLinksProps';
import { DesignTypes, ItemSize } from '../../common/Constants';
import { CalloutTriggers } from '@pnp/spfx-property-controls/lib/common/callout/Callout';
import SPService from '../../common/Helpers/spService';
import { PropertyFieldNumber } from '@pnp/spfx-property-controls';
import { PropertyPaneWebPartInformation } from '@pnp/spfx-property-controls/lib/PropertyPaneWebPartInformation';
import { IPropertyFieldList } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';
import { Alignment, ColorPicker, IColor } from '@fluentui/react';
import { _isSearchPositionAvailable } from '../../common/util';
import { restFields, restFieldTypes, fieldInfo } from '../../common/Constants';
import { filter } from 'lodash';
import { IFieldInfo } from '@pnp/sp/fields';
import { PropertyFieldCollectionData, CustomCollectionFieldType, ICustomDropdownOption } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';

export interface IModernQuickLinksWebPartProps {
    showTitle: boolean;
    title: string;
    titleIcon: string;
    //tilesInARowCount: number;
    customTileCount: number;
    enableCustomTiles: boolean;
    enableManageTiles: boolean;
    globalList: string;
    enableTileInfo: boolean;
    tileInfoPanelHeader: string;
    userList: string;
    designType: DesignTypes;
    tileSize: ItemSize;
    enableGLPCache: boolean;
    enableULPCache: boolean;
    freeFlow: boolean;
    wpHeight: number;
    glSortOrder: boolean;
    ulSortOrder: boolean;
    glSortBy: string;
    ulSortBy: string;
    // Search Settings
    enableSearch: boolean;
    searchPlaceholder: string;
    searchInContentArea: boolean;
    searchWidth: number;
    searchPosition: Alignment;
    keywordSearchField: string;
    enableKeywordSearch: boolean;
    // Color Settings
    useThemeColors: boolean;
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

export default class ModernQuickLinksWebPart extends BaseClientSideWebPart<IModernQuickLinksWebPartProps> {

    private _currentTheme: IReadonlyTheme;
    private _isDarkTheme: boolean = false;
    private _environmentMessage: string = '';
    private _webpartZoneWidth: number = 0;
    private _spService: SPService;
    private showLicDialog: boolean = false;
    private themeColors: any = {};
    private _allFields: any[] = [];
    private _fields: ICustomDropdownOption[] = [];
    /** Property Controls */
    private propertyFieldToggleWithCallout: any = undefined;
    private propertyFieldSliderWithCallout: any = undefined;
    private propertyFieldListPicker: any = undefined;
    private propertyFieldListPickerOrderBy: any = undefined;
    private propertyFieldIconPicker: any = undefined;
    private propertyFieldNumber: any = undefined;
    private propertyFieldColorPicker: any = undefined;
    private propertyFieldCollectionData: any = undefined;
    private isSearchPositionAvailable: boolean = false;


    public render(): void {
        const element: React.ReactElement<IModernQuickLinksProps> = React.createElement(
            ModernQuickLinks,
            {
                context: this.context,
                spService: this._spService,
                httpClient: this.context.httpClient,
                tName: this.context.pageContext.legacyPageContext.tenantDisplayName,
                siteurl: this.context.pageContext.legacyPageContext.webAbsoluteUrl,
                serRelativeUrl: this.context.pageContext.legacyPageContext.webServerRelativeUrl,
                isSiteAdmin: this.context.pageContext.legacyPageContext.isSiteAdmin,

                theme: this._currentTheme,
                isDarkTheme: this._isDarkTheme,
                environmentMessage: this._environmentMessage,
                hasTeamsContext: !!this.context.sdks.microsoftTeams,
                userDisplayName: this.context.pageContext.user.displayName,
                displayMode: this.displayMode,
                propertyPane: this.context.propertyPane,
                wpZoneWidth: this.width,
                wpInstanceId: this.instanceId,
                wpZoneWidthDynamic: this._webpartZoneWidth,
                designType: this.properties.designType,
                tileSize: this.properties.tileSize,

                showTitle: this.properties.showTitle,
                title: this.properties.title,
                titleIcon: this.properties.titleIcon,
                customTileCount: this.properties.customTileCount,
                enableCustomTiles: this.properties.enableCustomTiles,
                enableManageTiles: this.properties.enableManageTiles,
                globalList: this.properties.globalList,
                userList: this.properties.userList,
                enableGLPCache: this.properties.enableGLPCache,
                enableULPCache: this.properties.enableULPCache,
                freeFlow: this.properties.freeFlow,
                wpHeight: this.properties.wpHeight,
                glSortOrder: this.properties.glSortOrder,
                ulSortOrder: this.properties.ulSortOrder,
                glSortBy: this.properties.glSortBy,
                ulSortBy: this.properties.ulSortBy,
                enableTileInfo: this.properties.enableTileInfo,
                tileInfoPanelHeader: this.properties.tileInfoPanelHeader,

                enableSearch: this.properties.enableSearch,
                searchPlaceholder: this.properties.searchPlaceholder,
                enableSearchInContentArea: this.properties.searchInContentArea,
                searchWidth: this.properties.searchWidth,
                searchPosition: this.properties.searchPosition,
                enableKeywordSearch: this.properties.enableKeywordSearch,
                keywordSearchField: this.properties.keywordSearchField,

                useThemeColors: this.properties.useThemeColors,
                backgroundColor: this.properties.backgroundColor,
                fontColor: this.properties.fontColor,
                overflowBackgroundColor: this.properties.overflowBackgroundColor,
                overflowFontColor: this.properties.overflowFontColor,
                actionIconColor: this.properties.actionIconColor,
                actionIconHoverColor: this.properties.actionIconHoverColor,
                themeColors: this.themeColors,
                useTileColors: this.properties.useTileColors,
                listFilters: this.properties.listFilters,
                useGLFilter: this.properties.useGLFilter
            }
        );
        this.isSearchPositionAvailable = _isSearchPositionAvailable(this._webpartZoneWidth, this.width);
        ReactDom.render(element, this.domElement);
    }

    protected async onInit(): Promise<void> {
        this._environmentMessage = await this._getEnvironmentMessage();
        super.onInit();
        this._spService = new SPService(this.context);
    }

    private _getEnvironmentMessage(): Promise<string> {
        if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
            return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
                .then(context => {
                    let environmentMessage: string = '';
                    switch (context.app.host.name) {
                        case 'Teams': // running in Teams
                            environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
                            break;
                        default:
                            throw new Error('Unknown host');
                    }

                    return environmentMessage;
                });
        }

        return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
    }

    protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
        if (!currentTheme) {
            return;
        }
        this._currentTheme = currentTheme;
        this._isDarkTheme = !!currentTheme.isInverted;
        const {
            semanticColors
        } = currentTheme;

        if (semanticColors) {
            this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
            this.domElement.style.setProperty('--link', semanticColors.link || null);
            this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
            this.domElement.style.setProperty('--actionLinkHovered', semanticColors.actionLinkHovered || null);
            this.domElement.style.setProperty('--primaryButtonBackground', this._isDarkTheme ? semanticColors.buttonText : this._currentTheme.palette.themePrimary || null);
            this.domElement.style.setProperty('--primaryButtonText', this._isDarkTheme ? semanticColors.buttonBackground : semanticColors.primaryButtonText || null);
            this.domElement.style.setProperty('--primaryButtonBg', semanticColors.primaryButtonBackground || null);
            this.domElement.style.setProperty('--neutralDark', currentTheme.palette.neutralDark || null);
            this.domElement.style.setProperty('--neutralPrimary', currentTheme.palette.neutralPrimary || null);
            this.domElement.style.setProperty('--themeSecondary', currentTheme.palette.themeSecondary || null);
            this.domElement.style.setProperty('--bodyBg', semanticColors.bodyStandoutBackground || null);
            this.domElement.style.setProperty('--bodyBgHovered', this._isDarkTheme ? semanticColors.bodyBackgroundHovered : this._currentTheme.palette.themeLighter || null);
            this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
            this.domElement.style.setProperty('--actLink', semanticColors.actionLink || null);
            this.domElement.style.setProperty('--actLinkHovered', this._isDarkTheme ? semanticColors.variantBorderHovered : this._currentTheme.palette.themePrimary || null);
            this.domElement.style.setProperty('--buttonBgColor', this._currentTheme.palette.themePrimary || null);
            this.domElement.style.setProperty('--overflowBg', this._isDarkTheme ? semanticColors.primaryButtonBackgroundHovered : this._currentTheme.palette.black || null);
            //this.domElement.style.setProperty('--overflowFontColor', this.overflowFontColor || null);
            // if (this.properties.overflowFontColor) this.domElement.style.setProperty('--overflowFontColor', this.properties.overflowFontColor);
            this.domElement.style.setProperty('--overflowFontColor', this._isDarkTheme ? semanticColors.accentButtonText : this._currentTheme.palette.white || null);
            this.domElement.style.setProperty('--cardHoverBorder', this._isDarkTheme ? semanticColors.variantBorderHovered : semanticColors.primaryButtonBackground || null);
            this.domElement.style.setProperty('--cardPrevBg', this._isDarkTheme ? semanticColors.bodyBackgroundHovered : semanticColors.bodyStandoutBackground || null);
            this.domElement.style.setProperty('--cardPrevFontColor', this._isDarkTheme ? semanticColors.variantBorderHovered : this._currentTheme.palette.black || null);

            this.themeColors = {
                primaryButtonBackground: this._isDarkTheme ? semanticColors.buttonText : this._currentTheme.palette.themePrimary || null,
                primaryButtonText: this._isDarkTheme ? semanticColors.buttonBackground : semanticColors.primaryButtonText || null,
                primaryButtonBg: semanticColors.primaryButtonBackground || null,
                overflowBg: this._isDarkTheme ? semanticColors.primaryButtonBackgroundHovered : this._currentTheme.palette.black || null,
                overflowFontColor: this._isDarkTheme ? semanticColors.accentButtonText : this._currentTheme.palette.white || null,
                actLink: semanticColors.actionLink || null,
                actLinkHovered: this._isDarkTheme ? semanticColors.variantBorderHovered : this._currentTheme.palette.themePrimary || null,
                bodyText: semanticColors.bodyText || null,
                bodyBgHovered: this._isDarkTheme ? semanticColors.bodyBackgroundHovered : this._currentTheme.palette.themeLighter || null,
                buttonBgColor: this._currentTheme.palette.themePrimary || null,
                cardHoverBorder: this._isDarkTheme ? semanticColors.variantBorderHovered : semanticColors.primaryButtonBackground || null,
                cardPrevBg: this._isDarkTheme ? semanticColors.bodyBackgroundHovered : semanticColors.bodyStandoutBackground || null,
                cardPrevFontColor: this._isDarkTheme ? semanticColors.variantBorderHovered : this._currentTheme.palette.black || null
            };
        }
    }

    private getFields = async (): Promise<ICustomDropdownOption[]> => {
        let retOptions: ICustomDropdownOption[] = [];
        try {
            const lstFields = await this._spService.getFields('', this.properties.globalList);
            const results = filter(lstFields, (f) => {
                return restFields.filter(rf => rf.toLowerCase() === f.StaticName.toLowerCase()).length <= 0 &&
                    restFieldTypes.filter(rf => rf.toLowerCase() === f.TypeAsString.toLowerCase()).length <= 0;
            });
            this._allFields = results;
            if (results.length > 0) {
                retOptions = results.map((f) => {
                    return {
                        key: f.InternalName,
                        text: f.Title
                    };
                });
            }
        } catch (err) {
            console.error(err);
        }
        return retOptions;
    };

    protected onDispose(): void {
        ReactDom.unmountComponentAtNode(this.domElement);
    }

    protected get dataVersion(): Version {
        return Version.parse('1.0');
    }

    protected override get disableReactivePropertyChanges(): boolean {
        return true;
    }

    protected onAfterResize(newWidth: number): void {
        this._webpartZoneWidth = newWidth;
        this.render();
        //console.log("Webpart width: ", newWidth);
    }

    private onGlobalListPickerValidation(value: string | string[] | IPropertyFieldList | IPropertyFieldList[]): string | Promise<string> {
        if (!value) {
            return strings.WP_GL_VAL_NOTSEL;
        }
        if (this.properties.userList && this.properties.userList === value) return "Global List and User List cannot be same.";
        return '';
    }

    private onUserListPickerValidation(value: string | string[] | IPropertyFieldList | IPropertyFieldList[]): string | Promise<string> {
        if (!value) {
            return strings.WP_UL_VAL_NOTSEL;
        }
        if (this.properties.globalList && this.properties.globalList === value) return "Global List and User List cannot be same.";
        return '';
    }

    protected async loadPropertyPaneResources(): Promise<void> {
        const { PropertyFieldSliderWithCallout } = await import(
            /* webpackChunkName: 'pnp-propcontrols-sliderwithcallout' */
            '@pnp/spfx-property-controls/lib/PropertyFieldSliderWithCallout'
        );
        this.propertyFieldSliderWithCallout = PropertyFieldSliderWithCallout;

        const { PropertyFieldToggleWithCallout } = await import(
            /* webpackChunkName: 'pnp-propcontrols-togglewithcallout' */
            '@pnp/spfx-property-controls/lib/PropertyFieldToggleWithCallout'
        );
        this.propertyFieldToggleWithCallout = PropertyFieldToggleWithCallout;

        const { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } = await import(
            /* webpackChunkName: 'pnp-propcontrols-listpicker' */
            '@pnp/spfx-property-controls/lib/PropertyFieldListPicker'
        );
        this.propertyFieldListPicker = PropertyFieldListPicker;
        this.propertyFieldListPickerOrderBy = PropertyFieldListPickerOrderBy;

        const { PropertyFieldIconPicker } = await import(
            /* webpackChunkName: 'pnp-propcontrols-iconpicker' */
            '@pnp/spfx-property-controls/lib/PropertyFieldIconPicker'
        );
        this.propertyFieldIconPicker = PropertyFieldIconPicker;

        const { PropertyFieldNumber } = await import(
            /* webpackChunkName: 'pnp-propcontrols-propertyfieldnumber' */
            '@pnp/spfx-property-controls/lib/PropertyFieldNumber'
        );
        this.propertyFieldNumber = PropertyFieldNumber;

        const { PropertyFieldColorPicker } = await import(
            /* webpackChunkName: 'pnp-propcontrols-propertyfieldcolorpicker' */
            '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker'
        );
        this.propertyFieldColorPicker = PropertyFieldColorPicker;

        const { PropertyFieldCollectionData } = await import(
            /* webpackChunkName: 'pnp-propcontrols-propertyfieldcollectiondata' */
            '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData'
        );
        this.propertyFieldCollectionData = PropertyFieldCollectionData;
        this._fields = await this.getFields();
    }

    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
        let Ctrl_tileSize: any = [];
        if (this.properties.designType === 'Tile') {
            Ctrl_tileSize = PropertyPaneChoiceGroup('tileSize', {
                label: strings.WP_L_TileSize,
                options: [
                    {
                        key: 'Small',
                        text: strings.WP_TS_S,
                    },
                    {
                        key: 'Medium',
                        text: strings.WP_TS_M
                    },
                    {
                        key: 'Large',
                        text: strings.WP_TS_L
                    },
                    {
                        key: 'Extra Large',
                        text: strings.WP_TS_XL
                    }
                ]
            });
        }
        let Ctrl_title: any = [];
        let Ctrl_titleIcon: any = [];
        if (this.properties.showTitle) {
            Ctrl_title = PropertyPaneTextField('title', {
                label: strings.WP_L_Title,
                placeholder: strings.WP_PH_Title,
                value: this.properties.title
            });
            Ctrl_titleIcon = this.propertyFieldIconPicker('titleIcon', {
                currentIcon: this.properties.titleIcon,
                key: "iconPickerId",
                onSave: (icon: string) => { console.log(icon); this.properties.titleIcon = icon; },
                onChanged: (icon: string) => { console.log(icon); },
                buttonLabel: strings.WP_Btn_HeaderIco,
                renderOption: "panel",
                properties: this.properties,
                onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                label: strings.WP_L_HeaderIco
            });
        }
        let Ctrl_wpHeight: any = [];
        let Ctrl_wpHeightInfo: any = [];
        if (!this.properties.freeFlow) {
            Ctrl_wpHeight = this.propertyFieldNumber('wpHeight', {
                key: 'wpHeightField',
                label: 'Height (px)',
                description: 'Height is configured in px. Enter only number.',
                value: this.properties.wpHeight,
                minValue: 100,
                maxValue: 1000
            });
            Ctrl_wpHeightInfo = PropertyPaneWebPartInformation({
                key: 'heightWebpartInfo',
                description: `Scroll bars will be displayed if the links exceeded the height.`
            });
        }
        let Ctrl_glSort: any = [];
        let Ctrl_ulSort: any = [];
        Ctrl_glSort = PropertyPaneChoiceGroup('glSortBy', {
            label: 'Sort By',
            options: [
                {
                    key: 'ID',
                    text: 'ID',
                },
                {
                    key: 'Sequence',
                    text: 'Sequence'
                },
                {
                    key: 'Title',
                    text: 'Title'
                },
                {
                    key: 'Created',
                    text: 'Created'
                },
                {
                    key: 'Modified',
                    text: 'Modified'
                }
            ]
        });
        Ctrl_ulSort = PropertyPaneChoiceGroup('ulSortBy', {
            label: 'Sort By',
            options: [
                {
                    key: 'ID',
                    text: 'ID',
                },
                {
                    key: 'Sequence',
                    text: 'Sequence'
                },
                {
                    key: 'Title',
                    text: 'Title'
                },
                {
                    key: 'Created',
                    text: 'Created'
                },
                {
                    key: 'Modified',
                    text: 'Modified'
                }
            ]
        });
        let ctrl_search_position: any = [];
        ctrl_search_position = PropertyPaneChoiceGroup('searchPosition', {
            label: 'Search Position',
            options: [
                {
                    key: 'start',
                    text: 'Left',
                },
                {
                    key: 'center',
                    text: 'Center'
                },
                {
                    key: 'end',
                    text: 'Right'
                }
            ],
        });
        let ctrl_listFilters: any = [];
        if (this.properties.globalList && this.properties.useGLFilter) {
            ctrl_listFilters = this.propertyFieldCollectionData('listFilters', {
                key: 'collectionData',
                label: 'Item Filters',
                panelHeader: 'Filters',
                manageBtnLabel: 'Data Filter',
                value: this.properties.listFilters,
                panelDescription: 'Add, edit, and remove filters',
                enableSorting: true,
                fields: [
                    {
                        id: 'fieldname',
                        title: 'Field Name',
                        type: CustomCollectionFieldType.dropdown,
                        required: true,
                        options: this._fields
                    },
                    {
                        id: 'operator',
                        title: 'Operator',
                        type: CustomCollectionFieldType.dropdown,
                        required: true,
                        options: [
                            {
                                key: 'eq',
                                text: 'Equals'
                            },
                            {
                                key: 'ne',
                                text: 'Not Equals'
                            },
                            {
                                key: 'gt',
                                text: 'Greater Than'
                            },
                            {
                                key: 'lt',
                                text: 'Less Than'
                            },
                            {
                                key: 'ge',
                                text: 'Greater Than or Equals'
                            },
                            {
                                key: 'le',
                                text: 'Less Than or Equals'
                            },
                            {
                                key: 'contains',
                                text: 'Contains'
                            },
                            {
                                key: 'startswith',
                                text: 'Starts With'
                            },
                            {
                                key: 'endswith',
                                text: 'Ends With'
                            }
                        ]
                    },
                    {
                        id: 'value',
                        title: 'Value',
                        type: CustomCollectionFieldType.string,
                        required: true
                    },
                    {
                        id: 'andOr',
                        title: 'Condition',
                        type: CustomCollectionFieldType.dropdown,
                        options: [
                            {
                                key: 'and',
                                text: 'And'
                            },
                            {
                                key: 'or',
                                text: 'Or'
                            }
                        ]
                    }
                ]
            });
        }
        return {
            pages: [
                {
                    header: {
                        description: ''
                    },
                    displayGroupsAsAccordion: true,
                    groups: [
                        // General Settings
                        {
                            groupName: strings.WP_GN_Gen,
                            isCollapsed: false,
                            groupFields: [
                                PropertyPaneToggle('showTitle', {
                                    label: strings.WP_L_STitle,
                                    checked: this.properties.showTitle,
                                    onText: strings.WP_T_S,
                                    offText: strings.WP_T_H
                                }),
                                Ctrl_title,
                                Ctrl_titleIcon
                            ]
                        },
                        // Layout Settings
                        {
                            groupName: strings.WP_GN_Layout,
                            isCollapsed: false,
                            groupFields: [
                                PropertyPaneChoiceGroup('designType', {
                                    label: strings.WP_L_LayOpt,
                                    options: [
                                        {
                                            key: 'Tile',
                                            text: strings.WP_L_LO_Tile,
                                            iconProps: { officeFabricIconFontName: 'Tiles' },
                                        },
                                        {
                                            key: 'Button',
                                            text: strings.WP_L_LO_Button,
                                            iconProps: { officeFabricIconFontName: 'ButtonControl' }
                                        },
                                        {
                                            key: 'Compact',
                                            text: strings.WP_L_LO_Comp,
                                            iconProps: { officeFabricIconFontName: 'GroupedList' }
                                        },
                                        {
                                            key: 'Grid',
                                            text: strings.WP_L_LO_Grid,
                                            iconProps: { officeFabricIconFontName: 'PictureTile' }
                                        }
                                    ]
                                }),
                                Ctrl_tileSize,
                                this.propertyFieldToggleWithCallout('freeFlow', {
                                    key: 'freeFlowTilesFieldId',
                                    label: 'Free Flow',
                                    checked: this.properties.freeFlow
                                }),
                                Ctrl_wpHeight,
                                Ctrl_wpHeightInfo,
                            ]
                        },
                        // Color Settings
                        {
                            groupName: strings.WP_GN_Color,
                            isCollapsed: true,
                            groupFields: [
                                PropertyPaneToggle('useTileColors', {
                                    label: 'Use Individual Tile Colors',
                                    checked: this.properties.useTileColors,
                                    onText: 'On',
                                    offText: 'Off'
                                }),
                                PropertyPaneWebPartInformation({
                                    key: 'IndividualTileColorsInfo',
                                    description: `Individual link colors will be applied to each link and its applicable only for global links.`
                                }),
                                PropertyPaneToggle('useThemeColors', {
                                    label: 'Use Theme Colors',
                                    checked: this.properties.useThemeColors,
                                    onText: 'On',
                                    offText: 'Off'
                                }),
                                (this.properties.designType === 'Tile' || this.properties.designType === 'Compact'
                                    || this.properties.designType === 'Button' || this.properties.designType === 'Grid') &&
                                this.propertyFieldColorPicker('backgroundColor', {
                                    label: this.properties.designType === 'Compact' || this.properties.designType === 'Button' ? 'Border Color' : 'Background Color',
                                    selectedColor: this.properties.backgroundColor,
                                    onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                                    properties: this.properties,
                                    key: 'backgroundColorFieldId',
                                    style: PropertyFieldColorPickerStyle.Inline,
                                    disabled: this.properties.useThemeColors
                                }),
                                (this.properties.designType === 'Tile' || this.properties.designType === 'Compact'
                                    || this.properties.designType === 'Button' || this.properties.designType === 'Grid') &&
                                this.propertyFieldColorPicker('fontColor', {
                                    label: 'Icon & Text Color',
                                    selectedColor: this.properties.fontColor,
                                    onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                                    properties: this.properties,
                                    key: 'fontColorFieldId',
                                    style: PropertyFieldColorPickerStyle.Inline,
                                    disabled: this.properties.useThemeColors
                                }),
                                (this.properties.designType === 'Tile' || this.properties.designType === 'Compact'
                                    || this.properties.designType === 'Button') &&
                                this.propertyFieldColorPicker('overflowBackgroundColor', {
                                    label: this.properties.designType === 'Compact' ? 'Background Color' :
                                        this.properties.designType === 'Button' ? 'Hover Background Color' : 'Overflow Background Color',
                                    selectedColor: this.properties.overflowBackgroundColor,
                                    onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                                    properties: this.properties,
                                    key: 'overflowBackgroundColorFieldId',
                                    style: PropertyFieldColorPickerStyle.Inline,
                                    disabled: this.properties.useThemeColors
                                }),
                                (this.properties.designType === 'Tile' || this.properties.designType === 'Button' || this.properties.designType === 'Grid') &&
                                this.propertyFieldColorPicker('overflowFontColor', {
                                    label: this.properties.designType === 'Button' ? 'Hover Icon & Text Color' :
                                        this.properties.designType === 'Grid' ? 'Hover Color' : 'Overflow Font Color',
                                    selectedColor: this.properties.overflowFontColor,
                                    onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                                    properties: this.properties,
                                    key: 'overflowFontColorFieldId',
                                    style: PropertyFieldColorPickerStyle.Inline,
                                    disabled: this.properties.useThemeColors
                                }),
                                this.propertyFieldColorPicker('actionIconColor', {
                                    label: 'Action Icon Color',
                                    selectedColor: this.properties.actionIconColor,
                                    onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                                    properties: this.properties,
                                    key: 'actionIconColorFieldId',
                                    style: PropertyFieldColorPickerStyle.Inline,
                                    disabled: this.properties.useThemeColors
                                }),
                                this.propertyFieldColorPicker('actionIconHoverColor', {
                                    label: 'Action Icon Hover Color',
                                    selectedColor: this.properties.actionIconHoverColor,
                                    onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                                    properties: this.properties,
                                    key: 'actionIconHoverColorFieldId',
                                    style: PropertyFieldColorPickerStyle.Inline,
                                    disabled: this.properties.useThemeColors
                                }),
                            ]
                        }                        
                    ]
                },
                {
                    header: {
                        description: ''
                    },
                    groups: [
                        // Search Settings
                        {
                            groupName: 'Search Settings',
                            groupFields: [
                                PropertyPaneToggle('enableSearch', {
                                    label: strings.WP_Srch_Enable,
                                    checked: this.properties.enableSearch,
                                    onText: strings.WP_T_S,
                                    offText: strings.WP_T_H
                                }),
                                PropertyPaneTextField('searchPlaceholder', {
                                    label: strings.WP_Srch_Placeholder,
                                    placeholder: strings.WP_Srch_PlaceholderPH,
                                    value: this.properties.searchPlaceholder
                                }),
                                PropertyPaneToggle('searchInContentArea', {
                                    label: strings.WP_Srch_ContentEnable,
                                    checked: this.properties.searchInContentArea,
                                    onText: strings.WP_T_S,
                                    offText: strings.WP_T_H
                                }),
                                PropertyPaneWebPartInformation({
                                    key: 'SearchInfoDescription',
                                    description: strings.WP_Srch_Content_Info
                                }),
                                this.properties.searchInContentArea && this.isSearchPositionAvailable ? ctrl_search_position : [],
                                this.propertyFieldNumber('searchWidth', {
                                    key: 'searchWidthField',
                                    label: strings.WP_Srch_Width_Label,
                                    description: strings.WP_Srch_Width_Desc,
                                    value: this.properties.searchWidth,
                                    minValue: 200,
                                    maxValue: 300
                                }),
                                PropertyPaneToggle('enableKeywordSearch', {
                                    label: strings.WP_Srch_KWEnable,
                                    checked: this.properties.enableKeywordSearch,
                                    onText: strings.WP_T_S,
                                    offText: strings.WP_T_H
                                }),
                                PropertyPaneWebPartInformation({
                                    key: 'KeywordSearchDescription',
                                    description: strings.WP_KW_Desc
                                }),
                                PropertyPaneTextField('keywordSearchField', {
                                    label: strings.WP_KW_Field,
                                    placeholder: strings.WP_KW_FieldPH,
                                    value: this.properties.keywordSearchField,
                                    disabled: !this.properties.enableKeywordSearch
                                }),
                            ]
                        }
                    ]
                },
                {
                    header: {
                        description: ''
                    },
                    groups: [
                        {
                            groupName: strings.WP_GN_LstSett,
                            groupFields: [
                                this.propertyFieldListPicker('globalList', {
                                    label: strings.Settings_GTL,
                                    selectedList: this.properties.globalList,
                                    includeHidden: false,
                                    orderBy: this.propertyFieldListPickerOrderBy.Title,
                                    disabled: false,
                                    onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                                    properties: this.properties,
                                    context: this.context,
                                    onGetErrorMessage: this.onGlobalListPickerValidation.bind(this),
                                    deferredValidationTime: 0,
                                    key: 'globalListPickerId'
                                }),
                                PropertyPaneToggle('useGLFilter', {
                                    label: 'Use Filters',
                                    checked: this.properties.useGLFilter,
                                    onText: 'Yes',
                                    offText: 'No'
                                }),
                                ctrl_listFilters,
                                this.propertyFieldToggleWithCallout('enableGLPCache', {
                                    key: 'enableGLPCacheFieldId',
                                    label: strings.Settings_Cache_GT,
                                    checked: this.properties.enableGLPCache,
                                    calloutContent: React.createElement('div', {}, strings.Settings_Cache_GT_Desc),
                                    calloutTrigger: CalloutTriggers.Hover
                                }),
                                this.propertyFieldToggleWithCallout('glSortOrder', {
                                    key: 'glSortOrderFieldId',
                                    label: 'Sort Order',
                                    checked: this.properties.glSortOrder,
                                    onText: strings.WP_S_Asc,
                                    offText: strings.WP_S_Desc
                                }),
                                Ctrl_glSort,
                                PropertyPaneToggle('enableTileInfo', {
                                    label: strings.WP_GL_Info,
                                    checked: this.properties.enableTileInfo,
                                    onText: strings.WP_T_S,
                                    offText: strings.WP_T_H
                                }),
                                PropertyPaneTextField('tileInfoPanelHeader', {
                                    label: 'Information Panel Header Text',
                                    placeholder: 'Header Text...',
                                    value: this.properties.tileInfoPanelHeader,
                                    disabled: !this.properties.enableTileInfo
                                }),
                                PropertyPaneWebPartInformation({
                                    key: 'TileInfoDescription',
                                    description: strings.WP_GL_InfoDesc
                                }),
                                this.propertyFieldToggleWithCallout('enableCustomTiles', {
                                    key: 'enableCustomTilesFieldId',
                                    label: strings.Settings_EN_UserTile,
                                    checked: this.properties.enableCustomTiles
                                }),
                                this.propertyFieldListPicker('userList', {
                                    label: strings.Settings_UTL,
                                    selectedList: this.properties.userList,
                                    includeHidden: false,
                                    orderBy: this.propertyFieldListPickerOrderBy.Title,
                                    onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                                    properties: this.properties,
                                    context: this.context,
                                    onGetErrorMessage: this.onUserListPickerValidation.bind(this),
                                    deferredValidationTime: 0,
                                    key: 'userListPickerId',
                                    disabled: !this.properties.enableCustomTiles
                                }),
                                this.propertyFieldToggleWithCallout('enableULPCache', {
                                    key: 'enableULPCacheFieldId',
                                    label: strings.Settings_Cache_UT,
                                    checked: this.properties.enableULPCache,
                                    calloutContent: React.createElement('div', {}, strings.Settings_Cache_UT_Desc),
                                    calloutTrigger: CalloutTriggers.Hover,
                                    disabled: !this.properties.enableCustomTiles
                                }),
                                this.propertyFieldSliderWithCallout('customTileCount', {
                                    key: 'customTileCountFieldId',
                                    label: strings.Settings_UT_Limit,
                                    max: 18,
                                    min: 1,
                                    step: 1,
                                    showValue: true,
                                    value: this.properties.customTileCount,
                                    disabled: !this.properties.enableCustomTiles
                                }),
                                this.propertyFieldToggleWithCallout('ulSortOrder', {
                                    key: 'ulSortOrderFieldId',
                                    label: strings.WP_SO_Label,
                                    checked: this.properties.ulSortOrder,
                                    onText: strings.WP_S_Asc,
                                    offText: strings.WP_S_Desc,
                                }),
                                Ctrl_ulSort,
                                this.propertyFieldToggleWithCallout('enableManageTiles', {
                                    key: 'enableManageTilesFieldId',
                                    label: strings.Settings_Man_Title,
                                    checked: this.properties.enableManageTiles,
                                    disabled: !this.properties.enableCustomTiles
                                }),
                                PropertyPaneWebPartInformation({
                                    key: 'ManageTilesWebpartInfo',
                                    description: strings.WP_MT_Desc
                                })
                            ]
                        }
                    ]
                },
                {
                    header: {
                        description: ''
                    },
                    groups: [
                        {
                            groupName: 'Author Info',
                            groupFields: [
                                PropertyPaneLink('AuthorInfoLink', {
                                    text: 'Sudharsan K.',
                                    href: 'https://spknowledge.com/',
                                    target: '_blank'
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    }
}
