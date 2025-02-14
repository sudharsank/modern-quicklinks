export const secretKey = 'p1r2a3j4j5a6n7r8e9v10a11t12h13y14';
export const apiUrl: string = 'https://ks-app-license.azurewebsites.net/api';
export const appDB: string = 'app-license';
export const appCont: string = 'licenseInfo';

export const tPropertyKey: string = 'app-ik-mql';
export const tOobListsKey: string = 'app-ik-mql-ooblists';
export const productName: string = 'app_modernquicklinks';
export const licList: string = 'ik-lic';

export const defaultDateFormat = 'DD/MM/YYYY';

export const selectListInfo = ['BaseTemplate', 'EnableVersioning', 'EntityTypeName', 'Hidden', 'ItemCount', 'ListItemEntityTypeFullName', 'Title'];

export const fld_tileColors = 'TileColors';

// List Filters Constants
export const restFields = ["ItemChildCount", "FolderChildCount", "_ComplianceFlags", "_ComplianceTag", "_ComplianceTagWrittenTime", "_ComplianceTagUserId",
    "AppAuthor", "AppEditor", "Attachments", "_UIVersionString", "ComplianceAssetId", "_ColorTag", "Created", "Modified"];
export const restFieldTypes = ["Thumbnail", "DateTime", "File", "Lookup", "Computed"];
export const fieldInfo: string[] = ["EntityPropertyName", "InternalName", "StaticName", "Title", "TypeAsString", "TypeDisplayName", "TypeShortDescription"];
export const manualCondns: string[] = ['contains', 'startswith', 'endswith', 'eq', 'ne'];
export const excepFields: string[] = ['Boolean'];

export const enum ListTemplates {
    Announcements = 104,
    DocumentLibrary = 101,
    Events = 106,
    GenericList = 100,
    Links = 103,
    Tasks = 107
}

export const enum DialogTypes {
    LicDialog,
    LaunchPadSettings,
    NewsListSettings
}

export const enum PanelTypes {
    ManageTiles,
    AddUpdateTiles,
    TileInfo,
    TileColor
}

export const enum DesignTypes {
    Tiles = "Tile",
    Buttons = "Button",
    Compact = "Compact",
    Grid = "Grid"
}

export enum LoaderType {
    Spinner = 0,
    Indicator = 1
}

export enum MessageScope {
    Success = 0,
    Failure = 1,
    Warning = 2,
    Info = 3
}

export enum LicenseMessage {
    NotConfigured = 0,
    Trial = 1,
    Expired = 2,
    Valid = 3,
    ConfigError = 4
}

export enum ItemSize {
    Small = "Small",
    Medium = "Medium",
    Large = "Large",
    "Extra Large" = "Extra Large"
}

export const enum SettingsCategory {
    News = 'News'
}

export const enum SettingsList {
    NLP = 'News_ListPath'
}

export const enum LogSource {
    News = "News",
    NewsDisplay = "News Display",
    NewsList = "News List",
    LaunchPad = "Launch Pad",
    LaunchPadSettings = "Launch Pad Settings"
}

export const enum IKLists {
    ErrorLogs = "IK-ErrorLogs",
    Settings = "IK-Settings",
    Asset = 'SiteAssets',
    Pages = 'Site Pages'
}

export const enum IKListCacheKeys {
    LPGlobal = 'IK-LPGlobal',
    LPUsers = 'IK-LPUsers'
}