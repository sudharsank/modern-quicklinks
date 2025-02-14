declare interface ISpIntranetLibLibraryStrings {
    ErrorLstName: string;
    ErrorLogLstDesc: string;
    EL_FldTitle: string;
    EL_FldMessage: string;
    EL_FldStack: string;

    SettingsLstName: string;
    SettingsLstDesc: string;
    SL_FldTitle: string;
    SL_FldConVal: string;
    SL_FldDesc: string;
    SL_FldCat: string;

	LaunchPadGlobalLstDesc: string;
	LaunchPadUserLstDesc: string;
	Fld_LP_Url: string;
	Fld_LP_Icon: string;
    Fld_LP_Image: string;
	Fld_LP_Window: string;	

    Fld_Created: string;
    Fld_Author: string;
    Fld_Modified: string;
    Fld_Editor: string;
    VW_AllItems: string;
	Fld_IsActive: string;
	Fld_Sequence: string;
	Fld_News: string;
    Fld_Desc: string;

    Fld_News_CI: string;
    Fld_News_Desc: string;
    Fld_News_Cont: string;
}

declare interface ICommonStrings {
	Btn_Create: string;
	Btn_Cancel: string;
    Btn_Save: string;
	CloseAL: string;
    Btn_ChooseList: string;
    Btn_Verify: string;
    Btn_BackToList: string;
    Btn_BackToSite: string;
	Btn_Back: string;

    LicDialogTitle: string;
	SettingsDialogTitle: string;
    Desc_License: string;
	LicNotConfigured: string;
	LicExpired: string;
	LicLink: string;
	Invalid_LicKey: string;
	App_Act_Error: string;
	Desc_AppWeb: string;

    Msg_LoadingLists: string;
	Msg_Loading: string;
	Msg_Wait: string;
	Msg_LicSPAdminInfo: string;
	Msg_Mandatory: string;   
	Msg_SameList: string; 
	Msg_ConfigHeader: string;
	Msg_Config_CreateList: string;	
	Msg_Config_WPProp_Read: string;
	Msg_Config_WPProp_Edit: string;
    Msg_InvalidURL: string;
    Msg_SaveSuccess: string;
    Msg_ReqImage: string;
	Msg_NoData: string;

	Lbl_CreateLists: string;
	Lbl_Config_WPProp: string;

    NS_DialogTitle: string;
    Msg_LoadingSiteColl: string;
    Title_SitesContainer: string;
    Title_ListContainerBC: string;
    ListContainer_NoListDesc1: string;
    ListContainer_NoListDesc2: string;
    ListContainer_BCDesc: string;

    WP_GN_Gen: string;
    WP_L_Title: string;
    WP_PH_Title: string;    
    WP_L_LayOpt: string;
    WP_L_LO_Tile: string;
    WP_L_LO_Button: string;
    WP_L_LO_Comp: string;
    WP_L_LO_Grid: string;
    WP_L_TileSize: string;
    WP_TS_S: string;
    WP_TS_M: string;
    WP_TS_L: string;
    WP_TS_XL: string;
}

declare module 'SpIntranetLibLibraryStrings' {
    const strings: ISpIntranetLibLibraryStrings;
    export = strings;
}

declare module 'CommonStrings' {
    const strings: ICommonStrings;
    export = strings;
}
