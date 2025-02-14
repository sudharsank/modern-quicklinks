import { LicenseMessage } from "../../../common/Constants";

export interface IModernQuickLinksState {
    loading: boolean;
    isOpen: boolean;
    launchPadItems: any[];
    searchItems: any[];
    userTiles: any[];
    showAddLink: boolean;
    addTitle: string;
    addDesc: string;
    reload: boolean;
    hideDialog: boolean;
    delItemID: number;
    saving: boolean;
    showAdd: boolean;
    showEdit: boolean;
    showManage: boolean;
    showTileColor: boolean;
    item: any;
    editItem: any;
    showConfig: boolean;
    showConfigDialog: boolean;
    showSettings: boolean;
    showLicenseDialog: boolean;
    licenseMessage: LicenseMessage;
    licInfo: any;
}