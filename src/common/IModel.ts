import { MessageScope } from "./Constants";

export interface IResult {
	res?: any;
	status: boolean;
}

export interface IMessageInfo {
	msg: string;
	scope: MessageScope;
}

export interface IWebLists {
    WebUrl: string;
    ListTitle?: string;
    ListID?: string;
}

export interface ISelSiteInfo {
    SiteId: string;
    WebId: string;
    Url: string;
    Title: string;
    Author: string;
    WebTemplate: string;
    WebUrl: string;
}

export interface ISelListInfo {
    Id: string;
    EntityTypeName: string;
    Title: string;
    ItemCount: number;
    LastItemModifiedDate: string;
}

export interface ISelectionInfo {
    site: ISelSiteInfo;
    lists: ISelListInfo[];
}