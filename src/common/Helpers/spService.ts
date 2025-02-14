import * as strings from 'ModernQuickLinksWebPartStrings';
import { SPFI } from "@pnp/sp";
import { getSP } from "../../pnpjs-config";
import { PnPClientStorage } from "@pnp/core";
import { createBatch } from "@pnp/sp/batching";
import { fieldInfo, IKLists, licList, ListTemplates, LogSource, secretKey, tPropertyKey } from "../Constants";
import { IStorageEntity, IWeb, Web } from '@pnp/sp/webs/types';
import { ICamlQuery, IList, IListInfo, IRenderListDataAsStreamResult } from '@pnp/sp/lists/types';
import { IItemAddResult, IItemUpdateResult } from '@pnp/sp/items/types';
import { IFileAddResult } from '@pnp/sp/files/types';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
import { addDays } from 'office-ui-fabric-react';
import moment from 'moment';
import { FieldTypes, IField } from '@pnp/sp/fields';

export interface IQueryOption {
    select?: string[];
    filter?: string;
    expand?: string[];
    top?: number;
    skip?: number;
    orderby?: OrderByParams;
}

export interface OrderByParams {
    orderByCol: string;
    isAsc: boolean;
}


export default class SPService {
    private _sp: SPFI;
    private absUrl: string;
    private client: SPHttpClient;
    private storage = new PnPClientStorage();

    constructor(context: WebPartContext) {
        this._sp = getSP(context);
        this.absUrl = context.pageContext.web.absoluteUrl;
        this.client = context.spHttpClient;
    }

    public getStorageValue = (key: string, isSession?: boolean): any => {
        try {
            if (key) {
                if (isSession) return this.storage.session.get(key);
                else return this.storage.local.get(key);
            }
        } catch (e) {
            console.log(e);
        }
        return undefined;
    }

    public writeErrorLog = async (err: any, source: LogSource): Promise<void> => {
        try {
            await this._sp.web.lists.getByTitle(IKLists.ErrorLogs).items.add({
                Title: source,
                Message: err.message,
                Stack: err.stack
            });
        } catch (e) {
            console.log(e);
        }
    }

    public createStorageValue = (key: string, value: any, expire: Date, isSession?: boolean): boolean => {
        try {
            if (key) {
                if (isSession) this.storage.session.put(key, value, expire);
                else this.storage.local.put(key, value, expire);
            }
            return true;
        } catch (e) {
            console.log(e);
        }
        return false;
    }

    public deleteStorageValue = (key: string, isSession?: boolean): void => {
        if (key) {
            if (isSession) this.storage.session.delete(key);
            else this.storage.local.delete(key);
        }
    }

    public deleteExpiredStorage = (): void => {
        this.storage.local.deleteExpired();
        this.storage.session.deleteExpired();
    }

    public encryptData = (data: string): string => {
        try {
            if (data) {
                const crypto: any = require('crypto-js');
                return crypto.AES.encrypt(data, secretKey).toString();
            }
        } catch (e) {
            console.log(e);
        }
    }

    public decryptData = (data: string): any => {
        try {
            if (data) {
                const crypto: any = require('crypto-js');
                var bytes = crypto.AES.decrypt(data, secretKey);
                return bytes.toString(crypto.enc.Utf8);
            }
        } catch (e) {
            console.log(e);
        }
    }

    public getTenantProp = async (key: string): Promise<any> => {
        try {
            if (key) {
                const aw = await this._sp.getTenantAppCatalogWeb();
                const tprop: IStorageEntity = await aw.getStorageEntity(key);
                if (tprop && tprop.Value) {
                    return this.decryptData(tprop.Value);
                }
            }
        } catch (e) {
            console.log(e);
        }
    }

    public removeLicense = async (): Promise<void> => {
        try {
            await this._sp.web.lists.getByTitle(licList).delete();
            await this.deleteStorageValue(tPropertyKey)
        } catch (err) {
            console.error(err);
        }
    }

    public expireLicense = async (): Promise<void> => {
        try {
            const aw = await this._sp.web;
            const licitems: any[] = await aw.lists.getByTitle(licList).items.select('lic')();
            if (licitems[0].lic) {
                let licInfo = this.decryptData(licitems[0].lic);
                licInfo = JSON.parse(licInfo);
                licInfo.ed = addDays(new Date(), -1).toISOString();
                licInfo = JSON.stringify(licInfo);
                this.createStorageValue(tPropertyKey, this.encryptData(licInfo), new Date(moment().add(4, 'hours').toISOString()));
            }
        } catch (err) {
            console.error(err);
        }
    }

    public getSiteLicense = async (): Promise<any> => {
        try {
            const aw = await this._sp.web;
            //this.removeLicense();
            //this.expireLicense();
            const licitems: any[] = await aw.lists.getByTitle(licList).items.select('lic')();
            if (licitems[0].lic) return this.decryptData(licitems[0].lic);
        } catch (e) {
            console.log(e);
        }
    }

    public setTenantProp = async (key: string, propValue: any, appWeb: string): Promise<boolean> => {
        try {
            if (key && propValue) {
                const aw = await this._sp.getTenantAppCatalogWeb();
                let enVal: string = this.encryptData(propValue);
                await aw.setStorageEntity(key, enVal, strings.Desc_License);
                return true;
            } else return false;
        } catch (e) {
            console.log(e);
        }
    }

    public setSiteLicense = async (propValue: any): Promise<boolean> => {
        try {
            const aw = await this._sp.web;
            let enVal: string = this.encryptData(propValue);
            try {
                await aw.lists.add(licList, '', 100, false, { Hidden: true });
                await aw.lists.getByTitle(licList).fields.addMultilineText('lic');
            } catch (err) {
                console.log(err);
            }
            await aw.lists.getByTitle(licList).items.add({
                lic: enVal
            });
            return true;
        } catch (err) {
            console.error(err);
        }
    }

    public removeTenantProp = async (key: string, appWeb: string): Promise<void> => {
        try {
            if (key) {
                const aw = await this._sp.getTenantAppCatalogWeb();
                await aw.removeStorageEntity(key);
                await this._sp.web.removeStorageEntity(key);
            }
        } catch (e) {
            console.log(e);
        }
    }

    public getAppCatalogWeb = async (): Promise<any> => {
        const appCat = await this._sp.getTenantAppCatalogWeb();
        return appCat();
    }

    public createErrorLogList = async (weburl?: string): Promise<boolean> => {
        let retVal: boolean = false;
        try {
            const errorList = await this._sp.web.lists.add(strings.ErrorLstName, strings.ErrorLogLstDesc, ListTemplates.GenericList, false, {
                EnableAttachments: false, Hidden: false
            });
            const [batchedListUpdates, executeListUpdates] = createBatch(errorList.list);
            errorList.list.using(batchedListUpdates);
            errorList.list.fields.addMultilineText(strings.EL_FldMessage);
            errorList.list.fields.addMultilineText(strings.EL_FldStack);
            const lstview = errorList.list.views.getByTitle(strings.VW_AllItems);
            lstview.fields.add(strings.EL_FldMessage);
            lstview.fields.add(strings.EL_FldStack);
            lstview.fields.add(strings.Fld_Created);
            lstview.fields.add(strings.Fld_Author);
            await executeListUpdates();
            retVal = true;
        } catch (e) {
            console.log(e);
            throw e;
        }
        return retVal;
    }

    public createSettingsList = async (weburl?: string): Promise<boolean> => {
        let retVal: boolean = false;
        try {
            const settingsList = await this._sp.web.lists.add(strings.SettingsLstName, strings.SettingsLstDesc, ListTemplates.GenericList, false, {
                EnableAttachments: false, Hidden: false
            });
            const [batchedListUpdates, executeListUpdates] = createBatch(settingsList.list);
            settingsList.list.using(batchedListUpdates);
            settingsList.list.fields.addMultilineText(strings.SL_FldConVal, { Required: true });
            settingsList.list.fields.addMultilineText(strings.SL_FldDesc);
            settingsList.list.fields.addText(strings.SL_FldCat, { Required: true });
            const lstview = settingsList.list.views.getByTitle(strings.VW_AllItems);
            lstview.fields.add(strings.SL_FldConVal);
            lstview.fields.add(strings.SL_FldDesc);
            lstview.fields.add(strings.SL_FldCat);
            lstview.fields.add(strings.Fld_Modified);
            lstview.fields.add(strings.Fld_Editor);
            await executeListUpdates();
            retVal = true;
        } catch (e) {
            console.log(e);
            throw e;
        }
        return retVal;
    }

    public createLaunchPadList = async (title: string, isGlobal: boolean): Promise<boolean> => {
        let retVal: boolean = false;
        try {
            const launchPadList = await this._sp.web.lists.add(title, isGlobal ? strings.LaunchPadGlobalLstDesc : strings.LaunchPadUserLstDesc, ListTemplates.GenericList, true, {
                EnableAttachments: false, Hidden: false
            });
            const [batchedListUpdates, executeListUpdates] = createBatch(launchPadList.list);
            launchPadList.list.using(batchedListUpdates);
            launchPadList.list.fields.addUrl(strings.Fld_LP_Url, { Required: true });
            launchPadList.list.fields.addText(strings.Fld_LP_Icon, { Required: false });
            launchPadList.list.fields.addImageField(strings.Fld_LP_Image, { Required: false });
            launchPadList.list.fields.addBoolean(strings.Fld_LP_Window, { Required: true });
            launchPadList.list.fields.addBoolean(strings.Fld_IsActive);
            launchPadList.list.fields.addNumber(strings.Fld_Sequence, { MinimumValue: 1, Required: false });
            if (isGlobal) launchPadList.list.fields.addMultilineText(strings.Fld_LP_Desc, { Required: false, RichText: true, RestrictedMode: false });
            const lstview = launchPadList.list.views.getByTitle(strings.VW_AllItems);
            lstview.fields.add(strings.Fld_LP_Url);
            lstview.fields.add(strings.Fld_LP_Icon);
            lstview.fields.add(strings.Fld_LP_Image);
            lstview.fields.add(strings.Fld_LP_Window);
            lstview.fields.add(strings.Fld_IsActive);
            lstview.fields.add(strings.Fld_Sequence);
            if (isGlobal) lstview.fields.add(strings.Fld_LP_Desc);
            lstview.fields.add(strings.Fld_Modified);
            lstview.fields.add(strings.Fld_Editor);
            lstview.fields.add(strings.Fld_Created);
            lstview.fields.add(strings.Fld_Author);
            await executeListUpdates();
            retVal = true;
        } catch (e) {
            console.log(e);
            throw e;
        }
        return retVal;
    }

    public getLists = async (webUrl: string, query: IQueryOption): Promise<any[]> => {
        let currentWeb = null;
        try {
            if (webUrl) currentWeb = Web([this._sp.web, webUrl]);
            else currentWeb = this._sp.web;
            let result = await currentWeb.lists
            if (query.select) result = result.select(query.select.join(','));
            if (query.filter) result = result.filter(query.filter);
            return result();
        } catch (e) {
            console.log(e);
            throw e;
        }
    }

    public getList = async (listTitle: string, listid: string, fields?: string[]): Promise<IListInfo> => {
        try {
            let lst = null;
            if (listTitle) lst = this._sp.web.lists.getByTitle(listTitle);
            else if (listid) lst = this._sp.web.lists.getById(listid);
            if (lst) {
                if (fields) lst = lst.select(fields.join(','));
                return await lst();
            } else return null;
        } catch (e) {
            console.log(e);
            throw e;
        }
    }

    public getItemCount = async (listTitle: string, listid: string): Promise<number> => {
        try {
            if (listTitle) return (await this._sp.web.lists.getByTitle(listTitle).items.select("ID")()).length;
            else if (listid) return (await this._sp.web.lists.getById(listid).items.select("ID")()).length;
            else return 0;
        } catch (e) {
            console.log(e);
            throw e;
        }
    }

    public checkFieldExists = async (listTitle: string, listid: string, field: string): Promise<boolean> => {
        let retVal: boolean = false;
        try {
            let fld: IField = null;
            if (listTitle) fld = (await this._sp.web.lists.getByTitle(listTitle).fields.getByInternalNameOrTitle(field)());
            else if (listid) fld = (await this._sp.web.lists.getById(listid).fields.getByInternalNameOrTitle(field)());
            if (fld) retVal = true;
        } catch (e) {
            console.log(e);
        }
        return retVal;
    }

    public createField = async (listTitle: string, listid: string, field: string, fieldType: string, isRequired: boolean, isHidden: boolean): Promise<boolean> => {
        let retVal: boolean = false;
        try {
            let lst: IList = null;
            if (listTitle) lst = this._sp.web.lists.getByTitle(listTitle);
            else if (listid) lst = this._sp.web.lists.getById(listid);
            if (lst) {
                switch (fieldType) {
                    case 'multiline':
                        await lst.fields.add(field, FieldTypes.Note, { Required: false, Hidden: false });
                        break;
                }
            }
        } catch (e) {
            console.log(e);
            throw e;
        }
        return retVal;
    }

    public getFields = async (listTitle: string, listid: string): Promise<any[]> => {
        try {
            let lst: IList = null;
            if (listTitle) lst = this._sp.web.lists.getByTitle(listTitle);
            else if (listid) lst = this._sp.web.lists.getById(listid);
            if (lst) {
                return await lst.fields.filter(`Hidden eq false and TypeAsString ne 'Computed'`).select(fieldInfo.toString())();
            }
        } catch (err) {
            console.log(err);
        }
    }

    // Listitem Methods

    public async addListItem(item: any, listTitle: string, listid: string): Promise<IItemAddResult> {
        try {
            if (item) {
                let lst: IList = null;
                if (listTitle) lst = this._sp.web.lists.getByTitle(listTitle);
                else if (listid) lst = this._sp.web.lists.getById(listid);
                if (lst !== null) {
                    return await lst.items.add(item);
                }
            }
        } catch (e) {
            console.log(e);
            throw e;
        }
    }

    public async updateListItem(item: any, listTitle: string, listid: string): Promise<IItemUpdateResult> {
        try {
            if (item) {
                let lst: IList = null;
                if (listTitle) lst = this._sp.web.lists.getByTitle(listTitle);
                else if (listid) lst = this._sp.web.lists.getById(listid);
                if (lst !== null) {
                    return await lst.items.getById(item.Id).update(item);
                }
            }
        } catch (e) {
            console.log(e);
            throw e;
        }
    }

    public async deleteListItem(listTitle: string, listid: string, itemid: string): Promise<boolean> {
        try {
            if (itemid) {
                let lst: IList = null;
                if (listTitle) lst = this._sp.web.lists.getByTitle(listTitle);
                else if (listid) lst = this._sp.web.lists.getById(listid);
                if (lst !== null) {
                    await lst.items.getById(parseInt(itemid)).delete();
                    return true;
                }
            }
        } catch (e) {
            console.log(e);
            throw e;
        }
    }

    public async updateByListID(item: any, title: string): Promise<void> {
        try {
            if (title && item) {
                const list = this._sp.web.lists.getById(title);
                await list.items.getById(item.Id).update(item);
            }
        } catch (e) {
            console.log(e);
            throw e;
        }
    }

    public async updateByListTitle(item: any, title: string): Promise<void> {
        try {
            if (title && item) {
                const list = this._sp.web.lists.getByTitle(title);
                await list.items.getById(item.Id).update(item);
            }
        } catch (e) {
            console.log(e);
            throw e;
        }
    }

    public async addMultiple(items: any[], title: string): Promise<void> {
        try {
            if (title && items.length > 0) {
                const [batchedListItems, executeItemsAdd] = this._sp.batched();
                const list = batchedListItems.web.lists.getByTitle(title);
                items.forEach(item => {
                    list.items.add(item);
                });
                await executeItemsAdd();
            }
        } catch (e) {
            console.log(e);
            throw e;
        }
    }

    public async getItemById(listTitle: string, listid: string, itemid: string): Promise<any> {
        try {
            if (itemid) {
                let lst: IList = null;
                if (listTitle) lst = this._sp.web.lists.getByTitle(listTitle);
                else if (listid) lst = this._sp.web.lists.getById(listid);
                if (lst !== null) {
                    return await lst.items.getById(parseInt(itemid))();
                }
            }
        } catch (e) {
            console.log(e);
            throw e;
        }
    }

    public async getItemsByQuery(listTitle: string, listid: string, queryOptions?: IQueryOption): Promise<any[]> {
        const { filter, select, expand, top, skip, orderby } = queryOptions;
        try {
            let lst: IList = undefined;
            if (listTitle) {
                lst = await this._sp.web.lists.getByTitle(listTitle);
            }
            else if (listid) {
                lst = await this._sp.web.lists.getById(listid);
            }
            else return [];
            let result = lst.items;
            if (filter) result = result.filter(filter);
            if (select) result = result.select(...select);
            if (expand) result = result.expand(...expand);
            if (orderby) result = result.orderBy(orderby.orderByCol, orderby.isAsc);
            if (top) result = result.top(top); else result.top(5000);
            if (skip) result = result.skip(skip);
            return result();

            // const [batchedItems, executeBatchItems] = this._sp.batched();
            // let finalRes: any[] = [];
            // let globalRes = batchedItems.web.lists.getById('1ed1d781-6f65-41c8-aa1a-578effe96c91').items;
            // if (filter) globalRes = globalRes.filter(filter);
            // //if (select) globalRes = globalRes.select(...select);
            // if (expand) globalRes = globalRes.expand(...expand);
            // if (orderby) globalRes = globalRes.orderBy(orderby.orderByCol, orderby.isAsc);
            // if (top) globalRes = globalRes.top(top); else globalRes.top(5000);
            // if (skip) globalRes = globalRes.skip(skip);
            // globalRes().then(r => finalRes.push(r));

            // let userRes = batchedItems.web.lists.getById('3f23be05-b52f-414a-88a7-3cead8f2508f').items;
            // if (filter) userRes = userRes.filter(filter);
            // //if (select) userRes = userRes.select(...select);
            // if (expand) userRes = userRes.expand(...expand);
            // if (orderby) userRes = userRes.orderBy(orderby.orderByCol, orderby.isAsc);
            // if (top) userRes = userRes.top(top); else userRes.top(5000);
            // if (skip) userRes = userRes.skip(skip);
            // userRes().then(r => finalRes.push(r));

            // await executeBatchItems();
            // if (finalRes && finalRes.length > 0) {
            // 	finalRes.forEach((itemarr: any[]) => {
            // 		if (itemarr.length > 0) {
            // 			console.log(itemarr[0]['odata.editLink'])
            // 		}
            // 	})
            // }
        } catch (e) {
            console.log(e);
            throw e;
        }
    }

    public async getItemsByQueryForWeb(webUrl: string, listTitle: string, listid: string, queryOptions?: IQueryOption): Promise<any[]> {
        const { filter, select, expand, top, skip, orderby } = queryOptions;
        let currentWeb = null;
        try {
            if (webUrl) currentWeb = Web([this._sp.web, webUrl]);
            let lst: IList = undefined;
            if (listTitle) {
                lst = webUrl ? await currentWeb.lists.getByTitle(listTitle) : await this._sp.web.lists.getByTitle(listTitle);
            }
            else if (listid) {
                lst = webUrl ? await currentWeb.lists.getById(listid) : await this._sp.web.lists.getById(listid);
            }
            else return [];
            let result = lst.items;
            if (filter) result = result.filter(filter);
            if (select) result = result.select(...select);
            if (expand) result = result.expand(...expand);
            if (orderby) result = result.orderBy(orderby.orderByCol, orderby.isAsc);
            if (top) result = result.top(top); else result.top(5000);
            if (skip) result = result.skip(skip);
            return result();
        } catch (e) {
            console.log(e);
            throw e;
        }
    }

    public async getItemsAsDataStreamUsingCAMLQuery(query: ICamlQuery, listTitle: string, listid: string, webUrl?: string): Promise<IRenderListDataAsStreamResult> {
        let currentWeb = null;
        try {
            if (webUrl) currentWeb = Web([this._sp.web, webUrl]);
            let lst: IList = undefined;
            if (listTitle) {
                lst = webUrl ? await currentWeb.lists.getByTitle(listTitle) : await this._sp.web.lists.getByTitle(listTitle);
            }
            else if (listid) {
                lst = webUrl ? await currentWeb.lists.getById(listid) : await this._sp.web.lists.getById(listid);
            }
            return lst.renderListDataAsStream(query);
        } catch (e) {
            console.log(e);
            throw e;
        }
    }

    // File Methods

    public addFile = async (folderPath: string, fileinfo: any, weburl?: string): Promise<IFileAddResult> => {
        let tmpWeb: IWeb = weburl ? Web(weburl) : this._sp.web;
        return await tmpWeb.getFolderByServerRelativePath(folderPath).files.addUsingPath(encodeURI(`${folderPath}/${fileinfo.name}`), fileinfo.content, { Overwrite: true });
    }

    public getFile = async (filepath: string): Promise<any> => {
        return await this._sp.web.getFileByUrl(filepath).getBlob();
    }
}
