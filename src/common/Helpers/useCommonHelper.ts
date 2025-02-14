import { IListInfo } from "@pnp/sp/lists";
import { useCallback } from "react";
import { HttpClient, IHttpClientOptions, HttpClientResponse, SPHttpClientConfiguration, ODataVersion } from '@microsoft/sp-http';
import { IKLists, restFields, restFieldTypes, SettingsCategory, tOobListsKey } from "../Constants";
import * as moment from "moment";
import { IFilePickerResult } from "@pnp/spfx-controls-react";
import { IFileAddResult } from "@pnp/sp/files";
import { Guid } from "@microsoft/sp-core-library";
import { ISelListInfo, ISelSiteInfo } from "../IModel";
import { IItem, IItemAddResult, IItemUpdateResult } from "@pnp/sp/items";
import SPService from "./spService";
import { filter } from "lodash";

export const useCommonHelper = (spService: SPService, webabsurl?: string) => {

    const checkAndCreateOOBLists = useCallback(async (): Promise<boolean> => {
        let lstCount: number = 0;
        let lst: IListInfo = undefined;
        try {
            const lstProperty = spService.getStorageValue(tOobListsKey);
            if (lstProperty) {
                return true;
            } else {
                try {
                    // Check for Error Logs list
                    lst = await spService.getList(IKLists.ErrorLogs, '');
                    lstCount = lstCount + 1;
                } catch (e) {
                    await spService.createErrorLogList();
                    lstCount = lstCount + 1;
                }
                // try {
                // 	// Check for Settings list
                // 	lst = await spService.getList(IKLists.Settings, '');
                // 	lstCount = lstCount + 1;
                // } catch (e) {
                // 	await spService.createSettingsList();
                // 	lstCount = lstCount + 1;
                // }
                if (lstCount === 1) {
                    spService.createStorageValue(tOobListsKey, true, new Date(moment().add(4, "hours").toString()));
                    return true;
                } else return false;
            }
        } catch (e) {
            console.log(e);
        }
    }, [spService]);

    const returnImageFieldInfoForWrite = useCallback(async (imgFieldVal: IFilePickerResult, destFolderName: string, weburl: string): Promise<string> => {
        try {
            let filename: string = `${Guid.newGuid().toString()}.${imgFieldVal.fileName}`;
            let sourceFile = await imgFieldVal.downloadFileContent(); //await spService.getFile(imgFieldVal.fileAbsoluteUrl);
            let destFile: IFileAddResult = undefined;
            destFile = await spService.addFile(`${weburl}/${destFolderName}`, { name: filename, content: sourceFile });
            if (destFile) {
                return JSON.stringify({
                    type: 'thumbnail',
                    fileName: imgFieldVal.fileName,
                    serverRelativeUrl: destFile.data.ServerRelativeUrl
                });
            } else {
                return JSON.stringify({
                    type: 'thumbnail',
                    fileName: imgFieldVal.fileName,
                    serverRelativeUrl: imgFieldVal.fileAbsoluteUrl
                });
            }
        } catch (e) {
            throw e;
        }
    }, [spService]);

    // const getAllSiteCollections = useCallback(async (): Promise<ISelSiteInfo[]> => {
    // 	let finalSitecoll: ISelSiteInfo[] = [];
    // 	if (httpClient && webabsurl) {
    // 		let restApiUrl: string = webabsurl + "/_api/search/query?querytext='(contentclass=STS_Site OR contentclass=STS_Web)(-WebTemplate:SPSPERS AND -SiteTemplate:APPCATALOG)'&selectproperties='SiteId,WebId,Path,Title,Author,WebTemplate,SPWebUrl,ParentLink'&trimduplicates=true&rowlimit=1000";
    // 		let config: SPHttpClientConfiguration = new SPHttpClientConfiguration({
    // 			defaultODataVersion: ODataVersion.v3
    // 		});
    // 		let siteCollResponse = (await httpClient.get(restApiUrl, config, { headers: { Accept: "application/json;odata=minimalmetadata;charset=utf-8" } })) as HttpClientResponse;
    // 		let siteColljson = await siteCollResponse.json();
    // 		let resultsList: any[] = siteColljson.PrimaryQueryResult.RelevantResults.Table.Rows;
    // 		resultsList.map(res => {
    // 			finalSitecoll.push({
    // 				SiteId: res.Cells[0].Value,
    // 				WebId: res.Cells[1].Value,
    // 				Url: res.Cells[2].Value,
    // 				Title: res.Cells[3].Value,
    // 				Author: res.Cells[4].Value,
    // 				WebTemplate: res.Cells[5].Value,
    // 				WebUrl: res.Cells[6].Value ? res.Cells[6].Value : res.Cells[2].Value,
    // 				//ParentLink: res.Cells[7].Value
    // 			});
    // 		});
    // 		finalSitecoll = finalSitecoll.filter(r => r.WebTemplate !== null);
    // 	}
    // 	return finalSitecoll;
    // }, [httpClient, webabsurl]);

    const getSiteLists = useCallback(async (siteurl: string): Promise<ISelListInfo[]> => {
        if (siteurl) {
            let result = await spService.getLists(siteurl, {
                filter: `Hidden eq false and BaseTemplate eq 100`,
                select: ['Id', 'ItemCount', 'Title', 'EntityTypeName', 'LastItemModifiedDate']
            });
            return result as ISelListInfo[];
        }
    }, [spService]);

    // const getSitePages = useCallback(async (siteurl: string, query: string): Promise<ISelListInfo[]> => {
    // 	if (siteurl) {
    // 		return await spService.getSitePages(siteurl, query);
    // 	}
    // }, [spService]);

    const getStorageValue = useCallback((storageKey: string): any => {
        try {
            if (storageKey) return spService.getStorageValue(storageKey, false);
        } catch (e) {
            throw e;
        }
    }, [spService]);

    const putStorageValue = useCallback((storageKey: string, storageVal: any, expire: Date, isSession: boolean): void => {
        try {
            if (storageKey) spService.createStorageValue(storageKey, storageVal, expire, isSession);
        } catch (e) {
            throw e;
        }
    }, [spService]);

    const deleteStorageValue = useCallback((storageKey: string, isSession: boolean): void => {
        try {
            spService.deleteStorageValue(storageKey, isSession);
        } catch (e) {
            console.log(e);
        }
    }, [spService]);

    const getConfigSettings = useCallback(async (all: boolean, config?: string, category?: SettingsCategory, wpid?: string): Promise<IItem[]> => {
        try {
            if (all) {
                return await spService.getItemsByQuery(IKLists.Settings, '', {
                    select: ['Title', 'ConfigValue', 'Category', 'ExtraParam'],
                    top: 500
                });
            } else {
                let filQuery: string = '';
                if (category && filQuery.length > 0) filQuery = filQuery + `Category eq '${category}'`; else filQuery = `Category eq '${category}'`;
                //if (category) filQuery.length > 0 ? filQuery = filQuery + `Category eq '${category}'` : filQuery = `Category eq '${category}'`;
                if (config && filQuery.length > 0) filQuery = filQuery + ` or Title eq '${config}'`; else filQuery = `Title eq '${config}'`;
                //if (config) filQuery.length > 0 ? filQuery = filQuery + ` or Title eq '${config}'` : filQuery = `Title eq '${config}'`;
                if (wpid && filQuery.length > 0) filQuery = filQuery + ` and (WebpartID eq '${wpid}')`; else filQuery = `WebpartID eq '${wpid}'`;
                //if (wpid) filQuery.length > 0 ? filQuery = filQuery + ` and (WebpartID eq '${wpid}')` : filQuery = `WebpartID eq '${wpid}'`;
                return await spService.getItemsByQuery(IKLists.Settings, '', {
                    select: ['ID', 'Title', 'ConfigValue', 'Category', 'ExtraParam'],
                    filter: filQuery,
                    top: 100
                });
            }
        } catch (e) {
            throw e;
        }
    }, [spService]);

    const setConfigSettings = useCallback(async (configName: string, configValue: any, category: SettingsCategory, wpid?: string, xtraParam?: string): Promise<IItemAddResult | IItemUpdateResult> => {
        try {
            if (configName && configName.length > 0 && configValue) {
                let exitingItems: any[] = await getConfigSettings(false, configName, undefined, wpid);
                if (exitingItems.length > 0) {
                    return await spService.updateListItem({
                        Id: exitingItems[0].Id,
                        ConfigValue: configValue,
                        WebpartID: wpid,
                        ExtraParam: xtraParam
                    }, IKLists.Settings, '')
                } else {
                    return await spService.addListItem({
                        Title: configName,
                        ConfigValue: configValue,
                        Category: category,
                        WebpartID: wpid,
                        ExtraParam: xtraParam
                    }, IKLists.Settings, '')
                }
            }
        } catch (e) {
            throw e;
        }
    }, [spService]);

    const getListFields = useCallback(async (listName: string, listid: string, weburl: string): Promise<any[]> => {
        try {
            const lstFields = await spService.getFields(listName, listid);
            return filter(lstFields, (f) => {
                return restFields.filter(rf => rf.toLowerCase() === f.StaticName.toLowerCase()).length <= 0 &&
                    restFieldTypes.filter(rf => rf.toLowerCase() === f.TypeAsString.toLowerCase()).length <= 0;
            });
        } catch (e) {
            throw e;
        }
    }, [spService]);

    return {
        getStorageValue,
        putStorageValue,
        deleteStorageValue,
        checkAndCreateOOBLists,
        returnImageFieldInfoForWrite,
        getSiteLists,
        getConfigSettings,
        setConfigSettings,
        getListFields
        // getAllSiteCollections,
        // getSitePages,
    };
};