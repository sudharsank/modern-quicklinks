import { useCallback } from "react";
//import { HttpClient, IHttpClientOptions, HttpClientResponse } from '@microsoft/sp-http';
import { IResult } from "../IModel";
import { LPItems } from '../sampleData/LaunchPadItems';
import { fld_tileColors, IKListCacheKeys, manualCondns } from "../Constants";
import * as moment from "moment";
import SPService from "./spService";
import { IField, IFieldInfo } from "@pnp/sp/fields";
import { IList, initializeFocusRects } from "@fluentui/react";
import { IListInfo } from "@pnp/sp/lists";
import { endsWith } from "lodash";

export const useLaunchPadHelper = (spService: SPService, wpInstanceId?: string) => {

    const checkForListsAvailability = useCallback(async (title: string): Promise<boolean> => {
        let retVal: boolean = false;
        try {
            const lst = await spService.getList(title, '');
            retVal = true;
        } catch (e) {
            console.log(e);
        }
        return retVal;
    }, [spService]);

    const createGlobalLaunchPadLists = useCallback(async (title: string): Promise<IResult> => {
        let retVal: IResult = undefined;
        try {
            if (title) {
                try {
                    await spService.createLaunchPadList(title, true);
                    retVal = { status: true, res: 'Created' };
                } catch (e) {
                    await spService.createLaunchPadList(title, true);
                    retVal = { status: true, res: e };
                }
            }
        } catch (e) {
            retVal = { status: false, res: e.message };
            throw e;
        }
        return retVal;
    }, [spService]);

    const createUserLaunchPadLists = useCallback(async (title: string): Promise<IResult> => {
        let retVal: IResult = undefined;
        try {
            if (title) {
                try {
                    await spService.createLaunchPadList(title, false);
                    retVal = { status: true, res: 'Created' };
                } catch (e) {
                    retVal = { status: true, res: e };
                }
            }
        } catch (e) {
            retVal = { status: false, res: e.message };
            throw e;
        }
        return retVal;
    }, [spService]);

    const createSampleItemsForGlobalLinks = useCallback(async (title: string): Promise<void> => {
        try {
            let items: any[] = [];
            LPItems.forEach(item => {
                items.push(item);
            });
            await spService.addMultiple(items, title);
        } catch (e) {
            throw e;
        }
    }, [spService]);

    const _checkForValidSelection = (listFields: IFieldInfo[], condn: string, fieldName: string): boolean => {
        try {
            if (condn && fieldName && listFields.length > 0) {
                const fieldInfo: IFieldInfo[] = listFields.filter(f => f.InternalName.toLowerCase() === fieldName.toLowerCase());
                if (fieldInfo.length > 0) {
                    switch (fieldInfo[0].TypeAsString.toLowerCase()) {
                        case 'text':
                        case 'note':
                        case 'choice':
                        case 'url':
                            if (condn.toLowerCase() === 'eq' || condn.toLowerCase() === 'ne' || condn.toLowerCase() === 'contains'
                                || condn.toLowerCase() === 'startswith' || condn.toLowerCase() === 'endswith') return true;
                            else return false;
                        case 'number':
                        case 'currency':
                        case 'counter':
                            if (condn.toLowerCase() === 'eq' || condn.toLowerCase() === 'ne' || condn.toLowerCase() === 'gt' || condn.toLowerCase() === 'ge'
                                || condn.toLowerCase() === 'lt' || condn.toLowerCase() === 'le') return true;
                            else return false;
                        case 'boolean':
                        case 'bool':
                            if (condn.toLowerCase() === 'eq' || condn.toLowerCase() === 'ne') return true;
                            else return false;
                    }
                }
            }
        } catch (err) {
            console.error(err);
        }
        return false;
    };

    const getFilters = (listFields: any[], listfilters: any[]): string => {
        console.log('getFilters: ', listFields, listfilters);
        let retFilters: string = '';
        listfilters.forEach((filter) => {
            if (_checkForValidSelection(listFields, filter.operator, filter.fieldname)) {
                const filfield = listFields.filter(f => f.TypeAsString?.toLowerCase() === 'boolean' && f.InternalName?.toLowerCase() === filter.fieldname?.toLowerCase());
                if (filfield.length > 0) {
                    if (filter.value.toLowerCase() === 'true' || filter.value.toLowerCase() === 'yes' || filter.value == '1' || filter.value.toLowerCase() === '1') {
                        retFilters += `${filter.fieldname} ${filter.operator} 1 ${filter.andor ? filter.andor : ''} `;
                    } else retFilters += `${filter.fieldname} ${filter.operator} 0 ${filter.andor ? filter.andor : ''} `;
                } else {
                    if (manualCondns.indexOf(filter.operator) < 0)
                        retFilters += `${filter.fieldname} ${filter.operator} '${filter.value}' ${filter.andor ? filter.andor : ''} `;
                }
            }
        });
        // console.log('Filters: ', retFilters);
        return retFilters;
    };

    const getFinalItemsWithManualFilters = (items: any[], listFields: IFieldInfo[], listfilters: any[]): any[] => {
        // console.log('Manual Filters: ', items);
        let retItems: any[] = items;
        try {
            if (listfilters && listfilters.length > 0) {
                let filterCondn: string = undefined;
                listfilters.forEach((filter) => {
                    if (_checkForValidSelection(listFields, filter.operator, filter.fieldname)) {
                        const filfield = listFields.filter(f => f.TypeAsString?.toLowerCase() === 'boolean' && f.InternalName?.toLowerCase() === filter.fieldname?.toLowerCase());
                        if (manualCondns.indexOf(filter.operator) >= 0 && filfield.length <= 0) {
                            // console.log("Manual Filters: ", filter);
                            switch (filter.operator.toLowerCase()) {
                                case 'eq':
                                    if (filterCondn && filterCondn.toLowerCase() === 'or') {
                                        const filItems = items.filter((item) => {
                                            return item[filter.fieldname]?.toLowerCase() === filter.value.toLowerCase();
                                        });
                                        retItems = retItems.concat(filItems);
                                    }
                                    else {
                                        retItems = retItems.filter((item) => {
                                            return item[filter.fieldname]?.toLowerCase() === filter.value.toLowerCase();
                                        });
                                    }
                                    break;
                                case 'ne':
                                    if (filterCondn && filterCondn.toLowerCase() === 'or') {
                                        const filItems = items.filter((item) => {
                                            return item[filter.fieldname]?.toLowerCase() !== filter.value.toLowerCase();
                                        });
                                        retItems = retItems.concat(filItems);
                                    }
                                    else {
                                        retItems = retItems.filter((item) => {
                                            return item[filter.fieldname]?.toLowerCase() !== filter.value.toLowerCase();
                                        });
                                    }
                                    break;
                                case 'contains':
                                    if (filterCondn && filterCondn.toLowerCase() === 'or') {
                                        const filItems = items.filter((item) => {
                                            return item[filter.fieldname]?.toLowerCase().indexOf(filter.value.toLowerCase()) >= 0;
                                        });
                                        retItems = retItems.concat(filItems);
                                    }
                                    else {
                                        // console.log('Manual Filters - contains: ', retItems);                                    
                                        retItems = retItems.filter((item) => {
                                            return item[filter.fieldname]?.toLowerCase().indexOf(filter.value.toLowerCase()) >= 0;
                                        });
                                    }
                                    break;
                                case 'startswith':
                                    // console.log('Manual Filters - startswith: ', filter, filterCondn);
                                    if (filterCondn && filterCondn.toLowerCase() === 'or') {
                                        const filItems = items.filter((item) => {
                                            return item[filter.fieldname]?.toLowerCase().startsWith(filter.value.toLowerCase());
                                        });
                                        // console.log('Manual Filters - or - startswith: ', filItems);
                                        retItems = retItems.concat(filItems);
                                    }
                                    else {
                                        retItems = retItems.filter((item) => {
                                            return item[filter.fieldname]?.toLowerCase().startsWith(filter.value.toLowerCase());
                                        });
                                    }
                                    break;
                                case 'endswith':
                                    if (filterCondn && filterCondn.toLowerCase() === 'or') {
                                        const filItems = items.filter((item) => {
                                            return endsWith(item[filter.fieldname]?.toLowerCase(), filter.value.toLowerCase());
                                        });
                                        retItems = retItems.concat(filItems);
                                    }
                                    else {
                                        retItems = retItems.filter((item) => {
                                            return endsWith(item[filter.fieldname]?.toLowerCase(), filter.value.toLowerCase());
                                        });
                                    }
                                    break;
                            }
                            // console.log('Manual Filters after filtering: ', retItems);
                        }
                    }
                    filterCondn = filter.andOr;
                });
            }
        } catch (err) {
            console.error(err);
        }
        // remove duplicates
        retItems = retItems.filter((item, index, self) => self.findIndex((t) => t.ID === item.ID) === index);
        // console.log('Manual Filters after removing duplicates: ', retItems);
        return retItems;
    };

    const checkAndUpdateSelectFields = (selectFields: string[], listFilters: any[]): string[] => {
        let retFields: string[] = selectFields;
        try {
            listFilters.forEach((field) => {
                //if (field.InternalName === 'ID' || field.InternalName === 'Author' || field.InternalName === 'Editor') return;
                const filterFlds: string[] = selectFields.filter(f => f.toLowerCase() === field.fieldname?.toLowerCase());
                if (filterFlds.length === 0) retFields.push(field.fieldname);
            });
        } catch (err) {
            console.error(err);
        }
        return retFields;
    };

    const getGlobalLinks = useCallback(async (listTitle: string, listid: string, orderBy: string, isAsc: boolean, updateCache: boolean, isInfo: boolean, enableKeywordSearch: boolean,
        keywordField: string, useTileColors: boolean, listFields: IFieldInfo[], useGLFilter: boolean, useCache?: boolean, listFilters?: any[]): Promise<any[]> => {
        let retRes: any[] = [];
        let selectFields: string[] = isInfo ? ['ID', 'Title', 'URL', 'Sequence', 'NewWindow', 'IconName', 'Created', 'Modified', 'IconImage', 'Description','FieldValuesAsText/Description'] :
            ['ID', 'Title', 'URL', 'Sequence', 'NewWindow', 'IconName', 'Created', 'Modified', 'IconImage', 'FieldValuesAsText/Description'];
        if (enableKeywordSearch && keywordField) selectFields.push(keywordField);
        if (useTileColors) selectFields.push(fld_tileColors);
        selectFields = checkAndUpdateSelectFields(selectFields, listFilters);
        let storageKey: string = `${IKListCacheKeys.LPGlobal}.${wpInstanceId}`;
        try {
            if (useCache && !updateCache) {
                retRes = spService.getStorageValue(storageKey, false);
            }
            if (!retRes || retRes?.length === 0) {
                retRes = [];
                let lst: IListInfo = undefined;
                if (listTitle) lst = await spService.getList(listTitle, '');
                else if (listid) lst = await spService.getList('', listid);
                if (lst) {
                    // console.log('getGlobalLinks: ', listFields, listFilters);
                    let allItems: any[] = await spService.getItemsByQuery(lst.Title, undefined, {
                        select: selectFields,
                        expand: ['FieldValuesAsText'],
                        orderby: { orderByCol: orderBy, isAsc }
                    });
                    //console.log('All Items: ', allItems);
                    if (useGLFilter) {
                        let items: any[] = [];
                        let filterCondn: string = undefined;
                        if (listFilters && listFilters.length > 0) {
                            listFilters.forEach((filter) => {
                                const filField = listFields.filter(f => f.InternalName.toLowerCase() === filter.fieldname.toLowerCase());
                                switch (filter.operator.toLowerCase()) {
                                    case 'eq':
                                        switch (filField[0].TypeAsString.toLowerCase()) {
                                            case 'text':
                                            case 'note':
                                            case 'choice':
                                                if (!filterCondn || filterCondn.toLowerCase() === 'or') {
                                                    const filItems = allItems.filter((item) => {
                                                        return item[filter.fieldname]?.toLowerCase() === filter.value.toLowerCase();
                                                    });
                                                    // console.log('Text Filter - eq - Or: ', filItems);
                                                    if (filItems && filItems.length > 0) retRes = retRes.concat(filItems);
                                                } else {
                                                    retRes = retRes.filter((item) => {
                                                        return item[filter.fieldname]?.toLowerCase() === filter.value.toLowerCase();
                                                    });
                                                }
                                                break;
                                            case 'boolean': {
                                                let finalValue = false;
                                                if (filter.value.toLowerCase() === 'true' || filter.value.toLowerCase() === 'yes' || filter.value == '1' || filter.value.toLowerCase() === '1') finalValue = true;
                                                else finalValue = false;
                                                if (!filterCondn || filterCondn.toLowerCase() === 'or') {
                                                    const filItems = allItems.filter((item) => {
                                                        return item[filter.fieldname] === finalValue;
                                                    });
                                                    // console.log('Boolean Filter - eq - or: ', filItems);
                                                    if (filItems && filItems.length > 0) retRes = retRes.concat(filItems);
                                                } else {
                                                    // console.log('Boolean Filter - eq - and: ', retRes);
                                                    retRes = retRes.filter((item) => {
                                                        return item[filter.fieldname] === finalValue;
                                                    });
                                                }
                                            }
                                                break;
                                            case 'url':
                                                if (!filterCondn || filterCondn.toLowerCase() === 'or') {
                                                    const filItems = allItems.filter((item) => {
                                                        if (item[filter.fieldname]?.Url?.toLowerCase() === filter.value.toLowerCase() ||
                                                            item[filter.fieldname]?.Description?.toLowerCase() === filter.value.toLowerCase()) return item;
                                                    });
                                                    // console.log('Url Filter - eq - or: ', filItems);
                                                    if (filItems && filItems.length > 0) retRes = retRes.concat(filItems);
                                                } else {
                                                    // console.log('Url Filter - eq - and: ', retRes);
                                                    retRes = retRes.filter((item) => {
                                                        if (item[filter.fieldname]?.Url?.toLowerCase() === filter.value.toLowerCase() ||
                                                            item[filter.fieldname]?.Description?.toLowerCase() === filter.value.toLowerCase()) return item;
                                                    });
                                                }
                                                break;
                                            case 'multichoice':
                                                if (!filterCondn || filterCondn.toLowerCase() === 'or') {
                                                    const filItems = allItems.filter((item) => {
                                                        if (item[filter.fieldname] && item[filter.fieldname].length > 0) {
                                                            return item[filter.fieldname].indexOf(filter.value) >= 0;
                                                        }
                                                    });
                                                    // console.log('MultiChoice Filter - eq - or: ', filItems);
                                                    if (filItems && filItems.length > 0) retRes = retRes.concat(filItems);
                                                } else {
                                                    // console.log('MultiChoice Filter - eq - and: ', retRes);
                                                    retRes = retRes.filter((item) => {
                                                        if (item[filter.fieldname] && item[filter.fieldname].length > 0) {
                                                            return item[filter.fieldname].indexOf(filter.value) >= 0;
                                                        }
                                                    });
                                                }
                                                break;
                                            case 'number':
                                            case 'currency':
                                            case 'counter':
                                                if (!filterCondn || filterCondn.toLowerCase() === 'or') {
                                                    const filItems = allItems.filter((item) => {
                                                        return item[filter.fieldname] == filter.value;
                                                    });
                                                    // console.log('Number Filter - gt - or: ', filItems);
                                                    if (filItems && filItems.length > 0) retRes = retRes.concat(filItems);
                                                } else {
                                                    // console.log('Number Filter - gt - and: ', retRes);
                                                    retRes = retRes.filter((item) => {
                                                        return item[filter.fieldname] == filter.value;
                                                    });
                                                }
                                                break;
                                            case 'datetime':
                                                break;
                                            case 'lookup':
                                                break;
                                            case 'user':
                                                break;
                                        }
                                        break;
                                    case 'ne':
                                        break;
                                    case 'contains':
                                        switch (filField[0].TypeAsString.toLowerCase()) {
                                            case 'text':
                                            case 'note':
                                            case 'choice':
                                                if (!filterCondn || filterCondn.toLowerCase() === 'or') {
                                                    const filItems = allItems.filter((item) => {
                                                        return item[filter.fieldname]?.toLowerCase().indexOf(filter.value.toLowerCase()) >= 0;
                                                    });
                                                    // console.log('Text Filter - contains - or: ', filItems);
                                                    if (filItems && filItems.length > 0) retRes = retRes.concat(filItems);
                                                } else {
                                                    // console.log('Text Filter - contains - and: ', retRes);
                                                    retRes = retRes.filter((item) => {
                                                        return item[filter.fieldname]?.toLowerCase().indexOf(filter.value.toLowerCase()) >= 0;
                                                    });
                                                }
                                                break;
                                            case 'url':
                                                if (!filterCondn || filterCondn.toLowerCase() === 'or') {
                                                    const filItems = allItems.filter((item) => {
                                                        if (item[filter.fieldname]?.Url?.toLowerCase().indexOf(filter.value.toLowerCase()) >= 0 ||
                                                            item[filter.fieldname]?.Description?.toLowerCase().indexOf(filter.value.toLowerCase()) >= 0) return item;
                                                    });
                                                    // console.log('Url Filter - eq - or: ', filItems);
                                                    if (filItems && filItems.length > 0) retRes = retRes.concat(filItems);
                                                } else {
                                                    // console.log('Url Filter - eq - and: ', retRes);
                                                    retRes = retRes.filter((item) => {
                                                        if (item[filter.fieldname]?.Url?.toLowerCase().indexOf(filter.value.toLowerCase()) >= 0 ||
                                                            item[filter.fieldname]?.Description?.toLowerCase().indexOf(filter.value.toLowerCase()) >= 0) return item;
                                                    });
                                                }
                                                break;
                                            case 'multichoice':
                                                if (!filterCondn || filterCondn.toLowerCase() === 'or') {
                                                    const filItems = allItems.filter((item) => {
                                                        if (item[filter.fieldname] && item[filter.fieldname].length > 0) {
                                                            return item[filter.fieldname].indexOf(filter.value) >= 0;
                                                        }
                                                    });
                                                    // console.log('MultiChoice Filter - eq - or: ', filItems);
                                                    if (filItems && filItems.length > 0) retRes = retRes.concat(filItems);
                                                } else {
                                                    // console.log('MultiChoice Filter - eq - and: ', retRes);
                                                    retRes = retRes.filter((item) => {
                                                        if (item[filter.fieldname] && item[filter.fieldname].length > 0) {
                                                            return item[filter.fieldname].indexOf(filter.value) >= 0;
                                                        }
                                                    });
                                                }
                                                break;
                                        }
                                        break;
                                    case 'startswith':
                                        switch (filField[0].TypeAsString.toLowerCase()) {
                                            case 'text':
                                            case 'note':
                                            case 'choice':
                                                if (!filterCondn || filterCondn.toLowerCase() === 'or') {
                                                    const filItems = allItems.filter((item) => {
                                                        return item[filter.fieldname]?.toLowerCase().startsWith(filter.value.toLowerCase());
                                                    });
                                                    // console.log('Text Filter: ', filItems);
                                                    if (filItems && filItems.length > 0) retRes = retRes.concat(filItems);
                                                } else {
                                                    retRes = retRes.filter((item) => {
                                                        return item[filter.fieldname]?.toLowerCase().startsWith(filter.value.toLowerCase());
                                                    });
                                                }
                                                break;
                                            case 'url':
                                                if (!filterCondn || filterCondn.toLowerCase() === 'or') {
                                                    const filItems = allItems.filter((item) => {
                                                        if (item[filter.fieldname]?.Url?.toLowerCase().startsWith(filter.value.toLowerCase()) ||
                                                            item[filter.fieldname]?.Description?.toLowerCase().startsWith(filter.value.toLowerCase())) return item;
                                                    });
                                                    // console.log('Url Filter - eq - or: ', filItems);
                                                    if (filItems && filItems.length > 0) retRes = retRes.concat(filItems);
                                                } else {
                                                    // console.log('Url Filter - eq - and: ', retRes);
                                                    retRes = retRes.filter((item) => {
                                                        if (item[filter.fieldname]?.Url?.toLowerCase().startsWith(filter.value.toLowerCase()) ||
                                                            item[filter.fieldname]?.Description?.toLowerCase().startsWith(filter.value.toLowerCase())) return item;
                                                    });
                                                }
                                                break;
                                            case 'multichoice':
                                                if (!filterCondn || filterCondn.toLowerCase() === 'or') {
                                                    const filItems = allItems.filter((item) => {
                                                        if (item[filter.fieldname] && item[filter.fieldname].length > 0) {
                                                            item[filter.fieldname].forEach((val: string) => {
                                                                if (val?.toLowerCase().startsWith(filter.value.toLowerCase())) return item;
                                                            });
                                                        }
                                                    });
                                                    // console.log('MultiChoice Filter - eq - or: ', filItems);
                                                    if (filItems && filItems.length > 0) retRes = retRes.concat(filItems);
                                                } else {
                                                    // console.log('MultiChoice Filter - eq - and: ', retRes);
                                                    retRes = retRes.filter((item) => {
                                                        if (item[filter.fieldname] && item[filter.fieldname].length > 0) {
                                                            item[filter.fieldname].forEach((val: string) => {
                                                                if (val?.toLowerCase().startsWith(filter.value.toLowerCase())) return item;
                                                            });
                                                        }
                                                    });
                                                }
                                                break;
                                        }
                                        break;
                                    case 'endswith':
                                        switch (filField[0].TypeAsString.toLowerCase()) {
                                            case 'text':
                                            case 'note':
                                            case 'choice':
                                                if (!filterCondn || filterCondn.toLowerCase() === 'or') {
                                                    const filItems = allItems.filter((item) => {
                                                        return item[filter.fieldname]?.toLowerCase().endsWith(filter.value.toLowerCase());
                                                    });
                                                    // console.log('Text Filter: ', filItems);
                                                    if (filItems && filItems.length > 0) retRes = retRes.concat(filItems);
                                                } else {
                                                    retRes = retRes.filter((item) => {
                                                        return item[filter.fieldname]?.toLowerCase().endsWith(filter.value.toLowerCase());
                                                    });
                                                }
                                                break;
                                            case 'url':
                                                if (!filterCondn || filterCondn.toLowerCase() === 'or') {
                                                    const filItems = allItems.filter((item) => {
                                                        if (item[filter.fieldname]?.Url?.toLowerCase().endsWith(filter.value.toLowerCase()) ||
                                                            item[filter.fieldname]?.Description?.toLowerCase().endsWith(filter.value.toLowerCase())) return item;
                                                    });
                                                    // console.log('Url Filter - eq - or: ', filItems);
                                                    if (filItems && filItems.length > 0) retRes = retRes.concat(filItems);
                                                } else {
                                                    // console.log('Url Filter - eq - and: ', retRes);
                                                    retRes = retRes.filter((item) => {
                                                        if (item[filter.fieldname]?.Url?.toLowerCase().endsWith(filter.value.toLowerCase()) ||
                                                            item[filter.fieldname]?.Description?.toLowerCase().endsWith(filter.value.toLowerCase())) return item;
                                                    });
                                                }
                                                break;
                                            case 'multichoice':
                                                if (!filterCondn || filterCondn.toLowerCase() === 'or') {
                                                    const filItems = allItems.filter((item) => {
                                                        if (item[filter.fieldname] && item[filter.fieldname].length > 0) {
                                                            item[filter.fieldname].forEach((val: string) => {
                                                                if (val?.toLowerCase().endsWith(filter.value.toLowerCase())) return item;
                                                            });
                                                        }
                                                    });
                                                    // console.log('MultiChoice Filter - eq - or: ', filItems);
                                                    if (filItems && filItems.length > 0) retRes = retRes.concat(filItems);
                                                } else {
                                                    // console.log('MultiChoice Filter - eq - and: ', retRes);
                                                    retRes = retRes.filter((item) => {
                                                        if (item[filter.fieldname] && item[filter.fieldname].length > 0) {
                                                            item[filter.fieldname].forEach((val: string) => {
                                                                if (val?.toLowerCase().endsWith(filter.value.toLowerCase())) return item;
                                                            });
                                                        }
                                                    });
                                                }
                                                break;
                                        }
                                        break;
                                    case 'gt':
                                        switch (filField[0].TypeAsString.toLowerCase()) {
                                            case 'number':
                                            case 'currency':
                                            case 'counter':
                                                if (!filterCondn || filterCondn.toLowerCase() === 'or') {
                                                    const filItems = allItems.filter((item) => {
                                                        return item[filter.fieldname] > filter.value;
                                                    });
                                                    // console.log('Number Filter - gt - or: ', filItems);
                                                    if (filItems && filItems.length > 0) retRes = retRes.concat(filItems);
                                                } else {
                                                    // console.log('Number Filter - gt - and: ', retRes);
                                                    retRes = retRes.filter((item) => {
                                                        return item[filter.fieldname] > filter.value;
                                                    });
                                                }
                                                break;
                                        }
                                        break;
                                    case 'ge':
                                        switch (filField[0].TypeAsString.toLowerCase()) {
                                            case 'number':
                                            case 'currency':
                                            case 'counter':
                                                if (!filterCondn || filterCondn.toLowerCase() === 'or') {
                                                    const filItems = allItems.filter((item) => {
                                                        return item[filter.fieldname] >= filter.value;
                                                    });
                                                    // console.log('Number Filter - ge - or: ', filItems);
                                                    if (filItems && filItems.length > 0) retRes = retRes.concat(filItems);
                                                } else {
                                                    // console.log('Number Filter - ge - and: ', retRes);
                                                    retRes = retRes.filter((item) => {
                                                        return item[filter.fieldname] >= filter.value;
                                                    });
                                                }
                                                break;
                                        }
                                        break;
                                    case 'lt':
                                        switch (filField[0].TypeAsString.toLowerCase()) {
                                            case 'number':
                                            case 'currency':
                                            case 'counter':
                                                if (!filterCondn || filterCondn.toLowerCase() === 'or') {
                                                    const filItems = allItems.filter((item) => {
                                                        return item[filter.fieldname] < filter.value;
                                                    });
                                                    // console.log('Number Filter - lt - or: ', filItems);
                                                    if (filItems && filItems.length > 0) retRes = retRes.concat(filItems);
                                                } else {
                                                    // console.log('Number Filter - lt - and: ', retRes);
                                                    retRes = retRes.filter((item) => {
                                                        return item[filter.fieldname] < filter.value;
                                                    });
                                                }
                                                break;
                                        }
                                        break;
                                    case 'le':
                                        switch (filField[0].TypeAsString.toLowerCase()) {
                                            case 'number':
                                            case 'currency':
                                            case 'counter':
                                                if (!filterCondn || filterCondn.toLowerCase() === 'or') {
                                                    const filItems = allItems.filter((item) => {
                                                        return item[filter.fieldname] <= filter.value;
                                                    });
                                                    // console.log('Number Filter - le - or: ', filItems);
                                                    if (filItems && filItems.length > 0) retRes = retRes.concat(filItems);
                                                } else {
                                                    // console.log('Number Filter - le - and: ', retRes);
                                                    retRes = retRes.filter((item) => {
                                                        return item[filter.fieldname] <= filter.value;
                                                    });
                                                }
                                                break;
                                        }
                                        break;
                                    case 'in':
                                        break;
                                    case 'notin':
                                        break;
                                    case 'between':
                                        break;
                                    case 'notbetween':
                                        break;
                                    case 'isnull':
                                        break;
                                    case 'isnotnull':
                                        break;
                                }
                                filterCondn = filter.andOr;
                            });
                        }
                    } else {
                        retRes = await spService.getItemsByQuery(lst.Title, undefined, {
                            select: selectFields,
                            expand: ['FieldValuesAsText'],
                            filter: `IsActive eq 1`,
                            orderby: { orderByCol: orderBy, isAsc }
                        });
                    }
                }
                if (retRes && useCache)
                    spService.createStorageValue(storageKey, retRes, new Date(moment().add('minute', 15).toISOString()), false);
                if (!useCache) spService.deleteStorageValue(storageKey);
            }
        } catch (e) {
            throw e;
        }
        return retRes;
    }, [spService, wpInstanceId]);

    const getGlobalLinks1 = useCallback(async (listTitle: string, listid: string, orderBy: string, isAsc: boolean, updateCache: boolean, isInfo: boolean, enableKeywordSearch: boolean,
        keywordField: string, useTileColors: boolean, useCache?: boolean): Promise<any[]> => {
        let retRes: any[] = undefined;
        let selectFields: string[] = isInfo ? ['ID', 'Title', 'URL', 'Sequence', 'NewWindow', 'IconName', 'Created', 'Modified', 'IconImage', 'Description'] :
            ['ID', 'Title', 'URL', 'Sequence', 'NewWindow', 'IconName', 'Created', 'Modified', 'IconImage'];
        if (enableKeywordSearch && keywordField) selectFields.push(keywordField);
        if (useTileColors) selectFields.push(fld_tileColors);
        let storageKey: string = `${IKListCacheKeys.LPGlobal}.${wpInstanceId}`;
        try {
            if (useCache && !updateCache) {
                retRes = spService.getStorageValue(storageKey, false);
            }
            if (!retRes) {
                if (listTitle)
                    retRes = await spService.getItemsByQuery(listTitle, undefined, {
                        select: selectFields,
                        filter: `IsActive eq 1`,
                        orderby: { orderByCol: orderBy, isAsc }
                    });
                else if (listid)
                    retRes = await spService.getItemsByQuery(undefined, listid, {
                        select: selectFields,
                        filter: `IsActive eq 1`,
                        orderby: { orderByCol: orderBy, isAsc }
                    });
                if (retRes && useCache)
                    spService.createStorageValue(storageKey, retRes, new Date(moment().add('minute', 5).toISOString()), false);
            }
        } catch (e) {
            throw e;
        }
        return retRes;
    }, [spService, wpInstanceId]);

    const getUserLinks = useCallback(async (listTitle: string, listid: string, orderBy: string, isAsc: boolean, updateCache: boolean, useCache?: boolean): Promise<any[]> => {
        let retRes: any[] = undefined;
        let storageKey: string = `${IKListCacheKeys.LPUsers}.${wpInstanceId}`;
        try {
            if (useCache && !updateCache) {
                retRes = spService.getStorageValue(storageKey, false);
            }
            if (!retRes) {
                if (listTitle)
                    retRes = await spService.getItemsByQuery(listTitle, undefined, {
                        select: ['ID', 'Title', 'URL', 'Sequence', 'NewWindow', 'IconName', 'Created', 'Modified', 'IconImage'],
                        filter: `IsActive eq 1 `,
                        orderby: { orderByCol: orderBy, isAsc }
                    });
                else if (listid)
                    retRes = await spService.getItemsByQuery(undefined, listid, {
                        select: ['ID', 'Title', 'URL', 'Sequence', 'NewWindow', 'IconName', 'Created', 'Modified', 'IconImage'],
                        filter: `IsActive eq 1`,
                        orderby: { orderByCol: orderBy, isAsc }
                    });
                if (retRes && useCache)
                    spService.createStorageValue(storageKey, retRes, new Date(moment().add('minute', 5).toISOString()), false);
            }
        } catch (e) {
            throw e;
        }
        return retRes;
    }, [spService, wpInstanceId]);

    const manageTiles = useCallback(async (tileInfo: any, listTitle: string, listid: string): Promise<void> => {
        try {
            if (tileInfo) {
                if (tileInfo.Id) {
                    if (listTitle) await spService.updateListItem(tileInfo, listTitle, '');
                    else if (listid) await spService.updateListItem(tileInfo, '', listid);
                } else {
                    if (listTitle) await spService.addListItem(tileInfo, listTitle, '');
                    else if (listid) await spService.addListItem(tileInfo, '', listid);
                }
            }
        } catch (e) {
            throw e;
        }
    }, [spService]);

    const getTilesCount = useCallback(async (listTitle: string, listid: string): Promise<number> => {
        try {
            const itemCount: number = await spService.getItemCount(listTitle, listid);
            return itemCount;
        } catch (e) {
            throw e;
        }
    }, [spService]);

    const delUserTile = useCallback(async (listTitle: string, listid: string, itemid: string): Promise<boolean> => {
        try {
            return await spService.deleteListItem(listTitle, listid, itemid);
        } catch (e) {
            throw e;
        }
    }, [spService]);

    const checkForField = useCallback(async (listTitle: string, listid: string, field: string): Promise<boolean> => {
        try {
            return await spService.checkFieldExists(listTitle, listid, field);
        } catch (e) {
            console.log(e);
        }
    }, [spService]);

    const createField = useCallback(async (listTitle: string, listid: string, field: string, fieldType: string, isRequired: boolean, isHidden: boolean): Promise<void> => {
        try {
            await spService.createField(listTitle, listid, field, fieldType, isRequired, isHidden);
        } catch (e) {
            console.log(e);
        }
    }, [spService]);

    const checkAndCreateField = useCallback(async (listTitle: string, listid: string, field: string, fieldType: string, isRequired: boolean, isHidden: boolean): Promise<void> => {
        try {
            const isFieldExists = await checkForField(listTitle, listid, field);
            if (!isFieldExists) {
                await createField(listTitle, listid, field, fieldType, isRequired, isHidden);
            }
        } catch (e) {
            console.log(e);
        }
    }, [checkForField, createField]);

    const getItemById = useCallback(async (listTitle: string, listid: string, itemId: string): Promise<any> => {
        try {
            return await spService.getItemById(listTitle, listid, itemId);
        } catch (e) {
            throw e;
        }
    }, [spService]);

    return {
        checkForListsAvailability,
        createGlobalLaunchPadLists,
        createUserLaunchPadLists,
        createSampleItemsForGlobalLinks,
        getGlobalLinks,
        getUserLinks,
        manageTiles,
        getTilesCount,
        delUserTile,
        checkAndCreateField,
        checkForField,
        createField,
        getItemById
    };
};