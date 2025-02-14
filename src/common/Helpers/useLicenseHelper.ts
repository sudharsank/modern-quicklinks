import { useCallback } from "react";
import { HttpClient, IHttpClientOptions, HttpClientResponse } from '@microsoft/sp-http';
import { apiUrl, LicenseMessage, tPropertyKey } from "../Constants";
import * as moment from "moment";
import SPService from "./spService";


export const useLicenseHelper = (spService: SPService, httpClient?: HttpClient) => {

    const checkForValidLicDates = (ed: string): any => {
        console.log("Trial days remaining: ", moment(new Date(ed).toISOString()).diff(moment(new Date().toISOString()), 'days'));
        return moment(new Date(ed).toISOString()) >= moment(new Date().toISOString());
    };

    const getLicenseInfo = (): any => {
        try {
            let tprop = spService.getStorageValue(tPropertyKey);
            if(tprop) {
                tprop = spService.decryptData(tprop);
                tprop = JSON.parse(tprop.toString());
                //if(tprop?.ed) return moment(new Date(tprop.ed).toISOString()).diff(moment(new Date().toISOString()), 'days');
                return tprop;
            }
        } catch (err) {
            console.error(err);
        }
    };

    const checkLicenseStore = useCallback(async (lickey: string, priKey?: string): Promise<any> => {
        if (httpClient) {
            const requestHeaders: Headers = new Headers();
            requestHeaders.append("Content-type", "application/json");
            requestHeaders.append("Cache-Control", "no-cache");
            const postOptions: IHttpClientOptions = {
                headers: requestHeaders,
                body: `{"lickey":"${lickey}"}`
            };
            try {
                const response: HttpClientResponse = await httpClient.post(`${apiUrl}/checkForLicense`, HttpClient.configurations.v1, postOptions);
                if (response.ok) {
                    try {
                        return await response.text();
                    } catch (err) {
                        return undefined;
                    }
                }
            } catch (err) {
                console.log(err);
                throw new Error(err);
            }
        } else return undefined;
    }, [httpClient]);

    const checkForLicenseKey = useCallback(async (): Promise<LicenseMessage> => {
        let tprop: any = undefined;
        try {
            // Check for license details in the local browser storage.
            tprop = spService.getStorageValue(tPropertyKey);
            if (!tprop || tprop === null) {
                // Check for tenant license property
                let licProps: any = undefined;
                licProps = await spService.getTenantProp(tPropertyKey);
                if(!licProps) {
                    licProps = await spService.getSiteLicense(); 
                }
                const encryptedProp = spService.encryptData(licProps);
                if (encryptedProp) {
                    spService.createStorageValue(tPropertyKey, encryptedProp, new Date(moment().add(4, 'hours').toISOString()));
                }
                tprop = spService.decryptData(encryptedProp);
            } else {
                tprop = spService.decryptData(tprop);
            }
            if (tprop) {
                tprop = JSON.parse(tprop.toString());
                if (tprop.lk) {
                    const licp: string = tprop.licp;
                    const ed: string = tprop.ed;
                    if (licp.toLocaleLowerCase() === "trial") {
                        if (ed && ed.length > 0) {
                            if (!checkForValidLicDates(ed)) {
                                // Trial expired
                                return LicenseMessage.Expired;
                            } else {
                                return LicenseMessage.Valid;
                            }
                        } else {
                            return LicenseMessage.NotConfigured;
                        }
                    } else if (licp.toLocaleLowerCase() === "full") {
                        return LicenseMessage.Valid;
                    } else {
                        return LicenseMessage.ConfigError;
                    }
                } else {
                    return LicenseMessage.ConfigError;
                }
            } else {
                return LicenseMessage.NotConfigured;
            }
        } catch (err: any) {
            console.log(err);
            throw err;
        }
    }, [spService]);

    return {
        checkLicenseStore,
        checkForLicenseKey,
        checkForValidLicDates,
        getLicenseInfo
    };
};