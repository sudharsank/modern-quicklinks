import * as React from 'react';
import * as strings from 'ModernQuickLinksWebPartStrings';
import styles from '../common.module.scss';
import { HttpClient } from '@microsoft/sp-http';
import { useState, FC, useEffect, useContext } from 'react';
import { useLicenseHelper } from '../../Helpers/useLicenseHelper';
import { LoaderType, MessageScope, productName, tPropertyKey } from '../../Constants';
import { ITextFieldProps, TextField } from '@fluentui/react/lib/TextField';
import { Label } from '@fluentui/react/lib/Label';
import { addDays, Checkbox, DefaultButton, Icon, IIconStyles, IRenderFunction, IStackStyles, IStackTokens, PrimaryButton, SpinnerSize, Stack } from '@fluentui/react';
import { Text } from '@fluentui/react/lib/Text';
import { useBoolean } from '@fluentui/react-hooks';
import { delay, remove, trim } from 'lodash';
import ContentLoader from '../ContentLoader';
import MessageContainer from '../Message';
import AppContext, { IAppContextProps } from '../../AppContext';
import * as moment from 'moment';
import { IMessageInfo } from '../../IModel';
import { useCommonHelper } from '../../Helpers';

const controStackTokens: IStackTokens = { childrenGap: '10' };
const labelControlTokens: IStackTokens = { childrenGap: '5' };
const iconStyles: Partial<IIconStyles> = { root: { marginBottom: -3 } };
const actionStackStyle: IStackStyles = { root: { marginTop: '10px', borderTop: '1px solid #ccc', paddingTop: '10px' } };
const richErrorIconStyles: Partial<IIconStyles> = { root: { color: 'red' } };
const richErrorStackStyles: Partial<IStackStyles> = { root: { height: 24 } };
const richErrorStackTokens: IStackTokens = { childrenGap: 8 };

export interface ILicenseProps {
    //hideLicDialog: () => void;
    closeDialog: () => void;
    saveDialog: () => void;
    dispLicForm?: boolean;
    //tName: string;
    //httpClient: HttpClient;
    //userLoginName: string;
}

export const LicenseForm: FC<ILicenseProps> = (props) => {
    const appContext: IAppContextProps = useContext<IAppContextProps>(AppContext);
    const [TProp, setTProp] = useState<any>(undefined);
    const [dialogLoading, { setTrue: showDialogLoading, setFalse: hideDialogLoading }] = useBoolean(true);
    const [saving, { setTrue: showSaving, setFalse: hideSaving }] = useBoolean(false);
    const [saved, { setTrue: setSaved, setFalse: removeSaved }] = useBoolean(false);
    const [contactEmail, setContactEmail] = useState<string>('');
    const [lickey, setLicKey] = useState('');
    const [validEmail, { setTrue: emailValid, setFalse: emailInvalid }] = useBoolean(false);
    const [appWebUrl, setAppWebUrl] = useState<string>(undefined);
    const [message, setMessage] = useState<IMessageInfo>(undefined);
    const [trialExpired, setTrialExpired] = useState<boolean>(false);
    const [spAdmin, setSPAdmin] = useState<boolean>(false);
    const { getTenantProp, setTenantProp, getSiteLicense, decryptData, createStorageValue, encryptData, getAppCatalogWeb, setSiteLicense } = appContext.spService;
    const { checkLicenseStore, checkForValidLicDates } = useLicenseHelper(appContext.spService, appContext.httpClient);

    const onLicenseKeyLabelRenderer = (fieldProps: ITextFieldProps, defaultRender: IRenderFunction<ITextFieldProps>,) => {
        return (
            <>
                <Stack horizontal verticalAlign="center" tokens={labelControlTokens}>
                    <Icon iconName="Permissions" title="License Key" ariaLabel="Permissions" styles={iconStyles} />
                    <span>{defaultRender(fieldProps)}</span>
                </Stack>
            </>
        );
    };
    const onAppWebLabelRenderer = (fieldProps: ITextFieldProps, defaultRender: IRenderFunction<ITextFieldProps>,) => {
        return (
            <>
                <Stack horizontal verticalAlign="center" tokens={labelControlTokens}>
                    <Icon iconName="Video360Generic" title="App Catalog Url" ariaLabel="Video360Generic" styles={iconStyles} />
                    <span>{defaultRender(fieldProps)}</span>
                </Stack>
            </>
        );
    };
    const onContactEmailLabelRenderer = (
        fieldProps: ITextFieldProps,
        defaultRender: IRenderFunction<ITextFieldProps>,
    ) => {
        return (
            <>
                <Stack horizontal verticalAlign="center" tokens={labelControlTokens}>
                    <Icon iconName="Mail" title="Contact Email" ariaLabel="Mail" styles={iconStyles} />
                    <span>{defaultRender(fieldProps)}</span>
                </Stack>
            </>
        );
    };
    const getEmailErrorMessage = (value: string) => {
        let regexp = new RegExp(/^(([^<>()\[\]\\.,;:\s@"]+(\.[^<>()\[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/); // eslint-disable-line
        if (contactEmail && contactEmail.length > 0 && !regexp.test(contactEmail)) {
            emailInvalid();
            return (
                <Stack styles={richErrorStackStyles} verticalAlign="center" horizontal tokens={richErrorStackTokens}>
                    <Icon iconName="Error" styles={richErrorIconStyles} />
                    <Text variant="smallPlus">Invalid email address.</Text>
                </Stack>
            );
        }
        emailValid();
        return '';
    };
    const _onChangeEmail = (ev: any, newValue: string) => setContactEmail(newValue);
    const _onChangeLicKey = (ev: any, newValue: string) => setLicKey(newValue);
    const _onChangeAppWeb = (ev: any, newValue: string) => setAppWebUrl(newValue);
    const _onChangeSPAdmin = async (ev: any, checked: boolean) => {
        setSPAdmin(checked);
        let tc: any = await getAppCatalogWeb();
        setAppWebUrl(tc.Url);
    }

    const _checkForAppProperties = async () => {
        let appProp: any = undefined;
        let isTenant: boolean = true;
        appProp = await getTenantProp(tPropertyKey);
        if (!appProp) {
            appProp = await getSiteLicense();
            isTenant = false;
        }
        if (appProp) {
            appProp = JSON.parse(appProp);
            setTProp(appProp);
        } else hideDialogLoading();
    };

    const _checkAndSaveLicKey = async () => {
        setMessage(undefined);
        if (trim(lickey).length > 0) {
            showSaving();
            let licValid: any = undefined;
            try {
                licValid = await checkLicenseStore(lickey);
                if (licValid) {
                    licValid = decryptData(licValid);
                    licValid = JSON.parse(licValid);
                    if (licValid.customerInfo && licValid.product?.toLowerCase().indexOf(productName.toLowerCase()) >= 0) {
                        let propVal = undefined;
                        if (licValid.customerInfo?.email === 'demo' && licValid.customerInfo?.tName === 'demo') {
                            if (TProp) {
                                if (TProp?.licp === 'trial') propVal = undefined;
                                else setMessage({ msg: 'Oops, cannot downgrade the full license to trial!', scope: MessageScope.Warning });
                            } else {
                                propVal = {
                                    'sd': new Date().toISOString(),
                                    'lk': lickey,
                                    'cinfo': contactEmail,
                                    'licp': licValid.licperiod,
                                    'ed': addDays(new Date(), 10).toISOString(),
                                    'actScope': spAdmin ? 'tenant' : 'site'
                                };
                            }
                        } else if (licValid.customerInfo.tName.toLowerCase() === appContext.tName.toLowerCase()) {
                            propVal = {
                                'sd': licValid.licperiod && licValid.licperiod === 'trial' ? TProp && TProp.sd ? TProp.sd : new Date().toISOString() : new Date().toISOString(),
                                'lk': lickey,
                                'cinfo': contactEmail,
                                'licp': licValid.licperiod,
                                'ed': licValid.licperiod && licValid.licperiod === 'trial' ? TProp && TProp.ed ? new Date(TProp.ed).toISOString() : addDays(new Date(), 10).toISOString() : undefined,
                                'actScope': spAdmin ? 'tenant' : 'site'
                            };
                        } else {
                            setMessage({ msg: strings.Invalid_LicKey, scope: MessageScope.Failure });
                            removeSaved();
                            hideSaving();
                        }
                        if (propVal) {
                            let saveProp: boolean = false;
                            if (spAdmin) {
                                saveProp = await setTenantProp(tPropertyKey, JSON.stringify(propVal), appWebUrl);
                            } else {
                                saveProp = await setSiteLicense(JSON.stringify(propVal));
                            }
                            if (saveProp) {
                                setTProp(propVal)
                                createStorageValue(tPropertyKey, encryptData(JSON.stringify(propVal)), new Date(moment().add(4, 'hours').toISOString()))
                                delay(() => {
                                    if (saveProp) {
                                        _checkForAppProperties();
                                    }
                                    else setMessage({ msg: strings.App_Act_Error, scope: MessageScope.Failure });
                                    setSaved();
                                    hideSaving();
                                }, 2000);
                            } else {
                                setMessage({ msg: 'Error while saving the license key.', scope: MessageScope.Failure });
                                removeSaved();
                                hideSaving();
                            }
                        } else {
                            setMessage({ msg: 'Oops! Cannot activate the trial again.', scope: MessageScope.Failure });
                            removeSaved();
                            hideSaving();
                        }
                    }
                } else {
                    setMessage({ msg: strings.Invalid_LicKey, scope: MessageScope.Failure });
                    removeSaved();
                    hideSaving();
                }
            } catch (err) {
                setMessage({ msg: err.message, scope: MessageScope.Failure });
                removeSaved();
                hideSaving();
            }
        }
    };

    useEffect(() => {
        if (TProp) {
            if (TProp.lk) {
                setLicKey(TProp.lk);
                setContactEmail(TProp.cinfo);
                //setSPAdmin(TProp.actScope ? TProp.actScope === 'tenant' : undefined);
                let licp: string = TProp.licp;
                let ed: string = TProp.ed;
                if (licp.toLocaleLowerCase() === "trial") {
                    setTrialExpired(false);
                    if (ed && ed.length > 0) {
                        if (!checkForValidLicDates(ed)) {
                            setTrialExpired(true);
                        }
                    }
                }
            }
        }
        delay(() => hideDialogLoading(), 1000);
    }, [TProp]);

    useEffect(() => {
        _checkForAppProperties();
    }, [props]);

    return (
        <div className={styles.licenseCheck}>
            {dialogLoading &&
                <Stack verticalAlign={'center'}>
                    <ContentLoader loaderType={LoaderType.Spinner} loaderMsg={'Please wait...'} />
                </Stack>
            }
            {!dialogLoading &&
                <>
                    <Stack verticalAlign={'start'} tokens={controStackTokens}>
                        {/* <Stack.Item style={{ display: 'none' }}>
                            <TextField label="Primary Key" value={prikey} readOnly />
                        </Stack.Item> */}
                        <Stack.Item>
                            <TextField label={'Licence Key'} onRenderLabel={onLicenseKeyLabelRenderer} value={lickey}
                                maxLength={100} required onChange={_onChangeLicKey} />
                        </Stack.Item>
                        <Stack.Item>
                            <TextField label="Contact Email" onRenderLabel={onContactEmailLabelRenderer} value={contactEmail}
                                maxLength={255} onGetErrorMessage={getEmailErrorMessage} onChange={_onChangeEmail}
                                deferredValidationTime={500} required />
                        </Stack.Item>
                        <Stack.Item>
                            <Checkbox label='I am a SharePoint Administrator and I have access to the app catalog site.' onChange={_onChangeSPAdmin} checked={spAdmin} />
                        </Stack.Item>
                        {spAdmin &&
                            <Stack.Item>
                                <TextField label={'Tenanat App Catalog Site Url'} onRenderLabel={onAppWebLabelRenderer} value={appWebUrl}
                                    maxLength={255} onChange={_onChangeAppWeb} disabled />
                            </Stack.Item>
                        }
                        {TProp && TProp.sd && TProp.sd.length > 0 ? (
                            <>
                                <Stack.Item>
                                    <Stack horizontal horizontalAlign={'start'} tokens={controStackTokens}>
                                        <Stack.Item>
                                            <Label>Activated On:</Label>
                                        </Stack.Item>
                                        <Stack.Item>
                                            <Label>{new Date(TProp.sd).toLocaleString()}</Label>
                                        </Stack.Item>
                                    </Stack>
                                </Stack.Item>
                                <Stack.Item>
                                    <Stack horizontal horizontalAlign={'start'} tokens={controStackTokens}>
                                        <Stack.Item grow={2}>
                                            <Label>Activation Status:</Label>
                                        </Stack.Item>
                                        <Stack.Item>
                                            <Label style={{ color: 'green' }}>Activated {TProp?.actScope === 'site' ? '(Site)' : '(Tenant)'}</Label>
                                            <div style={{ fontStyle: 'italic' }}>
                                                <span className={styles.noteSpan}>Note:</span>
                                                If license is activated for the tenant, the web part can be used in any site collections. If activated for a site, the web part can be used only on the activated sites.
                                                Use the same license key to activate the web part on other sites.
                                            </div>
                                        </Stack.Item>
                                    </Stack>
                                </Stack.Item>
                            </>
                        ) : (
                            <Stack.Item>
                                <Stack horizontal horizontalAlign={'start'} tokens={controStackTokens}>
                                    <Stack.Item>
                                        <Label>Activation Status:</Label>
                                    </Stack.Item>
                                    <Stack.Item>
                                        <Label style={{ color: 'red' }}>Not Activated</Label>
                                    </Stack.Item>
                                </Stack>
                            </Stack.Item>
                        )}
                        {TProp && TProp.licp &&
                            <Stack.Item>
                                <Stack horizontal horizontalAlign={'start'} tokens={controStackTokens}>
                                    <Stack.Item>
                                        <Label>License:</Label>
                                    </Stack.Item>
                                    <Stack.Item>
                                        {TProp.licp.toLowerCase() === 'trial' &&
                                            <>
                                                {trialExpired ? (
                                                    <Label style={{ color: 'red' }}>Trial expired on {new Date(TProp.ed).toLocaleDateString()} </Label>
                                                ) : (
                                                    <Label style={{ color: 'red' }}>Trial {TProp.ed && TProp.ed.length > 0 ? `end by (${new Date(TProp.ed).toLocaleDateString()})` : ''} </Label>
                                                )}
                                            </>
                                        }
                                        {TProp.licp.toLowerCase() === 'full' &&
                                            <Label style={{ color: 'green' }}>Full</Label>
                                        }
                                    </Stack.Item>
                                </Stack>
                            </Stack.Item>
                        }
                        <Stack.Item>
                            <Stack horizontal horizontalAlign={'start'} tokens={controStackTokens} styles={actionStackStyle}>
                                <Stack.Item style={{ width: '100%' }}>
                                    <div>Send the below information to <b><a href="mailto:sudharsan_1985@live.in">SUDHARSAN_1985@LIVE.IN</a></b> to get full license key.</div>
                                </Stack.Item>
                            </Stack>
                        </Stack.Item>
                        <Stack.Item>
                            <Stack horizontal horizontalAlign={'start'} tokens={controStackTokens}>
                                <Stack.Item>
                                    <Label>Email:</Label>
                                </Stack.Item>
                                <Stack.Item>
                                    <Label>Contact person email address</Label>
                                </Stack.Item>
                            </Stack>
                        </Stack.Item>
                        <Stack.Item>
                            <Stack horizontal horizontalAlign={'start'} tokens={controStackTokens}>
                                <Stack.Item>
                                    <Label>Tenant Name:</Label>
                                </Stack.Item>
                                <Stack.Item>
                                    <Label>{appContext.tName ? appContext.tName : ' - '}</Label>
                                </Stack.Item>
                            </Stack>
                        </Stack.Item>
                        <Stack.Item>
                            <Stack horizontal horizontalAlign={'start'} tokens={controStackTokens}>
                                <Stack.Item style={{ width: '100%' }}>
                                    <div className={styles.adminKeyInfo}>{strings.Msg_LicSPAdminInfo}</div>
                                </Stack.Item>
                            </Stack>
                        </Stack.Item>
                    </Stack>
                    {message && message.msg &&
                        <Stack horizontal horizontalAlign={'start'} tokens={controStackTokens}>
                            <Stack.Item style={{ width: '100%' }}>
                                <MessageContainer MessageScope={message.scope} Message={message.msg} />
                            </Stack.Item>
                        </Stack>
                    }
                    <Stack horizontal horizontalAlign={'center'} tokens={controStackTokens} styles={actionStackStyle}>
                        {saving &&
                            <Stack.Item style={{ marginTop: '-10px', marginRight: '-15px' }}>
                                <ContentLoader loaderType={LoaderType.Spinner} spinSize={SpinnerSize.small} />
                            </Stack.Item>
                        }
                        <Stack.Item>
                            <PrimaryButton text='Save' iconProps={{ iconName: 'Save' }} onClick={_checkAndSaveLicKey}
                                disabled={!contactEmail || contactEmail.length <= 0 || !validEmail
                                    || lickey.length <= 0 || saving} />
                        </Stack.Item>
                        {/* {TProp && !trialExpired &&
							<Stack.Item>
								<DefaultButton text='Cancel' iconProps={{ iconName: 'ArrangeSendToBack' }} onClick={props.hideLicDialog}
									disabled={saving} />
							</Stack.Item>
						} */}
                        <Stack.Item>
                            <DefaultButton text='Close' iconProps={{ iconName: 'Blocked' }} onClick={saved ? props.saveDialog : props.closeDialog}
                                disabled={saving} />
                        </Stack.Item>
                    </Stack>
                </>
            }
        </div>
    );
};