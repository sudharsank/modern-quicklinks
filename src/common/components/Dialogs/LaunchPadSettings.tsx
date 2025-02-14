import * as React from 'react';
import * as strings from 'ModernQuickLinksWebPartStrings';
import styles from '../common.module.scss';
import { FC, useEffect, useState, useContext } from 'react';
import MessageContainer from '../Message';
import ContentLoader from '../ContentLoader';
import { IStackStyles, Stack } from 'office-ui-fabric-react/lib/Stack';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { css } from 'office-ui-fabric-react/lib/Utilities';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import { useBoolean } from '@fluentui/react-hooks';
import { DefaultButton, PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { LoaderType, LogSource, MessageScope } from '../../Constants';
import { SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { IMessageInfo, IResult } from '../../IModel';
import AppContext, { IAppContextProps } from '../../AppContext';
import { useLaunchPadHelper } from '../../Helpers/useLaunchPadHelper';

const stackItemStyles: Partial<IStackStyles> = { root: { width: '100%' } };

export interface ILaunchPadSettingsProps {
    closeDialog: () => void;
}

export const LaunchPadSettings: FC<ILaunchPadSettingsProps> = (props) => {
    const appContext: IAppContextProps = useContext<IAppContextProps>(AppContext);
    const [loading, { toggle: toggleLoading, setFalse: hideLoading }] = useBoolean(false);
    const [globalLinksList, setGlobalLinksList] = useState(undefined);
    const [userLinksList, setUserLinksList] = useState(undefined);
    const [sampleItems, { toggle: toggleSampleItems }] = useBoolean(false);
    const [message, setMessage] = useState<IMessageInfo>(undefined);
    const [disableActions, { toggle: toggleDisableActions }] = useBoolean(false);
    const { checkForListsAvailability, createGlobalLaunchPadLists, createUserLaunchPadLists, createSampleItemsForGlobalLinks } = useLaunchPadHelper(appContext.spService);

    const toggleLoadingActions = () => {
        toggleLoading();
        toggleDisableActions();
    };

    const clearFields = () => {
        setGlobalLinksList('');
        setUserLinksList('');
    };

    const _onCreateLists = async () => {
        try {
            setMessage(undefined);
            let listCreation: boolean = false;
            if (!globalLinksList && !userLinksList) setMessage({ msg: strings.Msg_ListReq, scope: MessageScope.Failure });
            else if (globalLinksList && userLinksList && globalLinksList.toLocaleLowerCase() === userLinksList.toLocaleLowerCase()) setMessage({ msg: strings.Msg_SameList, scope: MessageScope.Failure });
            else {
                setMessage(undefined);
                toggleLoadingActions();
                if (globalLinksList) {
                    const listExists = await checkForListsAvailability(globalLinksList);
                    if (listExists) {
                        setMessage({ msg: `List '${globalLinksList}' already exists.`, scope: MessageScope.Warning });
                    } else {
                        let result: IResult = await createGlobalLaunchPadLists(globalLinksList);
                        if (result.res !== 'Created') {
                            setMessage({ msg: result.res.message, scope: MessageScope.Failure });
                        } else {
                            listCreation = true;
                            if (sampleItems) {
                                setTimeout(async () => {
                                    await createSampleItemsForGlobalLinks(globalLinksList);
                                }, 2000);
                            }
                        }
                    }
                }
                if (userLinksList) {
                    const lstExists = await checkForListsAvailability(userLinksList);
                    if (lstExists) {
                        setMessage({ msg: `List '${globalLinksList}' already exists.`, scope: MessageScope.Warning });
                    } else {
                        let result = await createUserLaunchPadLists(userLinksList);
                        if (result.res !== 'Created') {
                            setMessage({ msg: result.res.message, scope: MessageScope.Failure });
                        } else listCreation = true;
                    }
                }
                toggleLoadingActions();
                if (listCreation) {
                    setMessage({ msg: 'List(s) created.', scope: MessageScope.Success });
                    clearFields();                    
                    if (sampleItems) toggleSampleItems();
                }
            }
        } catch (e) {
            setMessage({ msg: 'Something wrong while creating the lists. Please delete the lists from Site Contents and try creating the lists again from the settings screen.', scope: MessageScope.Failure });
            appContext.spService.writeErrorLog(e, LogSource.LaunchPadSettings);
            toggleLoadingActions();
        }
    };

    return (
        <Stack verticalAlign={'start'} tokens={{ childrenGap: 15 }}>
            <Stack.Item style={{ marginTop: '5px' }}>
                <div>{strings.Settings_Dialog_Note}</div>
            </Stack.Item>
            <Stack.Item>
                <Stack horizontal horizontalAlign={"space-evenly"}>
                    <Stack.Item styles={stackItemStyles}>
                        <div className={styles.divFormContainer}>
                            <div className={styles.divLabel}>
                                <Icon iconName={"TextField"} />
                                <label className={css(styles.fieldTitle)}>Global Links List:</label>
                            </div>
                            <div className={styles.divField}>
                                <TextField
                                    name={'lstGlobal'}
                                    value={globalLinksList}
                                    onChange={(ev, newValue) => { setGlobalLinksList(newValue); }}
                                    width={"100%"}
                                    style={{ resize: 'none' }}
                                    description='Links stored in this list are common to all the users.'
                                />
                            </div>
                        </div>
                    </Stack.Item>
                </Stack>
                {/* <Stack horizontal horizontalAlign={"space-evenly"}>
					<Stack.Item styles={stackItemStyles}>
						<div className={styles.divFormContainer}>
							<div className={styles.divLabel}>
								<Icon iconName={"CheckedOutByOther12"} />
								<label className={css(styles.fieldTitle)}>Enable User Links:</label>
							</div>
							<div className={styles.divField} style={{ marginLeft: '-4px' }}>
								<Checkbox checked={enableUserList} onChange={(ev, checked) => { toggleEnableUserList(); }}
									ariaLabel="Enable User Links" boxSide={'end'} />
							</div>
							<div style={{ fontSize: '11px', marginTop: '5px' }}>
								{"Links stored in this list are unique to the logged-in user."}
							</div>
						</div>
					</Stack.Item>
				</Stack> */}
                <Stack horizontal horizontalAlign={"space-evenly"}>
                    <Stack.Item styles={stackItemStyles}>
                        <div className={styles.divFormContainer}>
                            <div className={styles.divLabel}>
                                <Icon iconName={"TextField"} />
                                <label className={css(styles.fieldTitle)}>User Links List:</label>
                            </div>
                            <div className={styles.divField}>
                                <TextField
                                    name={'lstUser'}
                                    value={userLinksList}
                                    onChange={(ev, newValue) => { setUserLinksList(newValue); }}
                                    width={"100%"}
                                    style={{ resize: 'none' }}
                                />
                            </div>
                        </div>
                    </Stack.Item>
                </Stack>
                <Stack horizontal horizontalAlign={"space-evenly"}>
                    <Stack.Item styles={stackItemStyles}>
                        <div className={styles.divFormContainer}>
                            <div className={styles.divLabel}>
                                <Icon iconName={"CheckedOutByOther12"} />
                                <label className={css(styles.fieldTitle)}>Add Sample Items:</label>
                            </div>
                            <div className={styles.divField} style={{ marginLeft: '-4px' }}>
                                <Checkbox checked={sampleItems} onChange={(ev, checked) => { toggleSampleItems(); }}
                                    ariaLabel={"Sample Items"} boxSide={'end'} />
                            </div>
                            <div style={{ fontSize: '11px', marginTop: '5px' }}>
                                {"Sample items are created only for 'Global Links' list."}
                            </div>
                        </div>
                    </Stack.Item>
                </Stack>
            </Stack.Item>
            <Stack.Item style={{ marginBottom: '10px' }}>
                {message && message.msg &&
                    <Stack horizontal horizontalAlign={"space-evenly"}>
                        <MessageContainer MessageScope={message.scope} Message={message.msg} />
                    </Stack>
                }
            </Stack.Item>
            <Stack.Item>
                <Stack horizontal horizontalAlign={'end'} tokens={{ childrenGap: 10 }}>
                    <Stack.Item>
                        {loading &&
                            <div style={{ marginLeft: '20px', marginTop: '-10px' }}>
                                <ContentLoader loaderType={LoaderType.Spinner} spinSize={SpinnerSize.small} />
                            </div>
                        }
                    </Stack.Item>
                    <Stack.Item>
                        <PrimaryButton onClick={_onCreateLists} iconProps={{ iconName: "AddToShoppingList" }} text={strings.Btn_Create}
                            disabled={disableActions}></PrimaryButton>
                    </Stack.Item>
                    <Stack.Item>
                        <DefaultButton onClick={() => props.closeDialog()} iconProps={{ iconName: "Blocked" }} text={strings.Btn_Cancel}
                            disabled={disableActions}></DefaultButton>
                    </Stack.Item>
                </Stack>
            </Stack.Item>
        </Stack>
    );
};