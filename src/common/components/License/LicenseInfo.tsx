import * as React from 'react';
import styles from '../common.module.scss';
import * as strings from 'ModernQuickLinksWebPartStrings';
import { useState, FC, useEffect } from 'react';
import { Stack, IStackTokens, StackItem, Link } from '@fluentui/react';
import MessageContainer from '../Message';
import { DialogTypes, LicenseMessage, MessageScope } from '../../Constants';
import { AppDialog } from '../AppBaseDialog';

const controStackTokens: IStackTokens = { childrenGap: '10' };

export interface ILicenseInfoProps {
    licMsg: LicenseMessage;
    onLicValidated: () => void;
    onCloseCallback: () => void;
    showLicenseForm?: boolean;
}

export const LicenseInfo: FC<ILicenseInfoProps> = (props) => {
    const [showLicenseForm, setShowLicenseForm] = useState<boolean>(false);

    const _openLicenseDialog = () => {
        setShowLicenseForm(true);
    }
    const _closeLicenseDialog = () => {
        setShowLicenseForm(false);
        // eslint-disable-next-line @typescript-eslint/no-unused-expressions
        props.onCloseCallback ? props.onCloseCallback() : null;
    }

    useEffect(() => {
        if (props.showLicenseForm) _openLicenseDialog();
    }, [props.showLicenseForm])

    return (
        <div>
            <Stack tokens={controStackTokens} horizontal>
                <Stack.Item>
                    {props.licMsg === LicenseMessage.NotConfigured &&
                        <MessageContainer MessageScope={MessageScope.Info}>
                            <div>{strings.LicNotConfigured}
                                <Link className={styles.configLicLink} onClick={_openLicenseDialog}>{strings.LicLink}</Link>
                            </div>
                        </MessageContainer>
                    }
                    {props.licMsg === LicenseMessage.Expired &&
                        <MessageContainer MessageScope={MessageScope.Failure}>
                            <div>{strings.LicExpired}
                                <Link className={styles.configLicLink} onClick={_openLicenseDialog}>{strings.LicLink}</Link>
                            </div>
                        </MessageContainer>
                    }
                </Stack.Item>
            </Stack>
            {showLicenseForm &&
                <AppDialog dialogType={DialogTypes.LicDialog} closeCallback={_closeLicenseDialog} successCallback={props.onLicValidated} />
            }
        </div>
    );
};