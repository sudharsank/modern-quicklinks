import * as React from 'react';
import * as ReactDOM from 'react-dom';
import * as strings from 'ModernQuickLinksWebPartStrings';
import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import { Dialog, DialogType, IDialogContentProps } from '@fluentui/react/lib/Dialog';
import { useBoolean } from '@fluentui/react-hooks';
import { ResponsiveMode } from '@fluentui/react';
import { DialogTypes } from '../Constants';
import { LicenseForm } from './License/LicenseForm';
import { LaunchPadSettings } from './Dialogs/LaunchPadSettings';

const modelProps = {
    isBlocking: true
};

const licdialogContentProps: IDialogContentProps = {
    type: DialogType.largeHeader,
    title: strings.LicDialogTitle,
    subText: '',
    showCloseButton: true,
    styles: {
        content: {
            maxHeight: '850px',
            overflowY: 'auto'
        }
    }
};

const launchPadSettingsdialogContentProps: IDialogContentProps = {
    type: DialogType.largeHeader,
    title: strings.SettingsDialogTitle,
    subText: '',
    showCloseButton: true,
    styles: {
        content: {
            maxHeight: '850px',
            overflowY: 'auto'
        }
    }
};

const newsListSettingsDlgContProps: IDialogContentProps = {
    type: DialogType.largeHeader,
    title: strings.NS_DialogTitle,
    subText: '',
    showCloseButton: true,
    styles: {
        content: {
            maxHeight: '850px',
            overflowY: 'auto'
        }
    }
};

export interface IAppDialogProps {
    closeCallback?: () => void;
    successCallback?: () => void;
    dialogType: DialogTypes;
}

export const AppDialog: React.FC<IAppDialogProps> = (props) => {
    const [hideDialog, { toggle: toggleHideDialog }] = useBoolean(false);

    const _closeDialog = () => {
        // switch(props.dialogType) {
        // 	case DialogTypes.LicDialog:
        // 		toggleHideDialog();
        // 		break;
        // }
        toggleHideDialog();
        if (props.closeCallback) props.closeCallback();
    };

    const _saveDialog = () => {
        toggleHideDialog();
        if (props.successCallback) props.successCallback();
    }

    return (
        <>
            {props.dialogType === DialogTypes.LicDialog &&
                <Dialog
                    hidden={hideDialog}
                    onDismiss={_closeDialog}
                    dialogContentProps={licdialogContentProps}
                    modalProps={modelProps}
                    closeButtonAriaLabel={strings.CloseAL}
                    minWidth="500px"
                    responsiveMode={ResponsiveMode.large}>
                    <LicenseForm closeDialog={_closeDialog} saveDialog={_saveDialog}></LicenseForm>
                </Dialog>
            }
            {props.dialogType === DialogTypes.LaunchPadSettings &&
                <Dialog
                    hidden={hideDialog}
                    onDismiss={_closeDialog}
                    dialogContentProps={launchPadSettingsdialogContentProps}
                    modalProps={modelProps}
                    closeButtonAriaLabel={strings.CloseAL}
                    minWidth="500px"
                    responsiveMode={ResponsiveMode.large}>
                    <LaunchPadSettings closeDialog={_closeDialog} />
                </Dialog>
            }
        </>
    );
}