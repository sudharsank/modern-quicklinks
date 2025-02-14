import * as React from 'react';
import * as commonStrings from 'ModernQuickLinksWebPartStrings';
import commonStyles from './common.module.scss';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { DialogTypes, MessageScope } from '../Constants';
import MessageContainer from './Message';
import { DisplayMode } from '@microsoft/sp-core-library';
import { useBoolean } from '@fluentui/react-hooks';
import { AppDialog } from './AppBaseDialog';
import { IPropertyPaneAccessor } from '@microsoft/sp-webpart-base';

export interface IConfigMissing {
    displayMode: DisplayMode;
    propertyPane: IPropertyPaneAccessor;
    onLicManage: () => void;
}

export const ConfigMissing: React.FC<IConfigMissing> = (props) => {
    const { displayMode } = props;
    const [showConfigDialog, { toggle: toggleConfigDialog }] = useBoolean(false);

    const _openPropertyPane = () => props.propertyPane.open();

    return (
        <>
            <MessageContainer MessageScope={MessageScope.Warning}>
                <div className={commonStyles.configHeader}>{commonStrings.Msg_ConfigHeader}</div>
                <ol className={commonStyles.configList}>
                    <li>{commonStrings.Msg_Config_CreateList} <Link className={commonStyles.configLicLink} onClick={toggleConfigDialog}>{commonStrings.Lbl_CreateLists}</Link></li>
                    {displayMode === DisplayMode.Edit ? (
                        <li>{commonStrings.Msg_Config_WPProp_Edit} <Link className={commonStyles.configLicLink} onClick={_openPropertyPane}>{commonStrings.Lbl_Config_WPProp}</Link></li>
                    ) : (
                        <li>{commonStrings.Msg_Config_WPProp_Read}</li>
                    )}
                    <li><Link className={commonStyles.configLicLink} onClick={() => props.onLicManage()}>Manage License</Link></li>
                </ol>
            </MessageContainer>
            {showConfigDialog &&
                <AppDialog dialogType={DialogTypes.LaunchPadSettings} closeCallback={toggleConfigDialog} successCallback={toggleConfigDialog} />
            }
        </>
    );
};