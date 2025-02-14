import * as React from 'react';
import styles from './common.module.scss';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { MessageScope } from '../Constants';

export interface IMessageContainerProps {
    Message?: string;
    MessageScope: MessageScope;
}

const MessageContainer: React.FunctionComponent<IMessageContainerProps> = (props) => {
    return (
        <div className={styles.MessageContainer}>
            {
                props.MessageScope === MessageScope.Success &&
                <MessageBar messageBarType={MessageBarType.success} className={styles.successMessage}>{props.children? props.children : props.Message}</MessageBar>
            }
            {
                props.MessageScope === MessageScope.Failure &&
                <MessageBar messageBarType={MessageBarType.error} className={styles.errorMessage}>{props.children? props.children : props.Message}</MessageBar>
            }
            {
                props.MessageScope === MessageScope.Warning &&
                <MessageBar messageBarType={MessageBarType.warning} className={styles.warningMessage}>{props.children? props.children : props.Message}</MessageBar>
            }
            {
                props.MessageScope === MessageScope.Info &&
                <MessageBar messageBarType={MessageBarType.info} className={styles.infoMessage}>{props.children? props.children : props.Message}</MessageBar>
            }
        </div>
    );
};

export default MessageContainer;