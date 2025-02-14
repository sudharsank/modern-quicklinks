import * as React from 'react';
import * as strings from 'ModernQuickLinksWebPartStrings';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { ProgressIndicator } from '@fluentui/react/lib/ProgressIndicator';
import { LoaderType } from '../Constants';
import AppContext, { IAppContextProps } from '../AppContext';

export interface IContentLoaderProps {
    loaderMsg?: string;
    loaderType: LoaderType;
    spinSize?: SpinnerSize;
}

const ContentLoader: React.FunctionComponent<IContentLoaderProps> = (props) => {
    const appContext: IAppContextProps = React.useContext(AppContext);
    return (
        <div className="ms-Grid-row">
            {props.spinSize === SpinnerSize.xSmall ? (
                <div style={{ margin: "10px", marginRight: '14px' }}>
                    <Spinner label={props.loaderMsg} size={SpinnerSize.xSmall} ariaLive="assertive" labelPosition="top" />
                </div>
            ) : (
                <div style={{ margin: "20px" }}>
                    {props.loaderType == LoaderType.Spinner &&
                        <Spinner label={props.loaderMsg} size={props.spinSize ? props.spinSize : SpinnerSize.large} ariaLive="assertive" labelPosition="top"
                            theme={appContext?.theme as any} />
                    }
                    {props.loaderType == LoaderType.Indicator &&
                        <ProgressIndicator label={props.loaderMsg} description={strings.Msg_Wait} />
                    }
                </div>
            )}
        </div>
    );
};

export default ContentLoader;