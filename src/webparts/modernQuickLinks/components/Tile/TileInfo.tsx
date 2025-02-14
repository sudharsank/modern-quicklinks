import * as React from 'react';
import { DefaultButton } from '@fluentui/react/lib/Button';
import { Icon } from '@fluentui/react/lib/Icon';
import { IStackStyles, Stack } from '@fluentui/react/lib/Stack';
import strings from 'ModernQuickLinksWebPartStrings';

export interface ITileInfoProps {
    id: number;
    description: string;
    onDismissPanel: () => void;
}

const stackTokens = { childrenGap: 10 };
const stackStyles: Partial<IStackStyles> = { root: { width: '100%', marginTop: 20 } };

const TileInfo: React.FC<ITileInfoProps> = (props) => {
    return (
        <Stack tokens={stackTokens} styles={stackStyles} verticalAlign='start'>
            <Stack.Item>
                <div dangerouslySetInnerHTML={{ __html: props.description }}></div>
            </Stack.Item>
            <Stack.Item style={{ marginTop: '15px' }}>
                <DefaultButton onClick={props.onDismissPanel}>
                    {strings.CloseAL}
                </DefaultButton>
            </Stack.Item>
        </Stack>
    );
};

export default TileInfo;