import * as React from 'react';
import { useEffect, useState, useContext } from 'react';
import styles from './common.module.scss';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { DesignTypes, PanelTypes } from '../Constants';
import { BaseWebPartContext } from '@microsoft/sp-webpart-base';
import AppContext, { IAppContextProps } from '../AppContext';
import AddUpdateTile from '../../webparts/modernQuickLinks/components/Tile/AddUpdateTile';
import TilesList from '../../webparts/modernQuickLinks/components/Tile/TilesList';
import TileInfo from '../../webparts/modernQuickLinks/components/Tile/TileInfo';
import TileColor from '../../webparts/modernQuickLinks/components/Tile/TileColor';

export interface IAppPanelProps {
    panelType: PanelTypes;
    dismissCallback?: () => void;
    successCallback?: () => void;
    deleteCallback?: (itemid: number) => void;
    headerText: string;
    item?: any;
    customTileRowCount?: number;
    userTileList?: string;
    globalList?: string;
    items?: any[];
    designType?: DesignTypes;
    tileColors?: any | undefined;
}

export const AppPanel: React.FC<IAppPanelProps> = (props) => {
    const appContext: IAppContextProps = useContext<IAppContextProps>(AppContext);
    const [isOpen, setIsOpen] = useState(true);

    const _dismissPanel = () => {
        // eslint-disable-next-line @typescript-eslint/no-unused-expressions
        setIsOpen(false); props.dismissCallback ? props.dismissCallback() : undefined;
    };
    const _successCall = () => {
        // eslint-disable-next-line @typescript-eslint/no-unused-expressions
        props.successCallback ? props.successCallback() : undefined;
    };

    return (
        <>
            <Panel
                isOpen={isOpen}
                onDismiss={_dismissPanel}
                isBlocking={true}
                type={PanelType.medium}
                closeButtonAriaLabel="Close"
                headerText={props.headerText}
                className={styles.customPanel}
                headerClassName={styles.header}>
                {props.panelType === PanelTypes.AddUpdateTiles &&
                    <AddUpdateTile onDismissPanel={_dismissPanel} myTilesCallback={_successCall} item={props.item} customTileCount={props.customTileRowCount}
                        userTileList={props.userTileList} context={appContext.context} />
                }
                {props.panelType === PanelTypes.ManageTiles &&
                    <TilesList items={props.items} customTileCount={props.customTileRowCount} userTileList={props.userTileList} successCallback={props.successCallback}
                        deleteCallback={props.deleteCallback} />
                }
                {props.panelType === PanelTypes.TileInfo &&
                    <TileInfo onDismissPanel={_dismissPanel} id={props.item.ID} description={props.item.Description} />
                }
                {props.panelType === PanelTypes.TileColor &&
                    <TileColor item={props.item} onDismissPanel={_dismissPanel} myTilesCallback={_successCall} globalList={props.globalList} designType={props.designType}
                        tileColors={props.tileColors} />
                }
            </Panel>
        </>
    );
};