import * as React from 'react';
import * as strings from 'ModernQuickLinksWebPartStrings';
import { useEffect, useState, useContext } from 'react';
import styles from './Tile.module.scss';
import { Icon } from '@fluentui/react/lib/Icon';
import AddUpdateTile from './AddUpdateTile';
import AppContext, { IAppContextProps } from '../../../../common/AppContext';
import { IColumn } from '@fluentui/react/lib/DetailsList';
import { DefaultButton, IconButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { Stack } from 'office-ui-fabric-react';
import { Dialog, DialogFooter, DialogType, Image, Link, SpinnerSize } from '@fluentui/react';
import DetailsGrid from '../../../../common/components/DetailsGrid/DetailsGrid';
import { openURL } from '../../../../common/util';
import ContentLoader from '../../../../common/components/ContentLoader';
import { LoaderType } from '../../../../common/Constants';

export interface ITilesListProps {
    items: any[];
    customTileCount: number;
    userTileList: string;
    successCallback?: () => void;
    deleteCallback?: (itemid: number) => void;
}

const TilesList: React.FC<ITilesListProps> = (props) => {
    const appContext: IAppContextProps = useContext<IAppContextProps>(AppContext);
    const [allLinks, setAllLinks] = useState<any[]>([]);
    const [showAddUpdateTile, setShowAddUpdateTile] = useState<boolean>(false);
    const [item, setItem] = useState<any>(undefined);
    const [delItem, setDelItem] = useState<any>(undefined);
    const [columns, setcolumns] = useState<IColumn[]>([]);
    const [hideDialog, sethideDialog] = useState<boolean>(true);
    const [saving, setSaving] = useState<boolean>(false);
    const modalPropsStyles = { main: { minwidth: 400, maxWidth: 450 } };
    const dialogContentProps = {
        type: DialogType.close,
        title: 'Delete My Tile'
    };
    const modalProps = React.useMemo(
        () => ({
            isBlocking: true,
            styles: modalPropsStyles,
        }),
        [false],
    );

    const _dismissDelDialog = () => sethideDialog(true);

    const _editLink = (link: any) => {
        setItem(link);
        setShowAddUpdateTile(true);
    };

    const _invokeDelDialog = (item: any) => {
        setDelItem(item);
        sethideDialog(false);
    };

    const _buildColumns = (colValues: string[]) => {
        let cols: IColumn[] = [];
        cols.push({
            key: 'icoimg', name: '', fieldName: '', minWidth: 30, maxWidth: 50, 
            className: styles.gridIconCellStyle,
            onRender: (item: any, index: number, column: IColumn) => {
                return (
                    <div style={{ margin: '0 auto' }}>
                        {item.ImageUrl ? (
                            <Image src={item.ImageUrl} width={50} />
                        ) : (
                            <Icon iconName={item.IconName ? item.IconName : 'Link'} style={{ fontSize: '20px' }} />
                        )}
                    </div>
                );
            }
        });
        colValues.map(col => {
            if (col.toLowerCase() == "title") {
                cols.push({
                    key: col, name: 'Title', fieldName: col, minWidth: 200, maxWidth: 350,
                    onRender: (item: any, index: number, column: IColumn) => {
                        return (
                            <Link onClick={() => openURL(item.URL.Url, true)}
                                style={{ fontSize: '15px', fontWeight: '500', fontFamily: 'inherit' }}>{item.Title}</Link>
                        );
                    }
                } as IColumn);
            }
        });
        cols.push({
            key: 'action', name: 'Action', fieldName: 'ID', minWidth: 50, maxWidth: 50,
            onRender: (item: any, index: number, column: IColumn) => {
                return (
                    <Stack horizontal horizontalAlign='start'>
                        <Stack.Item>
                            <IconButton iconProps={{ iconName: 'PageEdit' }} onClick={() => { _editLink(item); }} />
                        </Stack.Item>
                        <Stack.Item>
                            <IconButton iconProps={{ iconName: 'Delete' }} onClick={() => _invokeDelDialog(item)} />
                        </Stack.Item>
                    </Stack>
                );
            }
        } as IColumn);
        setcolumns(cols);
    };

    const _loadAllLinks = async () => {
        let cols: string[] = ['Title'];
        _buildColumns(cols);
        setAllLinks(props.items);
    };

    const _dismissModify = () => {
        setItem(undefined);
        setShowAddUpdateTile(false);
    };

    const _deleteLink = async () => {
        if (delItem) {
            setSaving(true);
            await props.deleteCallback(delItem.ID);
            sethideDialog(true);
            setSaving(false);
            setDelItem(undefined);
        }
    };

    const _successCallback = () => {
        props.successCallback();
    };

    useEffect(() => {
        _loadAllLinks();
    }, [props]);

    return (
        <>
            {showAddUpdateTile && item ? (
                <AddUpdateTile onDismissPanel={_dismissModify} myTilesCallback={_successCallback} item={item} customTileCount={props.customTileCount}
                    userTileList={props.userTileList} context={appContext.context} />
            ) : (
                <div className={styles.tilesList}>
                    <DetailsGrid fields={columns} items={allLinks} enableSearch={true} searchKeys={['Title']} searchKeyPlaceholders={['Title']}
                        PagingSize={10} themeVariant={appContext.theme} />
                </div>
            )}
            <Dialog
                hidden={hideDialog}
                onDismiss={_dismissDelDialog}
                dialogContentProps={dialogContentProps}
                modalProps={modalProps} >
                <div className={styles.dialogContent}>
                    <div>{strings.DeleteDialogDescription}</div>
                    <div><b>Note:</b> {strings.DeleteDialogNote}</div>
                </div>
                <DialogFooter>
                    <Stack tokens={{ childrenGap: 10 }} horizontal horizontalAlign='end'>
                        <Stack.Item>
                            {saving &&
                                <div style={{ marginTop: '-13px', display: 'inline-block' }}><ContentLoader loaderType={LoaderType.Spinner} spinSize={SpinnerSize.medium} /></div>
                            }
                        </Stack.Item>
                        <Stack.Item>
                            <PrimaryButton onClick={_deleteLink} text="Yes" disabled={saving} />
                        </Stack.Item>
                        <Stack.Item>
                            <DefaultButton onClick={_dismissDelDialog} text="No" disabled={saving} />
                        </Stack.Item>
                    </Stack>
                </DialogFooter>
            </Dialog>
        </>
    );
};

export default TilesList;