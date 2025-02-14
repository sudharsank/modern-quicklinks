import * as React from 'react';
import { useEffect, useState } from 'react';
import styles from './DetailsGrid.module.scss';
import * as commonStrings from 'ModernQuickLinksWebPartStrings';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { DetailsList, IColumn, DetailsListLayoutMode, ConstrainMode, SelectionMode } from '@fluentui/react/lib/DetailsList';
import { SearchBox } from 'office-ui-fabric-react/lib/SearchBox';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { LoaderType } from '../../Constants';
import { debounce } from 'throttle-debounce';
import ContentLoader from '../ContentLoader';
import { SpinnerSize } from '@fluentui/react/lib/Spinner';
import { useBoolean } from '@uifabric/react-hooks';
import DGPaging from './DGPaging';

export interface IDetailsGridProps {
    themeVariant: IReadonlyTheme | undefined;
    items: any[];
    fields: IColumn[];
    enableSearch: boolean;
    searchKeys: string[];
    searchKeyPlaceholders: string[];
    PagingSize: number;
}

const slice: any = require('lodash/slice');
const filter: any = require('lodash/filter');
const map: any = require('lodash/map');

const DetailsGrid: React.FC<IDetailsGridProps> = (props) => {
    const [searchLoading, { setTrue: showLoading, setFalse: hideLoading }] = useBoolean(false);
    const [columns, setColumns] = React.useState<IColumn[]>([]);
    const [items, setItems] = useState<any[]>([]);
    const [filItems, setFilItems] = useState<any[]>([]);
    const [searchText, setSearchText] = useState<string>('');
    // Paging
    const [currentPage, setcurrentPage] = useState<number>(1);
    const [pagedItems, setpagedItems] = useState<any[]>([]);
    const [pageSize, setpageSize] = useState<number>(props.PagingSize ? props.PagingSize : 20);

    const _onPageUpdate = async (pageno?: number) => {
        var currentPge = (pageno) ? pageno : currentPage;
        var startItem = ((currentPge - 1) * pageSize);
        var endItem = currentPge * pageSize;
        let filPageItems = slice(searchText.length > 0 ? filItems : items, startItem, endItem);
        setcurrentPage(currentPge);
        setpagedItems(filPageItems);
    };

    let _searchForItems = (searchKey?: string) => {
        if (searchKey && searchKey.length > 0) {
            let filtered: any[] = filter(items, (o:any) => {
                let fil = map(props.searchKeys, (s:any) => {
                    return o[s].toLowerCase().indexOf(searchKey.toLowerCase()) >= 0;
                });
                if (fil.length > 0 && filter(fil, (f:any) => { if (f) return f; }).length > 0) return o;
            });
            setFilItems(filtered);
        } else setFilItems(items);
        hideLoading();
    };
    _searchForItems = debounce(500, _searchForItems);

    const _onChangeSearchBox = (event?: React.ChangeEvent<HTMLInputElement>, searchKey?: string) => {
        setcurrentPage(1);
        showLoading();
        setSearchText(searchKey);
        _searchForItems(searchKey);
    };

    const _loadDataList = () => {
        setColumns(props.fields);
        setItems(props.items);
        setFilItems(props.items);
    };

    useEffect(() => {
        setcurrentPage(1);
        _onPageUpdate();
    }, [items, filItems, props.PagingSize]);

    useEffect(() => {
        if (props.fields && props.fields.length > 0 && props.items && props.items.length > 0) _loadDataList();
    }, [props.items, props.fields]);

    return (
        <div>
            <div className="ms-Grid-row" style={{ marginBottom: '5px', paddingLeft: '18px', display: 'flex' }}>
                <div style={{ width: '95%' }}>
                    <SearchBox
                        placeholder={`Search using ${props.searchKeyPlaceholders.join(', ')}...`}
                        onChange={_onChangeSearchBox}
                        underlined={true}
                        iconProps={{ iconName: 'Filter' }}
                        value={searchText}
                        autoFocus={false}
                        className={styles.favSearch}
                    />
                </div>
                {searchLoading &&
                    <div className={styles.searchLoader}>
                        <ContentLoader loaderType={LoaderType.Spinner} spinSize={SpinnerSize.small} />
                    </div>
                }
            </div>
            {searchText.length > 0 && filItems.length <= 0 &&
                <MessageBar messageBarType={MessageBarType.error}>{commonStrings.Msg_NoData}</MessageBar>
            }
            {pagedItems.length > 0 &&
                <>
                    <div className="ms-Grid-row">
                        <DGPaging
                            totalItems={searchText.length > 0 ? filItems.length : items.length}
                            itemsCountPerPage={pageSize}
                            onPageUpdate={_onPageUpdate}
                            currentPage={currentPage} />
                    </div>
                    <div style={{ overflowX: 'auto' }}>
                        <DetailsList
                            items={pagedItems}
                            setKey="set"
                            columns={columns}
                            compact={true}
                            layoutMode={DetailsListLayoutMode.justified}
                            constrainMode={ConstrainMode.unconstrained}
                            isHeaderVisible={true}
                            selectionMode={SelectionMode.none}
                            className={styles.detailsList}
                        />
                    </div>
                    <div className="ms-Grid-row">
                        <DGPaging
                            totalItems={searchText.length > 0 ? filItems.length : items.length}
                            itemsCountPerPage={pageSize}
                            onPageUpdate={_onPageUpdate}
                            currentPage={currentPage} />
                    </div>
                </>
            }
        </div>
    );
};

export default DetailsGrid;