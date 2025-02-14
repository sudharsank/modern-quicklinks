import * as React from 'react';
import styles from './DetailsGrid.module.scss';
import { css } from 'office-ui-fabric-react/lib/Utilities';
import Pagination from 'react-js-pagination';
import { Icon } from 'office-ui-fabric-react';

export type PageUpdateCallback = (pageNumber: number) => void;

export interface IDGPagingProps {
    totalItems: number;
    itemsCountPerPage: number;
    onPageUpdate: PageUpdateCallback;
    currentPage: number;
    className?: string;
}

export interface IPagingState {
    currentPage: number;
}

export default class DGPaging extends React.Component<IDGPagingProps, IPagingState> {

    constructor(props: IDGPagingProps) {
        super(props);

        this.state = {
            currentPage: props.currentPage
        };
    }

    public componentDidUpdate = (prevProps: IDGPagingProps) => {
        if (prevProps.currentPage !== this.props.currentPage) {
            this.setState({
                currentPage: this.props.currentPage
            });
        }
    }

    public render(): React.ReactElement<IDGPagingProps> {

        return (
            <div className={styles.paginationContainer}>
                <div className={css(styles.searchWp__paginationContainer__pagination, this.props.className ? this.props.className : '')}>
                    <Pagination
                        activePage={this.state.currentPage}
                        firstPageText={<Icon iconName='DoubleChevronLeft' />}
                        lastPageText={<Icon iconName='DoubleChevronRight' />}
                        prevPageText={<Icon iconName='ChevronLeft' />}
                        nextPageText={<Icon iconName='ChevronRight' />}
                        activeLinkClass={styles.active}
                        itemsCountPerPage={this.props.itemsCountPerPage}
                        totalItemsCount={this.props.totalItems}
                        pageRangeDisplayed={5}
                        onChange={(pageNumber: number) => {
                            this.setState({
                                currentPage: pageNumber
                            });
                            this.props.onPageUpdate(pageNumber);
                        }}
                    />
                </div>
            </div>
        );
    }
}