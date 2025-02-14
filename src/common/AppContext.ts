import * as React from 'react';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { HttpClient } from '@microsoft/sp-http';
import SPService from './Helpers/spService';
import { BaseWebPartContext } from '@microsoft/sp-webpart-base';

export interface IAppContextProps {
    spService: SPService;
    httpClient: HttpClient;
	tName: string;
    wpInstanceId?: string;
	theme?: IReadonlyTheme;
    webabsurl?: string;
    isSiteAdmin?: boolean;
    serRelativeUrl?: string;
    context?: BaseWebPartContext;
}

const AppContext = React.createContext<IAppContextProps>(undefined);

export default AppContext;
