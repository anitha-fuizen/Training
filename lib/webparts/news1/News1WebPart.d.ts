import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import "@pnp/sp/sputilities";
export interface IGetListItemFromSharePointListWebPartProps {
    description: string;
}
export interface ISPLists {
    value: ISPList[];
}
export interface ISPList {
    Title: string;
    Description: String;
}
export interface Igetdetails {
    Reportingmanager: {
        Email: string;
    };
}
export default class GetListItemFromSharePointListWebPart extends BaseClientSideWebPart<IGetListItemFromSharePointListWebPartProps> {
    constructor(props: any);
    private _getListData;
    private _renderListAsync;
    private _getData;
    send_Email(title: any): Promise<void>;
    private _renderList;
    render(): void;
    protected get dataVersion(): Version;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
//# sourceMappingURL=News1WebPart.d.ts.map