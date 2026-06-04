import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
export interface IOrgChartWebPartProps {
    rootUserEmail: string;
}
export default class OrgChartWebPart extends BaseClientSideWebPart<IOrgChartWebPartProps> {
    render(): Promise<void>;
    protected onDispose(): void;
    protected get dataVersion(): Version;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
//# sourceMappingURL=OrgChartWebPart.d.ts.map