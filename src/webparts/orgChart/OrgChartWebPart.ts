import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import OrgChart, { IOrgChartProps } from './components/OrgChart';

export interface IOrgChartWebPartProps {
  rootUserEmail: string;
}

export default class OrgChartWebPart extends BaseClientSideWebPart<IOrgChartWebPartProps> {

  public async render(): Promise<void> {
    const client: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient('3');

    const element: React.ReactElement<IOrgChartProps> = React.createElement(OrgChart, {
      graphClient: client,
      rootUserEmail: this.properties.rootUserEmail || undefined,
    });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  // ── Property Pane ──────────────────────────────────────────────────────────
  // Lets SharePoint editors pin the chart to a specific root user.

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: 'Org Chart Settings' },
          groups: [
            {
              groupName: 'Configuration',
              groupFields: [
                PropertyPaneTextField('rootUserEmail', {
                  label: 'Root User Email (optional)',
                  description:
                    'Pin the chart to a specific person (e.g. ceo@yourcompany.com). ' +
                    'Leave blank to auto-detect the top of the org.',
                  placeholder: 'Leave blank for full org',
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
