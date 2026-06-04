import * as React from 'react';
import { MSGraphClientV3 } from '@microsoft/sp-http';
export interface IOrgChartProps {
    graphClient: MSGraphClientV3;
    rootUserEmail?: string;
}
declare const OrgChart: React.FC<IOrgChartProps>;
export default OrgChart;
