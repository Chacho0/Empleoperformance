import { SPHttpClient } from '@microsoft/sp-http';

export interface IEmpleoPerformanceProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  spHttpClient: SPHttpClient;
  webUrl: string;
  listEmployees: string;
  listDepartments: string;
  listEmployeeDepartments: string;
  listCycles: string;
  listDeliverables: string;
  listObjectives: string;
  listTasks: string;
  enableDesignMode: boolean;
}
