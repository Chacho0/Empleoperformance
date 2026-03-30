import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'EmpleoPerformanceWebPartStrings';
import EmpleoPerformance from './components/EmpleoPerformance';
import { IEmpleoPerformanceProps } from './components/IEmpleoPerformanceProps';

export interface IEmpleoPerformanceWebPartProps {
  description: string;
  listEmployees: string;
  listDepartments: string;
  listEmployeeDepartments: string;
  listCycles: string;
  listDeliverables: string;
  listObjectives: string;
  listTasks: string;
  enableDesignMode: boolean;
}

export default class EmpleoPerformanceWebPart extends BaseClientSideWebPart<IEmpleoPerformanceWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IEmpleoPerformanceProps> = React.createElement(
      EmpleoPerformance,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        spHttpClient: this.context.spHttpClient,
        webUrl: this.context.pageContext.web.absoluteUrl,
        listEmployees: this.properties.listEmployees || 'PM_Employee',
        listDepartments: this.properties.listDepartments || 'PM_Departments',
        listEmployeeDepartments: this.properties.listEmployeeDepartments || 'PM_EmployeeDepartments',
        listCycles: this.properties.listCycles || 'PM_Cycles',
        listDeliverables: this.properties.listDeliverables || 'PM_Deliverable',
        listObjectives: this.properties.listObjectives || 'PM_Objectives',
        listTasks: this.properties.listTasks || 'PM_Task',
        enableDesignMode: this.properties.enableDesignMode || false
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams':
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: 'Configure SharePoint lists for the Employee Performance module'
          },
          groups: [
            {
              groupName: 'Organizational Lists',
              groupFields: [
                PropertyPaneDropdown('listEmployees', {
                  label: 'Employees List',
                  options: [
                    { key: 'PM_Employee', text: 'PM_Employee' },
                    { key: 'Employees', text: 'Employees' },
                    { key: 'EmployeeList', text: 'EmployeeList' }
                  ],
                  selectedKey: this.properties.listEmployees || 'PM_Employee'
                }),
                PropertyPaneDropdown('listDepartments', {
                  label: 'Departments List',
                  options: [
                    { key: 'PM_Departments', text: 'PM_Departments' },
                    { key: 'Departments', text: 'Departments' },
                    { key: 'DepartmentList', text: 'DepartmentList' }
                  ],
                  selectedKey: this.properties.listDepartments || 'PM_Departments'
                }),
                PropertyPaneDropdown('listEmployeeDepartments', {
                  label: 'Employee-Department Relations List',
                  options: [
                    { key: 'PM_EmployeeDepartments', text: 'PM_EmployeeDepartments' },
                    { key: 'EmployeeDepartments', text: 'EmployeeDepartments' },
                    { key: 'EmployeeDeptRelations', text: 'EmployeeDeptRelations' }
                  ],
                  selectedKey: this.properties.listEmployeeDepartments || 'PM_EmployeeDepartments'
                }),
                PropertyPaneDropdown('listCycles', {
                  label: 'Evaluation Cycles List',
                  options: [
                    { key: 'PM_Cycles', text: 'PM_Cycles' },
                    { key: 'Cycles', text: 'Cycles' },
                    { key: 'EvaluationCycles', text: 'EvaluationCycles' }
                  ],
                  selectedKey: this.properties.listCycles || 'PM_Cycles'
                })
              ]
            },
            {
              groupName: 'Performance Lists',
              groupFields: [
                PropertyPaneDropdown('listDeliverables', {
                  label: 'Deliverables List',
                  options: [
                    { key: 'PM_Deliverable', text: 'PM_Deliverable' },
                    { key: 'Deliverables', text: 'Deliverables' },
                    { key: 'ProjectDeliverables', text: 'ProjectDeliverables' }
                  ],
                  selectedKey: this.properties.listDeliverables || 'PM_Deliverable'
                }),
                PropertyPaneDropdown('listObjectives', {
                  label: 'Objectives List',
                  options: [
                    { key: 'PM_Objectives', text: 'PM_Objectives' },
                    { key: 'Objectives', text: 'Objectives' },
                    { key: 'PerformanceObjectives', text: 'PerformanceObjectives' }
                  ],
                  selectedKey: this.properties.listObjectives || 'PM_Objectives'
                }),
                PropertyPaneDropdown('listTasks', {
                  label: 'Tasks List',
                  options: [
                    { key: 'PM_Task', text: 'PM_Task' },
                    { key: 'Tasks', text: 'Tasks' },
                    { key: 'PerformanceTasks', text: 'PerformanceTasks' }
                  ],
                  selectedKey: this.properties.listTasks || 'PM_Task'
                })
              ]
            },
            {
              groupName: 'Preview Options',
              groupFields: [
                PropertyPaneToggle('enableDesignMode', {
                  label: 'Design Mode',
                  onText: 'Active',
                  offText: 'Inactive'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
