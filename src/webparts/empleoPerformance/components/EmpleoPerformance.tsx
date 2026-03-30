import * as React from 'react';
import type { IEmpleoPerformanceProps } from './IEmpleoPerformanceProps';
import { EmployeeDashboard } from './Employee/EmployeeDashboard';

export default class EmpleoPerformance extends React.Component<IEmpleoPerformanceProps> {
  public render(): React.ReactElement<IEmpleoPerformanceProps> {
    const { spHttpClient, webUrl, listEmployees, listDepartments, listEmployeeDepartments, listCycles, listDeliverables, listObjectives, listTasks, enableDesignMode } = this.props;

    return (
      <div className="tasksWrap">
        <EmployeeDashboard
          spHttpClient={spHttpClient}
          webUrl={webUrl}
          listEmployees={listEmployees}
          listDepartments={listDepartments}
          listEmployeeDepartments={listEmployeeDepartments}
          listCycles={listCycles}
          listDeliverables={listDeliverables}
          listObjectives={listObjectives}
          listTasks={listTasks}
          enableDesignMode={enableDesignMode}
        />
      </div>
    );
  }
}
