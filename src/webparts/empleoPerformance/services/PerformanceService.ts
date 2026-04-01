import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import {
  IDeliverable,
  IDeliverableRelation,
  IObjective,
  IObjectiveRelation,
  ITask,
  ITaskRelation,
  DeliverableCategory,
  DeliverableStatus,
  Priority,
  ObjectiveSource,
  ObjectiveType,
  ObjectiveCategory,
  ObjectiveStatus,
  TaskSource,
  TaskType,
  TaskStatus,
  ApprovalStatus,
  IEmployee,
  IDepartment,
  IEvaluationCycle
} from '../interfaces/IOrganization';

export interface IPerformanceServiceConfig {
  deliverables: string;
  objectives: string;
  tasks: string;
  employees: string;
  departments: string;
  cycles: string;
}

export class PerformanceService {
  private spHttpClient: SPHttpClient;
  private webUrl: string;
  private listNames: IPerformanceServiceConfig;
  private entityTypes: Map<string, string> = new Map();

  constructor(spHttpClient: SPHttpClient, webUrl: string, config?: IPerformanceServiceConfig) {
    this.spHttpClient = spHttpClient;
    this.webUrl = webUrl;
    this.listNames = {
      deliverables: config?.deliverables || 'PM_Deliverable',
      objectives: config?.objectives || 'PM_Objectives',
      tasks: config?.tasks || 'PM_Task',
      employees: config?.employees || 'PM_Employee',
      departments: config?.departments || 'PM_Departments',
      cycles: config?.cycles || 'PM_Cycles'
    };
  }

  public updateListNames(config: IPerformanceServiceConfig): void {
    this.listNames = { ...this.listNames, ...config };
  }

  private getApiUrl(listName: string, query?: string): string {
    const baseUrl = `${this.webUrl}/_api/web/lists/getbytitle('${listName}')/items`;
    return query ? `${baseUrl}?${query}` : baseUrl;
  }

  private async getListEntityType(listName: string): Promise<string> {
    if (this.entityTypes.has(listName)) {
      return this.entityTypes.get(listName) || `SP.Data.${listName}ListItem`;
    }
    try {
      const url = `${this.webUrl}/_api/web/lists/getbytitle('${listName}')/entitytype`;
      const response = await this.spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (response.ok) {
        const data = await response.json();
        const entityType = data.EntityType || `SP.Data.${listName}ListItem`;
        this.entityTypes.set(listName, entityType);
        return entityType;
      }
    } catch (e) {
      console.log('Could not fetch entity type for', listName, e);
    }
    return `SP.Data.${listName}ListItem`;
  }

  private async fetchItems<T>(listName: string, query?: string): Promise<T[]> {
    const url = this.getApiUrl(listName, query);
    console.log(`Fetching from ${listName}:`, url);
    try {
      const response: SPHttpClientResponse = await this.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1
      );
      if (!response.ok) {
        const errorText = await response.text();
        console.error(`Error fetching from ${listName}:`, response.status, errorText);
        return [];
      }
      const data = await response.json();
      console.log(`Results from ${listName}:`, data.value?.length || 0, 'items');
      return data.value || [];
    } catch (error) {
      console.error(`Exception fetching from ${listName}:`, error);
      return [];
    }
  }

  private async fetchItemById<T>(listName: string, id: number): Promise<T | null> {
    const url = this.getApiUrl(listName) + `(${id})`;
    try {
      const response: SPHttpClientResponse = await this.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1
      );
      if (response.ok) {
        return await response.json();
      }
      return null;
    } catch {
      return null;
    }
  }

  private async createItem(listName: string, data: any): Promise<number | null> {
    const url = this.getApiUrl(listName);

    console.log('Creating item in', listName, 'with data:', JSON.stringify(data, null, 2));

    const response: SPHttpClientResponse = await this.spHttpClient.post(
      url,
      SPHttpClient.configurations.v1,
      {
        body: JSON.stringify(data),
        headers: {
          'Accept': 'application/json',
          'Content-Type': 'application/json'
        }
      }
    );

    if (response.ok) {
      const result = await response.json();
      return result.Id;
    }

    const errorText = await response.text();
    console.error(`Error creating item in ${listName}:`, response.status, errorText);
    return null;
  }

  private async updateItem(listName: string, id: number, data: any): Promise<boolean> {
    const url = this.getApiUrl(listName) + `(${id})`;

    const cleanData: any = {};
    for (const key of Object.keys(data)) {
      const value = data[key];
      if (value === 'True' || value === 'False') {
        cleanData[key] = value === 'True';
      } else if (value === 'Yes' || value === 'No') {
        cleanData[key] = value === 'Yes';
      } else if (typeof value === 'string' && (value.startsWith('{') || value.startsWith('['))) {
        try {
          cleanData[key] = JSON.parse(value);
        } catch {
          cleanData[key] = value;
        }
      } else {
        cleanData[key] = value;
      }
    }

    const response: SPHttpClientResponse = await this.spHttpClient.post(
      url,
      SPHttpClient.configurations.v1,
      {
        body: JSON.stringify(cleanData),
        headers: {
          'Accept': 'application/json',
          'Content-Type': 'application/json',
          'X-HTTP-Method': 'MERGE',
          'IF-Match': '*'
        }
      }
    );

    if (!response.ok) {
      const errorText = await response.text();
      console.error(`Error updating item in ${listName}:`, response.status, errorText);
    }
    return response.ok;
  }

  private async deleteItem(listName: string, id: number): Promise<boolean> {
    const url = this.getApiUrl(listName) + `(${id})`;
    const response: SPHttpClientResponse = await this.spHttpClient.post(
      url,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json',
          'Content-Type': 'application/json',
          'X-HTTP-Method': 'DELETE'
        }
      }
    );
    return response.ok;
  }

  // ==================== DELIVERABLES ====================

  public async getAllDeliverables(): Promise<IDeliverable[]> {
    const query = `$select=*,AssignedByEmployee/Id,AssignedByEmployee/Title,DepartmentManagerEmployee/Id,DepartmentManagerEmployee/Title,AssignedDepartment/Id,AssignedDepartment/Title,Cycle/Id,Cycle/Title&$expand=AssignedByEmployee,DepartmentManagerEmployee,AssignedDepartment,Cycle&$filter=IsActive eq true&$orderby=Created desc`;
    return this.fetchItems<IDeliverable>(this.listNames.deliverables, query);
  }

  public async getDeliverablesByDepartment(departmentId: number): Promise<IDeliverable[]> {
    if (!departmentId || departmentId <= 0) {
      return this.fetchItems<IDeliverable>(this.listNames.deliverables, '$orderby=Created desc&$top=50');
    }
    const query = `$select=*,AssignedByEmployee/Id,AssignedByEmployee/Title,DepartmentManagerEmployee/Id,DepartmentManagerEmployee/Title,AssignedDepartment/Id,AssignedDepartment/Title,Cycle/Id,Cycle/Title&$expand=AssignedByEmployee,DepartmentManagerEmployee,AssignedDepartment,Cycle&$filter=AssignedDepartment/Id eq ${departmentId} and IsActive eq true&$orderby=Created desc`;
    return this.fetchItems<IDeliverable>(this.listNames.deliverables, query);
  }

  public async getDeliverablesByEmployee(employeeId: number): Promise<IDeliverable[]> {
    if (!employeeId || employeeId <= 0) return [];
    const query = `$select=*,AssignedByEmployee/Id,AssignedByEmployee/Title,DepartmentManagerEmployee/Id,DepartmentManagerEmployee/Title,AssignedDepartment/Id,AssignedDepartment/Title,Cycle/Id,Cycle/Title&$expand=AssignedByEmployee,DepartmentManagerEmployee,AssignedDepartment,Cycle&$filter=AssignedByEmployee/Id eq ${employeeId} and IsActive eq true&$orderby=Created desc`;
    return this.fetchItems<IDeliverable>(this.listNames.deliverables, query);
  }

  public async getDeliverablesByCycle(cycleId: number): Promise<IDeliverable[]> {
    if (!cycleId || cycleId <= 0) return [];
    const query = `$select=*,AssignedByEmployee/Id,AssignedByEmployee/Title,DepartmentManagerEmployee/Id,DepartmentManagerEmployee/Title,AssignedDepartment/Id,AssignedDepartment/Title,Cycle/Id,Cycle/Title&$expand=AssignedByEmployee,DepartmentManagerEmployee,AssignedDepartment,Cycle&$filter=Cycle/Id eq ${cycleId} and IsActive eq true&$orderby=Created desc`;
    return this.fetchItems<IDeliverable>(this.listNames.deliverables, query);
  }

  public async getDeliverablesByManager(managerEmployeeId: number): Promise<IDeliverable[]> {
    if (!managerEmployeeId || managerEmployeeId <= 0) return [];
    const query = `$select=*,AssignedByEmployee/Id,AssignedByEmployee/Title,DepartmentManagerEmployee/Id,DepartmentManagerEmployee/Title,AssignedDepartment/Id,AssignedDepartment/Title,Cycle/Id,Cycle/Title&$expand=AssignedByEmployee,DepartmentManagerEmployee,AssignedDepartment,Cycle&$filter=DepartmentManagerEmployee/Id eq ${managerEmployeeId} and IsActive eq true&$orderby=Created desc`;
    return this.fetchItems<IDeliverable>(this.listNames.deliverables, query);
  }

  public async getDeliverableById(id: number): Promise<IDeliverable | null> {
    return this.fetchItemById<IDeliverable>(this.listNames.deliverables, id);
  }

  public async getDeliverableWithRelations(id: number): Promise<IDeliverableRelation | null> {
    const deliverable = await this.getDeliverableById(id);
    if (!deliverable) return null;

    const relations: IDeliverableRelation = {
      deliverable,
      objectives: [],
      tasks: []
    };

    try {
      const deptQuery = `$filter=Id eq ${deliverable.AssignedDepartmentId}&$top=1`;
      const depts = await this.fetchItems<IDepartment>(this.listNames.departments, deptQuery);
      relations.department = depts[0];

      const empQuery = `$filter=EmployeeEmail/Id eq ${deliverable.AssignedByEmployeeId}&$expand=EmployeeEmail&$top=1`;
      const assignedByEmps = await this.fetchItems<any>(this.listNames.employees, empQuery);
      if (assignedByEmps[0]) relations.assignedByEmployee = assignedByEmps[0];

      const mgrQuery = `$filter=EmployeeEmail/Id eq ${deliverable.DepartmentManagerEmployeeId}&$expand=EmployeeEmail&$top=1`;
      const mgrEmps = await this.fetchItems<any>(this.listNames.employees, mgrQuery);
      if (mgrEmps[0]) relations.departmentManager = mgrEmps[0];

      const cycleQuery = `$filter=Id eq ${deliverable.CycleId}&$top=1`;
      const cycles = await this.fetchItems<IEvaluationCycle>(this.listNames.cycles, cycleQuery);
      relations.cycle = cycles[0];

      const objQuery = `$filter=DeliverableId eq ${id} and IsActive eq true`;
      relations.objectives = await this.fetchItems<IObjective>(this.listNames.objectives, objQuery);

      const taskQuery = `$filter=DeliverableId eq ${id} and IsActive eq true`;
      relations.tasks = await this.fetchItems<ITask>(this.listNames.tasks, taskQuery);

      relations.childrenCount = {
        objectives: relations.objectives.length,
        tasks: relations.tasks.length
      };
    } catch (e) {
      console.log('Error loading deliverable relations:', e);
    }

    return relations;
  }

  public async createDeliverable(data: Partial<IDeliverable>): Promise<number | null> {
    const toBool = (val: any): boolean => val === true || val === 1 || val === '1' || val === 'Yes' || val === 'True';
    const toYesNo = (val: any): string => toBool(val) ? 'Yes' : 'No';

    const deliverableData: any = {
      Title: data.Title,
      AssignedDepartmentId: data.AssignedDepartmentId,
      AssignedByEmployeeId: data.AssignedByEmployeeId,
      DepartmentManagerEmployeeId: data.DepartmentManagerEmployeeId,
      Cycles: data.CycleId ? data.CycleId.toString() : undefined,
      Category: data.Category || 'Project Deliverable',
      Priority: data.Priority || 'Medium',
      Status: data.Status || 'Not Started',
      DueDate: data.DueDate,
      ProgressPercent: data.ProgressPercent || 0,
      CanCreateChildObjectives: toYesNo(data.CanCreateChildObjectives) || 'True',
      CanCreateChildTasks: toYesNo(data.CanCreateChildTasks) || 'True',
      IsActive: 'True'
    };

    if (data.DeliverableCode) deliverableData.DeliverableCode = data.DeliverableCode;
    if (data.StartDate) deliverableData.StartDate = data.StartDate;
    if (data.Description) deliverableData.Description = data.Description;
    if (data.ExpectedOutcome) deliverableData.ExpectedOutcome = data.ExpectedOutcome;
    if (data.Comments) deliverableData.Comments = data.Comments;

    return this.createItem(this.listNames.deliverables, deliverableData);
  }

  public async updateDeliverable(id: number, data: Partial<IDeliverable>): Promise<boolean> {
    const toBool = (val: any): boolean => val === true || val === 1 || val === '1' || val === 'Yes' || val === 'True';
    const toYesNo = (val: any): string => toBool(val) ? 'Yes' : 'No';
    const updateData: any = {};

    if (data.Title !== undefined) updateData.Title = data.Title;
    if (data.DeliverableCode !== undefined) updateData.DeliverableCode = data.DeliverableCode;
    if (data.AssignedDepartmentId !== undefined) updateData.AssignedDepartmentId = data.AssignedDepartmentId;
    if (data.Category !== undefined) updateData.Category = data.Category;
    if (data.Priority !== undefined) updateData.Priority = data.Priority;
    if (data.Status !== undefined) updateData.Status = data.Status;
    if (data.StartDate !== undefined) updateData.StartDate = data.StartDate;
    if (data.DueDate !== undefined) updateData.DueDate = data.DueDate;
    if (data.ProgressPercent !== undefined) updateData.ProgressPercent = data.ProgressPercent;
    if (data.Description !== undefined) updateData.Description = data.Description;
    if (data.ExpectedOutcome !== undefined) updateData.ExpectedOutcome = data.ExpectedOutcome;
    if (data.Comments !== undefined) updateData.Comments = data.Comments;
    if (data.CanCreateChildObjectives !== undefined) updateData.CanCreateChildObjectives = toYesNo(data.CanCreateChildObjectives);
    if (data.CanCreateChildTasks !== undefined) updateData.CanCreateChildTasks = toYesNo(data.CanCreateChildTasks);
    if (data.IsActive !== undefined) updateData.IsActive = toYesNo(data.IsActive);

    return this.updateItem(this.listNames.deliverables, id, updateData);
  }

  public async deleteDeliverable(id: number): Promise<boolean> {
    return this.deleteItem(this.listNames.deliverables, id);
  }

  // ==================== OBJECTIVES ====================

  public async getAllObjectives(): Promise<IObjective[]> {
    const query = `$filter=IsActive eq true&$orderby=Created desc`;
    return this.fetchItems<IObjective>(this.listNames.objectives, query);
  }

  public async getObjectivesByDepartments(departmentIds: number[]): Promise<IObjective[]> {
    if (!departmentIds || departmentIds.length === 0) return [];

    const uniqueIds: number[] = [];
    departmentIds.forEach((id: number) => { if (id > 0 && uniqueIds.indexOf(id) === -1) uniqueIds.push(id); });
    if (uniqueIds.length === 0) return [];

    const filter = uniqueIds.map(id => `AssignedDepartment/Id eq ${id}`).join(' or ');

    const query =
      `$select=*,AssignedEmployee/Id,AssignedEmployee/Title,AssignedByEmployee/Id,AssignedByEmployee/Title,AssignedDepartment/Id,AssignedDepartment/Title,Cycle/Id,Cycle/Title,Deliverable/Id,Deliverable/Title` +
      `&$expand=AssignedEmployee,AssignedByEmployee,AssignedDepartment,Cycle,Deliverable` +
      `&$filter=(${filter}) and IsActive eq true` +
      `&$orderby=Created desc`;

    return this.fetchItems<IObjective>(this.listNames.objectives, query);
  }

  public async getObjectivesByDeliverable(deliverableId: number): Promise<IObjective[]> {
    if (!deliverableId || deliverableId <= 0) return [];
    const query = `$filter=Deliverable/Id eq ${deliverableId} and IsActive eq true&$orderby=Created desc`;
    return this.fetchItems<IObjective>(this.listNames.objectives, query);
  }

  public async getObjectivesByEmployee(employeeId: number): Promise<IObjective[]> {
    if (!employeeId || employeeId <= 0) return [];
    const query = `$select=*,AssignedEmployee/Id,AssignedEmployee/Title,AssignedByEmployee/Id,AssignedByEmployee/Title,AssignedDepartment/Id,AssignedDepartment/Title,Cycle/Id,Cycle/Title,Deliverable/Id,Deliverable/Title&$expand=AssignedEmployee,AssignedByEmployee,AssignedDepartment,Cycle,Deliverable&$filter=AssignedEmployee/Id eq ${employeeId} and IsActive eq true&$orderby=Created desc`;
    return this.fetchItems<IObjective>(this.listNames.objectives, query);
  }

  public async getObjectivesByDepartment(departmentId: number): Promise<IObjective[]> {
    if (!departmentId || departmentId <= 0) {
      return this.fetchItems<IObjective>(this.listNames.objectives, '$orderby=Created desc&$top=50');
    }
    const query = `$select=*,AssignedEmployee/Id,AssignedEmployee/Title,AssignedByEmployee/Id,AssignedByEmployee/Title,AssignedDepartment/Id,AssignedDepartment/Title,Cycle/Id,Cycle/Title,Deliverable/Id,Deliverable/Title&$expand=AssignedEmployee,AssignedByEmployee,AssignedDepartment,Cycle,Deliverable&$filter=AssignedDepartment/Id eq ${departmentId} and IsActive eq true&$orderby=Created desc`;
    return this.fetchItems<IObjective>(this.listNames.objectives, query);
  }

  public async getObjectivesByCycle(cycleId: number): Promise<IObjective[]> {
    if (!cycleId || cycleId <= 0) return [];
    const query = `$select=*,AssignedEmployee/Id,AssignedEmployee/Title,AssignedByEmployee/Id,AssignedByEmployee/Title,AssignedDepartment/Id,AssignedDepartment/Title,Cycle/Id,Cycle/Title,Deliverable/Id,Deliverable/Title&$expand=AssignedEmployee,AssignedByEmployee,AssignedDepartment,Cycle,Deliverable&$filter=Cycle/Id eq ${cycleId} and IsActive eq true&$orderby=Created desc`;
    return this.fetchItems<IObjective>(this.listNames.objectives, query);
  }

  public async getObjectiveById(id: number): Promise<IObjective | null> {
    return this.fetchItemById<IObjective>(this.listNames.objectives, id);
  }

  public async getObjectiveWithRelations(id: number): Promise<IObjectiveRelation | null> {
    const objective = await this.getObjectiveById(id);
    if (!objective) return null;

    const relations: IObjectiveRelation = {
      objective,
      tasks: []
    };

    try {
      if (objective.DeliverableId) {
        const delQuery = `$filter=Id eq ${objective.DeliverableId}&$top=1`;
        const dels = await this.fetchItems<IDeliverable>(this.listNames.deliverables, delQuery);
        relations.deliverable = dels[0];
      }

      if (objective.AssignedEmployeeId) {
        const empQuery = `$filter=EmployeeEmail/Id eq ${objective.AssignedEmployeeId}&$expand=EmployeeEmail&$top=1`;
        const emps = await this.fetchItems<any>(this.listNames.employees, empQuery);
        if (emps[0]) relations.assignedEmployee = emps[0];
      }

      if (objective.AssignedDepartmentId) {
        const deptQuery = `$filter=Id eq ${objective.AssignedDepartmentId}&$top=1`;
        const depts = await this.fetchItems<IDepartment>(this.listNames.departments, deptQuery);
        relations.assignedDepartment = depts[0];
      }

      if (objective.AssignedByEmployeeId) {
        const byQuery = `$filter=EmployeeEmail/Id eq ${objective.AssignedByEmployeeId}&$expand=EmployeeEmail&$top=1`;
        const byEmps = await this.fetchItems<any>(this.listNames.employees, byQuery);
        if (byEmps[0]) relations.assignedByEmployee = byEmps[0];
      }

      if (objective.ManagerEmployeeId) {
        const mgrQuery = `$filter=EmployeeEmail/Id eq ${objective.ManagerEmployeeId}&$expand=EmployeeEmail&$top=1`;
        const mgrs = await this.fetchItems<any>(this.listNames.employees, mgrQuery);
        if (mgrs[0]) relations.managerEmployee = mgrs[0];
      }

      const taskQuery = `$filter=ObjectiveId eq ${id} and IsActive eq true`;
      relations.tasks = await this.fetchItems<ITask>(this.listNames.tasks, taskQuery);
      relations.childrenCount = { tasks: relations.tasks.length };
    } catch (e) {
      console.log('Error loading objective relations:', e);
    }

    return relations;
  }

  public async createObjective(data: Partial<IObjective>): Promise<number | null> {
    const toBool = (val: any): boolean => val === true || val === 1 || val === '1' || val === 'Yes' || val === 'True';
    const toYesNo = (val: any): string => toBool(val) ? 'Yes' : 'No';

    const objectiveData: any = {
      Title: data.Title,
      Cycles: data.CycleId ? data.CycleId.toString() : undefined,
      ObjectiveSource: data.ObjectiveSource || 'Standalone',
      ObjectiveType: data.ObjectiveType || 'Individual',
      Category: data.Category || 'Performance',
      Priority: data.Priority || 'Medium',
      Status: data.Status || 'Not Started',
      DueDate: data.DueDate,
      ProgressPercent: data.ProgressPercent || 0,
      CanCascadeTasks: toYesNo(data.CanCascadeTasks) || 'True',
      IsDepartmentObjective: toYesNo(data.IsDepartmentObjective),
      IsLinkedToDeliverable: toYesNo(data.IsLinkedToDeliverable),
      IsOverdue: 'False',
      IsActive: 'True'
    };

    if (data.ObjectiveCode) objectiveData.ObjectiveCode = data.ObjectiveCode;
    if (data.DeliverableId) objectiveData.DeliverableId = data.DeliverableId;
    if (data.AssignedEmployeeId) objectiveData.AssignedEmployeeId = data.AssignedEmployeeId;
    if (data.AssignedDepartmentId) objectiveData.AssignedDepartmentId = data.AssignedDepartmentId;
    if (data.AssignedByEmployeeId) objectiveData.AssignedByEmployeeId = data.AssignedByEmployeeId;
    if (data.ManagerEmployeeId) objectiveData.ManagerEmployeeId = data.ManagerEmployeeId;
    if (data.ParentObjectiveId) objectiveData.ParentObjectiveId = data.ParentObjectiveId;
    if (data.StartDate) objectiveData.StartDate = data.StartDate;
    if (data.Weight) objectiveData.Weight = data.Weight;
    if (data.Tags) objectiveData.Tags = data.Tags;
    if (data.Description) objectiveData.Description = data.Description;
    if (data.ExpectedResult) objectiveData.ExpectedResult = data.ExpectedResult;
    if (data.Comments) objectiveData.Comments = data.Comments;

    return this.createItem(this.listNames.objectives, objectiveData);
  }

  public async updateObjective(id: number, data: Partial<IObjective>): Promise<boolean> {
    const toBool = (val: any): boolean => val === true || val === 1 || val === '1' || val === 'Yes' || val === 'True';
    const toYesNo = (val: any): string => toBool(val) ? 'Yes' : 'No';
    const updateData: any = {};

    if (data.Title !== undefined) updateData.Title = data.Title;
    if (data.ObjectiveCode !== undefined) updateData.ObjectiveCode = data.ObjectiveCode;
    if (data.DeliverableId !== undefined) updateData.DeliverableId = data.DeliverableId;
    if (data.ObjectiveSource !== undefined) updateData.ObjectiveSource = data.ObjectiveSource;
    if (data.ObjectiveType !== undefined) updateData.ObjectiveType = data.ObjectiveType;
    if (data.AssignedEmployeeId !== undefined) updateData.AssignedEmployeeId = data.AssignedEmployeeId;
    if (data.AssignedDepartmentId !== undefined) updateData.AssignedDepartmentId = data.AssignedDepartmentId;
    if (data.ManagerEmployeeId !== undefined) updateData.ManagerEmployeeId = data.ManagerEmployeeId;
    if (data.ParentObjectiveId !== undefined) updateData.ParentObjectiveId = data.ParentObjectiveId;
    if (data.Category !== undefined) updateData.Category = data.Category;
    if (data.Priority !== undefined) updateData.Priority = data.Priority;
    if (data.Status !== undefined) updateData.Status = data.Status;
    if (data.StartDate !== undefined) updateData.StartDate = data.StartDate;
    if (data.DueDate !== undefined) updateData.DueDate = data.DueDate;
    if (data.ProgressPercent !== undefined) updateData.ProgressPercent = data.ProgressPercent;
    if (data.Weight !== undefined) updateData.Weight = data.Weight;
    if (data.Tags !== undefined) updateData.Tags = data.Tags;
    if (data.Description !== undefined) updateData.Description = data.Description;
    if (data.ExpectedResult !== undefined) updateData.ExpectedResult = data.ExpectedResult;
    if (data.Comments !== undefined) updateData.Comments = data.Comments;
    if (data.CanCascadeTasks !== undefined) updateData.CanCascadeTasks = toYesNo(data.CanCascadeTasks);
    if (data.IsDepartmentObjective !== undefined) updateData.IsDepartmentObjective = toYesNo(data.IsDepartmentObjective);
    if (data.IsLinkedToDeliverable !== undefined) updateData.IsLinkedToDeliverable = toYesNo(data.IsLinkedToDeliverable);
    if (data.IsOverdue !== undefined) updateData.IsOverdue = toYesNo(data.IsOverdue);
    if (data.IsActive !== undefined) updateData.IsActive = toYesNo(data.IsActive);

    return this.updateItem(this.listNames.objectives, id, updateData);
  }

  public async deleteObjective(id: number): Promise<boolean> {
    return this.deleteItem(this.listNames.objectives, id);
  }

  // ==================== TASKS ====================

  public async getAllTasks(): Promise<ITask[]> {
    const query = `$filter=IsActive eq true&$orderby=Created desc&$top=50`;
    return this.fetchItems<ITask>(this.listNames.tasks, query);
  }

  public async getTasksByDeliverable(deliverableId: number): Promise<ITask[]> {
    if (!deliverableId || deliverableId <= 0) return [];
    const query = `$filter=DeliverableId eq ${deliverableId} and IsActive eq true&$orderby=Created desc`;
    return this.fetchItems<ITask>(this.listNames.tasks, query);
  }

  public async getTasksByObjective(objectiveId: number): Promise<ITask[]> {
    if (!objectiveId || objectiveId <= 0) return [];
    const query = `$filter=ObjectiveId eq ${objectiveId} and IsActive eq true&$orderby=Created desc`;
    return this.fetchItems<ITask>(this.listNames.tasks, query);
  }

  public async getTasksByEmployee(employeeId: number): Promise<ITask[]> {
    if (!employeeId || employeeId <= 0) return [];
    const query = `$select=*,AssignedEmployee/Id,AssignedEmployee/Title,AssignedDepartment/Id,AssignedDepartment/Title,Deliverable/Id,Deliverable/Title,Objective/Id,Objective/Title,ApprovedByEmployee/Id,ApprovedByEmployee/Title&$expand=AssignedEmployee,AssignedDepartment,Deliverable,Objective,ApprovedByEmployee&$filter=AssignedEmployee/Id eq ${employeeId} and IsActive eq true&$orderby=Created desc`;
    return this.fetchItems<ITask>(this.listNames.tasks, query);
  }

  public async getTasksByDepartment(departmentId: number): Promise<ITask[]> {
    if (!departmentId || departmentId <= 0) {
      return this.fetchItems<ITask>(this.listNames.tasks, '$orderby=Created desc&$top=50');
    }
    const query = `$select=*,AssignedEmployee/Id,AssignedEmployee/Title,AssignedDepartment/Id,AssignedDepartment/Title,Deliverable/Id,Deliverable/Title,Objective/Id,Objective/Title,ApprovedByEmployee/Id,ApprovedByEmployee/Title&$expand=AssignedEmployee,AssignedDepartment,Deliverable,Objective,ApprovedByEmployee&$filter=AssignedDepartment/Id eq ${departmentId} and IsActive eq true&$orderby=Created desc`;
    return this.fetchItems<ITask>(this.listNames.tasks, query);
  }

  public async getTasksByDepartments(departmentIds: number[]): Promise<ITask[]> {
    if (!departmentIds || departmentIds.length === 0) return [];

    const uniqueIds: number[] = [];
    departmentIds.forEach((id: number) => { if (id > 0 && uniqueIds.indexOf(id) === -1) uniqueIds.push(id); });
    if (uniqueIds.length === 0) return [];

    const filter = uniqueIds.map(id => `AssignedDepartment/Id eq ${id}`).join(' or ');

    const query =
      `$select=*,AssignedEmployee/Id,AssignedEmployee/Title,AssignedDepartment/Id,AssignedDepartment/Title,Deliverable/Id,Deliverable/Title,Objective/Id,Objective/Title,ApprovedByEmployee/Id,ApprovedByEmployee/Title` +
      `&$expand=AssignedEmployee,AssignedDepartment,Deliverable,Objective,ApprovedByEmployee` +
      `&$filter=(${filter}) and IsActive eq true` +
      `&$orderby=Created desc`;

    return this.fetchItems<ITask>(this.listNames.tasks, query);
  }

  public async getTasksByCycle(cycleId: number): Promise<ITask[]> {
    if (!cycleId || cycleId <= 0) return [];
    const query = `$select=*,AssignedEmployee/Id,AssignedEmployee/Title,AssignedDepartment/Id,AssignedDepartment/Title,Deliverable/Id,Deliverable/Title,Objective/Id,Objective/Title,ApprovedByEmployee/Id,ApprovedByEmployee/Title&$expand=AssignedEmployee,AssignedDepartment,Deliverable,Objective,ApprovedByEmployee&$filter=CycleId eq ${cycleId} and IsActive eq true&$orderby=Created desc`;
    return this.fetchItems<ITask>(this.listNames.tasks, query);
  }

  public async getTaskById(id: number): Promise<ITask | null> {
    return this.fetchItemById<ITask>(this.listNames.tasks, id);
  }

  public async getTaskWithRelations(id: number): Promise<ITaskRelation | null> {
    const task = await this.getTaskById(id);
    if (!task) return null;

    const relations: ITaskRelation = { task };

    try {
      if (task.DeliverableId) {
        const delQuery = `$filter=Id eq ${task.DeliverableId}&$top=1`;
        const dels = await this.fetchItems<IDeliverable>(this.listNames.deliverables, delQuery);
        relations.deliverable = dels[0];
      }

      if (task.ObjectiveId) {
        const objQuery = `$filter=Id eq ${task.ObjectiveId}&$top=1`;
        const objs = await this.fetchItems<IObjective>(this.listNames.objectives, objQuery);
        relations.objective = objs[0];
      }

      if (task.ParentTaskId) {
        const parent = await this.getTaskById(task.ParentTaskId);
        relations.parentTask = parent || undefined;
      }

      if (task.AssignedEmployeeId) {
        const empQuery = `$filter=EmployeeEmail/Id eq ${task.AssignedEmployeeId}&$expand=EmployeeEmail&$top=1`;
        const emps = await this.fetchItems<any>(this.listNames.employees, empQuery);
        if (emps[0]) relations.assignedEmployee = emps[0];
      }

      if (task.AssignedDepartmentId) {
        const deptQuery = `$filter=Id eq ${task.AssignedDepartmentId}&$top=1`;
        const depts = await this.fetchItems<IDepartment>(this.listNames.departments, deptQuery);
        relations.assignedDepartment = depts[0];
      }

      if (task.ApprovedByEmployeeId) {
        const apprQuery = `$filter=EmployeeEmail/Id eq ${task.ApprovedByEmployeeId}&$expand=EmployeeEmail&$top=1`;
        const apprs = await this.fetchItems<any>(this.listNames.employees, apprQuery);
        if (apprs[0]) relations.approvedBy = apprs[0];
      }

      const childQuery = `$filter=ParentTaskId eq ${id} and IsActive eq true`;
      relations.childrenTasks = await this.fetchItems<ITask>(this.listNames.tasks, childQuery);
    } catch (e) {
      console.log('Error loading task relations:', e);
    }

    return relations;
  }

  public async createTask(data: Partial<ITask>): Promise<number | null> {
    const taskData: any = {
      Title: data.Title,
      TaskSource: data.TaskSource || 'Standalone',
      TaskType: data.TaskType || 'Operational',
      Priority: data.Priority || 'Medium',
      Status: data.Status || 'Pending',
      DueDate: data.DueDate,
      ProgressPercent: data.ProgressPercent || 0,
      RequiresApproval: Boolean(data.RequiresApproval),
      ApprovalStatus: data.ApprovalStatus || 'Not Required',
      IsDepartmentTask: Boolean(data.IsDepartmentTask),
      IsLinkedToDeliverable: Boolean(data.IsLinkedToDeliverable),
      IsLinkedToObjective: Boolean(data.IsLinkedToObjective),
      IsOverdue: Boolean(data.IsOverdue),
      IsActive: 'True'
    };

    if (data.CycleId) taskData.CycleId = data.CycleId;
    if (data.TaskCode) taskData.TaskCode = data.TaskCode;
    if (data.DeliverableId) taskData.DeliverableId = data.DeliverableId;
    if (data.ObjectiveId) taskData.ObjectiveId = data.ObjectiveId;
    if (data.ParentTaskId) taskData.ParentTaskId = data.ParentTaskId;
    if (data.AssignedEmployeeId) taskData.AssignedEmployeeId = data.AssignedEmployeeId;
    if (data.AssignedDepartmentId) taskData.AssignedDepartmentId = data.AssignedDepartmentId;
    if (data.StartDate) taskData.StartDate = data.StartDate;
    if (data.CompletedDate) taskData.CompletedDate = data.CompletedDate;
    if (data.Weight) taskData.Weight = data.Weight;
    if (data.PlannedHours) taskData.PlannedHours = data.PlannedHours;
    if (data.WorkedHours) taskData.WorkedHours = data.WorkedHours;
    if (data.Description) taskData.Description = data.Description;
    if (data.ExpectedResult) taskData.ExpectedResult = data.ExpectedResult;
    if (data.Comments) taskData.Comments = data.Comments;
    if (data.EvidenceLink) taskData.EvidenceLink = data.EvidenceLink;
    if (data.ApprovedByEmployeeId) taskData.ApprovedByEmployeeId = data.ApprovedByEmployeeId;
    if (data.ApprovalDate) taskData.ApprovalDate = data.ApprovalDate;

    console.log('Task data to create:', JSON.stringify(taskData, null, 2));
    return this.createItem(this.listNames.tasks, taskData);
  }

  public async updateTask(id: number, data: Partial<ITask>): Promise<boolean> {
    const toBool = (val: any): boolean => val === true || val === 1 || val === '1' || val === 'Yes' || val === 'True';
    const toYesNo = (val: any): string => toBool(val) ? 'True' : 'False';
    const updateData: any = {};

    if (data.Title !== undefined) updateData.Title = data.Title;
    if (data.TaskCode !== undefined) updateData.TaskCode = data.TaskCode;
    if (data.CycleId !== undefined) updateData.CycleId = data.CycleId;
    if (data.DeliverableId !== undefined) updateData.DeliverableId = data.DeliverableId;
    if (data.ObjectiveId !== undefined) updateData.ObjectiveId = data.ObjectiveId;
    if (data.ParentTaskId !== undefined) updateData.ParentTaskId = data.ParentTaskId;
    if (data.TaskSource !== undefined) updateData.TaskSource = data.TaskSource;
    if (data.TaskType !== undefined) updateData.TaskType = data.TaskType;
    if (data.AssignedEmployeeId !== undefined) updateData.AssignedEmployeeId = data.AssignedEmployeeId;
    if (data.AssignedDepartmentId !== undefined) updateData.AssignedDepartmentId = data.AssignedDepartmentId;
    if (data.Priority !== undefined) updateData.Priority = data.Priority;
    if (data.Status !== undefined) updateData.Status = data.Status;
    if (data.StartDate !== undefined) updateData.StartDate = data.StartDate;
    if (data.DueDate !== undefined) updateData.DueDate = data.DueDate;
    if (data.CompletedDate !== undefined) updateData.CompletedDate = data.CompletedDate;
    if (data.ProgressPercent !== undefined) updateData.ProgressPercent = data.ProgressPercent;
    if (data.Weight !== undefined) updateData.Weight = data.Weight;
    if (data.PlannedHours !== undefined) updateData.PlannedHours = data.PlannedHours;
    if (data.WorkedHours !== undefined) updateData.WorkedHours = data.WorkedHours;
    if (data.Description !== undefined) updateData.Description = data.Description;
    if (data.ExpectedResult !== undefined) updateData.ExpectedResult = data.ExpectedResult;
    if (data.Comments !== undefined) updateData.Comments = data.Comments;
    if (data.EvidenceLink !== undefined) updateData.EvidenceLink = data.EvidenceLink;
    if (data.RequiresApproval !== undefined) updateData.RequiresApproval = toYesNo(data.RequiresApproval);
    if (data.ApprovalStatus !== undefined) updateData.ApprovalStatus = data.ApprovalStatus;
    if (data.ApprovedByEmployeeId !== undefined) updateData.ApprovedByEmployeeId = data.ApprovedByEmployeeId;
    if (data.ApprovalDate !== undefined) updateData.ApprovalDate = data.ApprovalDate;
    if (data.IsDepartmentTask !== undefined) updateData.IsDepartmentTask = toYesNo(data.IsDepartmentTask);
    if (data.IsLinkedToDeliverable !== undefined) updateData.IsLinkedToDeliverable = toYesNo(data.IsLinkedToDeliverable);
    if (data.IsLinkedToObjective !== undefined) updateData.IsLinkedToObjective = toYesNo(data.IsLinkedToObjective);
    if (data.IsOverdue !== undefined) updateData.IsOverdue = toYesNo(data.IsOverdue);
    if (data.IsActive !== undefined) updateData.IsActive = data.IsActive ? 'True' : 'False';

    return this.updateItem(this.listNames.tasks, id, updateData);
  }

  public async deleteTask(id: number): Promise<boolean> {
    return this.deleteItem(this.listNames.tasks, id);
  }

  public async approveTask(id: number, approvedByEmployeeId: number): Promise<boolean> {
    return this.updateTask(id, {
      ApprovalStatus: 'Approved',
      ApprovedByEmployeeId: approvedByEmployeeId,
      ApprovalDate: new Date().toISOString()
    });
  }

  public async rejectTask(id: number, approvedByEmployeeId: number): Promise<boolean> {
    return this.updateTask(id, {
      ApprovalStatus: 'Rejected',
      ApprovedByEmployeeId: approvedByEmployeeId,
      ApprovalDate: new Date().toISOString()
    });
  }

  // ==================== LOOKUPS ====================

  public async getDepartments(): Promise<IDepartment[]> {
    const query = `$filter=IsActive eq true&$orderby=Title`;
    return this.fetchItems<IDepartment>(this.listNames.departments, query);
  }

  public async getEmployees(): Promise<IEmployee[]> {
    const query = `$filter=IsActive eq true&$orderby=Title`;
    return this.fetchItems<IEmployee>(this.listNames.employees, query);
  }

  public async getCycles(): Promise<IEvaluationCycle[]> {
    const query = `$filter=IsActive eq true&$orderby=Year desc, StartDate desc`;
    return this.fetchItems<IEvaluationCycle>(this.listNames.cycles, query);
  }

  public async getActiveCycles(): Promise<IEvaluationCycle[]> {
    const query = `$filter=Status eq 'Active' and IsActive eq true&$orderby=Year desc`;
    return this.fetchItems<IEvaluationCycle>(this.listNames.cycles, query);
  }
}

export default PerformanceService;
