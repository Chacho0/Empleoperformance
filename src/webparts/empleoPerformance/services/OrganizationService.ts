import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import {
  IEmployee,
  IDepartment,
  IEmployeeDepartment,
  IEmployeeDepartmentRelation,
  IOrganizationalContext,
  IUserPermissions,
  ISubordinateInfo,
  IEvaluationCycle,
  UserRoleType
} from '../interfaces/IOrganization';

export interface IOrganizationServiceConfig {
  employees: string;
  departments: string;
  employeeDepartments: string;
  cycles?: string;
}

export class OrganizationService {
  private spHttpClient: SPHttpClient;
  private webUrl: string;
  private listNames: IOrganizationServiceConfig;
  private entityTypes: Map<string, string> = new Map();

  constructor(spHttpClient: SPHttpClient, webUrl: string, config?: IOrganizationServiceConfig) {
    this.spHttpClient = spHttpClient;
    this.webUrl = webUrl;
    this.listNames = {
      employees: config?.employees || 'PM_Employee',
      departments: config?.departments || 'PM_Departments',
      employeeDepartments: config?.employeeDepartments || 'PM_EmployeeDepartments',
      cycles: config?.cycles || 'PM_Cycles'
    };
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

  public updateListNames(config: IOrganizationServiceConfig): void {
    this.listNames = { ...this.listNames, ...config };
  }

  private getApiUrl(listName: string, query?: string): string {
    const baseUrl = `${this.webUrl}/_api/web/lists/getbytitle('${listName}')/items`;
    return query ? `${baseUrl}?${query}` : baseUrl;
  }

  private async fetchItems<T>(listName: string, query?: string): Promise<T[]> {
    const url = this.getApiUrl(listName, query);
    const response: SPHttpClientResponse = await this.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1
    );
    const data = await response.json();
    return data.value || [];
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

  public async getCurrentUserEmail(): Promise<string> {
    const url = `${this.webUrl}/_api/web/currentuser`;
    const response: SPHttpClientResponse = await this.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1
    );
    const user = await response.json();
    return user.Email || user.LoginName || '';
  }

  public async getCurrentUserId(): Promise<number> {
    const url = `${this.webUrl}/_api/web/currentuser`;
    const response: SPHttpClientResponse = await this.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1
    );
    const user = await response.json();
    return user.Id || 0;
  }

  public async getEmployeeByUserEmail(email: string): Promise<IEmployee | null> {
    const query = `$filter=EmployeeEmail/EMail eq '${encodeURIComponent(email)}' and IsActive eq true&$expand=EmployeeEmail&$top=1`;
    const employees = await this.fetchItems<IEmployee>(this.listNames.employees, query);
    if (employees.length > 0) {
      console.log('Employee loaded - CanManageCycles:', employees[0].CanManageCycles, 'type:', typeof employees[0].CanManageCycles);
    }
    return employees.length > 0 ? employees[0] : null;
  }

  public async getEmployeeByUserId(userId: number): Promise<IEmployee | null> {
    const query = `$filter=EmployeeEmail/Id eq ${userId} and IsActive eq true&$top=1`;
    const employees = await this.fetchItems<IEmployee>(this.listNames.employees, query);
    if (employees.length === 0) return null;
    const emp = employees[0];
    if (emp.EmployeeEmail && typeof emp.EmployeeEmail === 'object') {
      const emailObj = emp.EmployeeEmail as any;
      emp.EmployeeEmail = emailObj.EMail || emailObj.Title || '';
    }
    return emp;
  }

  public async getAllEmployees(): Promise<IEmployee[]> {
    const query = '$filter=IsActive eq true&$top=50';
    const employees = await this.fetchItems<IEmployee>(this.listNames.employees, query);
    return employees.map(emp => {
      if (emp.EmployeeEmail && typeof emp.EmployeeEmail === 'object') {
        const emailObj = emp.EmployeeEmail as any;
        emp.EmployeeEmail = emailObj.EMail || emailObj.Title || '';
      }
      return emp;
    });
  }

  public async getEmployeeById(id: number): Promise<IEmployee | null> {
    return this.fetchItemById<IEmployee>(this.listNames.employees, id);
  }

  public async getEmployeeDepartments(employeeId: number): Promise<IEmployeeDepartment[]> {
    const query = `$filter=EmployeeId eq ${employeeId} and IsActive eq true&$orderby=IsPrimaryDepartment desc`;
    const items = await this.fetchItems<IEmployeeDepartment>(this.listNames.employeeDepartments, query);
    console.log('EmployeeDepartments for employee', employeeId, ':', items.length, 'items');
    return items;
  }

  public async getDepartmentEmployees(departmentId: number): Promise<IEmployeeDepartment[]> {
    const query = `$filter=DepartmentId eq ${departmentId} and IsActive eq true`;
    return this.fetchItems<IEmployeeDepartment>(this.listNames.employeeDepartments, query);
  }

  public async getDepartmentEmployeesByDepartments(departmentIds: number[]): Promise<IEmployeeDepartment[]> {
    if (departmentIds.length === 0) return [];
    const idsFilter = departmentIds.map(id => `DepartmentId eq ${id}`).join(' or ');
    const query = `$filter=(${idsFilter}) and IsActive eq true`;
    return this.fetchItems<IEmployeeDepartment>(this.listNames.employeeDepartments, query);
  }

  public async getDepartmentById(id: number): Promise<IDepartment | null> {
    return this.fetchItemById<IDepartment>(this.listNames.departments, id);
  }

  public async getAllActiveDepartments(): Promise<IDepartment[]> {
    const query = '$filter=IsActive eq true&$orderby=Title asc';
    return this.fetchItems<IDepartment>(this.listNames.departments, query);
  }

  public async getAllDepartments(): Promise<IDepartment[]> {
    const query = '$orderby=Title asc&$top=500';
    return this.fetchItems<IDepartment>(this.listNames.departments, query);
  }

  public async getSubordinatesByReportsTo(reportsToEmployeeId: number): Promise<IEmployeeDepartment[]> {
    const query = `$filter=ReportsToEmployeeId eq ${reportsToEmployeeId} and IsActive eq true`;
    return this.fetchItems<IEmployeeDepartment>(this.listNames.employeeDepartments, query);
  }

  public async getDepartmentEmployeesByDepartment(departmentId: number): Promise<IEmployeeDepartmentRelation[]> {
    const empDepts = await this.getDepartmentEmployees(departmentId);
    const relations: IEmployeeDepartmentRelation[] = [];

    for (const ed of empDepts) {
      const employee = await this.getEmployeeById(ed.EmployeeId);
      if (!employee) continue;

      const department = await this.getDepartmentById(ed.DepartmentId);
      if (!department) continue;

      let reportsToEmployee: IEmployee | null = null;
      if (ed.ReportsToEmployeeId) {
        reportsToEmployee = await this.getEmployeeById(ed.ReportsToEmployeeId);
      }

      relations.push({
        employeeDepartment: ed,
        employee,
        department,
        reportsToEmployee
      });
    }

    return relations;
  }

  public async getDepartmentRelations(departmentId: number): Promise<IEmployeeDepartmentRelation[]> {
    const empDepts = await this.getDepartmentEmployees(departmentId);
    const relations: IEmployeeDepartmentRelation[] = [];

    for (const ed of empDepts) {
      const employee = await this.getEmployeeById(ed.EmployeeId);
      if (!employee) continue;

      const department = await this.getDepartmentById(ed.DepartmentId);
      if (!department) continue;

      let reportsToEmployee: IEmployee | null = null;
      if (ed.ReportsToEmployeeId) {
        reportsToEmployee = await this.getEmployeeById(ed.ReportsToEmployeeId);
      }

      relations.push({
        employeeDepartment: ed,
        employee,
        department,
        reportsToEmployee
      });
    }

    return relations;
  }

  public async getEmployeeDepartmentsByDepartment(departmentId: number): Promise<IEmployeeDepartmentRelation[]> {
    return this.getDepartmentRelations(departmentId);
  }

  public async getDepartmentManager(departmentId: number): Promise<IEmployee | null> {
    const department = await this.getDepartmentById(departmentId);
    if (!department || !department.DepartmentManagerId) return null;
    return this.getEmployeeById(department.DepartmentManagerId);
  }

  public async buildOrganizationalContext(currentUserEmail: string): Promise<IOrganizationalContext | null> {
    const currentEmployee = await this.getEmployeeByUserEmail(currentUserEmail);
    if (!currentEmployee) return null;
    return this.buildContextFromEmployee(currentEmployee);
  }

  public async buildOrganizationalContextById(userId: number): Promise<IOrganizationalContext | null> {
    const currentEmployee = await this.getEmployeeByUserId(userId);
    if (!currentEmployee) return null;
    return this.buildContextFromEmployee(currentEmployee);
  }

  private async buildContextFromEmployee(currentEmployee: IEmployee): Promise<IOrganizationalContext> {
    const departmentRelations = await this.buildDepartmentRelations(currentEmployee.Id);
    const primaryRelation = departmentRelations.find(r => r.employeeDepartment.IsPrimaryDepartment) || null;
    
    const reportsTo = await this.resolveReportsTo(currentEmployee.Id);
    const subordinates = await this.resolveSubordinates(currentEmployee.Id);
    const isDeptManager = departmentRelations.some(r => r.employeeDepartment.RoleInDepartment === 'DepartmentManager');
    const managedDepts = departmentRelations
      .filter(r => r.employeeDepartment.RoleInDepartment === 'DepartmentManager')
      .map(r => r.department);
    
    const primaryDept = primaryRelation?.department || null;
    let executiveLeader: IEmployee | null = null;
    if (primaryDept?.ReportsToExecutiveId) {
      executiveLeader = await this.getEmployeeById(primaryDept.ReportsToExecutiveId);
    }

    const permissions = this.aggregatePermissions(departmentRelations, currentEmployee);

    return {
      currentEmployee,
      departmentRelations,
      primaryDepartmentRelation: primaryRelation,
      allDepartments: departmentRelations.map(r => r.department),
      reportsTo,
      subordinates,
      isDepartmentManager: isDeptManager,
      isPrimaryDepartmentManager: primaryRelation?.employeeDepartment.IsDepartmentManager || false,
      managedDepartments: managedDepts,
      executiveLeader,
      permissions
    };
  }

  private async buildDepartmentRelations(employeeId: number): Promise<IEmployeeDepartmentRelation[]> {
    const employeeDepartments = await this.getEmployeeDepartments(employeeId);
    console.log('Building department relations for employee', employeeId, 'RoleInDepartment values:', employeeDepartments.map(ed => ed.RoleInDepartment));
    const relations: IEmployeeDepartmentRelation[] = [];

    for (const ed of employeeDepartments) {
      const department = await this.getDepartmentById(ed.DepartmentId);
      if (!department) continue;

      const employee = await this.getEmployeeById(employeeId);
      if (!employee) continue;

      let reportsToEmployee: IEmployee | null = null;
      if (ed.ReportsToEmployeeId) {
        reportsToEmployee = await this.getEmployeeById(ed.ReportsToEmployeeId);
      }

      relations.push({
        employeeDepartment: ed,
        employee,
        department,
        reportsToEmployee
      });
    }

    return relations;
  }

  private async resolveReportsTo(employeeId: number): Promise<IEmployee | null> {
    const employeeDepts = await this.getEmployeeDepartments(employeeId);
    const primaryRelation = employeeDepts.find(ed => ed.IsPrimaryDepartment);
    
    if (primaryRelation?.ReportsToEmployeeId) {
      return this.getEmployeeById(primaryRelation.ReportsToEmployeeId);
    }
    return null;
  }

  private async resolveSubordinates(employeeId: number): Promise<IEmployee[]> {
    const subordinateRelations = await this.getSubordinatesByReportsTo(employeeId);
    const subordinates: IEmployee[] = [];
    const seenIds = new Set<number>();

    for (const sr of subordinateRelations) {
      if (!seenIds.has(sr.EmployeeId)) {
        seenIds.add(sr.EmployeeId);
        const emp = await this.getEmployeeById(sr.EmployeeId);
        if (emp) subordinates.push(emp);
      }
    }

    return subordinates;
  }

  private aggregatePermissions(
    departmentRelations: IEmployeeDepartmentRelation[],
    currentEmployee: IEmployee
  ): IUserPermissions {
    const toBool = (val: any): boolean => val === true || val === 1 || val === '1' || val === 'Yes';

    console.log('aggregatePermissions - currentEmployee.CanManageCycles:', currentEmployee.CanManageCycles, 'type:', typeof currentEmployee.CanManageCycles);
    console.log('aggregatePermissions - toBool result:', toBool(currentEmployee.CanManageCycles));

    const aggregated: IUserPermissions = {
      canViewDepartment: false,
      canAssignObjectives: false,
      canAssignTasks: false,
      canViewReports: false,
      canApprove: false,
      canManageCycles: toBool(currentEmployee.CanManageCycles),
      canManageSettings: toBool(currentEmployee.CanManageSettings)
    };

    console.log('aggregatePermissions - final canManageCycles:', aggregated.canManageCycles);

    for (const relation of departmentRelations) {
      const perms = relation.employeeDepartment;
      aggregated.canViewDepartment = aggregated.canViewDepartment || toBool(perms.CanViewDepartment);
      aggregated.canAssignObjectives = aggregated.canAssignObjectives || toBool(perms.CanAssignObjectives);
      aggregated.canAssignTasks = aggregated.canAssignTasks || toBool(perms.CanAssignTasks);
      aggregated.canViewReports = aggregated.canViewReports || toBool(perms.CanViewReports);
      aggregated.canApprove = aggregated.canApprove || toBool(perms.CanApprove);
    }

    return aggregated;
  }

  public async getSubordinatesWithDetails(
    employeeId: number
  ): Promise<ISubordinateInfo[]> {
    const subordinateRelations = await this.getSubordinatesByReportsTo(employeeId);
    const subordinateInfos: ISubordinateInfo[] = [];

    for (const sr of subordinateRelations) {
      const emp = await this.getEmployeeById(sr.EmployeeId);
      if (!emp) continue;

      const dept = await this.getDepartmentById(sr.DepartmentId);
      if (!dept) continue;

      const subordinatesCount = (await this.getSubordinatesByReportsTo(sr.EmployeeId)).length;

      subordinateInfos.push({
        employee: emp,
        departmentRelation: {
          employeeDepartment: sr,
          employee: emp,
          department: dept,
          reportsToEmployee: await this.getEmployeeById(employeeId)
        },
        subordinatesCount
      });
    }

    return subordinateInfos;
  }

  public getUserRole(context: IOrganizationalContext): UserRoleType {
    const canManageSettings = context.permissions.canManageSettings === true;
    const canManageCycles = context.permissions.canManageCycles === true;
    const isManager = context.isDepartmentManager === true || context.managedDepartments.length > 0;

    if (canManageSettings || canManageCycles) {
      return 'admin';
    }

    if (isManager) {
      const hasReports = (context.subordinates?.length ?? 0) > 0;
      const hasExecutiveRole = context.executiveLeader !== null;
      return hasReports && hasExecutiveRole ? 'executive' : 'department_manager';
    }

    return 'normal';
  }

  public canCreateObjectives(context: IOrganizationalContext): boolean {
    return context.permissions.canAssignObjectives;
  }

  public canCreateTasks(context: IOrganizationalContext): boolean {
    return context.permissions.canAssignTasks;
  }

  public canViewTeamReports(context: IOrganizationalContext): boolean {
    return context.permissions.canViewReports || context.isDepartmentManager;
  }

  public canApproveItems(context: IOrganizationalContext): boolean {
    return context.permissions.canApprove || context.isDepartmentManager;
  }

  public getDepartmentTeam(departmentId: number): Promise<IEmployeeDepartmentRelation[]> {
    return this.buildDepartmentTeamRelations(departmentId);
  }

  private async buildDepartmentTeamRelations(departmentId: number): Promise<IEmployeeDepartmentRelation[]> {
    const deptEmployees = await this.getDepartmentEmployees(departmentId);
    const relations: IEmployeeDepartmentRelation[] = [];

    for (const de of deptEmployees) {
      const emp = await this.getEmployeeById(de.EmployeeId);
      const dept = await this.getDepartmentById(departmentId);
      if (!emp || !dept) continue;

      let reportsTo: IEmployee | null = null;
      if (de.ReportsToEmployeeId) {
        reportsTo = await this.getEmployeeById(de.ReportsToEmployeeId);
      }

      relations.push({
        employeeDepartment: de,
        employee: emp,
        department: dept,
        reportsToEmployee: reportsTo
      });
    }

    return relations;
  }

  public async getAllCycles(): Promise<IEvaluationCycle[]> {
    const listName = this.listNames.cycles || 'PM_EvaluationCycles';
    const query = '$orderby=Year desc,StartDate desc&$top=50';
    return this.fetchItems<IEvaluationCycle>(listName, query);
  }

  public async getActiveCycles(): Promise<IEvaluationCycle[]> {
    const listName = this.listNames.cycles || 'PM_EvaluationCycles';
    const query = "$filter=Status eq 'Active'&$orderby=StartDate desc";
    return this.fetchItems<IEvaluationCycle>(listName, query);
  }

  public async getCycleById(id: number): Promise<IEvaluationCycle | null> {
    const listName = this.listNames.cycles || 'PM_EvaluationCycles';
    return this.fetchItemById<IEvaluationCycle>(listName, id);
  }

  public async createCycle(cycle: Partial<IEvaluationCycle>): Promise<IEvaluationCycle | null> {
    const listName = this.listNames.cycles || 'PM_Cycles';
    const url = this.getApiUrl(listName);

    const body = {
      Title: cycle.Title || '',
      Year: cycle.Year || new Date().getFullYear(),
      PeriodType: cycle.PeriodType || 'Annual',
      StartDate: cycle.StartDate || new Date().toISOString(),
      EndDate: cycle.EndDate || new Date().toISOString(),
      Status: cycle.Status || 'Draft',
      Description: cycle.Description || '',
      IsActive: cycle.IsActive ?? true
    };

    try {
      const response: SPHttpClientResponse = await this.spHttpClient.post(
        url,
        SPHttpClient.configurations.v1,
        {
          body: JSON.stringify(body),
          headers: {
            'Accept': 'application/json',
            'Content-Type': 'application/json'
          }
        }
      );

      if (response.ok) {
        return await response.json();
      }

      const errorText = await response.text();
      console.error('Error creating cycle:', response.status, errorText);
      return null;
    } catch (error) {
      console.error('Error creating cycle:', error);
      return null;
    }
  }

  public async updateCycle(id: number, cycle: Partial<IEvaluationCycle>): Promise<boolean> {
    const listName = this.listNames.cycles || 'PM_EvaluationCycles';
    const url = `${this.getApiUrl(listName)}(${id})`;
    const body: Record<string, unknown> = {};

    if (cycle.Title !== undefined) body.Title = cycle.Title;
    if (cycle.Year !== undefined) body.Year = cycle.Year;
    if (cycle.PeriodType !== undefined) body.PeriodType = cycle.PeriodType;
    if (cycle.StartDate !== undefined) body.StartDate = cycle.StartDate;
    if (cycle.EndDate !== undefined) body.EndDate = cycle.EndDate;
    if (cycle.Status !== undefined) body.Status = cycle.Status;
    if (cycle.Description !== undefined) body.Description = cycle.Description;
    if (cycle.IsActive !== undefined) body.IsActive = cycle.IsActive;

    try {
      const response: SPHttpClientResponse = await this.spHttpClient.post(
        url,
        SPHttpClient.configurations.v1,
        {
          body: JSON.stringify(body),
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
        console.error('Error updating cycle:', response.status, errorText);
      }

      return response.ok;
    } catch (error) {
      console.error('Error updating cycle:', error);
      return false;
    }
  }

  public async closeCycle(id: number): Promise<boolean> {
    return this.updateCycle(id, { Status: 'Closed', IsActive: false });
  }
}

export default OrganizationService;
