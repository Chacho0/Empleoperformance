import * as React from 'react';
import { SPHttpClient } from '@microsoft/sp-http';
import './styles.css';
import {
  IOrganizationalContext,
  ISubordinateInfo,
  IEmployee,
  IDepartment,
  IEmployeeDepartment,
  IEmployeeDepartmentRelation,
  IEvaluationCycle,
  UserRoleType,
  IDeliverable,
  IObjective,
  ITask
} from '../../interfaces/IOrganization';
import { OrganizationService, IOrganizationServiceConfig } from '../../services/OrganizationService';
import { PerformanceService, IPerformanceServiceConfig } from '../../services/PerformanceService';

export interface IEmployeeDashboardProps {
  spHttpClient: SPHttpClient;
  webUrl: string;
  listEmployees?: string;
  listDepartments?: string;
  listEmployeeDepartments?: string;
  listCycles?: string;
  listDeliverables?: string;
  listObjectives?: string;
  listTasks?: string;
  enableDesignMode?: boolean;
  onError?: (error: Error) => void;
}

export interface IEmployeeDashboardState {
  context: IOrganizationalContext | null;
  subordinates: ISubordinateInfo[];
  departmentEmployees: IEmployee[];
  deptEmployeeRelations: IEmployeeDepartment[];
  cycles: IEvaluationCycle[];
  deliverables: IDeliverable[];
  objectives: IObjective[];
  tasks: ITask[];
  allDepartments: IDepartment[];
  loading: boolean;
  error: string | null;
  userRole: UserRoleType;
  searchText: string;
  selectedDepartment: string;
  selectedDepartments: number[];
  showDeliverableModal: boolean;
  showObjectiveModal: boolean;
  showTaskModal: boolean;
  showEmployeeModal: boolean;
  showCycleModal: boolean;
  showCycleListModal: boolean;
  showDeliverableDetailModal: boolean;
  showObjectiveDetailModal: boolean;
  showObjectiveEditModal: boolean;
  selectedEmployee: ISubordinateInfo | null;
  selectedCycle: IEvaluationCycle | null;
  selectedDeliverable: IDeliverable | null;
  selectedObjective: IObjective | null;
  isDesignMode: boolean;
  designModeRole: 'employee' | 'manager' | 'director';
  activeTab: 'team' | 'deliverables' | 'objectives' | 'tasks';
  isSaving: boolean;
  isManagerOfDeliverables: boolean;
}

const createEmployeeData = (
  id: number,
  title: string,
  email: string,
  position: string,
  level: number,
  appRole: string,
  canManageCycles: boolean = false,
  canManageSettings: boolean = false
): IEmployee => ({
  Id: id,
  Title: title,
  EmployeeEmail: email,
  PositionTitle: position,
  PositionLevel: level,
  AppRole: appRole,
  CanManageCycles: canManageCycles,
  CanManageSettings: canManageSettings,
  IsActive: true
});

const createDepartmentData = (
  id: number,
  title: string,
  code: string,
  managerId: number | null = null,
  reportsToId: number | null = null
): IDepartment => ({
  Id: id,
  Title: title,
  DepartmentCode: code,
  DepartmentManagerId: managerId,
  ReportsToExecutiveId: reportsToId,
  IsActive: true
});

const createRelationData = (
  id: number,
  empId: number,
  deptId: number,
  role: string,
  reportsTo: number | null,
  isPrimary: boolean,
  isManager: boolean,
  canView: boolean,
  canAssignObj: boolean,
  canAssignTask: boolean,
  canViewReports: boolean,
  canApprove: boolean
): IEmployeeDepartment => ({
  Id: id,
  Title: `Relation ${id}`,
  EmployeeId: empId,
  DepartmentId: deptId,
  RoleInDepartment: role,
  ReportsToEmployeeId: reportsTo,
  IsPrimaryDepartment: isPrimary,
  IsDepartmentManager: isManager,
  CanViewDepartment: canView,
  CanAssignObjectives: canAssignObj,
  CanAssignTasks: canAssignTask,
  CanViewReports: canViewReports,
  CanApprove: canApprove,
  IsActive: true
});

const createSubordinateInfo = (
  emp: IEmployee,
  dept: IDepartment,
  relation: IEmployeeDepartment,
  reportsTo: IEmployee | null,
  subCount: number
): ISubordinateInfo => ({
  employee: emp,
  departmentRelation: {
    employeeDepartment: relation,
    employee: emp,
    department: dept,
    reportsToEmployee: reportsTo
  },
  subordinatesCount: subCount
});

const createDeliverableData = (
  id: number,
  title: string,
  deptId: number,
  assignedById: number,
  deptManagerId: number,
  cycleId: number,
  category: string,
  priority: string,
  status: string,
  progress: number,
  dueDate: string,
  canCreateObj: boolean = true,
  canCreateTasks: boolean = true
): IDeliverable => ({
  Id: id,
  Title: title,
  AssignedDepartmentId: deptId,
  AssignedByEmployeeId: assignedById,
  DepartmentManagerEmployeeId: deptManagerId,
  CycleId: cycleId,
  Category: category as any,
  Priority: priority as any,
  Status: status as any,
  ProgressPercent: progress,
  DueDate: dueDate,
  CanCreateChildObjectives: canCreateObj,
  CanCreateChildTasks: canCreateTasks,
  IsActive: true
});

const createObjectiveData = (
  id: number,
  title: string,
  deliverableId: number | undefined,
  source: string,
  type: string,
  cycleId: number,
  category: string,
  priority: string,
  status: string,
  progress: number,
  dueDate: string,
  assignedEmpId?: number,
  assignedDeptId?: number
): IObjective => ({
  Id: id,
  Title: title,
  DeliverableId: deliverableId,
  ObjectiveSource: source as any,
  ObjectiveType: type as any,
  CycleId: cycleId,
  Category: category as any,
  Priority: priority as any,
  Status: status as any,
  ProgressPercent: progress,
  DueDate: dueDate,
  AssignedEmployeeId: assignedEmpId,
  AssignedDepartmentId: assignedDeptId,
  CanCascadeTasks: true,
  IsDepartmentObjective: false,
  IsLinkedToDeliverable: !!deliverableId,
  IsOverdue: false,
  IsActive: true
});

const createTaskData = (
  id: number,
  title: string,
  objectiveId: number | undefined,
  deliverableId: number | undefined,
  source: string,
  type: string,
  cycleId: number,
  priority: string,
  status: string,
  progress: number,
  dueDate: string,
  assignedEmpId?: number
): ITask => ({
  Id: id,
  Title: title,
  ObjectiveId: objectiveId,
  DeliverableId: deliverableId,
  TaskSource: source as any,
  TaskType: type as any,
  CycleId: cycleId,
  Priority: priority as any,
  Status: status as any,
  ProgressPercent: progress,
  DueDate: dueDate,
  AssignedEmployeeId: assignedEmpId,
  RequiresApproval: false,
  ApprovalStatus: 'Not Required' as any,
  IsDepartmentTask: false,
  IsLinkedToDeliverable: !!deliverableId,
  IsLinkedToObjective: !!objectiveId,
  IsOverdue: false,
  IsActive: true
});

const DESIGN_MODE_DATA = {
  employee: {
    currentEmployee: createEmployeeData(1, 'John Smith', 'jsmith@company.com', 'Data Analyst', 2, 'Employee'),
    department: createDepartmentData(1, 'Technology', 'TECH-001'),
    departmentRelation: createRelationData(1, 1, 1, 'Analyst', 10, true, false, true, false, false, false, false),
    reportsTo: createEmployeeData(10, 'Carlos Mendoza', 'cmendoza@company.com', 'Technology Manager', 5, 'DepartmentManager'),
    subordinates: [] as ISubordinateInfo[],
    permissions: {
      canViewDepartment: true,
      canAssignObjectives: false,
      canAssignTasks: false,
      canViewReports: false,
      canApprove: false,
      canManageCycles: false,
      canManageSettings: false
    },
    role: 'normal' as UserRoleType,
    allDepartments: [
      createDepartmentData(1, 'Technology', 'TECH-001'),
      createDepartmentData(2, 'Human Resources', 'HR-001'),
      createDepartmentData(3, 'Finance', 'FIN-001'),
      createDepartmentData(4, 'Marketing', 'MKT-001'),
    ]
  },
  manager: {
    currentEmployee: createEmployeeData(10, 'Carlos Mendoza', 'cmendoza@company.com', 'Technology Manager', 5, 'DepartmentManager'),
    department: createDepartmentData(1, 'Technology', 'TECH-001', 10),
    departmentRelation: createRelationData(10, 10, 1, 'Manager', null, true, true, true, true, true, true, true),
    reportsTo: createEmployeeData(20, 'Maria Rodriguez CEO', 'mrodriguez@company.com', 'Chief Executive Officer', 7, 'Director'),
    subordinates: [
      createSubordinateInfo(
        createEmployeeData(1, 'John Smith', 'jsmith@company.com', 'Data Analyst', 2, 'Employee'),
        createDepartmentData(1, 'Technology', 'TECH-001'),
        createRelationData(1, 1, 1, 'Analyst', 10, true, false, true, false, false, false, false),
        null, 0
      ),
      createSubordinateInfo(
        createEmployeeData(2, 'Emily Davis', 'edavis@company.com', 'Senior Developer', 3, 'Employee'),
        createDepartmentData(1, 'Technology', 'TECH-001'),
        createRelationData(2, 2, 1, 'Developer', 10, true, false, true, false, false, false, false),
        null, 1
      ),
      createSubordinateInfo(
        createEmployeeData(3, 'Michael Brown', 'mbrown@company.com', 'Junior Developer', 1, 'Employee'),
        createDepartmentData(1, 'Technology', 'TECH-001'),
        createRelationData(3, 3, 1, 'Developer', 10, true, false, true, false, false, false, false),
        null, 0
      ),
      createSubordinateInfo(
        createEmployeeData(4, 'Sarah Wilson', 'swilson@company.com', 'QA Engineer', 2, 'Employee'),
        createDepartmentData(1, 'Technology', 'TECH-001'),
        createRelationData(4, 4, 1, 'QA Engineer', 10, true, false, true, false, false, false, false),
        null, 0
      ),
      createSubordinateInfo(
        createEmployeeData(5, 'David Martinez', 'dmartinez@company.com', 'Senior Developer', 3, 'Employee'),
        createDepartmentData(1, 'Technology', 'TECH-001'),
        createRelationData(5, 5, 1, 'Developer', 2, true, false, true, false, false, false, false),
        null, 0
      )
    ] as ISubordinateInfo[],
    permissions: {
      canViewDepartment: true,
      canAssignObjectives: true,
      canAssignTasks: true,
      canViewReports: true,
      canApprove: true,
      canManageCycles: false,
      canManageSettings: false
    },
    role: 'department_manager' as UserRoleType,
    allDepartments: [
      createDepartmentData(1, 'Technology', 'TECH-001'),
      createDepartmentData(2, 'Human Resources', 'HR-001'),
      createDepartmentData(3, 'Finance', 'FIN-001'),
      createDepartmentData(4, 'Marketing', 'MKT-001'),
    ]
  },
  director: {
    currentEmployee: createEmployeeData(20, 'Maria Rodriguez CEO', 'mrodriguez@company.com', 'Chief Executive Officer', 7, 'Director', true, true),
    department: createDepartmentData(0, 'Executive Direction', 'DIR-001'),
    departmentRelation: createRelationData(20, 20, 0, 'Director', null, true, true, true, true, true, true, true),
    reportsTo: null,
    subordinates: [
      createSubordinateInfo(
        createEmployeeData(10, 'Carlos Mendoza', 'cmendoza@company.com', 'Technology Manager', 5, 'DepartmentManager'),
        createDepartmentData(1, 'Technology', 'TECH-001', 10, 20),
        createRelationData(10, 10, 1, 'Manager', null, true, true, true, true, true, true, true),
        null, 5
      ),
      createSubordinateInfo(
        createEmployeeData(11, 'Patricia Vega', 'pvega@company.com', 'HR Manager', 5, 'DepartmentManager'),
        createDepartmentData(2, 'Human Resources', 'HR-001', 11, 20),
        createRelationData(11, 11, 2, 'Manager', null, true, true, true, true, true, true, true),
        null, 4
      ),
      createSubordinateInfo(
        createEmployeeData(12, 'Luis Hernandez', 'lhernandez@company.com', 'Finance Manager', 5, 'DepartmentManager'),
        createDepartmentData(3, 'Finance', 'FIN-001', 12, 20),
        createRelationData(12, 12, 3, 'Manager', null, true, true, true, true, true, true, true),
        null, 3
      ),
      createSubordinateInfo(
        createEmployeeData(13, 'Sofia Morales', 'smorales@company.com', 'Marketing Manager', 5, 'DepartmentManager'),
        createDepartmentData(4, 'Marketing', 'MKT-001', 13, 20),
        createRelationData(13, 13, 4, 'Manager', null, true, true, true, true, true, true, true),
        null, 2
      )
    ] as ISubordinateInfo[],
    permissions: {
      canViewDepartment: true,
      canAssignObjectives: true,
      canAssignTasks: true,
      canViewReports: true,
      canApprove: true,
      canManageCycles: true,
      canManageSettings: true
    },
    role: 'admin' as UserRoleType,
    allDepartments: [
      createDepartmentData(1, 'Technology', 'TECH-001'),
      createDepartmentData(2, 'Human Resources', 'HR-001'),
      createDepartmentData(3, 'Finance', 'FIN-001'),
      createDepartmentData(4, 'Marketing', 'MKT-001'),
    ]
  }
};

export class EmployeeDashboard extends React.Component<IEmployeeDashboardProps, IEmployeeDashboardState> {
  private organizationService: OrganizationService | null = null;
  private performanceService: PerformanceService | null = null;

  constructor(props: IEmployeeDashboardProps) {
    super(props);
    this.state = {
      context: null,
      subordinates: [],
      departmentEmployees: [],
      deptEmployeeRelations: [],
      cycles: [],
      deliverables: [],
      objectives: [],
      tasks: [],
      allDepartments: [],
      loading: true,
      error: null,
      userRole: 'normal',
      searchText: '',
      selectedDepartment: 'all',
      selectedDepartments: [],
      showDeliverableModal: false,
      showObjectiveModal: false,
      showTaskModal: false,
      showEmployeeModal: false,
      showCycleModal: false,
      showCycleListModal: false,
      showDeliverableDetailModal: false,
      showObjectiveDetailModal: false,
      showObjectiveEditModal: false,
      selectedEmployee: null,
      selectedCycle: null,
      selectedDeliverable: null,
      selectedObjective: null,
      isDesignMode: false,
      designModeRole: 'employee',
      activeTab: 'team',
      isSaving: false,
      isManagerOfDeliverables: false
    };
  }

  public async componentDidMount(): Promise<void> {
    await this.loadData();
  }

  private async loadData(): Promise<void> {
    const { enableDesignMode } = this.props;
    
    if (enableDesignMode) {
      this.loadDesignModeData('manager');
      return;
    }

    try {
      this.setState({ loading: true, error: null, isDesignMode: false });

      const { spHttpClient, webUrl, listEmployees, listDepartments, listEmployeeDepartments, listCycles, listDeliverables, listObjectives, listTasks } = this.props;
      
      const orgConfig: IOrganizationServiceConfig = {
        employees: listEmployees || 'PM_Employee',
        departments: listDepartments || 'PM_Departments',
        employeeDepartments: listEmployeeDepartments || 'PM_EmployeeDepartments',
        cycles: listCycles || 'PM_Cycles'
      };

      const perfConfig: IPerformanceServiceConfig = {
        deliverables: listDeliverables || 'PM_Deliverable',
        objectives: listObjectives || 'PM_Objectives',
        tasks: listTasks || 'PM_Task',
        employees: listEmployees || 'PM_Employee',
        departments: listDepartments || 'PM_Departments',
        cycles: listCycles || 'PM_Cycles'
      };

      this.organizationService = new OrganizationService(spHttpClient, webUrl, orgConfig);
      this.performanceService = new PerformanceService(spHttpClient, webUrl, perfConfig);

      const currentUserId = await this.organizationService.getCurrentUserId();
      
      console.log('=== DEBUG ===');
      console.log('User ID:', currentUserId);
      
      let orgContext = await this.organizationService.buildOrganizationalContextById(currentUserId);
      
      if (!orgContext) {
        this.setState({
          loading: false,
          error: 'This user is not mapped in PM_Employee or has no organizational context.',
          isDesignMode: false
        });
        return;
      }

      const subordinates = await this.organizationService.getSubordinatesWithDetails(orgContext.currentEmployee.Id);
      const userRole = this.organizationService.getUserRole(orgContext);
      
      let cycles: IEvaluationCycle[] = [];
      let deliverables: IDeliverable[] = [];
      let objectives: IObjective[] = [];
      let tasks: ITask[] = [];
      let departmentEmployees: IEmployee[] = [];
      let deptEmployeeRelations: IEmployeeDepartment[] = [];
      let isManagerOfDeliverables = false;
      let deliverablesDeptIds: number[] = [];
      
      const isDeptManager = orgContext.departmentRelations.some(rel => rel.employeeDepartment.RoleInDepartment === 'DepartmentManager');
      const managedDeptIds = orgContext.managedDepartments.map(d => d.Id);
      
      cycles = await this.organizationService.getAllCycles();
      
      let allDepartments: IDepartment[] = [];
      if (orgContext.allDepartments.length === 0) {
        orgContext.allDepartments = await this.organizationService.getAllActiveDepartments();
      }
      allDepartments = await this.organizationService.getAllDepartments();
      
      if (this.performanceService) {
        const primaryDeptId = orgContext.primaryDepartmentRelation?.department.Id;
        const currentEmployeeId = orgContext.currentEmployee.Id;
        const isDeptManagerContext = orgContext.isDepartmentManager;
        
        console.log('Loading data - primaryDeptId:', primaryDeptId, 'employeeId:', currentEmployeeId, 'AppRole:', orgContext.currentEmployee.AppRole, 'isDeptManager:', isDeptManagerContext, 'managedDepts:', managedDeptIds);
        
        if (currentEmployeeId) {
          const mgrDeliverables = await this.performanceService.getDeliverablesByManager(currentEmployeeId);
          console.log('Deliverables by manager:', mgrDeliverables.length);
          isManagerOfDeliverables = mgrDeliverables.length > 0;
          
          if (isManagerOfDeliverables) {
            deliverables = mgrDeliverables;
            const deptIdSet = new Set<number>();
            mgrDeliverables.forEach(d => deptIdSet.add(d.AssignedDepartmentId));
            deliverablesDeptIds = Array.from(deptIdSet);
          }
        }
        
        if (primaryDeptId) {
          const deptDeliverables = await this.performanceService.getDeliverablesByDepartment(primaryDeptId);
          console.log('Deliverables by department:', deptDeliverables.length);
          deliverables = [...deliverables, ...deptDeliverables];
          deliverablesDeptIds.push(primaryDeptId);
          objectives = await this.performanceService.getObjectivesByDepartment(primaryDeptId);
          tasks = await this.performanceService.getTasksByDepartment(primaryDeptId);
        } else if (isDeptManagerContext && managedDeptIds.length > 0) {
          for (const deptId of managedDeptIds) {
            const deptDeliverables = await this.performanceService.getDeliverablesByDepartment(deptId);
            const deptObjectives = await this.performanceService.getObjectivesByDepartment(deptId);
            const deptTasks = await this.performanceService.getTasksByDepartment(deptId);
            deliverables = [...deliverables, ...deptDeliverables];
            deliverablesDeptIds.push(deptId);
            objectives = [...objectives, ...deptObjectives];
            tasks = [...tasks, ...deptTasks];
          }
        }
        
        if (currentEmployeeId) {
          const empDeliverables = await this.performanceService.getDeliverablesByEmployee(currentEmployeeId);
          console.log('Deliverables by employee:', empDeliverables.length);
          const empObjectives = await this.performanceService.getObjectivesByEmployee(currentEmployeeId);
          const empTasks = await this.performanceService.getTasksByEmployee(currentEmployeeId);
          
          const allDeliverables = [...deliverables, ...empDeliverables];
          const delMap = new Map<number, IDeliverable>();
          allDeliverables.forEach(d => delMap.set(d.Id, d));
          deliverables = Array.from(delMap.values());
          console.log('Total deliverables:', deliverables.length);
          
          const allObjectives = [...objectives, ...empObjectives];
          const objMap = new Map<number, IObjective>();
          allObjectives.forEach(o => objMap.set(o.Id, o));
          objectives = Array.from(objMap.values());
          
          const allTasks = [...tasks, ...empTasks];
          const taskMap = new Map<number, ITask>();
          allTasks.forEach(t => taskMap.set(t.Id, t));
          tasks = Array.from(taskMap.values());
        }
        
        const uniqueDeptIds: number[] = [];
        deliverablesDeptIds.forEach((id: number) => { if (id > 0 && uniqueDeptIds.indexOf(id) === -1) uniqueDeptIds.push(id); });
        
        if (uniqueDeptIds.length > 0) {
          const allEmps = await this.organizationService.getAllEmployees();
          deptEmployeeRelations = await this.organizationService.getDepartmentEmployeesByDepartments(uniqueDeptIds);
          const empIds: number[] = deptEmployeeRelations.map((de: IEmployeeDepartment) => de.EmployeeId);
          const uniqueEmpIds: number[] = [];
          empIds.forEach((id: number) => { if (uniqueEmpIds.indexOf(id) === -1) uniqueEmpIds.push(id); });
          departmentEmployees = allEmps.filter(emp => uniqueEmpIds.indexOf(emp.Id) >= 0);
          
          const deptObjectives = await this.performanceService.getObjectivesByDepartments(uniqueDeptIds);
          const deptTasks = await this.performanceService.getTasksByDepartments(uniqueDeptIds);
          
          const allObjectives = [...objectives, ...deptObjectives];
          const objMap = new Map<number, IObjective>();
          allObjectives.forEach(o => objMap.set(o.Id, o));
          objectives = Array.from(objMap.values());
          
          const allTasks = [...tasks, ...deptTasks];
          const taskMap = new Map<number, ITask>();
          allTasks.forEach(t => taskMap.set(t.Id, t));
          tasks = Array.from(taskMap.values());
        }
      }

      this.setState({
        context: orgContext,
        subordinates,
        departmentEmployees,
        deptEmployeeRelations,
        cycles,
        deliverables,
        objectives,
        tasks,
        allDepartments,
        loading: false,
        userRole,
        isDesignMode: false,
        isManagerOfDeliverables
      });
    } catch (error) {
      console.error('Error loading real data:', error);
      this.setState({
        loading: false,
        error: error instanceof Error ? error.message : 'Unknown error loading dashboard',
        isDesignMode: false
      });
    }
  }

  private loadDesignModeData(role: 'employee' | 'manager' | 'director'): void {
    const data = DESIGN_MODE_DATA[role];
    
    const demoCycles: IEvaluationCycle[] = [
      {
        Id: 1,
        Title: 'Annual Evaluation 2024',
        Year: 2024,
        PeriodType: 'Annual',
        StartDate: '2024-01-01',
        EndDate: '2024-12-31',
        Status: 'Active',
        Description: 'Annual evaluation cycle for all departments',
        IsActive: true
      },
      {
        Id: 2,
        Title: 'First Semester 2025 Evaluation',
        Year: 2025,
        PeriodType: 'Semestral',
        StartDate: '2025-01-01',
        EndDate: '2025-06-30',
        Status: 'Active',
        Description: 'First semester evaluation of the year',
        IsActive: true
      },
      {
        Id: 3,
        Title: 'Annual Evaluation 2023',
        Year: 2023,
        PeriodType: 'Annual',
        StartDate: '2023-01-01',
        EndDate: '2023-12-31',
        Status: 'Closed',
        Description: 'Closed cycle - year 2023',
        IsActive: false
      }
    ];

    const demoDeliverables: IDeliverable[] = role !== 'employee' ? [
      createDeliverableData(1, 'Digital Transformation Project', 1, 10, 10, 2, 'Project Deliverable', 'High', 'In Progress', 65, '2025-06-30'),
      createDeliverableData(2, 'Q1 Sales Target', 1, 10, 10, 2, 'Operational Deliverable', 'Critical', 'In Progress', 80, '2025-03-31'),
      createDeliverableData(3, 'Customer Satisfaction Initiative', 1, 10, 10, 2, 'Strategic Deliverable', 'Medium', 'Pending', 25, '2025-06-30'),
      createDeliverableData(4, 'Team Training Program', 2, 11, 11, 2, 'Department Deliverable', 'Low', 'Completed', 100, '2025-02-28'),
    ] : [];

    const demoObjectives: IObjective[] = role !== 'employee' ? [
      createObjectiveData(1, 'Implement new CRM system', 1, 'Deliverable', 'Team', 2, 'Operational', 'High', 'In Progress', 50, '2025-04-30', 1, 1),
      createObjectiveData(2, 'Increase customer retention', 3, 'Deliverable', 'Strategic', 2, 'Strategic', 'High', 'In Progress', 35, '2025-06-30', 2, 1),
      createObjectiveData(3, 'Complete code review process', undefined, 'Department', 'Department', 2, 'Performance', 'Medium', 'In Progress', 75, '2025-03-15', 1, 1),
      createObjectiveData(4, 'Reduce technical debt', undefined, 'Standalone', 'Individual', 2, 'Development', 'Medium', 'Pending', 10, '2025-06-30', 3, 1),
    ] : role === 'employee' ? [
      createObjectiveData(5, 'Learn new framework', undefined, 'Individual', 'Individual', 2, 'Development', 'Low', 'In Progress', 30, '2025-05-01', 1),
    ] : [];

    const demoTasks: ITask[] = [
      createTaskData(1, 'Setup development environment', 1, 1, 'Deliverable', 'Technical', 2, 'High', 'In Progress', 100, '2025-02-15', 1),
      createTaskData(2, 'Create API documentation', 1, 1, 'Deliverable', 'Technical', 2, 'Medium', 'In Progress', 60, '2025-03-01', 1),
      createTaskData(3, 'Write unit tests', 1, 1, 'Deliverable', 'Technical', 2, 'Medium', 'Pending', 0, '2025-03-15', 2),
      createTaskData(4, 'Conduct customer survey', 2, 3, 'Deliverable', 'Report', 2, 'Low', 'Completed', 100, '2025-02-28', 2),
      createTaskData(5, 'Update knowledge base', undefined, undefined, 'Standalone', 'Administrative', 2, 'Low', 'Pending', 0, '2025-03-20', 1),
    ];
    
    const context: IOrganizationalContext = {
      currentEmployee: data.currentEmployee,
      departmentRelations: [{
        employeeDepartment: data.departmentRelation,
        employee: data.currentEmployee,
        department: data.department,
        reportsToEmployee: data.reportsTo
      }],
      primaryDepartmentRelation: {
        employeeDepartment: data.departmentRelation,
        employee: data.currentEmployee,
        department: data.department,
        reportsToEmployee: data.reportsTo
      },
      allDepartments: data.department.Id === 0 ? [] : [data.department],
      reportsTo: data.reportsTo,
      subordinates: data.subordinates.map(s => s.employee),
      isDepartmentManager: data.departmentRelation.IsDepartmentManager,
      isPrimaryDepartmentManager: data.departmentRelation.IsDepartmentManager,
      managedDepartments: data.departmentRelation.IsDepartmentManager ? [data.department] : [],
      executiveLeader: null,
      permissions: data.permissions
    };

    this.setState({
      context,
      subordinates: data.subordinates,
      cycles: demoCycles,
      deliverables: demoDeliverables,
      objectives: demoObjectives,
      tasks: demoTasks,
      allDepartments: data.allDepartments || [],
      loading: false,
      error: null,
      userRole: data.role,
      isDesignMode: true,
      designModeRole: role,
      activeTab: 'team'
    });
  }

  private getRoleBadgeClass(): string {
    const { userRole } = this.state;
    switch (userRole) {
      case 'admin': return 'chipGreen';
      case 'executive': return 'chipPurple';
      case 'department_manager': return 'chipBlue';
      default: return 'chipNeutral';
    }
  }

  private getRoleLabel(): string {
    const { userRole, context } = this.state;
    if (context?.currentEmployee.PositionTitle) {
      return context.currentEmployee.PositionTitle;
    }
    switch (userRole) {
      case 'admin': return 'Chief Executive Officer';
      case 'executive': return 'Executive Leader';
      case 'department_manager': return 'Department Manager';
      default: return 'Employee';
    }
  }

  private get filteredSubordinates(): ISubordinateInfo[] {
    const { subordinates, searchText, selectedDepartment } = this.state;
    
    return subordinates.filter(sub => {
      const matchesSearch = searchText === '' || 
        sub.employee.Title.toLowerCase().includes(searchText.toLowerCase()) ||
        sub.employee.PositionTitle.toLowerCase().includes(searchText.toLowerCase());
      
      const matchesDept = selectedDepartment === 'all' || 
        sub.departmentRelation.department.Id.toString() === selectedDepartment;
      
      return matchesSearch && matchesDept;
    });
  }

  private get departmentManagers(): { employee: IEmployee; department: IDepartment; deptRelation: IEmployeeDepartment }[] {
    const { deptEmployeeRelations, departmentEmployees, allDepartments } = this.state;
    
    const managers: { employee: IEmployee; department: IDepartment; deptRelation: IEmployeeDepartment }[] = [];
    
    deptEmployeeRelations.forEach(rel => {
      if (rel.RoleInDepartment === 'DepartmentManager') {
        const employee = departmentEmployees.find(emp => emp.Id === rel.EmployeeId);
        const department = allDepartments.find(dept => dept.Id === rel.DepartmentId);
        if (employee && department) {
          managers.push({ employee, department, deptRelation: rel });
        }
      }
    });
    
    return managers;
  }

  private get totalKPI(): { objectives: number; tasks: number; completed: number; pending: number } {
    return {
      objectives: this.state.subordinates.length + 3,
      tasks: this.state.subordinates.length * 2 + 8,
      completed: Math.floor((this.state.subordinates.length * 2 + 8) * 0.65),
      pending: Math.floor((this.state.subordinates.length * 2 + 8) * 0.35)
    };
  }

  private get isProjectLeader(): boolean {
    const { context } = this.state;
    const appRole = context?.currentEmployee?.AppRole;
    return appRole === 'ProjectLeader';
  }

  private get isDepartmentManager(): boolean {
    const { context } = this.state;
    if (!context) return false;
    return context.departmentRelations.some(rel => rel.employeeDepartment.RoleInDepartment === 'DepartmentManager');
  }

  private get isDepartmentManagerInAny(): boolean {
    const { context } = this.state;
    if (!context) return false;
    return context.departmentRelations.some(rel => rel.employeeDepartment.RoleInDepartment === 'DepartmentManager');
  }

  private get hasDeliverablesForManager(): boolean {
    const { deliverables, context } = this.state;
    if (!context) return false;
    return deliverables.some(d => d.DepartmentManagerEmployeeId === context.currentEmployee.Id);
  }

  private get selectedDepartments(): number[] {
    return this.state.selectedDepartments;
  }

  private get hasSelectedDepartments(): boolean {
    return this.state.selectedDepartments.length > 0;
  }

  private canAssignObjectivesInDepartment(departmentId: number): boolean {
    const { context } = this.state;
    if (!context) return false;
    const deptRelation = context.departmentRelations.find(
      rel => rel.department.Id === departmentId && rel.employeeDepartment.RoleInDepartment === 'DepartmentManager'
    );
    if (deptRelation) return true;
    return context.departmentRelations.some(
      rel => rel.department.Id === departmentId && rel.employeeDepartment.CanAssignObjectives === true
    );
  }

  private canAssignTasksInDepartment(departmentId: number): boolean {
    const { context } = this.state;
    if (!context) return false;
    const deptRelation = context.departmentRelations.find(
      rel => rel.department.Id === departmentId && rel.employeeDepartment.RoleInDepartment === 'DepartmentManager'
    );
    if (deptRelation) return true;
    return context.departmentRelations.some(
      rel => rel.department.Id === departmentId && rel.employeeDepartment.CanAssignTasks === true
    );
  }

  private get canCreateObjective(): boolean {
    if (this.isProjectLeader) return true;
    if (!this.hasSelectedDepartments) {
      return this.canCreateObjectiveInAnyDepartment();
    }
    return this.state.selectedDepartments.some(deptId => this.canAssignObjectivesInDepartment(deptId));
  }

  private get canCreateTask(): boolean {
    if (this.isProjectLeader) return true;
    if (!this.hasSelectedDepartments) {
      return this.canCreateTaskInAnyDepartment();
    }
    return this.state.selectedDepartments.some(deptId => this.canAssignTasksInDepartment(deptId));
  }

  private canCreateObjectiveInAnyDepartment(): boolean {
    const relations = this.state.context?.departmentRelations ?? [];
    const isDeptManager = relations.some(rel => rel?.employeeDepartment?.RoleInDepartment === 'DepartmentManager');
    if (isDeptManager) return true;
    if (this.state.isManagerOfDeliverables) return true;
    return relations.some(rel => rel?.employeeDepartment?.CanAssignObjectives === true);
  }

  private canCreateTaskInAnyDepartment(): boolean {
    const relations = this.state.context?.departmentRelations ?? [];
    const isDeptManager = relations.some(rel => rel?.employeeDepartment?.RoleInDepartment === 'DepartmentManager');
    if (isDeptManager) return true;
    if (this.state.isManagerOfDeliverables) return true;
    return relations.some(rel => rel?.employeeDepartment?.CanAssignTasks === true);
  }

  private get canCreateDeliverable(): boolean {
    return this.isProjectLeader;
  }

  private get filteredDeliverables(): IDeliverable[] {
    const { deliverables } = this.state;
    if (!this.hasSelectedDepartments) return deliverables;
    return deliverables.filter(d => this.state.selectedDepartments.indexOf(d.AssignedDepartmentId) >= 0);
  }

  private get filteredObjectives(): IObjective[] {
    const { objectives } = this.state;
    if (!this.hasSelectedDepartments) return objectives;
    return objectives.filter(o => !o.AssignedDepartmentId || this.state.selectedDepartments.indexOf(o.AssignedDepartmentId) >= 0);
  }

  private get filteredTasks(): ITask[] {
    const { tasks } = this.state;
    if (!this.hasSelectedDepartments) return tasks;
    return tasks.filter(t => !t.AssignedDepartmentId || this.state.selectedDepartments.indexOf(t.AssignedDepartmentId) >= 0);
  }

  private get filteredDepartmentEmployees(): IEmployee[] {
    const { departmentEmployees, context } = this.state;
    if (!context) return departmentEmployees;
    if (!this.hasSelectedDepartments) return departmentEmployees;
    const deptRelations = context.departmentRelations.filter(rel => this.state.selectedDepartments.indexOf(rel.department.Id) >= 0);
    const empIds = new Set<number>();
    deptRelations.forEach(rel => empIds.add(rel.employee.Id));
    return departmentEmployees.filter(emp => empIds.has(emp.Id));
  }

  private get isManagerView(): boolean {
    if (this.isProjectLeader) return true;
    if (this.state.userRole !== 'normal') return true;
    if (this.state.isManagerOfDeliverables) return true;
    const relations = this.state.context?.departmentRelations ?? [];
    const isDeptManager = relations.some(rel => rel?.employeeDepartment?.RoleInDepartment === 'DepartmentManager');
    return isDeptManager;
  }

  private get canCreateObjectiveOrTask(): boolean {
    return this.canCreateObjective || this.canCreateTask;
  }

  private openObjectiveModal = (): void => this.setState({ showObjectiveModal: true });
  private closeObjectiveModal = (): void => {
    if (confirm('Are you sure you want to close? Any unsaved changes will be lost.')) {
      this.setState({ showObjectiveModal: false });
    }
  };
  private openTaskModal = (): void => this.setState({ showTaskModal: true });
  private closeTaskModal = (): void => {
    if (confirm('Are you sure you want to close? Any unsaved changes will be lost.')) {
      this.setState({ showTaskModal: false });
    }
  };
  private openEmployeeModal = (emp: ISubordinateInfo): void => this.setState({ showEmployeeModal: true, selectedEmployee: emp });
  private closeEmployeeModal = (): void => this.setState({ showEmployeeModal: false, selectedEmployee: null });
  private openCycleModal = (cycle?: IEvaluationCycle): void => this.setState({ showCycleModal: true, selectedCycle: cycle || null });
  private closeCycleModal = (): void => {
    if (confirm('Are you sure you want to close? Any unsaved changes will be lost.')) {
      this.setState({ showCycleModal: false, selectedCycle: null });
    }
  };
  private openCycleListModal = (): void => this.setState({ showCycleListModal: true });
  private closeCycleListModal = (): void => this.setState({ showCycleListModal: false });
  private openDeliverableDetailModal = (del: IDeliverable): void => this.setState({ showDeliverableDetailModal: true, selectedDeliverable: del });
  private closeDeliverableDetailModal = (): void => this.setState({ showDeliverableDetailModal: false, selectedDeliverable: null });
  private openObjectiveDetailModal = (obj: IObjective): void => this.setState({ showObjectiveDetailModal: true, selectedObjective: obj });
  private closeObjectiveDetailModal = (): void => this.setState({ showObjectiveDetailModal: false, selectedObjective: null });
  private openObjectiveEditModal = (obj: IObjective): void => this.setState({ showObjectiveEditModal: true, selectedObjective: obj, showObjectiveDetailModal: false });
  private closeObjectiveEditModal = (): void => this.setState({ showObjectiveEditModal: false, selectedObjective: null });
  private openDeliverableModal = (): void => this.setState({ showDeliverableModal: true });
  private closeDeliverableModal = (): void => {
    if (confirm('Are you sure you want to close? Any unsaved changes will be lost.')) {
      this.setState({ showDeliverableModal: false });
    }
  };
  
  private switchDesignRole = (role: 'employee' | 'manager' | 'director'): void => {
    this.loadDesignModeData(role);
  };

  public renderDesignModeSelector(): React.ReactElement | null {
    if (!this.state.isDesignMode) return null;

    return (
      <div className="demoBanner" style={{ marginBottom: '14px', background: '#fff6df', borderColor: '#f6d58b', color: '#92400e' }}>
        <span>
          <strong>Design Mode:</strong> Preview without connection. Switch role:{' '}
          <button 
            className={`chip ${this.state.designModeRole === 'employee' ? 'chipBlue' : 'chipNeutral'}`}
            style={{ cursor: 'pointer', marginLeft: '5px' }}
            onClick={() => this.switchDesignRole('employee')}
          >
            Employee
          </button>
          <button 
            className={`chip ${this.state.designModeRole === 'manager' ? 'chipBlue' : 'chipNeutral'}`}
            style={{ cursor: 'pointer', marginLeft: '5px' }}
            onClick={() => this.switchDesignRole('manager')}
          >
            Manager
          </button>
          <button 
            className={`chip ${this.state.designModeRole === 'director' ? 'chipBlue' : 'chipNeutral'}`}
            style={{ cursor: 'pointer', marginLeft: '5px' }}
            onClick={() => this.switchDesignRole('director')}
          >
            Director
          </button>
        </span>
      </div>
    );
  }

  public renderHero(): React.ReactElement {
    const { context, isDesignMode } = this.state;
    if (!context) return <div />;

    return (
      <div className="hero">
        <div>
          <h1 className="title">
            Welcome, {context.currentEmployee.Title}
            {isDesignMode && <span className="chip chipGreen" style={{ marginLeft: '10px', fontSize: '11px' }}>Demo</span>}
          </h1>
          <p className="subtitle">
            {context.currentEmployee.PositionTitle}
            <br />
            <span className={`chip ${this.getRoleBadgeClass()}`} style={{ marginTop: '4px' }}>
              {this.getRoleLabel()}
            </span>
            {context.primaryDepartmentRelation && context.primaryDepartmentRelation.department.Title !== 'Executive Direction' && (
              <span className="chip chipNeutral" style={{ marginLeft: '4px' }}>
                {context.primaryDepartmentRelation.department.Title}
              </span>
            )}
            {this.isProjectLeader && (
              <span className="chip chipPurple" style={{ marginLeft: '4px' }}>
                Project Lead
              </span>
            )}
            {this.isDepartmentManagerInAny && !this.isProjectLeader && (
              <span className="chip chipBlue" style={{ marginLeft: '4px' }}>
                Department Manager
              </span>
            )}
          </p>
        </div>
        <div className="heroActions">
          {context.permissions.canManageCycles && (
            <button className="secondaryBtn" onClick={this.openCycleListModal}>📋 Cycles</button>
          )}
          {this.canCreateDeliverable && (
            <button className="secondaryBtn" onClick={this.openDeliverableModal}>📦 Deliverable</button>
          )}
          {this.canCreateObjectiveOrTask && (
            <button className="primaryBtn" onClick={this.openObjectiveModal}>+ Objective</button>
          )}
          {this.canCreateObjectiveOrTask && (
            <button className="secondaryBtn" onClick={this.openTaskModal}>+ Task</button>
          )}
        </div>
      </div>
    );
  }

  public renderToolbar(): React.ReactElement {
    const { searchText, selectedDepartment, selectedDepartments, context, allDepartments } = this.state;
    const departments = context?.departmentRelations || [];
    const uniqueDepts = departments.filter((v, i, a) => a.findIndex(t => t.department.Id === v.department.Id) === i);
    
    const availableDepts = this.isProjectLeader
      ? allDepartments
      : uniqueDepts.map(rel => rel.department);

    const handleDepartmentToggle = (deptId: number): void => {
      const currentDepts = this.state.selectedDepartments;
      let newDepts: number[];
      if (currentDepts.indexOf(deptId) >= 0) {
        newDepts = currentDepts.filter(id => id !== deptId);
      } else {
        newDepts = [...currentDepts, deptId];
      }
      this.setState({ selectedDepartments: newDepts });
    };

    const selectAllDepts = (): void => {
      const allDeptIds = availableDepts.map(d => d.Id);
      this.setState({ selectedDepartments: allDeptIds });
    };

    const clearDepts = (): void => {
      if (this.isProjectLeader) {
        this.setState({ selectedDepartments: [] });
      } else {
        const allDeptIds = availableDepts.map(d => d.Id);
        this.setState({ selectedDepartments: allDeptIds });
      }
    };

    return (
      <div className="toolbar">
        <input
          type="text"
          className="input"
          placeholder="Search by name or position..."
          value={searchText}
          onChange={(e) => this.setState({ searchText: e.target.value })}
        />
        <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
          <span style={{ fontSize: '12px', color: '#666' }}>Filter by Area:</span>
          <div style={{ display: 'flex', gap: '4px', flexWrap: 'wrap', maxWidth: '400px' }}>
            <button
              className={`chip ${selectedDepartments.length === 0 ? 'chipBlue' : 'chipNeutral'}`}
              style={{ cursor: 'pointer', fontSize: '11px' }}
              onClick={clearDepts}
            >
              All
            </button>
            {availableDepts.slice(0, 10).map(dept => (
              <button
                key={dept.Id}
                className={`chip ${selectedDepartments.indexOf(dept.Id) >= 0 ? 'chipBlue' : 'chipNeutral'}`}
                style={{ cursor: 'pointer', fontSize: '11px' }}
                onClick={() => handleDepartmentToggle(dept.Id)}
              >
                {dept.Title}
              </button>
            ))}
            {availableDepts.length > 10 && (
              <button
                className="chip chipNeutral"
                style={{ cursor: 'pointer', fontSize: '11px' }}
                onClick={selectAllDepts}
              >
                +{availableDepts.length - 10} more
              </button>
            )}
          </div>
          {selectedDepartments.length > 0 && (
            <span style={{ fontSize: '11px', color: '#0066cc' }}>
              ({selectedDepartments.length} selected)
            </span>
          )}
        </div>
      </div>
    );
  }

  public renderKPIGrid(): React.ReactElement {
    const kpi = this.totalKPI;
    const { userRole, context, cycles } = this.state;

    return (
      <div className="kpiGrid">
        {userRole !== 'normal' && (
          <div className="kpiCard">
            <div className="kpiValue">{this.state.subordinates.length}</div>
            <div className="kpiLabel">Members</div>
          </div>
        )}
        <div className="kpiCard">
          <div className="kpiValue">{kpi.objectives}</div>
          <div className="kpiLabel">Objectives</div>
        </div>
        <div className="kpiCard">
          <div className="kpiValue">{kpi.tasks}</div>
          <div className="kpiLabel">Tasks</div>
        </div>
        <div className="kpiCard">
          <div className="kpiValue">{kpi.completed}</div>
          <div className="kpiLabel">Completed</div>
        </div>
        <div className="kpiCard">
          <div className="kpiValue">{kpi.pending}</div>
          <div className="kpiLabel">Pending</div>
        </div>
        {context?.permissions.canManageCycles && (
          <div className="kpiCard">
            <div className="kpiValue">{cycles.filter(c => c.Status === 'Active').length}</div>
            <div className="kpiLabel">Active Cycles</div>
          </div>
        )}
      </div>
    );
  }

  public renderEmployeeView(): React.ReactElement {
    const { context } = this.state;
    if (!context) return <div />;

    return (
      <div className="mainGrid">
        <div className="tablePanel">
          <div className="panel">
            <div className="panelHeader">
              <h3>My Objectives & Tasks</h3>
            </div>
            <div className="barsList">
              <div className="barRow">
                <div className="barHeaderRow">
                  <span>Assigned Objectives</span>
                  <b>3/5</b>
                </div>
                <div className="barTrack">
                  <div className="barFill barBlue" style={{ width: '60%' }} />
                </div>
              </div>
              <div className="barRow">
                <div className="barHeaderRow">
                  <span>Completed Tasks</span>
                  <b>8/12</b>
                </div>
                <div className="barTrack">
                  <div className="barFill barGreen" style={{ width: '67%' }} />
                </div>
              </div>
              <div className="barRow">
                <div className="barHeaderRow">
                  <span>Deliverables</span>
                  <b>2/4</b>
                </div>
                <div className="barTrack">
                  <div className="barFill barOrange" style={{ width: '50%' }} />
                </div>
              </div>
            </div>
            <div className="noteBox" style={{ marginTop: '14px' }}>
              Your objectives and tasks are assigned by your supervisor: <strong>{context.reportsTo?.Title || 'Not assigned'}</strong>
            </div>
          </div>
        </div>

        <div className="sideCol">
          <div className="panel">
            <div className="panelHeader">
              <h3>My Information</h3>
            </div>
            <div className="detailGrid">
              <div className="detailBlock">
                <span className="detailLabel">Email</span>
                <span className="detailValue">{context.currentEmployee.EmployeeEmail}</span>
              </div>
              <div className="detailBlock">
                <span className="detailLabel">Level</span>
                <span className="detailValue">{context.currentEmployee.PositionLevel}</span>
              </div>
              <div className="detailBlock">
                <span className="detailLabel">Department</span>
                <span className="detailValue">
                  {context.primaryDepartmentRelation?.department.Title || 'N/A'}
                </span>
              </div>
              <div className="detailBlock">
                <span className="detailLabel">Reports To</span>
                <span className="detailValue">{context.reportsTo?.Title || 'N/A'}</span>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  public renderManagerView(): React.ReactElement {
    const { context, userRole } = this.state;
    const filteredSubs = this.filteredSubordinates;

    return (
      <div className="mainGrid">
        <div className="tablePanel">
          <div className="panel">
            <div className="panelHeader">
              <h3>My Team ({filteredSubs.length})</h3>
              {context?.permissions.canViewReports && (
                <span className="muted">With report permissions</span>
              )}
            </div>
            <div className="tableWrap">
              <table className="table">
                <thead>
                  <tr>
                    <th>Name</th>
                    <th>Position</th>
                    <th>Role</th>
                    <th>Progress</th>
                    <th>Status</th>
                    <th>Actions</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredSubs.length === 0 ? (
                    <tr>
                      <td colSpan={6} className="emptyCell">No team members found</td>
                    </tr>
                  ) : (
                    filteredSubs.map((sub, index) => this.renderTeamRow(sub, index))
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>

        <div className="sideCol">
          {this.renderDepartmentPanel()}
          {this.renderProgressPanel()}
          {userRole === 'admin' && this.renderAdminPanel()}
        </div>
      </div>
    );
  }

  public renderDirectorView(): React.ReactElement {
    const { context } = this.state;
    const deptManagers = this.departmentManagers;

    return (
      <div className="mainGrid">
        <div className="tablePanel">
          <div className="panel">
            <div className="panelHeader">
              <h3>General View - Department Managers ({deptManagers.length})</h3>
              {context?.permissions.canManageCycles && (
                <span className="badge badgeOk">Administrative Mode</span>
              )}
            </div>
            <div className="tableWrap">
              <table className="table">
                <thead>
                  <tr>
                    <th>Manager</th>
                    <th>Department</th>
                    <th>Position</th>
                    <th>Role</th>
                    <th>Status</th>
                    <th>Actions</th>
                  </tr>
                </thead>
                <tbody>
                  {deptManagers.length === 0 ? (
                    <tr>
                      <td colSpan={6} className="emptyCell">No department managers found</td>
                    </tr>
                  ) : (
                    deptManagers.map((mgr, index) => this.renderDepartmentManagerRow(mgr, index))
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>

        <div className="sideCol">
          <div className="panel">
            <div className="panelHeader">
              <h3>Departments</h3>
            </div>
            <div className="tagsRow">
              {deptManagers.map(mgr => (
                <span key={mgr.department.Id} className="chip chipBlue">
                  {mgr.department.Title}
                </span>
              ))}
            </div>
          </div>
          {this.renderAdminPanel()}
        </div>
      </div>
    );
  }

  private renderTeamRow(sub: ISubordinateInfo, index: number): React.ReactElement {
    const progress = 50 + (index * 8) % 45;
    const statusClass = progress >= 80 ? 'badgeOk' : progress >= 50 ? 'badgeWarn' : 'badgeNeutral';
    const statusText = progress >= 80 ? 'Excellent' : progress >= 50 ? 'In Progress' : 'Not Started';

    return (
      <tr key={`${sub.employee.Id}-${index}`} className="rowClick" onClick={() => this.openEmployeeModal(sub)}>
        <td className="titleCell">
          <span className="titleMain">{sub.employee.Title}</span>
        </td>
        <td>{sub.employee.PositionTitle}</td>
        <td>{sub.departmentRelation.employeeDepartment.RoleInDepartment}</td>
        <td>
          <div className="progressRow">
            <div className="progressTrack">
              <div className="progressFill" style={{ width: `${progress}%` }} />
            </div>
            <span className="progressText">{progress}%</span>
          </div>
        </td>
        <td>
          <span className={`badge ${statusClass}`}>{statusText}</span>
        </td>
        <td>
          <div className="actionBtns">
            <button className="iconBtnEdit" title="View details">👁</button>
          </div>
        </td>
      </tr>
    );
  }

  private renderDirectorRow(sub: ISubordinateInfo, index: number): React.ReactElement {
    const progress = 55 + (index * 10) % 40;
    const statusClass = progress >= 70 ? 'badgeOk' : progress >= 40 ? 'badgeWarn' : 'badgeNeutral';

    return (
      <tr key={`${sub.employee.Id}-${index}`} className="rowClick" onClick={() => this.openEmployeeModal(sub)}>
        <td className="titleCell">
          <span className="titleMain">{sub.employee.Title}</span>
        </td>
        <td>
          <span className="chip chipBlue">{sub.departmentRelation.department.Title}</span>
        </td>
        <td>{sub.employee.PositionTitle}</td>
        <td>
          <span className="badge badgeNeutral">{sub.subordinatesCount} members</span>
        </td>
        <td>
          <span className={`badge ${statusClass}`}>{progress}% progress</span>
        </td>
        <td>
          <div className="actionBtns">
            <button className="iconBtnEdit" title="View team">👥</button>
          </div>
        </td>
      </tr>
    );
  }

  private renderDepartmentManagerRow(mgr: { employee: IEmployee; department: IDepartment; deptRelation: IEmployeeDepartment }, index: number): React.ReactElement {
    return (
      <tr key={`${mgr.employee.Id}-${mgr.department.Id}-${index}`}>
        <td className="titleCell">
          <span className="titleMain">{mgr.employee.Title}</span>
        </td>
        <td>
          <span className="chip chipBlue">{mgr.department.Title}</span>
        </td>
        <td>{mgr.employee.PositionTitle}</td>
        <td>
          <span className="badge badgeOk">Manager</span>
        </td>
        <td>
          <span className="badge badgeOk">Active</span>
        </td>
        <td>
          <div className="actionBtns">
            <button className="iconBtnEdit" title="View team">👥</button>
          </div>
        </td>
      </tr>
    );
  }

  private renderDepartmentPanel(): React.ReactElement | null {
    const { context } = this.state;
    if (!context) return null;

    return (
      <div className="panel">
        <div className="panelHeader">
          <h3>My Department</h3>
        </div>
        <div className="tagsRow">
          {context.departmentRelations.map(rel => (
            <span
              key={rel.employeeDepartment.Id}
              className={`chip ${rel.employeeDepartment.IsPrimaryDepartment ? 'chipBlue' : 'chipNeutral'}`}
            >
              {rel.department.Title}
              {rel.employeeDepartment.IsDepartmentManager && ' ★'}
            </span>
          ))}
        </div>
        <div className="divider" />
        <div className="detailGrid" style={{ marginTop: '10px' }}>
          <div className="detailBlock">
            <span className="detailLabel">Reports To</span>
            <span className="detailValue">{context.reportsTo?.Title || 'N/A'}</span>
          </div>
          <div className="detailBlock">
            <span className="detailLabel">Senior Leader</span>
            <span className="detailValue">{context.executiveLeader?.Title || 'N/A'}</span>
          </div>
        </div>
      </div>
    );
  }

  private renderProgressPanel(): React.ReactElement | null {
    const { subordinates, departmentEmployees, objectives, tasks } = this.state;
    
    const membersToShow = subordinates.length > 0 ? subordinates : 
      departmentEmployees.map(emp => ({
        employee: emp,
        subordinatesCount: 0,
        departmentRelation: null as any
      }));

    if (membersToShow.length === 0) return null;

    const getMemberProgress = (employeeId: number): number => {
      const memberObjectives = objectives.filter(o => o.AssignedEmployeeId === employeeId);
      const memberTasks = tasks.filter(t => t.AssignedEmployeeId === employeeId);
      
      if (memberObjectives.length === 0 && memberTasks.length === 0) {
        return Math.floor(Math.random() * 40) + 30;
      }
      
      const completedObj = memberObjectives.filter(o => o.Status === 'Completed').length;
      const totalObj = memberObjectives.length;
      const completedTasks = memberTasks.filter(t => t.Status === 'Completed').length;
      const totalTasks = memberTasks.length;
      
      if (totalObj + totalTasks === 0) return 50;
      
      const totalItems = totalObj + totalTasks;
      const completedItems = completedObj + completedTasks;
      return Math.round((completedItems / totalItems) * 100);
    };

    return (
      <div className="panel">
        <div className="panelHeader">
          <h3>Team Progress</h3>
        </div>
        <div className="miniBars">
          {membersToShow.slice(0, 5).map((member: any, index: number) => {
            const progress = getMemberProgress(member.employee.Id);
            const name = member.employee.Title.split(' ')[0];
            return (
              <div key={member.employee.Id} className="miniBarItem">
                <div className="miniBarLabel">
                  <span>{name}</span>
                  <b>{progress}%</b>
                </div>
                <div className="miniBarTrack">
                  <div className={`miniBarFill ${progress >= 70 ? 'barGreen' : progress >= 40 ? 'barOrange' : 'barRed'}`} style={{ width: `${progress}%` }} />
                </div>
              </div>
            );
          })}
        </div>
      </div>
    );
  }

  private renderAdminPanel(): React.ReactElement | null {
    const { context } = this.state;
    if (!context) return null;

    return (
      <div className="panel">
        <div className="panelHeader">
          <h3>Admin Actions</h3>
        </div>
        <div className="quickActions">
          {context.permissions.canManageCycles && (
            <a href="#" onClick={this.openCycleListModal}>View Cycles</a>
          )}
          {context.permissions.canManageSettings && (
            <a href="#">Settings</a>
          )}
          {context.permissions.canViewReports && (
            <a href="#">General Reports</a>
          )}
        </div>
      </div>
    );
  }

  public renderObjectiveModal(): React.ReactElement | null {
    const { showObjectiveModal, cycles, deliverables, subordinates, departmentEmployees, deptEmployeeRelations, context, allDepartments, selectedDepartments } = this.state;
    if (!showObjectiveModal) return null;

    const isProjectLeader = this.isProjectLeader;
    const userDepartmentIds = context?.departmentRelations.map(r => r.department.Id) || [];
    
    const availableDepartments = isProjectLeader 
      ? allDepartments 
      : allDepartments.filter(d => userDepartmentIds.indexOf(d.Id) >= 0);
    
    const filteredDepts = selectedDepartments.length > 0 
      ? selectedDepartments.filter(id => availableDepartments.some(d => d.Id === id))
      : userDepartmentIds;
    
    const modalDeliverables = deliverables.filter(d => filteredDepts.indexOf(d.AssignedDepartmentId) >= 0);
    
    const modalDeptEmployees = departmentEmployees.filter(de => {
      const empDeptRelation = deptEmployeeRelations.find(r => r.EmployeeId === de.Id && filteredDepts.indexOf(r.DepartmentId) >= 0);
      return !!empDeptRelation;
    });

    const assignableEmployees = [
      ...subordinates.map(sub => sub.employee),
      ...modalDeptEmployees.filter(de => !subordinates.some(sub => sub.employee.Id === de.Id) && de.Id !== context?.currentEmployee.Id)
    ];

    return (
      <div className="modalBackdrop" onClick={this.closeObjectiveModal}>
        <div className="modalPanel" onClick={(e) => e.stopPropagation()}>
          <div className="modalHeader">
            <div>
              <h2 className="modalTitle">New Objective</h2>
              <p className="modalSubtitle">Create a new performance objective</p>
            </div>
            <button className="modalClose" onClick={this.closeObjectiveModal}>×</button>
          </div>
          <div className="formStack">
            <div className="formField">
              <label className="formLabel">Objective Title *</label>
              <input type="text" className="input" id="objectiveTitle" placeholder="e.g.: Improve communication skills" />
            </div>
            <div className="formGrid2">
              <div className="formField">
                <label className="formLabel">Parent Deliverable</label>
                <select className="select" id="objectiveDeliverable">
                  <option value="">-- None --</option>
                  {modalDeliverables.map(del => (
                    <option key={del.Id} value={del.Id.toString()}>{del.Title}</option>
                  ))}
                </select>
              </div>
              <div className="formField">
                <label className="formLabel">Cycle *</label>
                <select className="select" id="objectiveCycle">
                  <option value="">Select cycle...</option>
                  {cycles.map(cycle => (
                    <option key={cycle.Id} value={cycle.Id.toString()}>{cycle.Title} ({cycle.Status})</option>
                  ))}
                </select>
              </div>
            </div>
            <div className="formGrid2">
              <div className="formField">
                <label className="formLabel">Source</label>
                <select className="select" id="objectiveSource">
                  <option value="Standalone">Standalone</option>
                  <option value="Deliverable">Deliverable</option>
                  <option value="Department">Department</option>
                  <option value="Individual">Individual</option>
                </select>
              </div>
              <div className="formField">
                <label className="formLabel">Type</label>
                <select className="select" id="objectiveType">
                  <option value="Individual">Individual</option>
                  <option value="Team">Team</option>
                  <option value="Department">Department</option>
                  <option value="Strategic">Strategic</option>
                </select>
              </div>
            </div>
            <div className="formGrid2">
              <div className="formField">
                <label className="formLabel">Category</label>
                <select className="select" id="objectiveCategory">
                  <option value="Performance">Performance</option>
                  <option value="Operational">Operational</option>
                  <option value="Strategic">Strategic</option>
                  <option value="Development">Development</option>
                </select>
              </div>
              <div className="formField">
                <label className="formLabel">Priority</label>
                <select className="select" id="objectivePriority">
                  <option value="Low">Low</option>
                  <option value="Medium">Medium</option>
                  <option value="High">High</option>
                  <option value="Critical">Critical</option>
                </select>
              </div>
            </div>
            <div className="formGrid2">
              <div className="formField">
                <label className="formLabel">Assigned Employee</label>
                <select className="select" id="objectiveEmployee">
                  <option value="">-- None --</option>
                  {assignableEmployees.map(emp => (
                    <option key={emp.Id} value={emp.Id.toString()}>
                      {emp.Title}
                    </option>
                  ))}
                </select>
              </div>
              <div className="formField">
                <label className="formLabel">Assigned Department</label>
                <select className="select" id="objectiveDept">
                  <option value="">-- None --</option>
                  {availableDepartments.map(dept => (
                    <option key={dept.Id} value={dept.Id.toString()}>{dept.Title}</option>
                  ))}
                </select>
              </div>
            </div>
            <div className="formField">
              <label className="formLabel">Due Date *</label>
              <input type="date" className="input" id="objectiveDueDate" />
            </div>
            <div className="formField">
              <label className="formLabel">Description</label>
              <textarea className="textarea" id="objectiveDescription" placeholder="Describe the objective..." />
            </div>
            <div className="modalInfoBox">
              The objective will be visible to the assigned employee and their direct supervisor.
            </div>
            <div className="modalActions">
              <button className="secondaryBtn" onClick={this.closeObjectiveModal} disabled={this.state.isSaving}>Cancel</button>
              <button className="primaryBtn" onClick={this.handleSaveObjective} disabled={this.state.isSaving}>
                {this.state.isSaving ? 'Saving...' : 'Create Objective'}
              </button>
            </div>
          </div>
        </div>
      </div>
    );
  }

  public renderTaskModal(): React.ReactElement | null {
    const { showTaskModal, cycles, deliverables, objectives, subordinates, departmentEmployees, deptEmployeeRelations, context, allDepartments, selectedDepartments } = this.state;
    if (!showTaskModal) return null;

    const isProjectLeader = this.isProjectLeader;
    const userDepartmentIds = context?.departmentRelations.map(r => r.department.Id) || [];
    
    const availableDepartments = isProjectLeader 
      ? allDepartments 
      : allDepartments.filter(d => userDepartmentIds.indexOf(d.Id) >= 0);
    
    const filteredDepts = selectedDepartments.length > 0 
      ? selectedDepartments.filter(id => availableDepartments.some(d => d.Id === id))
      : userDepartmentIds;
    
    const modalDeliverables = deliverables.filter(d => filteredDepts.indexOf(d.AssignedDepartmentId) >= 0);
    const modalObjectives = objectives.filter(o => !o.AssignedDepartmentId || filteredDepts.indexOf(o.AssignedDepartmentId) >= 0);
    const modalDeptEmployees = departmentEmployees.filter(de => {
      const empDeptRelation = deptEmployeeRelations.find(r => r.EmployeeId === de.Id && filteredDepts.indexOf(r.DepartmentId) >= 0);
      return !!empDeptRelation;
    });

    const assignableEmployees = [
      ...subordinates.map(sub => sub.employee),
      ...modalDeptEmployees.filter(de => !subordinates.some(sub => sub.employee.Id === de.Id) && de.Id !== context?.currentEmployee.Id)
    ];

    return (
      <div className="modalBackdrop" onClick={this.closeTaskModal}>
        <div className="modalPanel" onClick={(e) => e.stopPropagation()}>
          <div className="modalHeader">
            <div>
              <h2 className="modalTitle">New Task</h2>
              <p className="modalSubtitle">Create a specific task</p>
            </div>
            <button className="modalClose" onClick={this.closeTaskModal}>×</button>
          </div>
          <div className="formStack">
            <div className="formField">
              <label className="formLabel">Task Title *</label>
              <input type="text" className="input" id="taskTitle" placeholder="e.g.: Complete monthly report" />
            </div>
            <div className="formGrid2">
              <div className="formField">
                <label className="formLabel">Parent Deliverable</label>
                <select className="select" id="taskDeliverable">
                  <option value="">-- None --</option>
                  {modalDeliverables.map(del => (
                    <option key={del.Id} value={del.Id.toString()}>{del.Title}</option>
                  ))}
                </select>
              </div>
              <div className="formField">
                <label className="formLabel">Parent Objective</label>
                <select className="select" id="taskObjective">
                  <option value="">-- None --</option>
                  {modalObjectives.map(obj => (
                    <option key={obj.Id} value={obj.Id.toString()}>{obj.Title}</option>
                  ))}
                </select>
              </div>
            </div>
            <div className="formGrid2">
              <div className="formField">
                <label className="formLabel">Cycle *</label>
                <select className="select" id="taskCycle">
                  <option value="">Select cycle...</option>
                  {cycles.map(cycle => (
                    <option key={cycle.Id} value={cycle.Id.toString()}>{cycle.Title} ({cycle.Status})</option>
                  ))}
                </select>
              </div>
              <div className="formField">
                <label className="formLabel">Priority</label>
                <select className="select" id="taskPriority">
                  <option value="Low">Low</option>
                  <option value="Medium">Medium</option>
                  <option value="High">High</option>
                  <option value="Critical">Critical</option>
                </select>
              </div>
            </div>
            <div className="formGrid2">
              <div className="formField">
                <label className="formLabel">Source</label>
                <select className="select" id="taskSource">
                  <option value="Standalone">Standalone</option>
                  <option value="Deliverable">Deliverable</option>
                  <option value="Objective">Objective</option>
                  <option value="Department">Department</option>
                  <option value="Support">Support</option>
                </select>
              </div>
              <div className="formField">
                <label className="formLabel">Type</label>
                <select className="select" id="taskType">
                  <option value="Operational">Operational</option>
                  <option value="Administrative">Administrative</option>
                  <option value="Technical">Technical</option>
                  <option value="Support">Support</option>
                  <option value="Report">Report</option>
                  <option value="FollowUp">FollowUp</option>
                </select>
              </div>
            </div>
            <div className="formGrid2">
              <div className="formField">
                <label className="formLabel">Assigned Employee</label>
                <select className="select" id="taskEmployee">
                  <option value="">-- None --</option>
                  {assignableEmployees.map(emp => (
                    <option key={emp.Id} value={emp.Id.toString()}>
                      {emp.Title}
                    </option>
                  ))}
                </select>
              </div>
              <div className="formField">
                <label className="formLabel">Due Date *</label>
                <input type="date" className="input" id="taskDueDate" />
              </div>
            </div>
            <div className="formField">
              <label className="formLabel">Description</label>
              <textarea className="textarea" id="taskDescription" placeholder="Describe the task..." />
            </div>
            <div className="modalActions">
              <button className="secondaryBtn" onClick={this.closeTaskModal} disabled={this.state.isSaving}>Cancel</button>
              <button className="primaryBtn" onClick={this.handleSaveTask} disabled={this.state.isSaving}>
                {this.state.isSaving ? 'Saving...' : 'Create Task'}
              </button>
            </div>
          </div>
        </div>
      </div>
    );
  }

  public renderEmployeeModal(): React.ReactElement | null {
    const { showEmployeeModal, selectedEmployee } = this.state;
    if (!showEmployeeModal || !selectedEmployee) return null;

    const progress = 72;

    return (
      <div className="modalBackdrop" onClick={this.closeEmployeeModal}>
        <div className="modalPanel" onClick={(e) => e.stopPropagation()}>
          <div className="modalHeader">
            <div>
              <h2 className="modalTitle">{selectedEmployee.employee.Title}</h2>
              <p className="modalSubtitle">{selectedEmployee.employee.PositionTitle}</p>
            </div>
            <button className="modalClose" onClick={this.closeEmployeeModal}>×</button>
          </div>
          <div className="detailGrid">
            <div className="detailBlock">
              <span className="detailLabel">Email</span>
              <span className="detailValue">{selectedEmployee.employee.EmployeeEmail}</span>
            </div>
            <div className="detailBlock">
              <span className="detailLabel">Department</span>
              <span className="detailValue">{selectedEmployee.departmentRelation.department.Title}</span>
            </div>
            <div className="detailBlock">
              <span className="detailLabel">Role in Dept.</span>
              <span className="detailValue">{selectedEmployee.departmentRelation.employeeDepartment.RoleInDepartment}</span>
            </div>
            <div className="detailBlock">
              <span className="detailLabel">Subordinates</span>
              <span className="detailValue">{selectedEmployee.subordinatesCount}</span>
            </div>
          </div>
          <div className="divider" />
          <div className="panelHeader">
            <h3>Task Progress</h3>
            <span className="badge badgeOk">{progress}% Completed</span>
          </div>
          <div className="barsList">
            <div className="barRow">
              <div className="barHeaderRow">
                <span>Objectives</span>
                <b>3/5</b>
              </div>
              <div className="barTrack">
                <div className="barFill barBlue" style={{ width: '60%' }} />
              </div>
            </div>
            <div className="barRow">
              <div className="barHeaderRow">
                <span>Tasks</span>
                <b>8/12</b>
              </div>
              <div className="barTrack">
                <div className="barFill barGreen" style={{ width: '67%' }} />
              </div>
            </div>
            <div className="barRow">
              <div className="barHeaderRow">
                <span>Deliverables</span>
                <b>2/4</b>
              </div>
              <div className="barTrack">
                <div className="barFill barOrange" style={{ width: '50%' }} />
              </div>
            </div>
          </div>
          <div className="modalActions">
            <button className="secondaryBtn" onClick={this.closeEmployeeModal}>Close</button>
            <button className="primaryBtn">View Full Details</button>
          </div>
        </div>
      </div>
    );
  }

  public renderCycleListModal(): React.ReactElement | null {
    const { showCycleListModal, cycles } = this.state;
    if (!showCycleListModal) return null;

    const getStatusBadge = (status: string) => {
      switch (status) {
        case 'Active': return 'badgeOk';
        case 'Closed': return 'badgeDanger';
        case 'Draft': return 'badgeWarn';
        default: return 'badgeNeutral';
      }
    };

    const getStatusText = (status: string) => {
      switch (status) {
        case 'Active': return 'Active';
        case 'Closed': return 'Closed';
        case 'Draft': return 'Draft';
        case 'Completed': return 'Completed';
        default: return status;
      }
    };

    return (
      <div className="modalBackdrop" onClick={this.closeCycleListModal}>
        <div className="modalPanel" onClick={(e) => e.stopPropagation()}>
          <div className="modalHeader">
            <div>
              <h2 className="modalTitle">Evaluation Cycles</h2>
              <p className="modalSubtitle">Manage performance evaluation cycles</p>
            </div>
            <button className="modalClose" onClick={this.closeCycleListModal}>×</button>
          </div>
          
          <div style={{ marginBottom: '14px' }}>
            <button className="primaryBtn" onClick={() => this.openCycleModal()}>+ New Cycle</button>
          </div>

          <div className="tableWrap">
            <table className="table">
              <thead>
                <tr>
                  <th>Name</th>
                  <th>Year</th>
                  <th>Type</th>
                  <th>Start Date</th>
                  <th>End Date</th>
                  <th>Status</th>
                  <th>Actions</th>
                </tr>
              </thead>
              <tbody>
                {cycles.length === 0 ? (
                  <tr>
                    <td colSpan={7} className="emptyCell">No cycles registered</td>
                  </tr>
                ) : (
                  cycles.map(cycle => (
                    <tr key={cycle.Id}>
                      <td className="titleCell">
                        <span className="titleMain">{cycle.Title}</span>
                      </td>
                      <td>{cycle.Year}</td>
                      <td>{cycle.PeriodType}</td>
                      <td>{new Date(cycle.StartDate).toLocaleDateString()}</td>
                      <td>{new Date(cycle.EndDate).toLocaleDateString()}</td>
                      <td>
                        <span className={`badge ${getStatusBadge(cycle.Status)}`}>
                          {getStatusText(cycle.Status)}
                        </span>
                      </td>
                      <td>
                        <div className="actionBtns">
                          <button 
                            className="iconBtnEdit" 
                            title="Edit"
                            onClick={() => this.openCycleModal(cycle)}
                          >
                            ✏️
                          </button>
                          {cycle.Status === 'Active' && (
                            <button 
                              className="iconBtnEdit dangerBtn" 
                              title="Close"
                              onClick={() => this.handleCloseCycle(cycle.Id)}
                            >
                              🔒
                            </button>
                          )}
                        </div>
                      </td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
          
          <div className="modalActions" style={{ marginTop: '14px' }}>
            <button className="secondaryBtn" onClick={this.closeCycleListModal}>Close</button>
          </div>
        </div>
      </div>
    );
  }

  private handleCloseCycle = async (cycleId: number): Promise<void> => {
    if (this.organizationService) {
      await this.organizationService.closeCycle(cycleId);
      await this.loadData();
    }
  };

  public renderCycleModal(): React.ReactElement | null {
    const { showCycleModal, selectedCycle } = this.state;
    if (!showCycleModal) return null;

    const isEditing = !!selectedCycle;

    return (
      <div className="modalBackdrop" onClick={this.closeCycleModal}>
        <div className="modalPanel" onClick={(e) => e.stopPropagation()}>
          <div className="modalHeader">
            <div>
              <h2 className="modalTitle">{isEditing ? 'Edit Cycle' : 'New Evaluation Cycle'}</h2>
              <p className="modalSubtitle">{isEditing ? `Editing: ${selectedCycle.Title}` : 'Create a new evaluation cycle'}</p>
            </div>
            <button className="modalClose" onClick={this.closeCycleModal}>×</button>
          </div>
          
          <div className="formStack">
            <div className="formField">
              <label className="formLabel">Cycle Name *</label>
              <input 
                type="text" 
                className="input" 
                id="cycleTitle"
                placeholder="e.g.: Annual Evaluation 2025"
                defaultValue={selectedCycle?.Title || ''}
              />
            </div>
            
            <div className="formGrid2">
              <div className="formField">
                <label className="formLabel">Year *</label>
                <input 
                  type="number" 
                  className="input" 
                  id="cycleYear"
                  defaultValue={selectedCycle?.Year || new Date().getFullYear()}
                />
              </div>
              <div className="formField">
                <label className="formLabel">Period Type *</label>
                <select className="select" id="cyclePeriodType" defaultValue={selectedCycle?.PeriodType || 'Annual'}>
                  <option value="Annual">Annual</option>
                  <option value="Semestral">Semestral</option>
                  <option value="Quarterly">Quarterly</option>
                  <option value="Monthly">Monthly</option>
                </select>
              </div>
            </div>

            <div className="formGrid2">
              <div className="formField">
                <label className="formLabel">Start Date *</label>
                <input 
                  type="date" 
                  className="input" 
                  id="cycleStartDate"
                  defaultValue={selectedCycle?.StartDate?.split('T')[0] || ''}
                />
              </div>
              <div className="formField">
                <label className="formLabel">End Date *</label>
                <input 
                  type="date" 
                  className="input" 
                  id="cycleEndDate"
                  defaultValue={selectedCycle?.EndDate?.split('T')[0] || ''}
                />
              </div>
            </div>

            <div className="formField">
              <label className="formLabel">Status</label>
              <select className="select" id="cycleStatus" defaultValue={selectedCycle?.Status || 'Draft'}>
                <option value="Draft">Draft</option>
                <option value="Active">Active</option>
                <option value="Completed">Completed</option>
                <option value="Closed">Closed</option>
              </select>
            </div>

            <div className="formField">
              <label className="formLabel">Description</label>
              <textarea 
                className="textarea" 
                id="cycleDescription"
                placeholder="Optional cycle description..."
                defaultValue={selectedCycle?.Description || ''}
              />
            </div>

            <div className="modalInfoBox">
              {isEditing 
                ? 'Changes will be saved when clicking "Save Changes"'
                : 'The cycle will be created with "Draft" status. You can activate it later.'
              }
            </div>

            <div className="modalActions">
              <button className="secondaryBtn" onClick={this.closeCycleModal} disabled={this.state.isSaving}>Cancel</button>
              <button className="primaryBtn" onClick={this.handleSaveCycle} disabled={this.state.isSaving}>
                {this.state.isSaving ? 'Saving...' : (isEditing ? 'Save Changes' : 'Create Cycle')}
              </button>
            </div>
          </div>
        </div>
      </div>
    );
  }

  private handleSaveCycle = async (): Promise<void> => {
    const title = (document.getElementById('cycleTitle') as HTMLInputElement).value;
    const year = parseInt((document.getElementById('cycleYear') as HTMLInputElement).value);
    const periodType = (document.getElementById('cyclePeriodType') as HTMLSelectElement).value;
    const startDate = (document.getElementById('cycleStartDate') as HTMLInputElement).value;
    const endDate = (document.getElementById('cycleEndDate') as HTMLInputElement).value;
    const status = (document.getElementById('cycleStatus') as HTMLSelectElement).value;
    const description = (document.getElementById('cycleDescription') as HTMLTextAreaElement).value;

    if (!title || !year || !startDate || !endDate) {
      alert('Please complete all required fields');
      return;
    }

    this.setState({ isSaving: true });

    const cycleData = {
      Title: title,
      Year: year,
      PeriodType: periodType,
      StartDate: startDate,
      EndDate: endDate,
      Status: status,
      Description: description,
      IsActive: status === 'Active' || status === 'Draft'
    };

    try {
      if (this.organizationService) {
        if (this.state.selectedCycle) {
          await this.organizationService.updateCycle(this.state.selectedCycle.Id, cycleData);
        } else {
          await this.organizationService.createCycle(cycleData);
        }
        await this.loadData();
        this.setState({ showCycleModal: false, selectedCycle: null, isSaving: false });
      }
    } catch (error) {
      console.error('Error saving cycle:', error);
      alert('Error saving cycle. Please try again.');
      this.setState({ isSaving: false });
    }
  };

  private handleSaveDeliverable = async (): Promise<void> => {
    const title = (document.getElementById('deliverableTitle') as HTMLInputElement).value;
    const departmentId = parseInt((document.getElementById('deliverableDepartment') as HTMLSelectElement).value);
    const cycleId = parseInt((document.getElementById('deliverableCycle') as HTMLSelectElement).value);
    const category = (document.getElementById('deliverableCategory') as HTMLSelectElement).value;
    const priority = (document.getElementById('deliverablePriority') as HTMLSelectElement).value;
    const dueDate = (document.getElementById('deliverableDueDate') as HTMLInputElement).value;
    const description = (document.getElementById('deliverableDescription') as HTMLTextAreaElement).value;

    if (!title || !departmentId || !cycleId || !dueDate) {
      alert('Please complete all required fields');
      return;
    }

    this.setState({ isSaving: true });

    try {
      if (this.performanceService && this.organizationService && this.state.context) {
        let deptManagerId = this.state.context.currentEmployee.Id;
        
        const empDeptRelations = await this.organizationService.getEmployeeDepartmentsByDepartment(departmentId);
        const deptManagerRelation = empDeptRelations.find(rel => 
          rel.employeeDepartment.RoleInDepartment === 'DepartmentManager'
        );
        
        if (deptManagerRelation) {
          deptManagerId = deptManagerRelation.employee.Id;
        }
        
        const deliverableData: any = {
          Title: title,
          AssignedDepartmentId: departmentId,
          AssignedByEmployeeId: this.state.context.currentEmployee.Id,
          DepartmentManagerEmployeeId: deptManagerId,
          CycleId: cycleId,
          Category: category,
          Priority: priority,
          Status: 'Not Started',
          DueDate: dueDate,
          ProgressPercent: 0,
          Description: description,
          CanCreateChildObjectives: true,
          CanCreateChildTasks: true,
          IsActive: true
        };
        
        deliverableData.DeliverableCode = this.generateCode('DEL');
        
        await this.performanceService.createDeliverable(deliverableData);
        await this.loadData();
        this.setState({ showDeliverableModal: false, isSaving: false });
      }
    } catch (error) {
      console.error('Error saving deliverable:', error);
      alert('Error saving deliverable. Please try again.');
      this.setState({ isSaving: false });
    }
  };

  private handleSaveObjective = async (): Promise<void> => {
    const title = (document.getElementById('objectiveTitle') as HTMLInputElement).value;
    const deliverableIdStr = (document.getElementById('objectiveDeliverable') as HTMLSelectElement).value;
    const cycleId = parseInt((document.getElementById('objectiveCycle') as HTMLSelectElement).value);
    const source = (document.getElementById('objectiveSource') as HTMLSelectElement).value;
    const type = (document.getElementById('objectiveType') as HTMLSelectElement).value;
    const category = (document.getElementById('objectiveCategory') as HTMLSelectElement).value;
    const priority = (document.getElementById('objectivePriority') as HTMLSelectElement).value;
    const dueDate = (document.getElementById('objectiveDueDate') as HTMLInputElement).value;
    const description = (document.getElementById('objectiveDescription') as HTMLTextAreaElement).value;
    const assignedEmpIdStr = (document.getElementById('objectiveEmployee') as HTMLSelectElement).value;
    const assignedDeptIdStr = (document.getElementById('objectiveDept') as HTMLSelectElement).value;

    if (!title || !cycleId || !dueDate) {
      alert('Please complete all required fields');
      return;
    }

    this.setState({ isSaving: true });

    try {
      if (this.performanceService && this.organizationService && this.state.context) {
        const deliverableId = deliverableIdStr ? parseInt(deliverableIdStr) : undefined;
        const selectedDeliverable = deliverableId
          ? this.state.deliverables.find(d => d.Id === deliverableId)
          : undefined;

        const assignedDeptId =
          assignedDeptIdStr
            ? parseInt(assignedDeptIdStr)
            : selectedDeliverable?.AssignedDepartmentId;

        let managerEmployeeId: number | undefined = undefined;

        if (selectedDeliverable?.DepartmentManagerEmployeeId) {
          managerEmployeeId = selectedDeliverable.DepartmentManagerEmployeeId;
        } else if (assignedDeptId) {
          const deptRelations = await this.organizationService.getDepartmentEmployeesByDepartment(assignedDeptId);
          const deptManagerRelation = deptRelations.find(
            rel => rel.employeeDepartment.RoleInDepartment === 'DepartmentManager'
          );

          if (deptManagerRelation) {
            managerEmployeeId = deptManagerRelation.employee.Id;
          }
        }

        const isDepartmentObjective =
          type === 'Department' ||
          !!assignedDeptId ||
          source === 'Department';

        await this.performanceService.createObjective({
          ObjectiveCode: this.generateCode('OBJ'),
          Title: title,
          DeliverableId: deliverableId,
          CycleId: cycleId,
          ObjectiveSource: source as any,
          ObjectiveType: type as any,
          Category: category as any,
          Priority: priority as any,
          Status: 'Not Started',
          DueDate: dueDate,
          ProgressPercent: 0,
          Description: description,
          AssignedEmployeeId: assignedEmpIdStr ? parseInt(assignedEmpIdStr) : undefined,
          AssignedDepartmentId: assignedDeptId,
          AssignedByEmployeeId: this.state.context.currentEmployee.Id,
          ManagerEmployeeId: managerEmployeeId,
          CanCascadeTasks: true,
          IsDepartmentObjective: isDepartmentObjective,
          IsLinkedToDeliverable: !!deliverableId,
          IsOverdue: false,
          IsActive: true
        });

        await this.loadData();
        this.setState({ showObjectiveModal: false, isSaving: false });
      }
    } catch (error) {
      console.error('Error saving objective:', error);
      alert('Error saving objective. Please try again.');
      this.setState({ isSaving: false });
    }
  };

  private handleUpdateObjective = async (): Promise<void> => {
    const { selectedObjective } = this.state;
    if (!selectedObjective) return;

    const code = (document.getElementById('objectiveCode') as HTMLInputElement).value;
    const title = (document.getElementById('objectiveTitle') as HTMLInputElement).value;
    const deliverableIdStr = (document.getElementById('objectiveDeliverable') as HTMLSelectElement).value;
    const cycleId = parseInt((document.getElementById('objectiveCycle') as HTMLSelectElement).value);
    const source = (document.getElementById('objectiveSource') as HTMLSelectElement).value;
    const type = (document.getElementById('objectiveType') as HTMLSelectElement).value;
    const category = (document.getElementById('objectiveCategory') as HTMLSelectElement).value;
    const priority = (document.getElementById('objectivePriority') as HTMLSelectElement).value;
    const status = (document.getElementById('objectiveStatus') as HTMLSelectElement).value;
    const dueDate = (document.getElementById('objectiveDueDate') as HTMLInputElement).value;
    const description = (document.getElementById('objectiveDescription') as HTMLTextAreaElement).value;
    const assignedEmpIdStr = (document.getElementById('objectiveEmployee') as HTMLSelectElement).value;
    const assignedDeptIdStr = (document.getElementById('objectiveDept') as HTMLSelectElement).value;
    const progressStr = (document.getElementById('objectiveProgress') as HTMLInputElement).value;

    if (!title || !cycleId || !dueDate) {
      alert('Please complete all required fields');
      return;
    }

    this.setState({ isSaving: true });

    try {
      if (this.performanceService && this.organizationService && this.state.context) {
        const deliverableId = deliverableIdStr ? parseInt(deliverableIdStr) : undefined;
        const selectedDeliverable = deliverableId
          ? this.state.deliverables.find(d => d.Id === deliverableId)
          : undefined;

        const assignedDeptId =
          assignedDeptIdStr
            ? parseInt(assignedDeptIdStr)
            : selectedDeliverable?.AssignedDepartmentId;

        let managerEmployeeId: number | undefined = selectedObjective.ManagerEmployeeId;

        if (selectedDeliverable?.DepartmentManagerEmployeeId) {
          managerEmployeeId = selectedDeliverable.DepartmentManagerEmployeeId;
        } else if (assignedDeptId && !managerEmployeeId && this.organizationService) {
          const deptRelations = await this.organizationService.getDepartmentEmployeesByDepartment(assignedDeptId);
          const deptManagerRelation = deptRelations.find(
            rel => rel.employeeDepartment.RoleInDepartment === 'DepartmentManager'
          );
          if (deptManagerRelation) {
            managerEmployeeId = deptManagerRelation.employee.Id;
          }
        }

        const isDepartmentObjective =
          type === 'Department' ||
          !!assignedDeptId ||
          source === 'Department';

        await this.performanceService.updateObjective(selectedObjective.Id, {
          ObjectiveCode: code || undefined,
          Title: title,
          DeliverableId: deliverableId,
          CycleId: cycleId,
          ObjectiveSource: source as any,
          ObjectiveType: type as any,
          Category: category as any,
          Priority: priority as any,
          Status: status as any,
          DueDate: dueDate,
          ProgressPercent: progressStr ? parseInt(progressStr) : selectedObjective.ProgressPercent,
          Description: description,
          AssignedEmployeeId: assignedEmpIdStr ? parseInt(assignedEmpIdStr) : undefined,
          AssignedDepartmentId: assignedDeptId,
          ManagerEmployeeId: managerEmployeeId,
          CanCascadeTasks: true,
          IsDepartmentObjective: isDepartmentObjective,
          IsLinkedToDeliverable: !!deliverableId,
          IsOverdue: new Date(dueDate) < new Date() && status !== 'Completed'
        });

        await this.loadData();
        this.setState({ showObjectiveEditModal: false, isSaving: false });
      }
    } catch (error) {
      console.error('Error updating objective:', error);
      alert('Error updating objective. Please try again.');
      this.setState({ isSaving: false });
    }
  };

  private handleSaveTask = async (): Promise<void> => {
    const title = (document.getElementById('taskTitle') as HTMLInputElement).value;
    const deliverableIdStr = (document.getElementById('taskDeliverable') as HTMLSelectElement).value;
    const objectiveIdStr = (document.getElementById('taskObjective') as HTMLSelectElement).value;
    const cycleId = parseInt((document.getElementById('taskCycle') as HTMLSelectElement).value);
    const source = (document.getElementById('taskSource') as HTMLSelectElement).value;
    const type = (document.getElementById('taskType') as HTMLSelectElement).value;
    const priority = (document.getElementById('taskPriority') as HTMLSelectElement).value;
    const dueDate = (document.getElementById('taskDueDate') as HTMLInputElement).value;
    const description = (document.getElementById('taskDescription') as HTMLTextAreaElement).value;
    const assignedEmpIdStr = (document.getElementById('taskEmployee') as HTMLSelectElement).value;

    if (!title || !cycleId || !dueDate) {
      alert('Please complete all required fields');
      return;
    }

    this.setState({ isSaving: true });

    try {
      if (this.performanceService && this.state.context) {
        await this.performanceService.createTask({
          TaskCode: this.generateCode('TASK'),
          Title: title,
          DeliverableId: deliverableIdStr ? parseInt(deliverableIdStr) : undefined,
          ObjectiveId: objectiveIdStr ? parseInt(objectiveIdStr) : undefined,
          CycleId: cycleId,
          TaskSource: source as any,
          TaskType: type as any,
          Priority: priority as any,
          Status: 'Pending',
          DueDate: dueDate,
          ProgressPercent: 0,
          Description: description,
          AssignedEmployeeId: assignedEmpIdStr ? parseInt(assignedEmpIdStr) : undefined,
          RequiresApproval: false,
          ApprovalStatus: 'Not Required' as any,
          IsDepartmentTask: false,
          IsLinkedToDeliverable: !!deliverableIdStr,
          IsLinkedToObjective: !!objectiveIdStr,
          IsOverdue: false,
          IsActive: true
        });
        await this.loadData();
        this.setState({ showTaskModal: false, isSaving: false });
      }
    } catch (error) {
      console.error('Error saving task:', error);
      alert('Error saving task. Please try again.');
      this.setState({ isSaving: false });
    }
  };

  private getStatusBadgeClass(status: string): string {
    switch (status) {
      case 'Completed': return 'badgeOk';
      case 'In Progress': return 'badgeBlue';
      case 'On Hold': return 'badgeWarn';
      case 'Cancelled': return 'badgeDanger';
      case 'Overdue': return 'badgeDanger';
      default: return 'badgeNeutral';
    }
  }

  private generateCode(prefix: string): string {
    const year = new Date().getFullYear();
    const random = Math.floor(Math.random() * 9000) + 1000;
    return `${prefix}-${year}-${random}`;
  }

  public renderManagerViews(): React.ReactElement {
    const { activeTab, userRole } = this.state;
    const isDirector = userRole === 'admin';

    return (
      <div>
        <div className="tabsContainer">
          <button 
            className={`tabBtn ${activeTab === 'team' ? 'tabBtnActive' : ''}`}
            onClick={() => this.setState({ activeTab: 'team' })}
          >
            My Team
          </button>
          <button 
            className={`tabBtn ${activeTab === 'deliverables' ? 'tabBtnActive' : ''}`}
            onClick={() => this.setState({ activeTab: 'deliverables' })}
          >
            Deliverables
          </button>
          <button 
            className={`tabBtn ${activeTab === 'objectives' ? 'tabBtnActive' : ''}`}
            onClick={() => this.setState({ activeTab: 'objectives' })}
          >
            Objectives
          </button>
          <button 
            className={`tabBtn ${activeTab === 'tasks' ? 'tabBtnActive' : ''}`}
            onClick={() => this.setState({ activeTab: 'tasks' })}
          >
            Tasks
          </button>
        </div>
        
        {activeTab === 'team' && (isDirector ? this.renderDirectorView() : this.renderManagerView())}
        {activeTab === 'deliverables' && this.renderDeliverablesView()}
        {activeTab === 'objectives' && this.renderObjectivesView()}
        {activeTab === 'tasks' && this.renderTasksView()}
      </div>
    );
  }

  public renderDeliverablesView(): React.ReactElement {
    const { deliverables, context, cycles, allDepartments, selectedDepartments } = this.state;
    const activeCycle = cycles.find(c => c.Status === 'Active');
    const displayDeliverables = this.filteredDeliverables;

    return (
      <div className="mainGrid">
        <div className="tablePanel" style={{ width: '100%' }}>
          <div className="panel">
            <div className="panelHeader">
              <h3>Project Deliverables</h3>
              <div style={{ display: 'flex', gap: '8px' }}>
                {this.canCreateDeliverable && (
                  <button className="primaryBtn" onClick={this.openDeliverableModal}>+ New Deliverable</button>
                )}
              </div>
            </div>
            {!this.canCreateDeliverable && !this.canCreateObjectiveOrTask && (
              <div className="noteBox" style={{ marginBottom: '14px', background: '#fef3c7', borderColor: '#fcd34d' }}>
                Only Project Leaders can create Deliverables. Only Department Managers can create Objectives and Tasks.
              </div>
            )}
            {!this.canCreateDeliverable && this.canCreateObjectiveOrTask && (
              <div className="noteBox" style={{ marginBottom: '14px', background: '#dbeafe', borderColor: '#93c5fd' }}>
                You can create Objectives and Tasks for your department's deliverables.
              </div>
            )}
            <div className="tableWrap">
              <table className="table">
                <thead>
                  <tr>
                    <th>Deliverable</th>
                    <th>Department</th>
                    <th>Category</th>
                    <th>Priority</th>
                    <th>Progress</th>
                    <th>Status</th>
                    <th>Due Date</th>
                    <th>Actions</th>
                  </tr>
                </thead>
                <tbody>
                  {displayDeliverables.length === 0 ? (
                    <tr>
                      <td colSpan={8} className="emptyCell">No deliverables found</td>
                    </tr>
                  ) : (
                    displayDeliverables.map(del => (
                      <tr key={del.Id} className="rowClick" onClick={() => this.openDeliverableDetailModal(del)}>
                        <td className="titleCell">
                          <span className="titleMain">{del.Title}</span>
                          {del.AssignedByEmployeeId === context?.currentEmployee.Id && (
                            <span className="badge badgePurple" style={{ marginLeft: '8px', fontSize: '10px' }}>My Deliverable</span>
                          )}
                        </td>
                        <td>
                          <span className="chip chipBlue">
                            {allDepartments.find(d => d.Id === del.AssignedDepartmentId)?.Title || 'N/A'}
                          </span>
                        </td>
                        <td>{del.Category}</td>
                        <td>
                          <span className={`priority priority${del.Priority}`}>{del.Priority}</span>
                        </td>
                        <td>
                          <div className="progressRow">
                            <div className="progressTrack">
                              <div className="progressFill" style={{ width: `${del.ProgressPercent}%` }} />
                            </div>
                            <span className="progressText">{del.ProgressPercent}%</span>
                          </div>
                        </td>
                        <td>
                          <span className={`badge ${this.getStatusBadgeClass(del.Status)}`}>{del.Status}</span>
                        </td>
                        <td>{new Date(del.DueDate).toLocaleDateString()}</td>
                        <td>
                          <div className="actionBtns">
                            <button className="iconBtnEdit" title="View details" onClick={(e) => { e.stopPropagation(); this.openDeliverableDetailModal(del); }}>👁</button>
                          </div>
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      </div>
    );
  }

  public renderObjectivesView(): React.ReactElement {
    const { objectives, context, cycles, deliverables } = this.state;

    return (
      <div className="mainGrid">
        <div className="tablePanel" style={{ width: '100%' }}>
          <div className="panel">
            <div className="panelHeader">
              <h3>Objectives</h3>
              <div style={{ display: 'flex', gap: '8px' }}>
                {this.canCreateObjectiveOrTask && (
                  <button className="primaryBtn" onClick={this.openObjectiveModal}>+ New Objective</button>
                )}
              </div>
            </div>
            {!this.canCreateObjectiveOrTask && (
              <div className="noteBox" style={{ marginBottom: '14px', background: '#fef3c7', borderColor: '#fcd34d' }}>
                Only Department Managers can create Objectives for their department's deliverables.
              </div>
            )}
            <div className="tableWrap">
              <table className="table">
                <thead>
                  <tr>
                    <th>Objective</th>
                    <th>Source</th>
                    <th>Type</th>
                    <th>Priority</th>
                    <th>Progress</th>
                    <th>Status</th>
                    <th>Due Date</th>
                    <th>Actions</th>
                  </tr>
                </thead>
                <tbody>
                  {objectives.length === 0 ? (
                    <tr>
                      <td colSpan={8} className="emptyCell">No objectives found</td>
                    </tr>
                  ) : (
                    objectives.map(obj => {
                      const parentDel = deliverables.find(d => d.Id === obj.DeliverableId);
                      return (
                        <tr key={obj.Id} className="rowClick" onClick={() => this.openObjectiveDetailModal(obj)}>
                          <td className="titleCell">
                            <span className="titleMain">{obj.Title}</span>
                            {parentDel && (
                              <span className="badge badgeBlue" style={{ marginLeft: '8px', fontSize: '10px' }}>
                                {parentDel.Title.substring(0, 20)}...
                              </span>
                            )}
                          </td>
                          <td>{obj.ObjectiveSource}</td>
                          <td>{obj.ObjectiveType}</td>
                          <td>
                            <span className={`priority priority${obj.Priority}`}>{obj.Priority}</span>
                          </td>
                          <td>
                            <div className="progressRow">
                              <div className="progressTrack">
                                <div className="progressFill" style={{ width: `${obj.ProgressPercent}%` }} />
                              </div>
                              <span className="progressText">{obj.ProgressPercent}%</span>
                            </div>
                          </td>
                          <td>
                            <span className={`badge ${this.getStatusBadgeClass(obj.Status)}`}>{obj.Status}</span>
                          </td>
                          <td>{new Date(obj.DueDate).toLocaleDateString()}</td>
                          <td>
                            <div className="actionBtns">
                              <button className="iconBtnEdit" title="View details" onClick={(e) => { e.stopPropagation(); this.openObjectiveDetailModal(obj); }}>👁</button>
                              <button className="iconBtnEdit" title="Add Task" onClick={(e) => { e.stopPropagation(); this.setState({ showTaskModal: true }); }}>➕</button>
                            </div>
                          </td>
                        </tr>
                      );
                    })
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      </div>
    );
  }

  public renderTasksView(): React.ReactElement {
    const { tasks, context, objectives, deliverables } = this.state;

    return (
      <div className="mainGrid">
        <div className="tablePanel" style={{ width: '100%' }}>
          <div className="panel">
            <div className="panelHeader">
              <h3>Tasks</h3>
              <div style={{ display: 'flex', gap: '8px' }}>
                {this.canCreateObjectiveOrTask && (
                  <button className="primaryBtn" onClick={this.openTaskModal}>+ New Task</button>
                )}
              </div>
            </div>
            {!this.canCreateObjectiveOrTask && (
              <div className="noteBox" style={{ marginBottom: '14px', background: '#fef3c7', borderColor: '#fcd34d' }}>
                Only Department Managers can create Tasks for their department's deliverables.
              </div>
            )}
            <div className="tableWrap">
              <table className="table">
                <thead>
                  <tr>
                    <th>Task</th>
                    <th>Source</th>
                    <th>Type</th>
                    <th>Priority</th>
                    <th>Progress</th>
                    <th>Status</th>
                    <th>Due Date</th>
                    <th>Actions</th>
                  </tr>
                </thead>
                <tbody>
                  {tasks.length === 0 ? (
                    <tr>
                      <td colSpan={8} className="emptyCell">No tasks found</td>
                    </tr>
                  ) : (
                    tasks.map(task => {
                      const parentObj = objectives.find(o => o.Id === task.ObjectiveId);
                      const parentDel = deliverables.find(d => d.Id === task.DeliverableId);
                      return (
                        <tr key={task.Id}>
                          <td className="titleCell">
                            <span className="titleMain">{task.Title}</span>
                            <div style={{ display: 'flex', gap: '4px', marginTop: '2px' }}>
                              {parentObj && (
                                <span className="badge badgeBlue" style={{ fontSize: '9px' }}>
                                  OBJ: {parentObj.Title.substring(0, 15)}...
                                </span>
                              )}
                              {parentDel && !parentObj && (
                                <span className="badge badgeGreen" style={{ fontSize: '9px' }}>
                                  DEL: {parentDel.Title.substring(0, 15)}...
                                </span>
                              )}
                            </div>
                          </td>
                          <td>{task.TaskSource}</td>
                          <td>{task.TaskType}</td>
                          <td>
                            <span className={`priority priority${task.Priority}`}>{task.Priority}</span>
                          </td>
                          <td>
                            <div className="progressRow">
                              <div className="progressTrack">
                                <div className="progressFill" style={{ width: `${task.ProgressPercent}%` }} />
                              </div>
                              <span className="progressText">{task.ProgressPercent}%</span>
                            </div>
                          </td>
                          <td>
                            <span className={`badge ${this.getStatusBadgeClass(task.Status)}`}>{task.Status}</span>
                          </td>
                          <td>{new Date(task.DueDate).toLocaleDateString()}</td>
                          <td>
                            <div className="actionBtns">
                              <button className="iconBtnEdit" title="View">👁</button>
                            </div>
                          </td>
                        </tr>
                      );
                    })
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      </div>
    );
  }

  public renderDeliverableModal(): React.ReactElement | null {
    const { showDeliverableModal, cycles, context, allDepartments } = this.state;
    if (!showDeliverableModal) return null;

    const isProjectLeader = this.isProjectLeader;
    const userDepartmentIds = context?.departmentRelations.map(r => r.department.Id) || [];
    
    const availableDepartments = isProjectLeader 
      ? allDepartments 
      : allDepartments.filter(d => userDepartmentIds.indexOf(d.Id) >= 0);

    return (
      <div className="modalBackdrop" onClick={this.closeDeliverableModal}>
        <div className="modalPanel" onClick={(e) => e.stopPropagation()}>
          <div className="modalHeader">
            <div>
              <h2 className="modalTitle">New Deliverable</h2>
              <p className="modalSubtitle">Create a new project deliverable</p>
            </div>
            <button className="modalClose" onClick={this.closeDeliverableModal}>×</button>
          </div>
          <div className="formStack">
            <div className="formField">
              <label className="formLabel">Deliverable Title *</label>
              <input type="text" className="input" id="deliverableTitle" placeholder="e.g.: Digital Transformation Project" />
            </div>
            <div className="formGrid2">
              <div className="formField">
                <label className="formLabel">Department *</label>
                <select className="select" id="deliverableDepartment">
                  <option value="">Select department...</option>
                  {availableDepartments.map(dept => (
                    <option key={dept.Id} value={dept.Id.toString()}>{dept.Title}</option>
                  ))}
                </select>
              </div>
              <div className="formField">
                <label className="formLabel">Cycle *</label>
                <select className="select" id="deliverableCycle">
                  <option value="">Select cycle...</option>
                  {cycles.map(cycle => (
                    <option key={cycle.Id} value={cycle.Id.toString()}>{cycle.Title} ({cycle.Status})</option>
                  ))}
                </select>
              </div>
            </div>
            <div className="formGrid2">
              <div className="formField">
                <label className="formLabel">Category</label>
                <select className="select" id="deliverableCategory">
                  <option value="Project Deliverable">Project Deliverable</option>
                  <option value="Operational Deliverable">Operational Deliverable</option>
                  <option value="Strategic Deliverable">Strategic Deliverable</option>
                  <option value="Department Deliverable">Department Deliverable</option>
                </select>
              </div>
              <div className="formField">
                <label className="formLabel">Priority</label>
                <select className="select" id="deliverablePriority">
                  <option value="Low">Low</option>
                  <option value="Medium">Medium</option>
                  <option value="High">High</option>
                  <option value="Critical">Critical</option>
                </select>
              </div>
            </div>
            <div className="formField">
              <label className="formLabel">Due Date *</label>
              <input type="date" className="input" id="deliverableDueDate" />
            </div>
            <div className="formField">
              <label className="formLabel">Description</label>
              <textarea className="textarea" id="deliverableDescription" placeholder="Describe the deliverable..." />
            </div>
            <div className="modalInfoBox">
              The deliverable will be assigned to the selected department. The department manager will be notified.
            </div>
            <div className="modalActions">
              <button className="secondaryBtn" onClick={this.closeDeliverableModal} disabled={this.state.isSaving}>Cancel</button>
              <button className="primaryBtn" onClick={this.handleSaveDeliverable} disabled={this.state.isSaving}>
                {this.state.isSaving ? 'Saving...' : 'Create Deliverable'}
              </button>
            </div>
          </div>
        </div>
      </div>
    );
  }

  public renderDeliverableDetailModal(): React.ReactElement | null {
    const { showDeliverableDetailModal, selectedDeliverable, context, objectives, tasks, allDepartments } = this.state;
    if (!showDeliverableDetailModal || !selectedDeliverable) return null;

    const relatedObjectives = objectives.filter(o => o.DeliverableId === selectedDeliverable.Id);
    const relatedTasks = tasks.filter(t => t.DeliverableId === selectedDeliverable.Id);

    return (
      <div className="modalBackdrop" onClick={this.closeDeliverableDetailModal}>
        <div className="modalPanel modalLarge" onClick={(e) => e.stopPropagation()}>
          <div className="modalHeader">
            <div>
              <h2 className="modalTitle">{selectedDeliverable.Title}</h2>
              <p className="modalSubtitle">{selectedDeliverable.Category} - {selectedDeliverable.Status}</p>
            </div>
            <button className="modalClose" onClick={this.closeDeliverableDetailModal}>×</button>
          </div>
          <div className="detailGrid">
            <div className="detailBlock">
              <span className="detailLabel">Department</span>
              <span className="detailValue">
                {allDepartments.find(d => d.Id === selectedDeliverable.AssignedDepartmentId)?.Title || 'N/A'}
              </span>
            </div>
            <div className="detailBlock">
              <span className="detailLabel">Priority</span>
              <span className="detailValue">
                <span className={`priority priority${selectedDeliverable.Priority}`}>{selectedDeliverable.Priority}</span>
              </span>
            </div>
            <div className="detailBlock">
              <span className="detailLabel">Progress</span>
              <span className="detailValue">{selectedDeliverable.ProgressPercent}%</span>
            </div>
            <div className="detailBlock">
              <span className="detailLabel">Due Date</span>
              <span className="detailValue">{new Date(selectedDeliverable.DueDate).toLocaleDateString()}</span>
            </div>
          </div>
          {selectedDeliverable.Description && (
            <div style={{ marginBottom: '14px' }}>
              <span className="detailLabel">Description</span>
              <p style={{ margin: '4px 0 0 0' }}>{selectedDeliverable.Description}</p>
            </div>
          )}
          <div className="divider" />
          <div className="panelHeader">
            <h3>Objectives ({relatedObjectives.length})</h3>
            {selectedDeliverable.CanCreateChildObjectives && this.canCreateObjectiveOrTask && (
              <button className="secondaryBtn" style={{ padding: '4px 8px', fontSize: '12px' }} onClick={() => { this.closeDeliverableDetailModal(); this.openObjectiveModal(); }}>
                + Add Objective
              </button>
            )}
          </div>
          {relatedObjectives.length === 0 ? (
            <p className="muted">No objectives linked to this deliverable</p>
          ) : (
            <div className="miniList">
              {relatedObjectives.slice(0, 5).map(obj => (
                <div key={obj.Id} className="miniItem">
                  <span>{obj.Title}</span>
                  <span className={`badge ${this.getStatusBadgeClass(obj.Status)}`}>{obj.Status}</span>
                </div>
              ))}
            </div>
          )}
          <div className="panelHeader" style={{ marginTop: '14px' }}>
            <h3>Tasks ({relatedTasks.length})</h3>
            {selectedDeliverable.CanCreateChildTasks && this.canCreateObjectiveOrTask && (
              <button className="secondaryBtn" style={{ padding: '4px 8px', fontSize: '12px' }} onClick={() => { this.closeDeliverableDetailModal(); this.openTaskModal(); }}>
                + Add Task
              </button>
            )}
          </div>
          {relatedTasks.length === 0 ? (
            <p className="muted">No tasks linked to this deliverable</p>
          ) : (
            <div className="miniList">
              {relatedTasks.slice(0, 5).map(task => (
                <div key={task.Id} className="miniItem">
                  <span>{task.Title}</span>
                  <span className={`badge ${this.getStatusBadgeClass(task.Status)}`}>{task.Status}</span>
                </div>
              ))}
            </div>
          )}
          <div className="modalActions" style={{ marginTop: '14px' }}>
            <button className="secondaryBtn" onClick={this.closeDeliverableDetailModal}>Close</button>
          </div>
        </div>
      </div>
    );
  }

  public renderObjectiveDetailModal(): React.ReactElement | null {
    const { showObjectiveDetailModal, selectedObjective, context, tasks, deliverables } = this.state;
    if (!showObjectiveDetailModal || !selectedObjective) return null;

    const relatedTasks = tasks.filter(t => t.ObjectiveId === selectedObjective.Id);
    const parentDeliverable = deliverables.find(d => d.Id === selectedObjective.DeliverableId);

    const canEdit = context?.currentEmployee.Id === selectedObjective.AssignedByEmployeeId ||
                    context?.currentEmployee.Id === selectedObjective.ManagerEmployeeId ||
                    this.isProjectLeader;

    return (
      <div className="modalBackdrop" onClick={this.closeObjectiveDetailModal}>
        <div className="modalPanel modalLarge" onClick={(e) => e.stopPropagation()}>
          <div className="modalHeader">
            <div>
              <h2 className="modalTitle">{selectedObjective.Title}</h2>
              <p className="modalSubtitle">
                {selectedObjective.ObjectiveSource} - {selectedObjective.ObjectiveType} - {selectedObjective.Status}
              </p>
            </div>
            <div style={{ display: 'flex', gap: '8px' }}>
              {canEdit && (
                <button className="secondaryBtn" style={{ padding: '4px 12px', fontSize: '12px' }} onClick={(e) => { e.stopPropagation(); this.openObjectiveEditModal(selectedObjective); }}>
                  Edit
                </button>
              )}
              <button className="modalClose" onClick={this.closeObjectiveDetailModal}>×</button>
            </div>
          </div>
          <div className="detailGrid">
            <div className="detailBlock">
              <span className="detailLabel">Category</span>
              <span className="detailValue">{selectedObjective.Category}</span>
            </div>
            <div className="detailBlock">
              <span className="detailLabel">Priority</span>
              <span className="detailValue">
                <span className={`priority priority${selectedObjective.Priority}`}>{selectedObjective.Priority}</span>
              </span>
            </div>
            <div className="detailBlock">
              <span className="detailLabel">Progress</span>
              <span className="detailValue">{selectedObjective.ProgressPercent}%</span>
            </div>
            <div className="detailBlock">
              <span className="detailLabel">Due Date</span>
              <span className="detailValue">{new Date(selectedObjective.DueDate).toLocaleDateString()}</span>
            </div>
          </div>
          {parentDeliverable && (
            <div style={{ marginBottom: '14px' }}>
              <span className="detailLabel">Parent Deliverable</span>
              <p style={{ margin: '4px 0 0 0' }}>
                <span className="chip chipBlue">{parentDeliverable.Title}</span>
              </p>
            </div>
          )}
          {selectedObjective.Description && (
            <div style={{ marginBottom: '14px' }}>
              <span className="detailLabel">Description</span>
              <p style={{ margin: '4px 0 0 0' }}>{selectedObjective.Description}</p>
            </div>
          )}
          <div className="divider" />
          <div className="panelHeader">
            <h3>Tasks ({relatedTasks.length})</h3>
            {selectedObjective.CanCascadeTasks && this.canCreateObjectiveOrTask && (
              <button className="secondaryBtn" style={{ padding: '4px 8px', fontSize: '12px' }} onClick={() => { this.closeObjectiveDetailModal(); this.openTaskModal(); }}>
                + Add Task
              </button>
            )}
          </div>
          {relatedTasks.length === 0 ? (
            <p className="muted">No tasks linked to this objective</p>
          ) : (
            <div className="miniList">
              {relatedTasks.map(task => (
                <div key={task.Id} className="miniItem">
                  <span>{task.Title}</span>
                  <span className={`badge ${this.getStatusBadgeClass(task.Status)}`}>{task.Status}</span>
                </div>
              ))}
            </div>
          )}
          <div className="modalActions" style={{ marginTop: '14px' }}>
            <button className="secondaryBtn" onClick={this.closeObjectiveDetailModal}>Close</button>
          </div>
        </div>
      </div>
    );
  }

  public renderObjectiveEditModal(): React.ReactElement | null {
    const { showObjectiveEditModal, selectedObjective, cycles, deliverables, subordinates, departmentEmployees, deptEmployeeRelations, context, allDepartments } = this.state;
    if (!showObjectiveEditModal || !selectedObjective) return null;

    const isProjectLeader = this.isProjectLeader;
    const userDepartmentIds = context?.departmentRelations.map(r => r.department.Id) || [];
    
    const availableDepartments = isProjectLeader 
      ? allDepartments 
      : allDepartments.filter(d => userDepartmentIds.indexOf(d.Id) >= 0);
    
    const filteredDepts = availableDepartments.map(d => d.Id);
    
    const modalDeliverables = deliverables.filter(d => filteredDepts.indexOf(d.AssignedDepartmentId) >= 0);
    
    const modalDeptEmployees = departmentEmployees.filter(de => {
      const empDeptRelation = deptEmployeeRelations.find(r => r.EmployeeId === de.Id && filteredDepts.indexOf(r.DepartmentId) >= 0);
      return !!empDeptRelation;
    });

    const assignableEmployees = [
      ...subordinates.map(sub => sub.employee),
      ...modalDeptEmployees.filter(de => !subordinates.some(sub => sub.employee.Id === de.Id) && de.Id !== context?.currentEmployee.Id)
    ];

    const canEdit = context?.currentEmployee.Id === selectedObjective.AssignedByEmployeeId ||
                    context?.currentEmployee.Id === selectedObjective.ManagerEmployeeId ||
                    this.isProjectLeader;

    if (!canEdit) {
      return null;
    }

    return (
      <div className="modalBackdrop" onClick={this.closeObjectiveEditModal}>
        <div className="modalPanel" onClick={(e) => e.stopPropagation()}>
          <div className="modalHeader">
            <div>
              <h2 className="modalTitle">Edit Objective</h2>
              <p className="modalSubtitle">Update objective details</p>
            </div>
            <button className="modalClose" onClick={this.closeObjectiveEditModal}>×</button>
          </div>
          <div className="formStack">
            <div className="formField">
              <label className="formLabel">Objective Code</label>
              <input type="text" className="input" id="objectiveCode" defaultValue={selectedObjective.ObjectiveCode || ''} placeholder="e.g.: OBJ-2024-001" />
            </div>
            <div className="formField">
              <label className="formLabel">Objective Title *</label>
              <input type="text" className="input" id="objectiveTitle" defaultValue={selectedObjective.Title} placeholder="e.g.: Improve communication skills" />
            </div>
            <div className="formGrid2">
              <div className="formField">
                <label className="formLabel">Parent Deliverable</label>
                <select className="select" id="objectiveDeliverable" defaultValue={selectedObjective.DeliverableId?.toString() || ''}>
                  <option value="">-- None --</option>
                  {modalDeliverables.map(del => (
                    <option key={del.Id} value={del.Id.toString()}>{del.Title}</option>
                  ))}
                </select>
              </div>
              <div className="formField">
                <label className="formLabel">Cycle *</label>
                <select className="select" id="objectiveCycle" defaultValue={selectedObjective.CycleId?.toString() || ''}>
                  <option value="">Select cycle...</option>
                  {cycles.map(cycle => (
                    <option key={cycle.Id} value={cycle.Id.toString()}>{cycle.Title}</option>
                  ))}
                </select>
              </div>
            </div>
            <div className="formGrid2">
              <div className="formField">
                <label className="formLabel">Source</label>
                <select className="select" id="objectiveSource" defaultValue={selectedObjective.ObjectiveSource}>
                  <option value="Standalone">Standalone</option>
                  <option value="Deliverable">Deliverable</option>
                  <option value="Department">Department</option>
                  <option value="Individual">Individual</option>
                </select>
              </div>
              <div className="formField">
                <label className="formLabel">Type</label>
                <select className="select" id="objectiveType" defaultValue={selectedObjective.ObjectiveType}>
                  <option value="Individual">Individual</option>
                  <option value="Team">Team</option>
                  <option value="Department">Department</option>
                  <option value="Strategic">Strategic</option>
                </select>
              </div>
            </div>
            <div className="formGrid2">
              <div className="formField">
                <label className="formLabel">Category</label>
                <select className="select" id="objectiveCategory" defaultValue={selectedObjective.Category}>
                  <option value="Performance">Performance</option>
                  <option value="Operational">Operational</option>
                  <option value="Strategic">Strategic</option>
                  <option value="Development">Development</option>
                </select>
              </div>
              <div className="formField">
                <label className="formLabel">Priority</label>
                <select className="select" id="objectivePriority" defaultValue={selectedObjective.Priority}>
                  <option value="Low">Low</option>
                  <option value="Medium">Medium</option>
                  <option value="High">High</option>
                  <option value="Critical">Critical</option>
                </select>
              </div>
            </div>
            <div className="formGrid2">
              <div className="formField">
                <label className="formLabel">Status</label>
                <select className="select" id="objectiveStatus" defaultValue={selectedObjective.Status}>
                  <option value="Not Started">Not Started</option>
                  <option value="Pending">Pending</option>
                  <option value="In Progress">In Progress</option>
                  <option value="Completed">Completed</option>
                  <option value="On Hold">On Hold</option>
                  <option value="Cancelled">Cancelled</option>
                </select>
              </div>
              <div className="formField">
                <label className="formLabel">Progress (%)</label>
                <input type="number" className="input" id="objectiveProgress" defaultValue={selectedObjective.ProgressPercent?.toString() || '0'} min="0" max="100" />
              </div>
            </div>
            <div className="formGrid2">
              <div className="formField">
                <label className="formLabel">Assigned Employee</label>
                <select className="select" id="objectiveEmployee" defaultValue={selectedObjective.AssignedEmployeeId?.toString() || ''}>
                  <option value="">-- None --</option>
                  {assignableEmployees.map(emp => (
                    <option key={emp.Id} value={emp.Id.toString()}>
                      {emp.Title}
                    </option>
                  ))}
                </select>
              </div>
              <div className="formField">
                <label className="formLabel">Assigned Department</label>
                <select className="select" id="objectiveDept" defaultValue={selectedObjective.AssignedDepartmentId?.toString() || ''}>
                  <option value="">-- None --</option>
                  {availableDepartments.map(dept => (
                    <option key={dept.Id} value={dept.Id.toString()}>{dept.Title}</option>
                  ))}
                </select>
              </div>
            </div>
            <div className="formField">
              <label className="formLabel">Due Date *</label>
              <input type="date" className="input" id="objectiveDueDate" defaultValue={selectedObjective.DueDate ? selectedObjective.DueDate.split('T')[0] : ''} />
            </div>
            <div className="formField">
              <label className="formLabel">Description</label>
              <textarea className="textarea" id="objectiveDescription" defaultValue={selectedObjective.Description || ''} placeholder="Describe the objective..." />
            </div>
            <div className="modalActions">
              <button className="secondaryBtn" onClick={this.closeObjectiveEditModal} disabled={this.state.isSaving}>Cancel</button>
              <button className="primaryBtn" onClick={this.handleUpdateObjective} disabled={this.state.isSaving}>
                {this.state.isSaving ? 'Saving...' : 'Update Objective'}
              </button>
            </div>
          </div>
        </div>
      </div>
    );
  }

  public renderLoading(): React.ReactElement {
    return (
      <div className="demoBanner">
        <span>Loading organizational context...</span>
      </div>
    );
  }

  public renderError(): React.ReactElement {
    return (
      <div className="demoBanner" style={{ background: '#feebee', borderColor: '#f5b5bd', color: '#991b1b' }}>
        <span>{this.state.error}</span>
        <button className="bannerClose" onClick={() => this.loadData()}>↻</button>
      </div>
    );
  }

  public render(): React.ReactElement {
    const { loading, error, userRole, context } = this.state;
    
    if (loading) {
      return (
        <div>
          {this.renderLoading()}
        </div>
      );
    }

    if (error) {
      return (
        <div>
          {this.renderError()}
        </div>
      );
    }

    if (!context) {
      return (
        <div className="demoBanner" style={{ background: '#feebee', borderColor: '#f5b5bd', color: '#991b1b' }}>
          <span>No organizational context was found for this user.</span>
        </div>
      );
    }

    return (
      <div>
        {this.renderDesignModeSelector()}
        {this.renderHero()}
        
        {this.isManagerView ? this.renderManagerViews() : this.renderEmployeeView()}
        
        {this.renderToolbar()}
        {this.renderKPIGrid()}
        {this.renderDeliverableModal()}
        {this.renderObjectiveModal()}
        {this.renderTaskModal()}
        {this.renderEmployeeModal()}
        {this.renderCycleListModal()}
        {this.renderCycleModal()}
        {this.renderDeliverableDetailModal()}
        {this.renderObjectiveDetailModal()}
        {this.renderObjectiveEditModal()}
      </div>
    );
  }
}

export default EmployeeDashboard;
