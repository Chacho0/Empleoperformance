export interface IEmployee {
  Id: number;
  Title: string;
  EmployeeEmail: string;
  PositionTitle: string;
  PositionLevel: number;
  AppRole: string;
  CanManageCycles: boolean;
  CanManageSettings: boolean;
  IsActive: boolean;
}

export interface IDepartment {
  Id: number;
  Title: string;
  DepartmentCode: string;
  DepartmentManagerId: number | null;
  ReportsToExecutiveId: number | null;
  IsActive: boolean;
}

export interface IEmployeeDepartment {
  Id: number;
  Title: string;
  EmployeeId: number;
  DepartmentId: number;
  RoleInDepartment: string;
  ReportsToEmployeeId: number | null;
  IsPrimaryDepartment: boolean;
  IsDepartmentManager: boolean;
  CanViewDepartment: boolean;
  CanAssignObjectives: boolean;
  CanAssignTasks: boolean;
  CanViewReports: boolean;
  CanApprove: boolean;
  IsActive: boolean;
}

export interface IEmployeeDepartmentRelation {
  employeeDepartment: IEmployeeDepartment;
  employee: IEmployee;
  department: IDepartment;
  reportsToEmployee: IEmployee | null;
}

export interface IOrganizationalContext {
  currentEmployee: IEmployee;
  departmentRelations: IEmployeeDepartmentRelation[];
  primaryDepartmentRelation: IEmployeeDepartmentRelation | null;
  allDepartments: IDepartment[];
  reportsTo: IEmployee | null;
  subordinates: IEmployee[];
  isDepartmentManager: boolean;
  isPrimaryDepartmentManager: boolean;
  managedDepartments: IDepartment[];
  executiveLeader: IEmployee | null;
  permissions: IUserPermissions;
}

export interface IUserPermissions {
  canViewDepartment: boolean;
  canAssignObjectives: boolean;
  canAssignTasks: boolean;
  canViewReports: boolean;
  canApprove: boolean;
  canManageCycles: boolean;
  canManageSettings: boolean;
}

export interface ISubordinateInfo {
  employee: IEmployee;
  departmentRelation: IEmployeeDepartmentRelation;
  subordinatesCount: number;
}

export type UserRoleType = 'normal' | 'department_manager' | 'executive' | 'admin';

export interface IEvaluationCycle {
  Id: number;
  Title: string;
  Year: number;
  PeriodType: string;
  StartDate: string;
  EndDate: string;
  Status: string;
  Description: string;
  IsActive: boolean;
}

export type CycleStatus = 'Draft' | 'Active' | 'Closed' | 'Completed';
export type PeriodType = 'Annual' | 'Semestral' | 'Trimestral' | 'Monthly';

export interface IDeliverable {
  Id: number;
  Title: string;
  DeliverableCode?: string;
  AssignedDepartmentId: number;
  AssignedByEmployeeId: number;
  DepartmentManagerEmployeeId: number;
  CycleId: number;
  Category: DeliverableCategory;
  Priority: Priority;
  Status: DeliverableStatus;
  StartDate?: string;
  DueDate: string;
  ProgressPercent: number;
  Description?: string;
  ExpectedOutcome?: string;
  Comments?: string;
  CanCreateChildObjectives: boolean;
  CanCreateChildTasks: boolean;
  IsActive: boolean;
}

export interface IDeliverableRelation {
  deliverable: IDeliverable;
  department?: IDepartment;
  assignedByEmployee?: IEmployee;
  departmentManager?: IEmployee;
  cycle?: IEvaluationCycle;
  objectives?: IObjective[];
  tasks?: ITask[];
  childrenCount?: { objectives: number; tasks: number };
}

export type DeliverableCategory = 'Department Deliverable' | 'Strategic Deliverable' | 'Operational Deliverable' | 'Project Deliverable';
export type DeliverableStatus = 'Not Started' | 'Pending' | 'In Progress' | 'Completed' | 'On Hold' | 'Cancelled';
export type Priority = 'Low' | 'Medium' | 'High' | 'Critical';

export interface IObjective {
  Id: number;
  Title: string;
  ObjectiveCode?: string;
  DeliverableId?: number;
  ObjectiveSource: ObjectiveSource;
  ObjectiveType: ObjectiveType;
  AssignedEmployeeId?: number;
  AssignedDepartmentId?: number;
  AssignedByEmployeeId?: number;
  ManagerEmployeeId?: number;
  ParentObjectiveId?: number;
  CycleId: number;
  Category: ObjectiveCategory;
  Priority: Priority;
  Status: ObjectiveStatus;
  StartDate?: string;
  DueDate: string;
  ProgressPercent: number;
  Weight?: number;
  Tags?: string;
  Description?: string;
  ExpectedResult?: string;
  Comments?: string;
  CanCascadeTasks: boolean;
  IsDepartmentObjective: boolean;
  IsLinkedToDeliverable: boolean;
  IsOverdue: boolean;
  IsActive: boolean;
}

export interface IObjectiveRelation {
  objective: IObjective;
  deliverable?: IDeliverable;
  assignedEmployee?: IEmployee;
  assignedDepartment?: IDepartment;
  assignedByEmployee?: IEmployee;
  managerEmployee?: IEmployee;
  parentObjective?: IObjective;
  cycle?: IEvaluationCycle;
  tasks?: ITask[];
  childrenCount?: { tasks: number };
}

export type ObjectiveSource = 'Deliverable' | 'Department' | 'Individual' | 'Standalone';
export type ObjectiveType = 'Individual' | 'Department' | 'Team' | 'Strategic';
export type ObjectiveCategory = 'Performance' | 'Operational' | 'Strategic' | 'Development';
export type ObjectiveStatus = 'Not Started' | 'Pending' | 'In Progress' | 'Completed' | 'On Hold' | 'Cancelled';

export interface ITask {
  Id: number;
  Title: string;
  TaskCode?: string;
  DeliverableId?: number;
  ObjectiveId?: number;
  ParentTaskId?: number;
  TaskSource: TaskSource;
  TaskType: TaskType;
  AssignedEmployeeId?: number;
  AssignedDepartmentId?: number;
  CycleId: number;
  Priority: Priority;
  Status: TaskStatus;
  StartDate?: string;
  DueDate: string;
  CompletedDate?: string;
  ProgressPercent: number;
  Weight?: number;
  PlannedHours?: number;
  WorkedHours?: number;
  Description?: string;
  ExpectedResult?: string;
  Comments?: string;
  EvidenceLink?: string;
  RequiresApproval: boolean;
  ApprovalStatus: ApprovalStatus;
  ApprovedByEmployeeId?: number;
  ApprovalDate?: string;
  IsDepartmentTask: boolean;
  IsLinkedToDeliverable: boolean;
  IsLinkedToObjective: boolean;
  IsOverdue: boolean;
  IsActive: boolean;
}

export interface ITaskRelation {
  task: ITask;
  deliverable?: IDeliverable;
  objective?: IObjective;
  parentTask?: ITask;
  assignedEmployee?: IEmployee;
  assignedDepartment?: IDepartment;
  cycle?: IEvaluationCycle;
  approvedBy?: IEmployee;
  childrenTasks?: ITask[];
}

export type TaskSource = 'Deliverable' | 'Objective' | 'Department' | 'Standalone' | 'Support';
export type TaskType = 'Operational' | 'Administrative' | 'Technical' | 'Support' | 'Report' | 'FollowUp';
export type TaskStatus = 'Pending' | 'In Progress' | 'Completed' | 'On Hold' | 'Cancelled' | 'Overdue';
export type ApprovalStatus = 'Not Required' | 'Pending' | 'Approved' | 'Rejected';
