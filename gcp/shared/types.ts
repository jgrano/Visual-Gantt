/**
 * Visual Gantt - Shared Types
 * Types shared between backend and frontend
 */

// ============================================================================
// TASK TYPES
// ============================================================================

export interface Task {
  id: string;
  name: string;
  startDate: string; // ISO date string
  endDate: string; // ISO date string
  duration?: number;
  owner?: string;
  percentComplete: number;
  priority: Priority;
  jiraTicket?: string;
  smartsheetRowId?: string;
  parentTask?: string;
  dependencies: string[];
  taskType: TaskType;
  swimlane?: string;
  isSubtask?: boolean;
  color?: string;
  // Internal tracking
  createdAt?: string;
  updatedAt?: string;
  modified?: boolean; // Tracks local changes for sync
}

export type Priority = 'High' | 'Medium' | 'Low';
export type TaskType = 'Task' | 'Milestone';

// ============================================================================
// PROJECT TYPES
// ============================================================================

export interface Project {
  id: string;
  name: string;
  description?: string;
  ownerId: string;
  members: ProjectMember[];
  createdAt: string;
  updatedAt: string;
  config: ProjectConfig;
  jiraConfig?: JiraConfig;
  smartsheetConfig?: SmartsheetConfig;
}

export interface ProjectMember {
  userId: string;
  email: string;
  role: 'owner' | 'editor' | 'viewer';
  addedAt: string;
}

// ============================================================================
// CONFIGURATION TYPES
// ============================================================================

export interface ProjectConfig {
  startDate?: string;
  endDate?: string;
  swimlaneGrouping: SwimlaneGrouping;
  colorScheme: ColorScheme;
  showDependencies: boolean;
  showPercentComplete: boolean;
  showMilestones: boolean;
  showTodayMarker: boolean;
  barHeight: number;
  barSpacing: number;
  fontSize: number;
  chartWidth: number;
  chartHeight: number;
}

export type SwimlaneGrouping = 'Owner' | 'Category' | 'None';
export type ColorScheme = 'priority' | 'taskType' | 'swimlane';

export const DEFAULT_PROJECT_CONFIG: ProjectConfig = {
  swimlaneGrouping: 'Category',
  colorScheme: 'priority',
  showDependencies: true,
  showPercentComplete: true,
  showMilestones: true,
  showTodayMarker: true,
  barHeight: 24,
  barSpacing: 8,
  fontSize: 12,
  chartWidth: 1200,
  chartHeight: 800,
};

// ============================================================================
// JIRA INTEGRATION TYPES
// ============================================================================

export interface JiraConfig {
  baseUrl: string;
  email: string;
  apiToken: string; // Encrypted at rest
  pullMethod: JiraPullMethod;
  jqlQuery?: string;
  projectKey?: string;
  filterId?: string;
  startDateField?: string;
  storyPointsField?: string;
  lastSyncTime?: string;
}

export type JiraPullMethod = 'jql' | 'project' | 'filter';

export interface JiraIssue {
  key: string;
  fields: {
    summary: string;
    status?: {
      name: string;
      statusCategory?: {
        key: string;
      };
    };
    priority?: {
      name: string;
    };
    assignee?: {
      displayName: string;
      emailAddress?: string;
    };
    duedate?: string;
    issuetype?: {
      name: string;
    };
    parent?: {
      key: string;
    };
    issuelinks?: JiraIssueLink[];
    labels?: string[];
    components?: { name: string }[];
    [key: string]: unknown; // For custom fields
  };
}

export interface JiraIssueLink {
  type: {
    inward: string;
    outward: string;
  };
  inwardIssue?: {
    key: string;
  };
  outwardIssue?: {
    key: string;
  };
}

export interface JiraSyncResult {
  success: boolean;
  issuesProcessed: number;
  issuesCreated: number;
  issuesUpdated: number;
  errors: string[];
}

// ============================================================================
// SMARTSHEET INTEGRATION TYPES
// ============================================================================

export interface SmartsheetConfig {
  apiToken: string; // Encrypted at rest
  sheetId: string;
  sheetName: string;
  includeChildRows: boolean;
  syncDependencies: boolean;
  columnMapping: SmartsheetColumnMapping;
  lastSyncTime?: string;
}

export interface SmartsheetColumnMapping {
  colTaskId?: string;
  colTaskName?: string;
  colStartDate?: string;
  colEndDate?: string;
  colDuration?: string;
  colOwner?: string;
  colPercentComplete?: string;
  colPriority?: string;
  colPredecessors?: string;
  colCategory?: string;
}

export interface SmartsheetSheet {
  id: string;
  name: string;
  columns: SmartsheetColumn[];
}

export interface SmartsheetColumn {
  id: string;
  title: string;
  type: string;
}

export interface SmartsheetRow {
  id: string;
  rowNumber: number;
  parentId?: string;
  cells: SmartsheetCell[];
}

export interface SmartsheetCell {
  columnId: string;
  value?: string | number | boolean;
  displayValue?: string;
}

export interface SmartsheetSyncResult {
  success: boolean;
  rowsProcessed: number;
  rowsCreated: number;
  rowsUpdated: number;
  errors: string[];
}

// ============================================================================
// SYNC TYPES
// ============================================================================

export interface SyncConflict {
  taskId: string;
  field: string;
  localValue: string;
  remoteValue: string;
  localModifiedAt?: string;
  remoteModifiedAt?: string;
}

export interface SyncResolution {
  taskId: string;
  resolution: 'local' | 'remote' | 'merge';
  mergedValue?: string;
}

// ============================================================================
// CHART DATA TYPES
// ============================================================================

export interface ChartData {
  tasks: Task[];
  config: ProjectConfig;
  startDate: string;
  endDate: string;
  today: string;
}

// ============================================================================
// API RESPONSE TYPES
// ============================================================================

export interface ApiResponse<T = unknown> {
  success: boolean;
  data?: T;
  error?: string;
  message?: string;
}

export interface PaginatedResponse<T> extends ApiResponse<T[]> {
  page: number;
  pageSize: number;
  total: number;
  hasMore: boolean;
}

// ============================================================================
// USER TYPES
// ============================================================================

export interface User {
  id: string;
  email: string;
  displayName?: string;
  photoUrl?: string;
  createdAt: string;
  lastLoginAt?: string;
}

// ============================================================================
// COLOR CONSTANTS
// ============================================================================

export const PRIORITY_COLORS: Record<Priority | 'default', string> = {
  High: '#DC3545',
  Medium: '#FFC107',
  Low: '#28A745',
  default: '#6C757D',
};

export const TASK_TYPE_COLORS: Record<TaskType | 'default', string> = {
  Milestone: '#9C27B0',
  Task: '#2196F3',
  default: '#607D8B',
};

export const SWIMLANE_COLORS: string[] = [
  '#4285F4', // Blue
  '#34A853', // Green
  '#FBBC05', // Yellow
  '#EA4335', // Red
  '#9C27B0', // Purple
  '#00ACC1', // Cyan
  '#FF7043', // Deep Orange
  '#5C6BC0', // Indigo
  '#26A69A', // Teal
  '#AB47BC', // Purple
];

// ============================================================================
// MAPPING CONSTANTS
// ============================================================================

export const JIRA_STATUS_MAPPING: Record<string, number> = {
  new: 0,
  indeterminate: 50,
  done: 100,
};

export const JIRA_PRIORITY_MAPPING: Record<string, Priority> = {
  Highest: 'High',
  High: 'High',
  Medium: 'Medium',
  Low: 'Low',
  Lowest: 'Low',
};

export const SMARTSHEET_PRIORITY_MAPPING: Record<string, Priority> = {
  Critical: 'High',
  High: 'High',
  Medium: 'Medium',
  Normal: 'Medium',
  Low: 'Low',
};
