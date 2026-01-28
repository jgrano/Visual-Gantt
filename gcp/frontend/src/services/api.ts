import axios, { AxiosInstance, AxiosError } from 'axios';
import { getIdToken } from './firebase';
import type {
  Project,
  Task,
  ProjectConfig,
  JiraConfig,
  SmartsheetConfig,
  ChartData,
  ApiResponse,
  JiraSyncResult,
  SmartsheetSyncResult,
  SyncConflict,
  SyncResolution,
  SmartsheetSheet,
  SmartsheetColumn,
} from '@shared/types';

// API base URL - adjust for production
const API_BASE_URL = import.meta.env.VITE_API_URL || '/api';

// Create axios instance
const apiClient: AxiosInstance = axios.create({
  baseURL: API_BASE_URL,
  headers: {
    'Content-Type': 'application/json',
  },
});

// Request interceptor to add auth token
apiClient.interceptors.request.use(async (config) => {
  const token = await getIdToken();
  if (token) {
    config.headers.Authorization = `Bearer ${token}`;
  }
  return config;
});

// Response interceptor for error handling
apiClient.interceptors.response.use(
  (response) => response,
  (error: AxiosError<ApiResponse>) => {
    const message = error.response?.data?.error || error.message || 'An error occurred';
    throw new Error(message);
  }
);

// ============================================================================
// Project APIs
// ============================================================================

export const projectApi = {
  list: async (): Promise<Project[]> => {
    const response = await apiClient.get<ApiResponse<Project[]>>('/projects');
    return response.data.data || [];
  },

  get: async (projectId: string): Promise<Project> => {
    const response = await apiClient.get<ApiResponse<Project>>(`/projects/${projectId}`);
    if (!response.data.data) throw new Error('Project not found');
    return response.data.data;
  },

  create: async (data: { name: string; description?: string }): Promise<Project> => {
    const response = await apiClient.post<ApiResponse<Project>>('/projects', data);
    if (!response.data.data) throw new Error('Failed to create project');
    return response.data.data;
  },

  update: async (projectId: string, data: Partial<Project>): Promise<Project> => {
    const response = await apiClient.put<ApiResponse<Project>>(`/projects/${projectId}`, data);
    if (!response.data.data) throw new Error('Failed to update project');
    return response.data.data;
  },

  delete: async (projectId: string): Promise<void> => {
    await apiClient.delete(`/projects/${projectId}`);
  },

  updateConfig: async (projectId: string, config: Partial<ProjectConfig>): Promise<ProjectConfig> => {
    const response = await apiClient.put<ApiResponse<ProjectConfig>>(`/projects/${projectId}/config`, config);
    if (!response.data.data) throw new Error('Failed to update config');
    return response.data.data;
  },
};

// ============================================================================
// Task APIs
// ============================================================================

export const taskApi = {
  list: async (projectId: string): Promise<Task[]> => {
    const response = await apiClient.get<ApiResponse<Task[]>>(`/projects/${projectId}/tasks`);
    return response.data.data || [];
  },

  get: async (projectId: string, taskId: string): Promise<Task> => {
    const response = await apiClient.get<ApiResponse<Task>>(`/projects/${projectId}/tasks/${taskId}`);
    if (!response.data.data) throw new Error('Task not found');
    return response.data.data;
  },

  create: async (projectId: string, task: Partial<Task>): Promise<Task> => {
    const response = await apiClient.post<ApiResponse<Task>>(`/projects/${projectId}/tasks`, task);
    if (!response.data.data) throw new Error('Failed to create task');
    return response.data.data;
  },

  createBulk: async (projectId: string, tasks: Partial<Task>[], replaceAll = false): Promise<Task[]> => {
    const response = await apiClient.post<ApiResponse<Task[]>>(`/projects/${projectId}/tasks/bulk`, {
      tasks,
      replaceAll,
    });
    return response.data.data || [];
  },

  update: async (projectId: string, taskId: string, task: Partial<Task>): Promise<Task> => {
    const response = await apiClient.put<ApiResponse<Task>>(`/projects/${projectId}/tasks/${taskId}`, task);
    if (!response.data.data) throw new Error('Failed to update task');
    return response.data.data;
  },

  delete: async (projectId: string, taskId: string): Promise<void> => {
    await apiClient.delete(`/projects/${projectId}/tasks/${taskId}`);
  },

  validate: async (projectId: string): Promise<{ valid: boolean; errors: string[]; warnings: string[] }> => {
    const response = await apiClient.post<ApiResponse<{ valid: boolean; errors: string[]; warnings: string[] }>>(
      `/projects/${projectId}/tasks/validate`
    );
    return response.data.data || { valid: true, errors: [], warnings: [] };
  },
};

// ============================================================================
// Jira APIs
// ============================================================================

export const jiraApi = {
  testConnection: async (
    projectId: string,
    credentials: { baseUrl: string; email: string; apiToken: string }
  ): Promise<string> => {
    const response = await apiClient.post<ApiResponse<never>>(`/projects/${projectId}/jira/test`, credentials);
    return response.data.message || 'Connected';
  },

  saveConfig: async (projectId: string, config: Partial<JiraConfig>): Promise<void> => {
    await apiClient.put(`/projects/${projectId}/jira/config`, config);
  },

  getConfig: async (projectId: string): Promise<JiraConfig | null> => {
    const response = await apiClient.get<ApiResponse<JiraConfig>>(`/projects/${projectId}/jira/config`);
    return response.data.data || null;
  },

  pull: async (projectId: string, replaceAll = false): Promise<JiraSyncResult> => {
    const response = await apiClient.post<ApiResponse<JiraSyncResult>>(`/projects/${projectId}/jira/pull`, {
      replaceAll,
    });
    if (!response.data.data) throw new Error('Failed to pull from Jira');
    return response.data.data;
  },

  push: async (projectId: string, taskIds?: string[]): Promise<JiraSyncResult> => {
    const response = await apiClient.post<ApiResponse<JiraSyncResult>>(`/projects/${projectId}/jira/push`, {
      taskIds,
    });
    if (!response.data.data) throw new Error('Failed to push to Jira');
    return response.data.data;
  },

  detectConflicts: async (projectId: string): Promise<SyncConflict[]> => {
    const response = await apiClient.post<ApiResponse<SyncConflict[]>>(`/projects/${projectId}/jira/conflicts`);
    return response.data.data || [];
  },

  resolveConflicts: async (projectId: string, resolutions: SyncResolution[]): Promise<void> => {
    await apiClient.post(`/projects/${projectId}/jira/resolve`, { resolutions });
  },
};

// ============================================================================
// Smartsheet APIs
// ============================================================================

export const smartsheetApi = {
  testConnection: async (projectId: string, apiToken: string): Promise<string> => {
    const response = await apiClient.post<ApiResponse<never>>(`/projects/${projectId}/smartsheet/test`, {
      apiToken,
    });
    return response.data.message || 'Connected';
  },

  saveConfig: async (projectId: string, config: Partial<SmartsheetConfig>): Promise<void> => {
    await apiClient.put(`/projects/${projectId}/smartsheet/config`, config);
  },

  getConfig: async (projectId: string): Promise<SmartsheetConfig | null> => {
    const response = await apiClient.get<ApiResponse<SmartsheetConfig>>(`/projects/${projectId}/smartsheet/config`);
    return response.data.data || null;
  },

  listSheets: async (projectId: string): Promise<SmartsheetSheet[]> => {
    const response = await apiClient.get<ApiResponse<SmartsheetSheet[]>>(`/projects/${projectId}/smartsheet/sheets`);
    return response.data.data || [];
  },

  getSheetColumns: async (projectId: string, sheetId: string): Promise<SmartsheetColumn[]> => {
    const response = await apiClient.get<ApiResponse<SmartsheetColumn[]>>(
      `/projects/${projectId}/smartsheet/sheets/${sheetId}/columns`
    );
    return response.data.data || [];
  },

  pull: async (projectId: string, replaceAll = false): Promise<SmartsheetSyncResult> => {
    const response = await apiClient.post<ApiResponse<SmartsheetSyncResult>>(
      `/projects/${projectId}/smartsheet/pull`,
      { replaceAll }
    );
    if (!response.data.data) throw new Error('Failed to pull from Smartsheet');
    return response.data.data;
  },

  push: async (projectId: string, taskIds?: string[]): Promise<SmartsheetSyncResult> => {
    const response = await apiClient.post<ApiResponse<SmartsheetSyncResult>>(
      `/projects/${projectId}/smartsheet/push`,
      { taskIds }
    );
    if (!response.data.data) throw new Error('Failed to push to Smartsheet');
    return response.data.data;
  },

  detectConflicts: async (projectId: string): Promise<SyncConflict[]> => {
    const response = await apiClient.post<ApiResponse<SyncConflict[]>>(
      `/projects/${projectId}/smartsheet/conflicts`
    );
    return response.data.data || [];
  },

  resolveConflicts: async (projectId: string, resolutions: SyncResolution[]): Promise<void> => {
    await apiClient.post(`/projects/${projectId}/smartsheet/resolve`, { resolutions });
  },
};

// ============================================================================
// Chart APIs
// ============================================================================

export const chartApi = {
  getData: async (projectId: string): Promise<ChartData> => {
    const response = await apiClient.get<ApiResponse<ChartData>>(`/projects/${projectId}/chart`);
    if (!response.data.data) throw new Error('Failed to get chart data');
    return response.data.data;
  },

  exportPng: async (projectId: string, imageData: string): Promise<{ url: string; filename: string }> => {
    const response = await apiClient.post<ApiResponse<{ url: string; filename: string }>>(
      `/projects/${projectId}/chart/export/png`,
      { imageData }
    );
    if (!response.data.data) throw new Error('Failed to export PNG');
    return response.data.data;
  },

  exportPdf: async (projectId: string, pdfData: string): Promise<{ url: string; filename: string }> => {
    const response = await apiClient.post<ApiResponse<{ url: string; filename: string }>>(
      `/projects/${projectId}/chart/export/pdf`,
      { pdfData }
    );
    if (!response.data.data) throw new Error('Failed to export PDF');
    return response.data.data;
  },
};
