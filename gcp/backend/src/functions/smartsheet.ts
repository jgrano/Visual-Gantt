/**
 * Visual Gantt - Smartsheet Integration Functions
 * Cloud Functions for Smartsheet integration
 */

import { Response } from 'express';
import * as admin from 'firebase-admin';
import axios, { AxiosError } from 'axios';
import { AuthenticatedRequest, checkProjectAccess } from '../services/auth';
import {
  SmartsheetConfig,
  SmartsheetSheet,
  SmartsheetColumn,
  SmartsheetRow,
  SmartsheetCell,
  SmartsheetSyncResult,
  Task,
  SyncConflict,
  SyncResolution,
  SMARTSHEET_PRIORITY_MAPPING,
  Priority,
} from '../../../shared/types';

const db = admin.firestore();
const SMARTSHEET_API_BASE = 'https://api.smartsheet.com/2.0';

interface SmartsheetGetSheetResponse {
  id: string;
  name: string;
  columns: SmartsheetColumn[];
  rows: SmartsheetRow[];
}

/**
 * Test Smartsheet connection
 */
export async function testConnection(req: AuthenticatedRequest, res: Response): Promise<void> {
  try {
    const { projectId } = req.params;
    const userId = req.userId!;
    const { apiToken } = req.body;

    const hasAccess = await checkProjectAccess(userId, projectId, 'editor');
    if (!hasAccess) {
      res.status(403).json({ success: false, error: 'Access denied' });
      return;
    }

    if (!apiToken) {
      res.status(400).json({
        success: false,
        error: 'API token is required',
      });
      return;
    }

    const response = await smartsheetApiRequest(apiToken, '/users/me');

    res.json({
      success: true,
      message: `Connected as ${response.firstName} ${response.lastName} (${response.email})`,
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : 'Connection failed';
    res.status(400).json({
      success: false,
      error: message,
    });
  }
}

/**
 * Save Smartsheet configuration
 */
export async function saveConfig(req: AuthenticatedRequest, res: Response): Promise<void> {
  try {
    const { projectId } = req.params;
    const userId = req.userId!;
    const config = req.body as Partial<SmartsheetConfig>;

    const hasAccess = await checkProjectAccess(userId, projectId, 'editor');
    if (!hasAccess) {
      res.status(403).json({ success: false, error: 'Access denied' });
      return;
    }

    // Get existing config to preserve API token if not provided
    const projectDoc = await db.collection('projects').doc(projectId).get();
    const project = projectDoc.data();
    const existingConfig = project?.smartsheetConfig as SmartsheetConfig | undefined;

    // If token is masked, keep existing token
    let apiToken = config.apiToken;
    if (apiToken === '••••••••' && existingConfig?.apiToken) {
      apiToken = existingConfig.apiToken;
    }

    const smartsheetConfig: SmartsheetConfig = {
      apiToken: apiToken || '',
      sheetId: config.sheetId || '',
      sheetName: config.sheetName || '',
      includeChildRows: config.includeChildRows ?? true,
      syncDependencies: config.syncDependencies ?? true,
      columnMapping: config.columnMapping || {
        colTaskId: '',
        colTaskName: '',
        colStartDate: '',
        colEndDate: '',
        colDuration: '',
        colOwner: '',
        colPercentComplete: '',
        colPriority: '',
        colPredecessors: '',
        colCategory: '',
      },
      lastSyncTime: existingConfig?.lastSyncTime,
    };

    await db.collection('projects').doc(projectId).update({
      smartsheetConfig,
      updatedAt: new Date().toISOString(),
    });

    res.json({ success: true, message: 'Smartsheet configuration saved' });
  } catch (error) {
    console.error('Error saving Smartsheet config:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to save Smartsheet configuration',
    });
  }
}

/**
 * Get Smartsheet configuration (with masked token)
 */
export async function getConfig(req: AuthenticatedRequest, res: Response): Promise<void> {
  try {
    const { projectId } = req.params;
    const userId = req.userId!;

    const hasAccess = await checkProjectAccess(userId, projectId);
    if (!hasAccess) {
      res.status(403).json({ success: false, error: 'Access denied' });
      return;
    }

    const projectDoc = await db.collection('projects').doc(projectId).get();
    const project = projectDoc.data();
    const config = project?.smartsheetConfig as SmartsheetConfig | undefined;

    if (!config) {
      res.json({
        success: true,
        data: null,
      });
      return;
    }

    // Mask the API token
    res.json({
      success: true,
      data: {
        ...config,
        apiToken: config.apiToken ? '••••••••' : '',
      },
    });
  } catch (error) {
    console.error('Error getting Smartsheet config:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to get Smartsheet configuration',
    });
  }
}

/**
 * List available Smartsheet sheets
 */
export async function listSheets(req: AuthenticatedRequest, res: Response): Promise<void> {
  try {
    const { projectId } = req.params;
    const userId = req.userId!;

    const hasAccess = await checkProjectAccess(userId, projectId, 'editor');
    if (!hasAccess) {
      res.status(403).json({ success: false, error: 'Access denied' });
      return;
    }

    // Get Smartsheet config
    const projectDoc = await db.collection('projects').doc(projectId).get();
    const project = projectDoc.data();
    const config = project?.smartsheetConfig as SmartsheetConfig | undefined;

    if (!config?.apiToken) {
      res.status(400).json({
        success: false,
        error: 'Smartsheet is not configured for this project',
      });
      return;
    }

    const response = await smartsheetApiRequest(config.apiToken, '/sheets');

    const sheets: SmartsheetSheet[] = response.data.map((sheet: any) => ({
      id: String(sheet.id),
      name: sheet.name,
      columns: [],
    }));

    res.json({
      success: true,
      data: sheets,
    });
  } catch (error) {
    console.error('Error listing Smartsheet sheets:', error);
    const message = error instanceof Error ? error.message : 'Failed to list sheets';
    res.status(500).json({
      success: false,
      error: message,
    });
  }
}

/**
 * Get columns for a specific sheet
 */
export async function getSheetColumns(req: AuthenticatedRequest, res: Response): Promise<void> {
  try {
    const { projectId, sheetId } = req.params;
    const userId = req.userId!;

    const hasAccess = await checkProjectAccess(userId, projectId, 'editor');
    if (!hasAccess) {
      res.status(403).json({ success: false, error: 'Access denied' });
      return;
    }

    // Get Smartsheet config
    const projectDoc = await db.collection('projects').doc(projectId).get();
    const project = projectDoc.data();
    const config = project?.smartsheetConfig as SmartsheetConfig | undefined;

    if (!config?.apiToken) {
      res.status(400).json({
        success: false,
        error: 'Smartsheet is not configured for this project',
      });
      return;
    }

    const response = await smartsheetApiRequest(
      config.apiToken,
      `/sheets/${sheetId}?include=columns`
    ) as SmartsheetGetSheetResponse;

    const columns: SmartsheetColumn[] = response.columns.map(col => ({
      id: String(col.id),
      title: col.title,
      type: col.type,
    }));

    res.json({
      success: true,
      data: columns,
    });
  } catch (error) {
    console.error('Error getting Smartsheet columns:', error);
    const message = error instanceof Error ? error.message : 'Failed to get columns';
    res.status(500).json({
      success: false,
      error: message,
    });
  }
}

/**
 * Pull data from Smartsheet
 */
export async function pullFromSmartsheet(req: AuthenticatedRequest, res: Response): Promise<void> {
  try {
    const { projectId } = req.params;
    const userId = req.userId!;
    const { replaceAll } = req.body;

    const hasAccess = await checkProjectAccess(userId, projectId, 'editor');
    if (!hasAccess) {
      res.status(403).json({ success: false, error: 'Access denied' });
      return;
    }

    // Get Smartsheet config
    const projectDoc = await db.collection('projects').doc(projectId).get();
    const project = projectDoc.data();
    const config = project?.smartsheetConfig as SmartsheetConfig | undefined;

    if (!config?.apiToken || !config.sheetId) {
      res.status(400).json({
        success: false,
        error: 'Smartsheet is not configured for this project',
      });
      return;
    }

    // Fetch sheet data
    const sheetData = await smartsheetApiRequest(
      config.apiToken,
      `/sheets/${config.sheetId}`
    ) as SmartsheetGetSheetResponse;

    // Convert rows to tasks
    const tasks = smartsheetRowsToTasks(sheetData.rows, sheetData.columns, config);

    // Save tasks to Firestore
    const batch = db.batch();
    const tasksRef = db.collection('projects').doc(projectId).collection('tasks');
    const now = new Date().toISOString();

    // If replaceAll, delete existing tasks first
    if (replaceAll) {
      const existing = await tasksRef.get();
      existing.docs.forEach(doc => {
        batch.delete(doc.ref);
      });
    }

    let created = 0;
    let updated = 0;

    for (const task of tasks) {
      const existingDoc = await tasksRef.doc(task.id).get();

      if (existingDoc.exists && !replaceAll) {
        batch.update(tasksRef.doc(task.id), {
          ...task,
          updatedAt: now,
          modified: false,
        });
        updated++;
      } else {
        batch.set(tasksRef.doc(task.id), {
          ...task,
          createdAt: now,
          updatedAt: now,
          modified: false,
        });
        created++;
      }
    }

    // Update last sync time
    batch.update(db.collection('projects').doc(projectId), {
      'smartsheetConfig.lastSyncTime': now,
      updatedAt: now,
    });

    await batch.commit();

    const result: SmartsheetSyncResult = {
      success: true,
      rowsProcessed: sheetData.rows.length,
      rowsCreated: created,
      rowsUpdated: updated,
      errors: [],
    };

    res.json({
      success: true,
      data: result,
    });
  } catch (error) {
    console.error('Error pulling from Smartsheet:', error);
    const message = error instanceof Error ? error.message : 'Failed to pull from Smartsheet';
    res.status(500).json({
      success: false,
      error: message,
    });
  }
}

/**
 * Push changes to Smartsheet
 */
export async function pushToSmartsheet(req: AuthenticatedRequest, res: Response): Promise<void> {
  try {
    const { projectId } = req.params;
    const userId = req.userId!;
    const { taskIds } = req.body;

    const hasAccess = await checkProjectAccess(userId, projectId, 'editor');
    if (!hasAccess) {
      res.status(403).json({ success: false, error: 'Access denied' });
      return;
    }

    // Get Smartsheet config
    const projectDoc = await db.collection('projects').doc(projectId).get();
    const project = projectDoc.data();
    const config = project?.smartsheetConfig as SmartsheetConfig | undefined;

    if (!config?.apiToken || !config.sheetId) {
      res.status(400).json({
        success: false,
        error: 'Smartsheet is not configured for this project',
      });
      return;
    }

    // Get tasks to push
    const tasksRef = db.collection('projects').doc(projectId).collection('tasks');
    let tasksSnapshot;

    if (taskIds && taskIds.length > 0) {
      tasksSnapshot = await tasksRef.where(admin.firestore.FieldPath.documentId(), 'in', taskIds).get();
    } else {
      tasksSnapshot = await tasksRef.where('modified', '==', true).get();
    }

    const tasks = tasksSnapshot.docs.map(doc => ({
      id: doc.id,
      ...doc.data(),
    })) as Task[];

    // Get current sheet columns for mapping
    const sheetData = await smartsheetApiRequest(
      config.apiToken,
      `/sheets/${config.sheetId}?include=columns`
    ) as SmartsheetGetSheetResponse;

    const errors: string[] = [];
    let updated = 0;

    for (const task of tasks) {
      // Only push tasks that have a Smartsheet row ID
      if (!task.smartsheetRowId) {
        continue;
      }

      try {
        // Build row update
        const cells = taskToSmartsheetCells(task, sheetData.columns, config);

        await smartsheetApiRequest(
          config.apiToken,
          `/sheets/${config.sheetId}/rows`,
          'PUT',
          {
            id: task.smartsheetRowId,
            cells,
          }
        );

        // Clear modified flag
        await tasksRef.doc(task.id).update({
          modified: false,
          updatedAt: new Date().toISOString(),
        });

        updated++;
      } catch (error) {
        const message = error instanceof Error ? error.message : 'Unknown error';
        errors.push(`Failed to update row ${task.smartsheetRowId}: ${message}`);
      }
    }

    // Update last sync time
    await db.collection('projects').doc(projectId).update({
      'smartsheetConfig.lastSyncTime': new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    });

    const result: SmartsheetSyncResult = {
      success: errors.length === 0,
      rowsProcessed: tasks.length,
      rowsCreated: 0,
      rowsUpdated: updated,
      errors,
    };

    res.json({
      success: true,
      data: result,
    });
  } catch (error) {
    console.error('Error pushing to Smartsheet:', error);
    const message = error instanceof Error ? error.message : 'Failed to push to Smartsheet';
    res.status(500).json({
      success: false,
      error: message,
    });
  }
}

/**
 * Detect conflicts between local and Smartsheet data
 */
export async function detectConflicts(req: AuthenticatedRequest, res: Response): Promise<void> {
  try {
    const { projectId } = req.params;
    const userId = req.userId!;

    const hasAccess = await checkProjectAccess(userId, projectId, 'editor');
    if (!hasAccess) {
      res.status(403).json({ success: false, error: 'Access denied' });
      return;
    }

    // Get Smartsheet config
    const projectDoc = await db.collection('projects').doc(projectId).get();
    const project = projectDoc.data();
    const config = project?.smartsheetConfig as SmartsheetConfig | undefined;

    if (!config?.apiToken || !config.sheetId) {
      res.status(400).json({
        success: false,
        error: 'Smartsheet is not configured for this project',
      });
      return;
    }

    // Fetch current Smartsheet data
    const sheetData = await smartsheetApiRequest(
      config.apiToken,
      `/sheets/${config.sheetId}`
    ) as SmartsheetGetSheetResponse;

    const smartsheetTasks = smartsheetRowsToTasks(sheetData.rows, sheetData.columns, config);
    const smartsheetTasksMap = new Map<string, Task>();
    smartsheetTasks.forEach(task => smartsheetTasksMap.set(task.id, task));

    // Get local modified tasks
    const tasksRef = db.collection('projects').doc(projectId).collection('tasks');
    const modifiedSnapshot = await tasksRef.where('modified', '==', true).get();
    const localTasks = modifiedSnapshot.docs.map(doc => ({
      id: doc.id,
      ...doc.data(),
    })) as Task[];

    // Detect conflicts
    const conflicts: SyncConflict[] = [];

    for (const localTask of localTasks) {
      const remoteTask = smartsheetTasksMap.get(localTask.id);
      if (!remoteTask) {
        continue;
      }

      const fields = ['name', 'startDate', 'endDate', 'percentComplete', 'priority', 'owner'] as const;

      for (const field of fields) {
        const localValue = String(localTask[field] || '');
        const remoteValue = String(remoteTask[field] || '');

        if (localValue !== remoteValue) {
          conflicts.push({
            taskId: localTask.id,
            field,
            localValue,
            remoteValue,
            localModifiedAt: localTask.updatedAt,
          });
        }
      }
    }

    res.json({
      success: true,
      data: conflicts,
    });
  } catch (error) {
    console.error('Error detecting Smartsheet conflicts:', error);
    const message = error instanceof Error ? error.message : 'Failed to detect conflicts';
    res.status(500).json({
      success: false,
      error: message,
    });
  }
}

/**
 * Resolve sync conflicts
 */
export async function resolveConflicts(req: AuthenticatedRequest, res: Response): Promise<void> {
  try {
    const { projectId } = req.params;
    const userId = req.userId!;
    const { resolutions } = req.body as { resolutions: SyncResolution[] };

    const hasAccess = await checkProjectAccess(userId, projectId, 'editor');
    if (!hasAccess) {
      res.status(403).json({ success: false, error: 'Access denied' });
      return;
    }

    if (!resolutions || !Array.isArray(resolutions)) {
      res.status(400).json({
        success: false,
        error: 'Resolutions array is required',
      });
      return;
    }

    // Get Smartsheet config
    const projectDoc = await db.collection('projects').doc(projectId).get();
    const project = projectDoc.data();
    const config = project?.smartsheetConfig as SmartsheetConfig | undefined;

    if (!config) {
      res.status(400).json({
        success: false,
        error: 'Smartsheet is not configured for this project',
      });
      return;
    }

    const batch = db.batch();
    const tasksRef = db.collection('projects').doc(projectId).collection('tasks');

    for (const resolution of resolutions) {
      const taskRef = tasksRef.doc(resolution.taskId);

      if (resolution.resolution === 'remote') {
        // Fetch fresh data from Smartsheet
        const sheetData = await smartsheetApiRequest(
          config.apiToken,
          `/sheets/${config.sheetId}`
        ) as SmartsheetGetSheetResponse;

        const tasks = smartsheetRowsToTasks(sheetData.rows, sheetData.columns, config);
        const remoteTask = tasks.find(t => t.id === resolution.taskId);

        if (remoteTask) {
          batch.update(taskRef, {
            ...remoteTask,
            modified: false,
            updatedAt: new Date().toISOString(),
          });
        }
      } else if (resolution.resolution === 'local') {
        batch.update(taskRef, {
          modified: true,
          updatedAt: new Date().toISOString(),
        });
      }
    }

    await batch.commit();

    res.json({
      success: true,
      message: 'Conflicts resolved',
    });
  } catch (error) {
    console.error('Error resolving Smartsheet conflicts:', error);
    const message = error instanceof Error ? error.message : 'Failed to resolve conflicts';
    res.status(500).json({
      success: false,
      error: message,
    });
  }
}

// ============================================================================
// Helper Functions
// ============================================================================

/**
 * Make authenticated request to Smartsheet API
 */
async function smartsheetApiRequest(
  apiToken: string,
  endpoint: string,
  method: 'GET' | 'POST' | 'PUT' | 'DELETE' = 'GET',
  data?: unknown
): Promise<any> {
  const url = `${SMARTSHEET_API_BASE}${endpoint}`;

  try {
    const response = await axios({
      method,
      url,
      headers: {
        Authorization: `Bearer ${apiToken}`,
        'Content-Type': 'application/json',
      },
      data,
    });

    return response.data;
  } catch (error) {
    if (error instanceof AxiosError) {
      const message = error.response?.data?.message || error.message;
      throw new Error(message);
    }
    throw error;
  }
}

/**
 * Convert Smartsheet rows to tasks
 */
function smartsheetRowsToTasks(
  rows: SmartsheetRow[],
  columns: SmartsheetColumn[],
  config: SmartsheetConfig
): Task[] {
  const columnMap = new Map<string, SmartsheetColumn>();
  columns.forEach(col => columnMap.set(String(col.id), col));

  const mapping = config.columnMapping;
  const tasks: Task[] = [];

  // Build parent row mapping for hierarchy
  const rowIdToTaskId = new Map<string, string>();

  for (const row of rows) {
    const getCellValue = (columnId: string | undefined): string | number | undefined => {
      if (!columnId) return undefined;
      const cell = row.cells.find(c => String(c.columnId) === columnId);
      return cell?.displayValue || cell?.value;
    };

    const taskId = String(getCellValue(mapping.colTaskId) || `SS-${row.id}`);
    rowIdToTaskId.set(String(row.id), taskId);

    // Get values from mapped columns
    const name = String(getCellValue(mapping.colTaskName) || '');
    const startDate = String(getCellValue(mapping.colStartDate) || '');
    const endDate = String(getCellValue(mapping.colEndDate) || '');
    const duration = Number(getCellValue(mapping.colDuration)) || undefined;
    const owner = String(getCellValue(mapping.colOwner) || '');
    const percentComplete = Number(getCellValue(mapping.colPercentComplete)) || 0;
    const priorityValue = String(getCellValue(mapping.colPriority) || 'Medium');
    const swimlane = String(getCellValue(mapping.colCategory) || 'Default');

    // Map priority
    const priority: Priority = SMARTSHEET_PRIORITY_MAPPING[priorityValue] || 'Medium';

    // Parse predecessors/dependencies
    const dependencies: string[] = [];
    if (config.syncDependencies && mapping.colPredecessors) {
      const predecessorsValue = getCellValue(mapping.colPredecessors);
      if (predecessorsValue) {
        // Smartsheet predecessors format: "1, 2, 3" or "1FS, 2SS"
        const parts = String(predecessorsValue).split(',');
        for (const part of parts) {
          const match = part.trim().match(/^(\d+)/);
          if (match) {
            // We'll resolve these to task IDs after all rows are processed
            dependencies.push(`ROW_${match[1]}`);
          }
        }
      }
    }

    // Determine parent task from row hierarchy
    let parentTask = '';
    if (config.includeChildRows && row.parentId) {
      parentTask = rowIdToTaskId.get(String(row.parentId)) || '';
    }

    if (!name) {
      continue; // Skip rows without a name
    }

    tasks.push({
      id: taskId,
      name,
      startDate: startDate || new Date().toISOString().split('T')[0],
      endDate: endDate || startDate || new Date().toISOString().split('T')[0],
      duration,
      owner,
      percentComplete: Math.max(0, Math.min(100, percentComplete)),
      priority,
      smartsheetRowId: String(row.id),
      parentTask,
      dependencies,
      taskType: 'Task',
      swimlane,
    });
  }

  // Resolve row number dependencies to task IDs
  const rowNumberToTaskId = new Map<string, string>();
  rows.forEach((row, index) => {
    const taskId = rowIdToTaskId.get(String(row.id));
    if (taskId) {
      rowNumberToTaskId.set(String(index + 1), taskId);
    }
  });

  for (const task of tasks) {
    task.dependencies = task.dependencies.map(dep => {
      if (dep.startsWith('ROW_')) {
        const rowNum = dep.replace('ROW_', '');
        return rowNumberToTaskId.get(rowNum) || dep;
      }
      return dep;
    }).filter(dep => !dep.startsWith('ROW_'));
  }

  return tasks;
}

/**
 * Convert task to Smartsheet cells
 */
function taskToSmartsheetCells(
  task: Task,
  columns: SmartsheetColumn[],
  config: SmartsheetConfig
): SmartsheetCell[] {
  const mapping = config.columnMapping;
  const cells: SmartsheetCell[] = [];

  const addCell = (columnId: string | undefined, value: unknown) => {
    if (columnId && value !== undefined && value !== null && value !== '') {
      cells.push({
        columnId,
        value: value as string | number | boolean,
      });
    }
  };

  addCell(mapping.colTaskName, task.name);
  addCell(mapping.colStartDate, task.startDate);
  addCell(mapping.colEndDate, task.endDate);
  addCell(mapping.colDuration, task.duration);
  addCell(mapping.colOwner, task.owner);
  addCell(mapping.colPercentComplete, task.percentComplete);
  addCell(mapping.colPriority, task.priority);
  addCell(mapping.colCategory, task.swimlane);

  return cells;
}

export const smartsheetFunctions = {
  testConnection,
  saveConfig,
  getConfig,
  listSheets,
  getSheetColumns,
  pullFromSmartsheet,
  pushToSmartsheet,
  detectConflicts,
  resolveConflicts,
};
