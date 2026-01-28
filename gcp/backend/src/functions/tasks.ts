/**
 * Visual Gantt - Task Functions
 * Cloud Functions for task management
 */

import { Response } from 'express';
import * as admin from 'firebase-admin';
import { v4 as uuidv4 } from 'uuid';
import { AuthenticatedRequest, checkProjectAccess } from '../services/auth';
import { Task, Priority, TaskType } from '../../../shared/types';

const db = admin.firestore();

interface ValidationResult {
  valid: boolean;
  errors: string[];
  warnings: string[];
}

/**
 * List all tasks for a project
 */
export async function listTasks(req: AuthenticatedRequest, res: Response): Promise<void> {
  try {
    const { projectId } = req.params;
    const userId = req.userId!;

    const hasAccess = await checkProjectAccess(userId, projectId);
    if (!hasAccess) {
      res.status(403).json({
        success: false,
        error: 'Access denied',
      });
      return;
    }

    const snapshot = await db
      .collection('projects')
      .doc(projectId)
      .collection('tasks')
      .orderBy('startDate', 'asc')
      .get();

    const tasks = snapshot.docs.map(doc => ({
      id: doc.id,
      ...doc.data(),
    })) as Task[];

    res.json({
      success: true,
      data: tasks,
    });
  } catch (error) {
    console.error('Error listing tasks:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to list tasks',
    });
  }
}

/**
 * Get a single task
 */
export async function getTask(req: AuthenticatedRequest, res: Response): Promise<void> {
  try {
    const { projectId, taskId } = req.params;
    const userId = req.userId!;

    const hasAccess = await checkProjectAccess(userId, projectId);
    if (!hasAccess) {
      res.status(403).json({
        success: false,
        error: 'Access denied',
      });
      return;
    }

    const doc = await db
      .collection('projects')
      .doc(projectId)
      .collection('tasks')
      .doc(taskId)
      .get();

    if (!doc.exists) {
      res.status(404).json({
        success: false,
        error: 'Task not found',
      });
      return;
    }

    res.json({
      success: true,
      data: { id: doc.id, ...doc.data() },
    });
  } catch (error) {
    console.error('Error getting task:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to get task',
    });
  }
}

/**
 * Create a new task
 */
export async function createTask(req: AuthenticatedRequest, res: Response): Promise<void> {
  try {
    const { projectId } = req.params;
    const userId = req.userId!;
    const taskData = req.body as Partial<Task>;

    const hasAccess = await checkProjectAccess(userId, projectId, 'editor');
    if (!hasAccess) {
      res.status(403).json({
        success: false,
        error: 'Access denied',
      });
      return;
    }

    // Validate required fields
    if (!taskData.name || typeof taskData.name !== 'string' || taskData.name.trim().length === 0) {
      res.status(400).json({
        success: false,
        error: 'Task name is required',
      });
      return;
    }

    if (!taskData.startDate) {
      res.status(400).json({
        success: false,
        error: 'Start date is required',
      });
      return;
    }

    const taskId = taskData.id || uuidv4();
    const now = new Date().toISOString();

    // Calculate end date from duration if not provided
    let endDate = taskData.endDate;
    if (!endDate && taskData.duration) {
      const start = new Date(taskData.startDate);
      start.setDate(start.getDate() + taskData.duration - 1);
      endDate = start.toISOString().split('T')[0];
    }

    // For milestones, end date equals start date
    if (taskData.taskType === 'Milestone') {
      endDate = taskData.startDate;
    }

    if (!endDate) {
      res.status(400).json({
        success: false,
        error: 'End date or duration is required',
      });
      return;
    }

    const task: Omit<Task, 'id'> = {
      name: taskData.name.trim(),
      startDate: taskData.startDate,
      endDate: endDate,
      duration: taskData.duration,
      owner: taskData.owner?.trim() || '',
      percentComplete: Math.max(0, Math.min(100, taskData.percentComplete || 0)),
      priority: validatePriority(taskData.priority),
      jiraTicket: taskData.jiraTicket?.trim() || '',
      smartsheetRowId: taskData.smartsheetRowId || '',
      parentTask: taskData.parentTask?.trim() || '',
      dependencies: taskData.dependencies || [],
      taskType: validateTaskType(taskData.taskType),
      swimlane: taskData.swimlane?.trim() || 'Default',
      createdAt: now,
      updatedAt: now,
      modified: false,
    };

    await db
      .collection('projects')
      .doc(projectId)
      .collection('tasks')
      .doc(taskId)
      .set(task);

    res.status(201).json({
      success: true,
      data: { id: taskId, ...task },
    });
  } catch (error) {
    console.error('Error creating task:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to create task',
    });
  }
}

/**
 * Create multiple tasks in bulk
 */
export async function createTasksBulk(req: AuthenticatedRequest, res: Response): Promise<void> {
  try {
    const { projectId } = req.params;
    const userId = req.userId!;
    const { tasks: tasksData, replaceAll } = req.body as { tasks: Partial<Task>[]; replaceAll?: boolean };

    const hasAccess = await checkProjectAccess(userId, projectId, 'editor');
    if (!hasAccess) {
      res.status(403).json({
        success: false,
        error: 'Access denied',
      });
      return;
    }

    if (!Array.isArray(tasksData) || tasksData.length === 0) {
      res.status(400).json({
        success: false,
        error: 'Tasks array is required',
      });
      return;
    }

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

    const createdTasks: Task[] = [];

    for (const taskData of tasksData) {
      if (!taskData.name || !taskData.startDate) {
        continue; // Skip invalid tasks
      }

      const taskId = taskData.id || uuidv4();

      let endDate = taskData.endDate;
      if (!endDate && taskData.duration) {
        const start = new Date(taskData.startDate);
        start.setDate(start.getDate() + taskData.duration - 1);
        endDate = start.toISOString().split('T')[0];
      }
      if (taskData.taskType === 'Milestone') {
        endDate = taskData.startDate;
      }
      if (!endDate) {
        endDate = taskData.startDate;
      }

      const task: Task = {
        id: taskId,
        name: taskData.name.trim(),
        startDate: taskData.startDate,
        endDate: endDate,
        duration: taskData.duration,
        owner: taskData.owner?.trim() || '',
        percentComplete: Math.max(0, Math.min(100, taskData.percentComplete || 0)),
        priority: validatePriority(taskData.priority),
        jiraTicket: taskData.jiraTicket?.trim() || '',
        smartsheetRowId: taskData.smartsheetRowId || '',
        parentTask: taskData.parentTask?.trim() || '',
        dependencies: taskData.dependencies || [],
        taskType: validateTaskType(taskData.taskType),
        swimlane: taskData.swimlane?.trim() || 'Default',
        createdAt: now,
        updatedAt: now,
        modified: false,
      };

      batch.set(tasksRef.doc(taskId), task);
      createdTasks.push(task);
    }

    await batch.commit();

    res.status(201).json({
      success: true,
      data: createdTasks,
      message: `${createdTasks.length} tasks created`,
    });
  } catch (error) {
    console.error('Error creating tasks in bulk:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to create tasks',
    });
  }
}

/**
 * Update a task
 */
export async function updateTask(req: AuthenticatedRequest, res: Response): Promise<void> {
  try {
    const { projectId, taskId } = req.params;
    const userId = req.userId!;
    const updates = req.body as Partial<Task>;

    const hasAccess = await checkProjectAccess(userId, projectId, 'editor');
    if (!hasAccess) {
      res.status(403).json({
        success: false,
        error: 'Access denied',
      });
      return;
    }

    const taskRef = db
      .collection('projects')
      .doc(projectId)
      .collection('tasks')
      .doc(taskId);

    const doc = await taskRef.get();
    if (!doc.exists) {
      res.status(404).json({
        success: false,
        error: 'Task not found',
      });
      return;
    }

    // Build update object
    const updateData: Partial<Task> = {
      updatedAt: new Date().toISOString(),
      modified: true, // Mark as modified for sync
    };

    if (updates.name !== undefined) {
      updateData.name = updates.name.trim();
    }
    if (updates.startDate !== undefined) {
      updateData.startDate = updates.startDate;
    }
    if (updates.endDate !== undefined) {
      updateData.endDate = updates.endDate;
    }
    if (updates.duration !== undefined) {
      updateData.duration = updates.duration;
    }
    if (updates.owner !== undefined) {
      updateData.owner = updates.owner.trim();
    }
    if (updates.percentComplete !== undefined) {
      updateData.percentComplete = Math.max(0, Math.min(100, updates.percentComplete));
    }
    if (updates.priority !== undefined) {
      updateData.priority = validatePriority(updates.priority);
    }
    if (updates.jiraTicket !== undefined) {
      updateData.jiraTicket = updates.jiraTicket.trim();
    }
    if (updates.smartsheetRowId !== undefined) {
      updateData.smartsheetRowId = updates.smartsheetRowId;
    }
    if (updates.parentTask !== undefined) {
      updateData.parentTask = updates.parentTask.trim();
    }
    if (updates.dependencies !== undefined) {
      updateData.dependencies = updates.dependencies;
    }
    if (updates.taskType !== undefined) {
      updateData.taskType = validateTaskType(updates.taskType);
    }
    if (updates.swimlane !== undefined) {
      updateData.swimlane = updates.swimlane.trim();
    }

    await taskRef.update(updateData);

    const updatedDoc = await taskRef.get();

    res.json({
      success: true,
      data: { id: updatedDoc.id, ...updatedDoc.data() },
    });
  } catch (error) {
    console.error('Error updating task:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to update task',
    });
  }
}

/**
 * Delete a task
 */
export async function deleteTask(req: AuthenticatedRequest, res: Response): Promise<void> {
  try {
    const { projectId, taskId } = req.params;
    const userId = req.userId!;

    const hasAccess = await checkProjectAccess(userId, projectId, 'editor');
    if (!hasAccess) {
      res.status(403).json({
        success: false,
        error: 'Access denied',
      });
      return;
    }

    const taskRef = db
      .collection('projects')
      .doc(projectId)
      .collection('tasks')
      .doc(taskId);

    const doc = await taskRef.get();
    if (!doc.exists) {
      res.status(404).json({
        success: false,
        error: 'Task not found',
      });
      return;
    }

    await taskRef.delete();

    res.json({
      success: true,
      message: 'Task deleted successfully',
    });
  } catch (error) {
    console.error('Error deleting task:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to delete task',
    });
  }
}

/**
 * Validate tasks (check dependencies, dates, circular references)
 */
export async function validateTasks(req: AuthenticatedRequest, res: Response): Promise<void> {
  try {
    const { projectId } = req.params;
    const userId = req.userId!;

    const hasAccess = await checkProjectAccess(userId, projectId);
    if (!hasAccess) {
      res.status(403).json({
        success: false,
        error: 'Access denied',
      });
      return;
    }

    const snapshot = await db
      .collection('projects')
      .doc(projectId)
      .collection('tasks')
      .get();

    const tasks = snapshot.docs.map(doc => ({
      id: doc.id,
      ...doc.data(),
    })) as Task[];

    const result = validateTaskList(tasks);

    res.json({
      success: true,
      data: result,
    });
  } catch (error) {
    console.error('Error validating tasks:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to validate tasks',
    });
  }
}

/**
 * Validate a list of tasks
 */
function validateTaskList(tasks: Task[]): ValidationResult {
  const errors: string[] = [];
  const warnings: string[] = [];

  // Build task map for dependency checking
  const taskMap = new Map<string, Task>();
  tasks.forEach(task => taskMap.set(task.id, task));

  // Validate each task
  for (const task of tasks) {
    // Check required fields
    if (!task.name) {
      errors.push(`Task ${task.id}: Name is required`);
    }
    if (!task.startDate) {
      errors.push(`Task ${task.id}: Start date is required`);
    }
    if (!task.endDate) {
      errors.push(`Task ${task.id}: End date is required`);
    }

    // Check date order
    if (task.startDate && task.endDate) {
      const start = new Date(task.startDate);
      const end = new Date(task.endDate);
      if (start > end) {
        errors.push(`Task "${task.name}": End date must be after start date`);
      }
    }

    // Check dependencies exist
    if (task.dependencies) {
      for (const depId of task.dependencies) {
        if (!taskMap.has(depId)) {
          errors.push(`Task "${task.name}": Dependency "${depId}" not found`);
        }
      }
    }

    // Check parent exists
    if (task.parentTask && !taskMap.has(task.parentTask)) {
      warnings.push(`Task "${task.name}": Parent task "${task.parentTask}" not found`);
    }

    // Check percent complete range
    if (task.percentComplete < 0 || task.percentComplete > 100) {
      warnings.push(`Task "${task.name}": Percent complete should be between 0 and 100`);
    }
  }

  // Check for circular dependencies
  const circularErrors = detectCircularDependencies(tasks, taskMap);
  errors.push(...circularErrors);

  return {
    valid: errors.length === 0,
    errors,
    warnings,
  };
}

/**
 * Detect circular dependencies
 */
function detectCircularDependencies(tasks: Task[], taskMap: Map<string, Task>): string[] {
  const errors: string[] = [];
  const visited = new Set<string>();
  const recursionStack = new Set<string>();

  function hasCycle(taskId: string, path: string[]): boolean {
    if (recursionStack.has(taskId)) {
      errors.push(`Circular dependency detected: ${path.join(' -> ')} -> ${taskId}`);
      return true;
    }

    if (visited.has(taskId)) {
      return false;
    }

    visited.add(taskId);
    recursionStack.add(taskId);

    const task = taskMap.get(taskId);
    if (task?.dependencies) {
      for (const depId of task.dependencies) {
        if (taskMap.has(depId) && hasCycle(depId, [...path, taskId])) {
          return true;
        }
      }
    }

    recursionStack.delete(taskId);
    return false;
  }

  for (const task of tasks) {
    if (!visited.has(task.id)) {
      hasCycle(task.id, []);
    }
  }

  return errors;
}

/**
 * Validate and normalize priority
 */
function validatePriority(priority?: string): Priority {
  const validPriorities: Priority[] = ['High', 'Medium', 'Low'];
  if (priority && validPriorities.includes(priority as Priority)) {
    return priority as Priority;
  }
  return 'Medium';
}

/**
 * Validate and normalize task type
 */
function validateTaskType(taskType?: string): TaskType {
  const validTypes: TaskType[] = ['Task', 'Milestone'];
  if (taskType && validTypes.includes(taskType as TaskType)) {
    return taskType as TaskType;
  }
  return 'Task';
}

export const taskFunctions = {
  listTasks,
  getTask,
  createTask,
  createTasksBulk,
  updateTask,
  deleteTask,
  validateTasks,
};
