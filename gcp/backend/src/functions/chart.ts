/**
 * Visual Gantt - Chart Functions
 * Cloud Functions for chart data preparation and export
 */

import { Response } from 'express';
import * as admin from 'firebase-admin';
import { AuthenticatedRequest, checkProjectAccess } from '../services/auth';
import {
  Task,
  Project,
  ProjectConfig,
  ChartData,
  PRIORITY_COLORS,
  TASK_TYPE_COLORS,
  SWIMLANE_COLORS,
  DEFAULT_PROJECT_CONFIG,
} from '../../../shared/types';

const db = admin.firestore();

/**
 * Get chart data for rendering
 */
export async function getChartData(req: AuthenticatedRequest, res: Response): Promise<void> {
  try {
    const { projectId } = req.params;
    const userId = req.userId!;

    const hasAccess = await checkProjectAccess(userId, projectId);
    if (!hasAccess) {
      res.status(403).json({ success: false, error: 'Access denied' });
      return;
    }

    // Get project config
    const projectDoc = await db.collection('projects').doc(projectId).get();
    const project = projectDoc.data() as Project;
    const config = project?.config || DEFAULT_PROJECT_CONFIG;

    // Get all tasks
    const tasksSnapshot = await db
      .collection('projects')
      .doc(projectId)
      .collection('tasks')
      .orderBy('startDate', 'asc')
      .get();

    const tasks = tasksSnapshot.docs.map(doc => ({
      id: doc.id,
      ...doc.data(),
    })) as Task[];

    if (tasks.length === 0) {
      res.status(400).json({
        success: false,
        error: 'No tasks found in this project',
      });
      return;
    }

    // Prepare tasks for chart
    const preparedTasks = prepareTasksForChart(tasks, config);

    // Calculate date range
    let startDate = config.startDate ? new Date(config.startDate) : null;
    let endDate = config.endDate ? new Date(config.endDate) : null;

    if (!startDate || !endDate) {
      const dates = tasks
        .flatMap(t => [new Date(t.startDate), new Date(t.endDate)])
        .filter(d => !isNaN(d.getTime()));

      if (!startDate) {
        startDate = new Date(Math.min(...dates.map(d => d.getTime())));
        startDate.setDate(startDate.getDate() - 7); // Add padding
      }
      if (!endDate) {
        endDate = new Date(Math.max(...dates.map(d => d.getTime())));
        endDate.setDate(endDate.getDate() + 7); // Add padding
      }
    }

    const chartData: ChartData = {
      tasks: preparedTasks,
      config,
      startDate: startDate.toISOString(),
      endDate: endDate.toISOString(),
      today: new Date().toISOString(),
    };

    res.json({
      success: true,
      data: chartData,
    });
  } catch (error) {
    console.error('Error getting chart data:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to get chart data',
    });
  }
}

/**
 * Export chart as PNG
 * Note: Actual rendering is done client-side, this provides the data
 */
export async function exportPng(req: AuthenticatedRequest, res: Response): Promise<void> {
  try {
    const { projectId } = req.params;
    const userId = req.userId!;
    const { imageData } = req.body;

    const hasAccess = await checkProjectAccess(userId, projectId);
    if (!hasAccess) {
      res.status(403).json({ success: false, error: 'Access denied' });
      return;
    }

    if (!imageData) {
      res.status(400).json({
        success: false,
        error: 'Image data is required',
      });
      return;
    }

    // Remove data URL prefix if present
    const base64Content = imageData.replace(/^data:image\/png;base64,/, '');

    // Store in Cloud Storage
    const bucket = admin.storage().bucket();
    const filename = `exports/${projectId}/${Date.now()}_gantt.png`;
    const file = bucket.file(filename);

    const buffer = Buffer.from(base64Content, 'base64');
    await file.save(buffer, {
      metadata: {
        contentType: 'image/png',
        metadata: {
          projectId,
          userId,
          exportedAt: new Date().toISOString(),
        },
      },
    });

    // Generate signed URL for download (valid for 1 hour)
    const [signedUrl] = await file.getSignedUrl({
      action: 'read',
      expires: Date.now() + 60 * 60 * 1000, // 1 hour
    });

    res.json({
      success: true,
      data: {
        url: signedUrl,
        filename: `visual_gantt_${projectId}.png`,
      },
    });
  } catch (error) {
    console.error('Error exporting PNG:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to export PNG',
    });
  }
}

/**
 * Export chart as PDF
 * Note: Actual rendering is done client-side, this provides the data
 */
export async function exportPdf(req: AuthenticatedRequest, res: Response): Promise<void> {
  try {
    const { projectId } = req.params;
    const userId = req.userId!;
    const { pdfData } = req.body;

    const hasAccess = await checkProjectAccess(userId, projectId);
    if (!hasAccess) {
      res.status(403).json({ success: false, error: 'Access denied' });
      return;
    }

    if (!pdfData) {
      res.status(400).json({
        success: false,
        error: 'PDF data is required',
      });
      return;
    }

    // Remove data URL prefix if present
    const base64Content = pdfData.replace(/^data:application\/pdf;base64,/, '');

    // Store in Cloud Storage
    const bucket = admin.storage().bucket();
    const filename = `exports/${projectId}/${Date.now()}_gantt.pdf`;
    const file = bucket.file(filename);

    const buffer = Buffer.from(base64Content, 'base64');
    await file.save(buffer, {
      metadata: {
        contentType: 'application/pdf',
        metadata: {
          projectId,
          userId,
          exportedAt: new Date().toISOString(),
        },
      },
    });

    // Generate signed URL for download (valid for 1 hour)
    const [signedUrl] = await file.getSignedUrl({
      action: 'read',
      expires: Date.now() + 60 * 60 * 1000, // 1 hour
    });

    res.json({
      success: true,
      data: {
        url: signedUrl,
        filename: `visual_gantt_${projectId}.pdf`,
      },
    });
  } catch (error) {
    console.error('Error exporting PDF:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to export PDF',
    });
  }
}

// ============================================================================
// Helper Functions
// ============================================================================

/**
 * Prepare tasks for chart rendering
 */
function prepareTasksForChart(tasks: Task[], config: ProjectConfig): Task[] {
  // Build task map for hierarchy detection
  const taskMap = new Map<string, Task>();
  tasks.forEach(task => taskMap.set(task.id, task));

  // Mark subtasks
  tasks.forEach(task => {
    if (task.parentTask && taskMap.has(task.parentTask)) {
      task.isSubtask = true;
    } else {
      task.isSubtask = false;
    }
  });

  // Sort tasks
  let sortedTasks = [...tasks];

  if (config.swimlaneGrouping !== 'None') {
    const groupField = config.swimlaneGrouping === 'Owner' ? 'owner' : 'swimlane';

    sortedTasks.sort((a, b) => {
      const groupA = (a[groupField as keyof Task] as string) || 'Other';
      const groupB = (b[groupField as keyof Task] as string) || 'Other';

      if (groupA !== groupB) {
        return groupA.localeCompare(groupB);
      }

      // Within same group, parent tasks first, then by start date
      if (a.isSubtask !== b.isSubtask) {
        return a.isSubtask ? 1 : -1;
      }

      return new Date(a.startDate).getTime() - new Date(b.startDate).getTime();
    });
  } else {
    sortedTasks.sort((a, b) =>
      new Date(a.startDate).getTime() - new Date(b.startDate).getTime()
    );
  }

  // Assign colors based on color scheme
  const swimlaneColorMap = new Map<string, string>();
  let swimlaneColorIndex = 0;

  sortedTasks.forEach(task => {
    switch (config.colorScheme) {
      case 'priority':
        task.color = PRIORITY_COLORS[task.priority] || PRIORITY_COLORS.default;
        break;
      case 'taskType':
        task.color = TASK_TYPE_COLORS[task.taskType] || TASK_TYPE_COLORS.default;
        break;
      case 'swimlane':
        const swimlane = task.swimlane || 'Default';
        if (!swimlaneColorMap.has(swimlane)) {
          swimlaneColorMap.set(swimlane, SWIMLANE_COLORS[swimlaneColorIndex % SWIMLANE_COLORS.length]);
          swimlaneColorIndex++;
        }
        task.color = swimlaneColorMap.get(swimlane);
        break;
      default:
        task.color = '#2196F3';
    }
  });

  return sortedTasks;
}

export const chartFunctions = {
  getChartData,
  exportPng,
  exportPdf,
};
