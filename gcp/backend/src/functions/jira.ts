/**
 * Visual Gantt - Jira Integration Functions
 * Cloud Functions for Jira Cloud integration
 */

import { Response } from 'express';
import * as admin from 'firebase-admin';
import axios, { AxiosError } from 'axios';
import { AuthenticatedRequest, checkProjectAccess } from '../services/auth';
import {
  JiraConfig,
  JiraIssue,
  JiraSyncResult,
  Task,
  SyncConflict,
  SyncResolution,
  JIRA_STATUS_MAPPING,
  JIRA_PRIORITY_MAPPING,
  Priority,
} from '../../../shared/types';

const db = admin.firestore();

interface JiraSearchResponse {
  issues: JiraIssue[];
  total: number;
  startAt: number;
  maxResults: number;
}

/**
 * Test Jira connection
 */
export async function testConnection(req: AuthenticatedRequest, res: Response): Promise<void> {
  try {
    const { projectId } = req.params;
    const userId = req.userId!;
    const { baseUrl, email, apiToken } = req.body;

    const hasAccess = await checkProjectAccess(userId, projectId, 'editor');
    if (!hasAccess) {
      res.status(403).json({ success: false, error: 'Access denied' });
      return;
    }

    if (!baseUrl || !email || !apiToken) {
      res.status(400).json({
        success: false,
        error: 'Base URL, email, and API token are required',
      });
      return;
    }

    const response = await jiraApiRequest(
      baseUrl,
      email,
      apiToken,
      '/rest/api/3/myself'
    );

    res.json({
      success: true,
      message: `Connected as ${response.displayName} (${response.emailAddress})`,
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
 * Save Jira configuration
 */
export async function saveConfig(req: AuthenticatedRequest, res: Response): Promise<void> {
  try {
    const { projectId } = req.params;
    const userId = req.userId!;
    const config = req.body as Partial<JiraConfig>;

    const hasAccess = await checkProjectAccess(userId, projectId, 'editor');
    if (!hasAccess) {
      res.status(403).json({ success: false, error: 'Access denied' });
      return;
    }

    // Get existing config to preserve API token if not provided
    const projectDoc = await db.collection('projects').doc(projectId).get();
    const project = projectDoc.data();
    const existingConfig = project?.jiraConfig as JiraConfig | undefined;

    // If token is masked, keep existing token
    let apiToken = config.apiToken;
    if (apiToken === '••••••••' && existingConfig?.apiToken) {
      apiToken = existingConfig.apiToken;
    }

    const jiraConfig: JiraConfig = {
      baseUrl: config.baseUrl || '',
      email: config.email || '',
      apiToken: apiToken || '',
      pullMethod: config.pullMethod || 'jql',
      jqlQuery: config.jqlQuery || '',
      projectKey: config.projectKey || '',
      filterId: config.filterId || '',
      startDateField: config.startDateField || '',
      storyPointsField: config.storyPointsField || '',
      lastSyncTime: existingConfig?.lastSyncTime,
    };

    await db.collection('projects').doc(projectId).update({
      jiraConfig,
      updatedAt: new Date().toISOString(),
    });

    res.json({ success: true, message: 'Jira configuration saved' });
  } catch (error) {
    console.error('Error saving Jira config:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to save Jira configuration',
    });
  }
}

/**
 * Get Jira configuration (with masked token)
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
    const config = project?.jiraConfig as JiraConfig | undefined;

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
    console.error('Error getting Jira config:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to get Jira configuration',
    });
  }
}

/**
 * Pull issues from Jira
 */
export async function pullFromJira(req: AuthenticatedRequest, res: Response): Promise<void> {
  try {
    const { projectId } = req.params;
    const userId = req.userId!;
    const { replaceAll } = req.body;

    const hasAccess = await checkProjectAccess(userId, projectId, 'editor');
    if (!hasAccess) {
      res.status(403).json({ success: false, error: 'Access denied' });
      return;
    }

    // Get Jira config
    const projectDoc = await db.collection('projects').doc(projectId).get();
    const project = projectDoc.data();
    const config = project?.jiraConfig as JiraConfig | undefined;

    if (!config || !config.baseUrl || !config.email || !config.apiToken) {
      res.status(400).json({
        success: false,
        error: 'Jira is not configured for this project',
      });
      return;
    }

    // Fetch issues from Jira
    const issues = await fetchJiraIssues(config);

    // Convert issues to tasks
    const tasks = issues.map(issue => jiraIssueToTask(issue, config));

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
        // Update existing task
        batch.update(tasksRef.doc(task.id), {
          ...task,
          updatedAt: now,
          modified: false, // Clear modified flag after sync
        });
        updated++;
      } else {
        // Create new task
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
      'jiraConfig.lastSyncTime': now,
      updatedAt: now,
    });

    await batch.commit();

    const result: JiraSyncResult = {
      success: true,
      issuesProcessed: issues.length,
      issuesCreated: created,
      issuesUpdated: updated,
      errors: [],
    };

    res.json({
      success: true,
      data: result,
    });
  } catch (error) {
    console.error('Error pulling from Jira:', error);
    const message = error instanceof Error ? error.message : 'Failed to pull from Jira';
    res.status(500).json({
      success: false,
      error: message,
    });
  }
}

/**
 * Push changes to Jira
 */
export async function pushToJira(req: AuthenticatedRequest, res: Response): Promise<void> {
  try {
    const { projectId } = req.params;
    const userId = req.userId!;
    const { taskIds } = req.body;

    const hasAccess = await checkProjectAccess(userId, projectId, 'editor');
    if (!hasAccess) {
      res.status(403).json({ success: false, error: 'Access denied' });
      return;
    }

    // Get Jira config
    const projectDoc = await db.collection('projects').doc(projectId).get();
    const project = projectDoc.data();
    const config = project?.jiraConfig as JiraConfig | undefined;

    if (!config || !config.baseUrl || !config.email || !config.apiToken) {
      res.status(400).json({
        success: false,
        error: 'Jira is not configured for this project',
      });
      return;
    }

    // Get tasks to push
    const tasksRef = db.collection('projects').doc(projectId).collection('tasks');
    let tasksSnapshot;

    if (taskIds && taskIds.length > 0) {
      // Push specific tasks
      tasksSnapshot = await tasksRef.where(admin.firestore.FieldPath.documentId(), 'in', taskIds).get();
    } else {
      // Push all modified tasks
      tasksSnapshot = await tasksRef.where('modified', '==', true).get();
    }

    const tasks = tasksSnapshot.docs.map(doc => ({
      id: doc.id,
      ...doc.data(),
    })) as Task[];

    const errors: string[] = [];
    let updated = 0;

    for (const task of tasks) {
      // Only push tasks that have a Jira ticket
      if (!task.jiraTicket && !task.id.includes('-')) {
        continue;
      }

      const issueKey = task.id;

      try {
        // Update issue in Jira
        await updateJiraIssue(config, issueKey, task);

        // Clear modified flag
        await tasksRef.doc(task.id).update({
          modified: false,
          updatedAt: new Date().toISOString(),
        });

        updated++;
      } catch (error) {
        const message = error instanceof Error ? error.message : 'Unknown error';
        errors.push(`Failed to update ${issueKey}: ${message}`);
      }
    }

    // Update last sync time
    await db.collection('projects').doc(projectId).update({
      'jiraConfig.lastSyncTime': new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    });

    const result: JiraSyncResult = {
      success: errors.length === 0,
      issuesProcessed: tasks.length,
      issuesCreated: 0,
      issuesUpdated: updated,
      errors,
    };

    res.json({
      success: true,
      data: result,
    });
  } catch (error) {
    console.error('Error pushing to Jira:', error);
    const message = error instanceof Error ? error.message : 'Failed to push to Jira';
    res.status(500).json({
      success: false,
      error: message,
    });
  }
}

/**
 * Detect conflicts between local and Jira data
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

    // Get Jira config
    const projectDoc = await db.collection('projects').doc(projectId).get();
    const project = projectDoc.data();
    const config = project?.jiraConfig as JiraConfig | undefined;

    if (!config || !config.baseUrl || !config.email || !config.apiToken) {
      res.status(400).json({
        success: false,
        error: 'Jira is not configured for this project',
      });
      return;
    }

    // Fetch current Jira issues
    const jiraIssues = await fetchJiraIssues(config);
    const jiraTasksMap = new Map<string, Task>();
    jiraIssues.forEach(issue => {
      const task = jiraIssueToTask(issue, config);
      jiraTasksMap.set(task.id, task);
    });

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
      const jiraTask = jiraTasksMap.get(localTask.id);
      if (!jiraTask) {
        continue;
      }

      // Check for field conflicts
      const fields = ['name', 'endDate', 'percentComplete', 'priority', 'owner'] as const;

      for (const field of fields) {
        const localValue = String(localTask[field] || '');
        const remoteValue = String(jiraTask[field] || '');

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
    console.error('Error detecting conflicts:', error);
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

    // Get Jira config
    const projectDoc = await db.collection('projects').doc(projectId).get();
    const project = projectDoc.data();
    const config = project?.jiraConfig as JiraConfig | undefined;

    if (!config) {
      res.status(400).json({
        success: false,
        error: 'Jira is not configured for this project',
      });
      return;
    }

    const batch = db.batch();
    const tasksRef = db.collection('projects').doc(projectId).collection('tasks');

    for (const resolution of resolutions) {
      const taskRef = tasksRef.doc(resolution.taskId);

      if (resolution.resolution === 'remote') {
        // Fetch fresh data from Jira and update local
        const jiraIssues = await fetchJiraIssues(config);
        const issue = jiraIssues.find(i => i.key === resolution.taskId);
        if (issue) {
          const task = jiraIssueToTask(issue, config);
          batch.update(taskRef, {
            ...task,
            modified: false,
            updatedAt: new Date().toISOString(),
          });
        }
      } else if (resolution.resolution === 'local') {
        // Keep local, but mark as needing push
        batch.update(taskRef, {
          modified: true,
          updatedAt: new Date().toISOString(),
        });
      }
      // For 'merge', the frontend should have already sent the merged value
    }

    await batch.commit();

    res.json({
      success: true,
      message: 'Conflicts resolved',
    });
  } catch (error) {
    console.error('Error resolving conflicts:', error);
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
 * Make authenticated request to Jira API
 */
async function jiraApiRequest(
  baseUrl: string,
  email: string,
  apiToken: string,
  endpoint: string,
  method: 'GET' | 'POST' | 'PUT' | 'DELETE' = 'GET',
  data?: unknown
): Promise<any> {
  const url = baseUrl.startsWith('http')
    ? `${baseUrl.replace(/\/$/, '')}${endpoint}`
    : `https://${baseUrl.replace(/\/$/, '')}${endpoint}`;

  const auth = Buffer.from(`${email}:${apiToken}`).toString('base64');

  try {
    const response = await axios({
      method,
      url,
      headers: {
        Authorization: `Basic ${auth}`,
        'Content-Type': 'application/json',
        Accept: 'application/json',
      },
      data,
    });

    return response.data;
  } catch (error) {
    if (error instanceof AxiosError) {
      const message = error.response?.data?.errorMessages?.join(', ') ||
        error.response?.data?.message ||
        error.message;
      throw new Error(message);
    }
    throw error;
  }
}

/**
 * Fetch issues from Jira based on configuration
 */
async function fetchJiraIssues(config: JiraConfig): Promise<JiraIssue[]> {
  let jql = '';

  switch (config.pullMethod) {
    case 'jql':
      jql = config.jqlQuery || '';
      break;
    case 'project':
      jql = `project = "${config.projectKey}" ORDER BY key ASC`;
      break;
    case 'filter':
      // Get filter JQL
      const filter = await jiraApiRequest(
        config.baseUrl,
        config.email,
        config.apiToken,
        `/rest/api/3/filter/${config.filterId}`
      );
      jql = filter.jql;
      break;
    default:
      throw new Error('Invalid pull method configured');
  }

  if (!jql) {
    throw new Error('No JQL query, project key, or filter ID configured');
  }

  // Fetch issues with pagination
  const allIssues: JiraIssue[] = [];
  let startAt = 0;
  const maxResults = 100;

  const fields = [
    'key', 'summary', 'status', 'priority', 'assignee', 'duedate',
    'issuetype', 'parent', 'issuelinks', 'labels', 'components',
  ];

  if (config.startDateField) {
    fields.push(config.startDateField);
  }

  do {
    const response = await jiraApiRequest(
      config.baseUrl,
      config.email,
      config.apiToken,
      `/rest/api/3/search?jql=${encodeURIComponent(jql)}&startAt=${startAt}&maxResults=${maxResults}&fields=${fields.join(',')}`
    ) as JiraSearchResponse;

    allIssues.push(...response.issues);
    startAt += response.issues.length;

    if (startAt >= response.total) {
      break;
    }
  } while (true);

  return allIssues;
}

/**
 * Convert Jira issue to Task
 */
function jiraIssueToTask(issue: JiraIssue, config: JiraConfig): Task {
  const fields = issue.fields;

  // Extract dates
  const startDate = config.startDateField && fields[config.startDateField]
    ? String(fields[config.startDateField])
    : undefined;
  const endDate = fields.duedate || undefined;

  // Map status to percent complete
  const statusCategory = fields.status?.statusCategory?.key || 'new';
  const percentComplete = JIRA_STATUS_MAPPING[statusCategory] ?? 0;

  // Map priority
  const jiraPriority = fields.priority?.name || 'Medium';
  const priority: Priority = JIRA_PRIORITY_MAPPING[jiraPriority] || 'Medium';

  // Build Jira ticket URL
  const baseUrl = config.baseUrl.replace(/\/$/, '');
  const jiraUrl = baseUrl.startsWith('http')
    ? `${baseUrl}/browse/${issue.key}`
    : `https://${baseUrl}/browse/${issue.key}`;

  // Extract dependencies from issue links
  const dependencies: string[] = [];
  if (fields.issuelinks) {
    fields.issuelinks.forEach(link => {
      if (link.inwardIssue &&
          (link.type.inward === 'is blocked by' ||
           link.type.inward === 'depends on' ||
           link.type.inward === 'is caused by')) {
        dependencies.push(link.inwardIssue.key);
      }
    });
  }

  // Determine task type
  const issueTypeName = fields.issuetype?.name || 'Task';
  const taskType = issueTypeName.toLowerCase().includes('milestone') ? 'Milestone' : 'Task';

  // Extract swimlane from labels or components
  const swimlane = fields.labels?.[0] || fields.components?.[0]?.name || 'Default';

  return {
    id: issue.key,
    name: fields.summary || '',
    startDate: startDate || new Date().toISOString().split('T')[0],
    endDate: endDate || new Date().toISOString().split('T')[0],
    owner: fields.assignee?.displayName || '',
    percentComplete,
    priority,
    jiraTicket: jiraUrl,
    parentTask: fields.parent?.key || '',
    dependencies,
    taskType,
    swimlane,
  };
}

/**
 * Update Jira issue with task data
 */
async function updateJiraIssue(config: JiraConfig, issueKey: string, task: Task): Promise<void> {
  const updateData: Record<string, unknown> = {
    fields: {
      duedate: task.endDate || null,
    },
  };

  // Update priority if changed
  const priorityMap: Record<Priority, string> = {
    High: 'High',
    Medium: 'Medium',
    Low: 'Low',
  };
  if (task.priority) {
    updateData.fields = {
      ...updateData.fields as Record<string, unknown>,
      priority: { name: priorityMap[task.priority] },
    };
  }

  // Update start date if custom field is configured
  if (config.startDateField && task.startDate) {
    updateData.fields = {
      ...updateData.fields as Record<string, unknown>,
      [config.startDateField]: task.startDate,
    };
  }

  await jiraApiRequest(
    config.baseUrl,
    config.email,
    config.apiToken,
    `/rest/api/3/issue/${issueKey}`,
    'PUT',
    updateData
  );

  // Update assignee if changed
  if (task.owner) {
    try {
      // Search for user by display name
      const users = await jiraApiRequest(
        config.baseUrl,
        config.email,
        config.apiToken,
        `/rest/api/3/user/search?query=${encodeURIComponent(task.owner)}`
      );

      if (users && users.length > 0) {
        await jiraApiRequest(
          config.baseUrl,
          config.email,
          config.apiToken,
          `/rest/api/3/issue/${issueKey}/assignee`,
          'PUT',
          { accountId: users[0].accountId }
        );
      }
    } catch (error) {
      // Ignore assignee errors
      console.warn(`Could not update assignee for ${issueKey}`);
    }
  }
}

export const jiraFunctions = {
  testConnection,
  saveConfig,
  getConfig,
  pullFromJira,
  pushToJira,
  detectConflicts,
  resolveConflicts,
};
