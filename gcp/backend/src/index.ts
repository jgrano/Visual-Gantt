/**
 * Visual Gantt - GCP Cloud Functions Entry Point
 * Main entry point for all Cloud Functions
 */

import * as functions from 'firebase-functions';
import * as admin from 'firebase-admin';
import express from 'express';
import cors from 'cors';

// Initialize Firebase Admin
admin.initializeApp();

// Import function modules
import { projectFunctions } from './functions/projects';
import { taskFunctions } from './functions/tasks';
import { jiraFunctions } from './functions/jira';
import { smartsheetFunctions } from './functions/smartsheet';
import { chartFunctions } from './functions/chart';
import { authMiddleware } from './services/auth';

// Create Express app
const app = express();

// Middleware
app.use(cors({ origin: true }));
app.use(express.json({ limit: '10mb' }));

// Auth middleware for all routes
app.use(authMiddleware);

// Health check
app.get('/health', (_req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

// ============================================================================
// PROJECT ROUTES
// ============================================================================

// Get all projects for user
app.get('/projects', projectFunctions.listProjects);

// Get single project
app.get('/projects/:projectId', projectFunctions.getProject);

// Create project
app.post('/projects', projectFunctions.createProject);

// Update project
app.put('/projects/:projectId', projectFunctions.updateProject);

// Delete project
app.delete('/projects/:projectId', projectFunctions.deleteProject);

// Update project config
app.put('/projects/:projectId/config', projectFunctions.updateProjectConfig);

// ============================================================================
// TASK ROUTES
// ============================================================================

// Get all tasks for project
app.get('/projects/:projectId/tasks', taskFunctions.listTasks);

// Get single task
app.get('/projects/:projectId/tasks/:taskId', taskFunctions.getTask);

// Create task
app.post('/projects/:projectId/tasks', taskFunctions.createTask);

// Create multiple tasks (bulk import)
app.post('/projects/:projectId/tasks/bulk', taskFunctions.createTasksBulk);

// Update task
app.put('/projects/:projectId/tasks/:taskId', taskFunctions.updateTask);

// Delete task
app.delete('/projects/:projectId/tasks/:taskId', taskFunctions.deleteTask);

// Validate tasks (check dependencies, dates, etc.)
app.post('/projects/:projectId/tasks/validate', taskFunctions.validateTasks);

// ============================================================================
// JIRA ROUTES
// ============================================================================

// Test Jira connection
app.post('/projects/:projectId/jira/test', jiraFunctions.testConnection);

// Save Jira config
app.put('/projects/:projectId/jira/config', jiraFunctions.saveConfig);

// Get Jira config (masked token)
app.get('/projects/:projectId/jira/config', jiraFunctions.getConfig);

// Pull from Jira
app.post('/projects/:projectId/jira/pull', jiraFunctions.pullFromJira);

// Push to Jira
app.post('/projects/:projectId/jira/push', jiraFunctions.pushToJira);

// Detect conflicts
app.post('/projects/:projectId/jira/conflicts', jiraFunctions.detectConflicts);

// Resolve conflicts
app.post('/projects/:projectId/jira/resolve', jiraFunctions.resolveConflicts);

// ============================================================================
// SMARTSHEET ROUTES
// ============================================================================

// Test Smartsheet connection
app.post('/projects/:projectId/smartsheet/test', smartsheetFunctions.testConnection);

// Save Smartsheet config
app.put('/projects/:projectId/smartsheet/config', smartsheetFunctions.saveConfig);

// Get Smartsheet config (masked token)
app.get('/projects/:projectId/smartsheet/config', smartsheetFunctions.getConfig);

// List available sheets
app.get('/projects/:projectId/smartsheet/sheets', smartsheetFunctions.listSheets);

// Get sheet columns
app.get('/projects/:projectId/smartsheet/sheets/:sheetId/columns', smartsheetFunctions.getSheetColumns);

// Pull from Smartsheet
app.post('/projects/:projectId/smartsheet/pull', smartsheetFunctions.pullFromSmartsheet);

// Push to Smartsheet
app.post('/projects/:projectId/smartsheet/push', smartsheetFunctions.pushToSmartsheet);

// Detect conflicts
app.post('/projects/:projectId/smartsheet/conflicts', smartsheetFunctions.detectConflicts);

// Resolve conflicts
app.post('/projects/:projectId/smartsheet/resolve', smartsheetFunctions.resolveConflicts);

// ============================================================================
// CHART ROUTES
// ============================================================================

// Get chart data (prepared tasks for rendering)
app.get('/projects/:projectId/chart', chartFunctions.getChartData);

// Export chart as PNG (returns base64)
app.post('/projects/:projectId/chart/export/png', chartFunctions.exportPng);

// Export chart as PDF (returns base64)
app.post('/projects/:projectId/chart/export/pdf', chartFunctions.exportPdf);

// ============================================================================
// ERROR HANDLING
// ============================================================================

// 404 handler
app.use((_req, res) => {
  res.status(404).json({
    success: false,
    error: 'Endpoint not found',
  });
});

// Error handler
app.use((err: Error, _req: express.Request, res: express.Response, _next: express.NextFunction) => {
  console.error('API Error:', err);
  res.status(500).json({
    success: false,
    error: err.message || 'Internal server error',
  });
});

// Export the Express app as a Cloud Function
export const api = functions.https.onRequest(app);

// ============================================================================
// SCHEDULED FUNCTIONS
// ============================================================================

// Cleanup old exports (runs daily)
export const cleanupExports = functions.pubsub
  .schedule('every 24 hours')
  .onRun(async () => {
    const bucket = admin.storage().bucket();
    const [files] = await bucket.getFiles({ prefix: 'exports/' });

    const oneDayAgo = Date.now() - 24 * 60 * 60 * 1000;

    for (const file of files) {
      const [metadata] = await file.getMetadata();
      const created = new Date(metadata.timeCreated as string).getTime();

      if (created < oneDayAgo) {
        await file.delete();
        console.log(`Deleted old export: ${file.name}`);
      }
    }

    return null;
  });

// ============================================================================
// FIRESTORE TRIGGERS
// ============================================================================

// Update project timestamp when tasks change
export const onTaskWrite = functions.firestore
  .document('projects/{projectId}/tasks/{taskId}')
  .onWrite(async (change, context) => {
    const { projectId } = context.params;

    await admin.firestore()
      .collection('projects')
      .doc(projectId)
      .update({
        updatedAt: admin.firestore.FieldValue.serverTimestamp(),
      });

    return null;
  });
