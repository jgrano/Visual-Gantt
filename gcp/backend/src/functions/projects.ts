/**
 * Visual Gantt - Project Functions
 * Cloud Functions for project management
 */

import { Response } from 'express';
import * as admin from 'firebase-admin';
import { v4 as uuidv4 } from 'uuid';
import { AuthenticatedRequest, checkProjectAccess } from '../services/auth';
import { Project, ProjectConfig, DEFAULT_PROJECT_CONFIG } from '../../../shared/types';

const db = admin.firestore();

/**
 * List all projects for the authenticated user
 */
export async function listProjects(req: AuthenticatedRequest, res: Response): Promise<void> {
  try {
    const userId = req.userId!;

    // Query projects where user is owner or member
    const ownedProjects = await db
      .collection('projects')
      .where('ownerId', '==', userId)
      .orderBy('updatedAt', 'desc')
      .get();

    const memberProjects = await db
      .collection('projects')
      .where('members', 'array-contains-any', [
        { userId, role: 'editor' },
        { userId, role: 'viewer' },
      ])
      .orderBy('updatedAt', 'desc')
      .get();

    const projectsMap = new Map<string, Project>();

    ownedProjects.docs.forEach(doc => {
      projectsMap.set(doc.id, { id: doc.id, ...doc.data() } as Project);
    });

    memberProjects.docs.forEach(doc => {
      if (!projectsMap.has(doc.id)) {
        projectsMap.set(doc.id, { id: doc.id, ...doc.data() } as Project);
      }
    });

    const projects = Array.from(projectsMap.values());

    res.json({
      success: true,
      data: projects,
    });
  } catch (error) {
    console.error('Error listing projects:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to list projects',
    });
  }
}

/**
 * Get a single project by ID
 */
export async function getProject(req: AuthenticatedRequest, res: Response): Promise<void> {
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

    const doc = await db.collection('projects').doc(projectId).get();

    if (!doc.exists) {
      res.status(404).json({
        success: false,
        error: 'Project not found',
      });
      return;
    }

    res.json({
      success: true,
      data: { id: doc.id, ...doc.data() },
    });
  } catch (error) {
    console.error('Error getting project:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to get project',
    });
  }
}

/**
 * Create a new project
 */
export async function createProject(req: AuthenticatedRequest, res: Response): Promise<void> {
  try {
    const userId = req.userId!;
    const { name, description } = req.body;

    if (!name || typeof name !== 'string' || name.trim().length === 0) {
      res.status(400).json({
        success: false,
        error: 'Project name is required',
      });
      return;
    }

    const projectId = uuidv4();
    const now = new Date().toISOString();

    const project: Omit<Project, 'id'> = {
      name: name.trim(),
      description: description?.trim() || '',
      ownerId: userId,
      members: [],
      createdAt: now,
      updatedAt: now,
      config: { ...DEFAULT_PROJECT_CONFIG },
    };

    await db.collection('projects').doc(projectId).set(project);

    res.status(201).json({
      success: true,
      data: { id: projectId, ...project },
    });
  } catch (error) {
    console.error('Error creating project:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to create project',
    });
  }
}

/**
 * Update a project
 */
export async function updateProject(req: AuthenticatedRequest, res: Response): Promise<void> {
  try {
    const { projectId } = req.params;
    const userId = req.userId!;
    const { name, description, members } = req.body;

    const hasAccess = await checkProjectAccess(userId, projectId, 'editor');
    if (!hasAccess) {
      res.status(403).json({
        success: false,
        error: 'Access denied',
      });
      return;
    }

    const updates: Partial<Project> = {
      updatedAt: new Date().toISOString(),
    };

    if (name !== undefined) {
      if (typeof name !== 'string' || name.trim().length === 0) {
        res.status(400).json({
          success: false,
          error: 'Project name cannot be empty',
        });
        return;
      }
      updates.name = name.trim();
    }

    if (description !== undefined) {
      updates.description = description?.trim() || '';
    }

    // Only owner can modify members
    if (members !== undefined) {
      const projectDoc = await db.collection('projects').doc(projectId).get();
      const project = projectDoc.data() as Project;

      if (project.ownerId !== userId) {
        res.status(403).json({
          success: false,
          error: 'Only project owner can modify members',
        });
        return;
      }

      updates.members = members;
    }

    await db.collection('projects').doc(projectId).update(updates);

    const updatedDoc = await db.collection('projects').doc(projectId).get();

    res.json({
      success: true,
      data: { id: updatedDoc.id, ...updatedDoc.data() },
    });
  } catch (error) {
    console.error('Error updating project:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to update project',
    });
  }
}

/**
 * Delete a project
 */
export async function deleteProject(req: AuthenticatedRequest, res: Response): Promise<void> {
  try {
    const { projectId } = req.params;
    const userId = req.userId!;

    const hasAccess = await checkProjectAccess(userId, projectId, 'owner');
    if (!hasAccess) {
      res.status(403).json({
        success: false,
        error: 'Only project owner can delete the project',
      });
      return;
    }

    // Delete all tasks first
    const tasksSnapshot = await db
      .collection('projects')
      .doc(projectId)
      .collection('tasks')
      .get();

    const batch = db.batch();

    tasksSnapshot.docs.forEach(doc => {
      batch.delete(doc.ref);
    });

    // Delete the project
    batch.delete(db.collection('projects').doc(projectId));

    await batch.commit();

    res.json({
      success: true,
      message: 'Project deleted successfully',
    });
  } catch (error) {
    console.error('Error deleting project:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to delete project',
    });
  }
}

/**
 * Update project configuration
 */
export async function updateProjectConfig(req: AuthenticatedRequest, res: Response): Promise<void> {
  try {
    const { projectId } = req.params;
    const userId = req.userId!;
    const config = req.body as Partial<ProjectConfig>;

    const hasAccess = await checkProjectAccess(userId, projectId, 'editor');
    if (!hasAccess) {
      res.status(403).json({
        success: false,
        error: 'Access denied',
      });
      return;
    }

    // Validate config values
    if (config.barHeight !== undefined && (config.barHeight < 16 || config.barHeight > 48)) {
      res.status(400).json({
        success: false,
        error: 'Bar height must be between 16 and 48',
      });
      return;
    }

    if (config.barSpacing !== undefined && (config.barSpacing < 4 || config.barSpacing > 20)) {
      res.status(400).json({
        success: false,
        error: 'Bar spacing must be between 4 and 20',
      });
      return;
    }

    if (config.fontSize !== undefined && (config.fontSize < 10 || config.fontSize > 16)) {
      res.status(400).json({
        success: false,
        error: 'Font size must be between 10 and 16',
      });
      return;
    }

    if (config.chartWidth !== undefined && (config.chartWidth < 800 || config.chartWidth > 3000)) {
      res.status(400).json({
        success: false,
        error: 'Chart width must be between 800 and 3000',
      });
      return;
    }

    if (config.chartHeight !== undefined && (config.chartHeight < 400 || config.chartHeight > 2000)) {
      res.status(400).json({
        success: false,
        error: 'Chart height must be between 400 and 2000',
      });
      return;
    }

    // Get current config
    const projectDoc = await db.collection('projects').doc(projectId).get();
    const project = projectDoc.data() as Project;
    const currentConfig = project.config || DEFAULT_PROJECT_CONFIG;

    // Merge configs
    const newConfig: ProjectConfig = {
      ...currentConfig,
      ...config,
    };

    await db.collection('projects').doc(projectId).update({
      config: newConfig,
      updatedAt: new Date().toISOString(),
    });

    res.json({
      success: true,
      data: newConfig,
    });
  } catch (error) {
    console.error('Error updating project config:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to update project configuration',
    });
  }
}

export const projectFunctions = {
  listProjects,
  getProject,
  createProject,
  updateProject,
  deleteProject,
  updateProjectConfig,
};
