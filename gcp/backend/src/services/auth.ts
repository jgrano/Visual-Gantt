/**
 * Visual Gantt - Authentication Service
 * Handles Firebase Authentication verification
 */

import { Request, Response, NextFunction } from 'express';
import * as admin from 'firebase-admin';

export interface AuthenticatedRequest extends Request {
  user?: admin.auth.DecodedIdToken;
  userId?: string;
}

/**
 * Express middleware to verify Firebase ID tokens
 */
export const authMiddleware = async (
  req: AuthenticatedRequest,
  res: Response,
  next: NextFunction
): Promise<void> => {
  // Skip auth for health check
  if (req.path === '/health') {
    next();
    return;
  }

  const authHeader = req.headers.authorization;

  if (!authHeader || !authHeader.startsWith('Bearer ')) {
    res.status(401).json({
      success: false,
      error: 'Missing or invalid authorization header',
    });
    return;
  }

  const idToken = authHeader.split('Bearer ')[1];

  try {
    const decodedToken = await admin.auth().verifyIdToken(idToken);
    req.user = decodedToken;
    req.userId = decodedToken.uid;
    next();
  } catch (error) {
    console.error('Auth error:', error);
    res.status(401).json({
      success: false,
      error: 'Invalid or expired token',
    });
  }
};

/**
 * Checks if user has access to a project
 */
export const checkProjectAccess = async (
  userId: string,
  projectId: string,
  requiredRole: 'owner' | 'editor' | 'viewer' = 'viewer'
): Promise<boolean> => {
  const projectDoc = await admin.firestore()
    .collection('projects')
    .doc(projectId)
    .get();

  if (!projectDoc.exists) {
    return false;
  }

  const project = projectDoc.data();
  if (!project) {
    return false;
  }

  // Check if user is owner
  if (project.ownerId === userId) {
    return true;
  }

  // Check members
  const member = project.members?.find((m: { userId: string }) => m.userId === userId);
  if (!member) {
    return false;
  }

  // Check role hierarchy
  const roleHierarchy = ['viewer', 'editor', 'owner'];
  const memberRoleIndex = roleHierarchy.indexOf(member.role);
  const requiredRoleIndex = roleHierarchy.indexOf(requiredRole);

  return memberRoleIndex >= requiredRoleIndex;
};

/**
 * Express middleware to check project access
 */
export const requireProjectAccess = (requiredRole: 'owner' | 'editor' | 'viewer' = 'viewer') => {
  return async (
    req: AuthenticatedRequest,
    res: Response,
    next: NextFunction
  ): Promise<void> => {
    const { projectId } = req.params;
    const userId = req.userId;

    if (!userId) {
      res.status(401).json({
        success: false,
        error: 'User not authenticated',
      });
      return;
    }

    if (!projectId) {
      res.status(400).json({
        success: false,
        error: 'Project ID is required',
      });
      return;
    }

    const hasAccess = await checkProjectAccess(userId, projectId, requiredRole);

    if (!hasAccess) {
      res.status(403).json({
        success: false,
        error: 'Access denied to this project',
      });
      return;
    }

    next();
  };
};
