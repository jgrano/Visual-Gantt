/**
 * DriveStorage.gs - Google Drive Project Data Management
 * Stores project data in Google Drive for access across Sheets and Slides
 *
 * Storage Structure:
 * - Google Drive/Visual Gantt Projects/
 *   - project-001.json (project data + metadata)
 *   - project-002.json
 */

// ============================================================================
// CONSTANTS
// ============================================================================

const PROJECTS_FOLDER_NAME = 'Visual Gantt Projects';
const USER_PROP_PROJECTS_FOLDER_ID = 'visualGanttProjectsFolderId';

// ============================================================================
// PROJECT FOLDER MANAGEMENT
// ============================================================================

/**
 * Gets or creates the projects folder in Google Drive
 * @returns {GoogleAppsScript.Drive.Folder} The projects folder
 */
function getOrCreateProjectsFolder() {
  // Try cached folder ID first
  const storedFolderId = getStoredId_(USER_PROP_PROJECTS_FOLDER_ID);
  if (storedFolderId) {
    const folder = getFolderById_(storedFolderId);
    if (folder) {
      return folder;
    }
    // Stored ID is invalid, clear it
    setStoredId_(USER_PROP_PROJECTS_FOLDER_ID, null);
  }

  // Create new projects folder
  const folder = DriveApp.createFolder(PROJECTS_FOLDER_NAME);
  const folderId = folder.getId();
  setStoredId_(USER_PROP_PROJECTS_FOLDER_ID, folderId);
  Logger.log('Created new projects folder: ' + folderId);
  return folder;
}

/**
 * Opens the projects folder in a new browser tab
 */
function openProjectsFolder() {
  try {
    const folder = getOrCreateProjectsFolder();
    const url = folder.getUrl();

    const html = '<html><body><script>window.open("' + url + '", "_blank");google.script.host.close();</script></body></html>';
    const output = HtmlService.createHtmlOutput(html)
      .setWidth(1)
      .setHeight(1);

    // Try to get UI context
    try {
      SpreadsheetApp.getUi().showModalDialog(output, 'Opening...');
    } catch (e) {
      // If in Slides context
      try {
        SlidesApp.getUi().showModalDialog(output, 'Opening...');
      } catch (e2) {
        return { success: false, message: 'Could not open folder: ' + e2.message, url: url };
      }
    }

    return { success: true, url: url };
  } catch (e) {
    return { success: false, message: 'Error opening projects folder: ' + e.message };
  }
}

// ============================================================================
// PROJECT DATA MANAGEMENT
// ============================================================================

/**
 * Saves a project to Google Drive
 * @param {Object} projectData - The project data to save
 * @param {string} projectData.id - Unique project ID (optional, will be generated)
 * @param {string} projectData.name - Project name
 * @param {Array} projectData.tasks - Array of task objects
 * @param {Object} projectData.metadata - Project metadata (created, modified, etc.)
 * @returns {Object} Result with project ID and file URL
 */
function saveProjectToDrive(projectData) {
  try {
    const folder = getOrCreateProjectsFolder();

    // Generate ID if not provided
    if (!projectData.id) {
      projectData.id = 'project-' + Utilities.getUuid().substring(0, 8);
    }

    // Add/update metadata
    const now = new Date().toISOString();
    projectData.metadata = projectData.metadata || {};
    projectData.metadata.modified = now;
    if (!projectData.metadata.created) {
      projectData.metadata.created = now;
    }

    const filename = projectData.id + '.json';
    const jsonContent = JSON.stringify(projectData, null, 2);

    // Check if file already exists
    const existingFiles = folder.getFilesByName(filename);
    let file;

    if (existingFiles.hasNext()) {
      // Update existing file
      file = existingFiles.next();
      file.setContent(jsonContent);
      Logger.log('Updated project file: ' + filename);
    } else {
      // Create new file
      file = folder.createFile(filename, jsonContent, 'application/json');
      Logger.log('Created project file: ' + filename);
    }

    return {
      success: true,
      projectId: projectData.id,
      fileId: file.getId(),
      url: file.getUrl(),
      message: 'Project saved successfully'
    };
  } catch (e) {
    Logger.log('Error saving project: ' + e.message);
    return {
      success: false,
      message: 'Failed to save project: ' + e.message
    };
  }
}

/**
 * Loads a project from Google Drive by ID
 * @param {string} projectId - The project ID
 * @returns {Object} The project data or null if not found
 */
function loadProjectFromDrive(projectId) {
  try {
    const folder = getOrCreateProjectsFolder();
    const filename = projectId + '.json';
    const files = folder.getFilesByName(filename);

    if (!files.hasNext()) {
      return null;
    }

    const file = files.next();
    const content = file.getBlob().getDataAsString();
    const projectData = JSON.parse(content);

    return projectData;
  } catch (e) {
    Logger.log('Error loading project: ' + e.message);
    return null;
  }
}

/**
 * Lists all projects in Google Drive
 * @returns {Array} Array of project metadata objects
 */
function listProjectsInDrive() {
  try {
    const folder = getOrCreateProjectsFolder();
    const files = folder.getFiles();
    const projects = [];

    while (files.hasNext()) {
      const file = files.next();

      // Skip non-JSON files
      if (!file.getName().endsWith('.json')) {
        continue;
      }

      try {
        const content = file.getBlob().getDataAsString();
        const projectData = JSON.parse(content);

        // Return lightweight metadata
        projects.push({
          id: projectData.id,
          name: projectData.name || 'Untitled Project',
          taskCount: projectData.tasks ? projectData.tasks.length : 0,
          created: projectData.metadata ? projectData.metadata.created : null,
          modified: projectData.metadata ? projectData.metadata.modified : null,
          fileId: file.getId()
        });
      } catch (e) {
        Logger.log('Error parsing project file: ' + file.getName() + ' - ' + e.message);
      }
    }

    // Sort by modified date (newest first)
    projects.sort((a, b) => {
      const dateA = a.modified ? new Date(a.modified) : new Date(0);
      const dateB = b.modified ? new Date(b.modified) : new Date(0);
      return dateB - dateA;
    });

    return projects;
  } catch (e) {
    Logger.log('Error listing projects: ' + e.message);
    return [];
  }
}

/**
 * Deletes a project from Google Drive
 * @param {string} projectId - The project ID to delete
 * @returns {Object} Result of the deletion
 */
function deleteProjectFromDrive(projectId) {
  try {
    const folder = getOrCreateProjectsFolder();
    const filename = projectId + '.json';
    const files = folder.getFilesByName(filename);

    if (!files.hasNext()) {
      return {
        success: false,
        message: 'Project not found'
      };
    }

    const file = files.next();
    file.setTrashed(true);

    return {
      success: true,
      message: 'Project deleted successfully'
    };
  } catch (e) {
    Logger.log('Error deleting project: ' + e.message);
    return {
      success: false,
      message: 'Failed to delete project: ' + e.message
    };
  }
}

// ============================================================================
// SHEETS INTEGRATION
// ============================================================================

/**
 * Exports current sheet data to Drive as a project
 * @returns {Object} Result with project ID
 */
function exportSheetToProject() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(DATA_SHEET_NAME);

    if (!sheet) {
      return {
        success: false,
        message: 'Data sheet not found. Please create it first using Visual Gantt > Setup Data Sheet.'
      };
    }

    // Read tasks from sheet
    const tasks = readProjectData(ss);

    if (!tasks || tasks.length === 0) {
      return {
        success: false,
        message: 'No tasks found in data sheet'
      };
    }

    // Create project object
    const projectData = {
      name: ss.getName() + ' - ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      tasks: tasks,
      config: getConfig(), // Include current configuration
      metadata: {
        source: 'Google Sheets',
        sourceId: ss.getId(),
        sourceUrl: ss.getUrl()
      }
    };

    return saveProjectToDrive(projectData);
  } catch (e) {
    Logger.log('Error exporting sheet to project: ' + e.message);
    return {
      success: false,
      message: 'Failed to export: ' + e.message
    };
  }
}

/**
 * Imports a project from Drive into the current sheet
 * @param {string} projectId - The project ID to import
 * @returns {Object} Result of the import
 */
function importProjectToSheet(projectId) {
  try {
    const projectData = loadProjectFromDrive(projectId);

    if (!projectData) {
      return {
        success: false,
        message: 'Project not found'
      };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(DATA_SHEET_NAME);

    // Create data sheet if it doesn't exist
    if (!sheet) {
      sheet = ss.insertSheet(DATA_SHEET_NAME);
      // Add headers
      sheet.getRange(1, 1, 1, COLUMN_HEADERS.length).setValues([COLUMN_HEADERS]);
      sheet.getRange(1, 1, 1, COLUMN_HEADERS.length).setFontWeight('bold');
    }

    // Clear existing data (except headers)
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, COLUMN_HEADERS.length).clearContent();
    }

    // Write task data
    if (projectData.tasks && projectData.tasks.length > 0) {
      const taskRows = projectData.tasks.map(task => [
        task.taskId || '',
        task.taskName || '',
        task.startDate || '',
        task.endDate || '',
        task.duration || '',
        task.owner || '',
        task.percentComplete || 0,
        task.priority || '',
        task.jiraTicket || '',
        task.parentTask || '',
        task.dependencies || '',
        task.taskType || 'Task',
        task.swimlane || '',
        task.project || '',
        task.originalStart || '',
        task.originalEnd || '',
        task.slip || '',
        task.modified || ''
      ]);

      sheet.getRange(2, 1, taskRows.length, COLUMN_HEADERS.length).setValues(taskRows);
    }

    // Apply configuration if included
    if (projectData.config) {
      saveConfig(projectData.config);
    }

    return {
      success: true,
      taskCount: projectData.tasks.length,
      message: 'Project imported successfully'
    };
  } catch (e) {
    Logger.log('Error importing project to sheet: ' + e.message);
    return {
      success: false,
      message: 'Failed to import: ' + e.message
    };
  }
}
