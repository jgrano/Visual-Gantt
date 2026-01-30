/**
 * SlidesDataManager.gs - Embedded Data Storage for Slides
 * Stores project data directly in the presentation file
 *
 * Storage Strategy:
 * - Primary: DocumentProperties (JSON in presentation metadata)
 * - Fallback: Hidden "Data" slide with JSON in speaker notes
 *
 * Benefits:
 * - Single file (no external dependencies)
 * - Share presentation = share data
 * - No Drive/Sheets required
 */

// ============================================================================
// CONSTANTS
// ============================================================================

const SLIDES_DATA_KEY = 'visualGanttProjectData';
const DATA_SLIDE_NAME = '[Visual Gantt Data - Do Not Delete]';

// ============================================================================
// EMBEDDED DATA STORAGE
// ============================================================================

/**
 * Saves project data directly into the presentation
 * @param {Object} projectData - The project data to save
 * @returns {Object} Result of the save operation
 */
function saveProjectToPresentation(projectData) {
  try {
    const presentation = SlidesApp.getActivePresentation();

    // Add metadata
    const now = new Date().toISOString();
    projectData.metadata = projectData.metadata || {};
    projectData.metadata.modified = now;
    if (!projectData.metadata.created) {
      projectData.metadata.created = now;
    }
    projectData.metadata.presentationId = presentation.getId();
    projectData.metadata.presentationName = presentation.getName();

    const jsonData = JSON.stringify(projectData);

    // Try DocumentProperties first (best approach)
    try {
      const props = PropertiesService.getDocumentProperties();
      props.setProperty(SLIDES_DATA_KEY, jsonData);

      return {
        success: true,
        storage: 'properties',
        message: 'Project saved to presentation'
      };
    } catch (e) {
      Logger.log('DocumentProperties failed, using slide fallback: ' + e.message);
    }

    // Fallback: Use hidden slide with speaker notes
    const dataSlide = getOrCreateDataSlide(presentation);
    dataSlide.getNotesPage().getSpeakerNotesShape().getText().setText(jsonData);

    return {
      success: true,
      storage: 'slide',
      message: 'Project saved to presentation (using hidden slide)'
    };
  } catch (e) {
    Logger.log('Error saving project to presentation: ' + e.message);
    return {
      success: false,
      message: 'Failed to save: ' + e.message
    };
  }
}

/**
 * Loads project data from the presentation
 * @returns {Object} The project data or null if not found
 */
function loadProjectFromPresentation() {
  try {
    const presentation = SlidesApp.getActivePresentation();

    // Try DocumentProperties first
    try {
      const props = PropertiesService.getDocumentProperties();
      const jsonData = props.getProperty(SLIDES_DATA_KEY);

      if (jsonData) {
        return JSON.parse(jsonData);
      }
    } catch (e) {
      Logger.log('DocumentProperties read failed, trying slide: ' + e.message);
    }

    // Fallback: Read from hidden slide
    const dataSlide = findDataSlide(presentation);
    if (dataSlide) {
      const notesText = dataSlide.getNotesPage().getSpeakerNotesShape().getText().asString();
      if (notesText) {
        return JSON.parse(notesText);
      }
    }

    return null;
  } catch (e) {
    Logger.log('Error loading project from presentation: ' + e.message);
    return null;
  }
}

/**
 * Checks if the presentation has embedded project data
 * @returns {boolean} True if data exists
 */
function hasEmbeddedData() {
  try {
    // Check DocumentProperties
    try {
      const props = PropertiesService.getDocumentProperties();
      const jsonData = props.getProperty(SLIDES_DATA_KEY);
      if (jsonData) return true;
    } catch (e) {
      // Ignore
    }

    // Check for data slide
    const presentation = SlidesApp.getActivePresentation();
    const dataSlide = findDataSlide(presentation);
    return dataSlide !== null;
  } catch (e) {
    return false;
  }
}

/**
 * Deletes project data from the presentation
 * @returns {Object} Result of the deletion
 */
function deleteProjectFromPresentation() {
  try {
    const presentation = SlidesApp.getActivePresentation();
    let deleted = false;

    // Delete from DocumentProperties
    try {
      const props = PropertiesService.getDocumentProperties();
      props.deleteProperty(SLIDES_DATA_KEY);
      deleted = true;
    } catch (e) {
      Logger.log('DocumentProperties delete failed: ' + e.message);
    }

    // Delete data slide if it exists
    const dataSlide = findDataSlide(presentation);
    if (dataSlide) {
      dataSlide.remove();
      deleted = true;
    }

    return {
      success: deleted,
      message: deleted ? 'Project data deleted' : 'No data found'
    };
  } catch (e) {
    Logger.log('Error deleting project from presentation: ' + e.message);
    return {
      success: false,
      message: 'Failed to delete: ' + e.message
    };
  }
}

// ============================================================================
// DATA SLIDE MANAGEMENT
// ============================================================================

/**
 * Finds the data slide in the presentation
 * @param {GoogleAppsScript.Slides.Presentation} presentation
 * @returns {GoogleAppsScript.Slides.Slide|null} The data slide or null
 */
function findDataSlide(presentation) {
  const slides = presentation.getSlides();

  for (const slide of slides) {
    try {
      const shapes = slide.getShapes();
      for (const shape of shapes) {
        if (shape.getText().asString().includes(DATA_SLIDE_NAME)) {
          return slide;
        }
      }
    } catch (e) {
      // Ignore errors on individual slides
    }
  }

  return null;
}

/**
 * Gets or creates the hidden data slide
 * @param {GoogleAppsScript.Slides.Presentation} presentation
 * @returns {GoogleAppsScript.Slides.Slide} The data slide
 */
function getOrCreateDataSlide(presentation) {
  let dataSlide = findDataSlide(presentation);

  if (!dataSlide) {
    // Append at the end
    dataSlide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);

    // Mark as data slide
    const textBox = dataSlide.insertTextBox(DATA_SLIDE_NAME);
    textBox.setLeft(50);
    textBox.setTop(50);
    textBox.getText().getTextStyle().setFontSize(24);

    // Add warning
    const warning = dataSlide.insertTextBox(
      'This slide contains project data.\n\n' +
      'Do not delete unless you want to remove all project information.\n\n' +
      'You can hide this slide from the presentation.'
    );
    warning.setLeft(50);
    warning.setTop(150);
    warning.setWidth(600);
    warning.getText().getTextStyle().setFontSize(14);
  }

  return dataSlide;
}

// ============================================================================
// TASK MANAGEMENT
// ============================================================================

/**
 * Adds a task to the embedded project data
 * @param {Object} task - The task to add
 * @returns {Object} Result with updated project data
 */
function addTaskToPresentation(task) {
  try {
    let projectData = loadProjectFromPresentation();

    if (!projectData) {
      // Initialize new project
      const presentation = SlidesApp.getActivePresentation();
      projectData = {
        name: presentation.getName(),
        tasks: [],
        config: getDefaultSlidesConfig()
      };
    }

    // Generate task ID if not provided
    if (!task.taskId) {
      task.taskId = 'TASK-' + (projectData.tasks.length + 1).toString().padStart(3, '0');
    }

    projectData.tasks.push(task);

    const result = saveProjectToPresentation(projectData);

    if (result.success) {
      result.task = task;
      result.taskCount = projectData.tasks.length;
    }

    return result;
  } catch (e) {
    Logger.log('Error adding task: ' + e.message);
    return {
      success: false,
      message: 'Failed to add task: ' + e.message
    };
  }
}

/**
 * Updates a task in the embedded project data
 * @param {string} taskId - The task ID to update
 * @param {Object} updates - The fields to update
 * @returns {Object} Result of the update
 */
function updateTaskInPresentation(taskId, updates) {
  try {
    const projectData = loadProjectFromPresentation();

    if (!projectData || !projectData.tasks) {
      return {
        success: false,
        message: 'No project data found'
      };
    }

    const taskIndex = projectData.tasks.findIndex(t => t.taskId === taskId);

    if (taskIndex === -1) {
      return {
        success: false,
        message: 'Task not found: ' + taskId
      };
    }

    // Update task
    projectData.tasks[taskIndex] = {
      ...projectData.tasks[taskIndex],
      ...updates
    };

    const result = saveProjectToPresentation(projectData);

    if (result.success) {
      result.task = projectData.tasks[taskIndex];
    }

    return result;
  } catch (e) {
    Logger.log('Error updating task: ' + e.message);
    return {
      success: false,
      message: 'Failed to update task: ' + e.message
    };
  }
}

/**
 * Deletes a task from the embedded project data
 * @param {string} taskId - The task ID to delete
 * @returns {Object} Result of the deletion
 */
function deleteTaskFromPresentation(taskId) {
  try {
    const projectData = loadProjectFromPresentation();

    if (!projectData || !projectData.tasks) {
      return {
        success: false,
        message: 'No project data found'
      };
    }

    const initialLength = projectData.tasks.length;
    projectData.tasks = projectData.tasks.filter(t => t.taskId !== taskId);

    if (projectData.tasks.length === initialLength) {
      return {
        success: false,
        message: 'Task not found: ' + taskId
      };
    }

    const result = saveProjectToPresentation(projectData);

    if (result.success) {
      result.taskCount = projectData.tasks.length;
    }

    return result;
  } catch (e) {
    Logger.log('Error deleting task: ' + e.message);
    return {
      success: false,
      message: 'Failed to delete task: ' + e.message
    };
  }
}

/**
 * Gets all tasks from the embedded project data
 * @returns {Array} Array of tasks
 */
function getTasksFromPresentation() {
  const projectData = loadProjectFromPresentation();
  return (projectData && projectData.tasks) ? projectData.tasks : [];
}

// ============================================================================
// IMPORT/EXPORT
// ============================================================================

/**
 * Imports project data from a JSON string
 * @param {string} jsonString - The JSON data to import
 * @returns {Object} Result of the import
 */
function importProjectJSON(jsonString) {
  try {
    const projectData = JSON.parse(jsonString);

    // Validate basic structure
    if (!projectData.tasks || !Array.isArray(projectData.tasks)) {
      return {
        success: false,
        message: 'Invalid project data format'
      };
    }

    const result = saveProjectToPresentation(projectData);

    if (result.success) {
      result.taskCount = projectData.tasks.length;
    }

    return result;
  } catch (e) {
    Logger.log('Error importing JSON: ' + e.message);
    return {
      success: false,
      message: 'Failed to import: ' + e.message
    };
  }
}

/**
 * Exports project data as JSON string
 * @returns {string} The project data as JSON
 */
function exportProjectJSON() {
  const projectData = loadProjectFromPresentation();
  return projectData ? JSON.stringify(projectData, null, 2) : null;
}

/**
 * Imports tasks from a CSV string
 * @param {string} csvString - The CSV data
 * @returns {Object} Result of the import
 */
function importTasksFromCSV(csvString) {
  try {
    const lines = csvString.split('\n');
    const headers = lines[0].split(',').map(h => h.trim());

    const tasks = [];

    for (let i = 1; i < lines.length; i++) {
      if (!lines[i].trim()) continue;

      const values = lines[i].split(',').map(v => v.trim());
      const task = {};

      headers.forEach((header, index) => {
        task[header] = values[index] || '';
      });

      tasks.push(task);
    }

    let projectData = loadProjectFromPresentation();

    if (!projectData) {
      const presentation = SlidesApp.getActivePresentation();
      projectData = {
        name: presentation.getName(),
        tasks: [],
        config: getDefaultSlidesConfig()
      };
    }

    projectData.tasks = tasks;

    const result = saveProjectToPresentation(projectData);

    if (result.success) {
      result.taskCount = tasks.length;
    }

    return result;
  } catch (e) {
    Logger.log('Error importing CSV: ' + e.message);
    return {
      success: false,
      message: 'Failed to import CSV: ' + e.message
    };
  }
}

// ============================================================================
// DEFAULT CONFIG
// ============================================================================

/**
 * Gets default configuration for Slides-based charts
 * @returns {Object} Default configuration
 */
function getDefaultSlidesConfig() {
  return {
    startDate: null,
    endDate: null,
    swimlaneGrouping: 'swimlane',
    colorBy: 'priority',
    showDependencies: true,
    showProgress: true,
    showTodayMarker: true,
    barHeight: 24,
    barSpacing: 8,
    chartWidth: 800,
    chartHeight: 600,
    fontSize: 11
  };
}

/**
 * Gets configuration from embedded project data
 * @returns {Object} Configuration object
 */
function getConfigFromPresentation() {
  const projectData = loadProjectFromPresentation();
  return (projectData && projectData.config) ? projectData.config : getDefaultSlidesConfig();
}

/**
 * Updates configuration in embedded project data
 * @param {Object} config - The configuration to save
 * @returns {Object} Result of the update
 */
function updateConfigInPresentation(config) {
  try {
    let projectData = loadProjectFromPresentation();

    if (!projectData) {
      const presentation = SlidesApp.getActivePresentation();
      projectData = {
        name: presentation.getName(),
        tasks: [],
        config: config
      };
    } else {
      projectData.config = config;
    }

    return saveProjectToPresentation(projectData);
  } catch (e) {
    Logger.log('Error updating config: ' + e.message);
    return {
      success: false,
      message: 'Failed to update config: ' + e.message
    };
  }
}
