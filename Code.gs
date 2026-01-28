/**
 * Visual Gantt - Google Sheets Timeline Add-on
 * Creates professional Gantt chart timelines for project management
 *
 * @author Visual Gantt Team
 * @version 1.0.0
 */

// ============================================================================
// CONSTANTS
// ============================================================================

const CONFIG_KEY = 'visualGanttConfig';
const DATA_SHEET_NAME = 'Project Data';
const TIMELINE_SHEET_NAME = 'Timeline View';

const COLUMN_HEADERS = [
  'Task ID',
  'Task Name',
  'Start Date',
  'End Date',
  'Duration (Days)',
  'Owner',
  '% Complete',
  'Priority',
  'Jira Ticket',
  'Parent Task',
  'Dependencies',
  'Task Type',
  'Swimlane/Category'
];

const DEFAULT_CONFIG = {
  startDate: null,
  endDate: null,
  swimlaneGrouping: 'Category', // 'Owner', 'Category', 'None'
  colorScheme: 'priority', // 'priority', 'taskType', 'swimlane'
  showDependencies: true,
  showPercentComplete: true,
  showMilestones: true,
  showTodayMarker: true,
  barHeight: 24,
  barSpacing: 8,
  fontSize: 12,
  chartWidth: 1200,
  chartHeight: 800
};

const PRIORITY_COLORS = {
  'High': '#DC3545',
  'Medium': '#FFC107',
  'Low': '#28A745',
  'default': '#6C757D'
};

const TASK_TYPE_COLORS = {
  'Milestone': '#9C27B0',
  'Task': '#2196F3',
  'default': '#607D8B'
};

const SWIMLANE_COLORS = [
  '#4285F4', // Blue
  '#34A853', // Green
  '#FBBC05', // Yellow
  '#EA4335', // Red
  '#9C27B0', // Purple
  '#00ACC1', // Cyan
  '#FF7043', // Deep Orange
  '#5C6BC0', // Indigo
  '#26A69A', // Teal
  '#AB47BC'  // Purple
];

// ============================================================================
// MENU & TRIGGERS
// ============================================================================

/**
 * Creates the add-on menu when the spreadsheet opens
 */
function onOpen(e) {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Visual Gantt')
    .addItem('Generate Timeline', 'generateTimeline')
    .addItem('Refresh Timeline', 'refreshTimeline')
    .addSeparator()
    .addItem('Configure Settings', 'showConfigDialog')
    .addSeparator()
    .addItem('Export as PNG', 'exportAsPng')
    .addItem('Export as PDF', 'exportAsPdf')
    .addSeparator()
    .addItem('Setup Data Sheet', 'setupDataSheet')
    .addItem('Create Sample Data', 'createSampleData')
    .addSeparator()
    .addItem('Help', 'showHelp')
    .addItem('About', 'showAbout')
    .addToUi();
}

/**
 * Runs when the add-on is installed
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Runs when file scope is granted
 */
function onFileScopeGranted(e) {
  onOpen(e);
}

// ============================================================================
// SETUP FUNCTIONS
// ============================================================================

/**
 * Creates the data sheet template with proper column headers
 */
function setupDataSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let dataSheet = ss.getSheetByName(DATA_SHEET_NAME);

  if (dataSheet) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Sheet Exists',
      `A sheet named "${DATA_SHEET_NAME}" already exists. Do you want to replace it?`,
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      return;
    }
    ss.deleteSheet(dataSheet);
  }

  // Create new data sheet
  dataSheet = ss.insertSheet(DATA_SHEET_NAME, 0);

  // Set up headers
  const headerRange = dataSheet.getRange(1, 1, 1, COLUMN_HEADERS.length);
  headerRange.setValues([COLUMN_HEADERS]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4285F4');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setHorizontalAlignment('center');

  // Set column widths
  const columnWidths = [80, 200, 100, 100, 100, 120, 80, 80, 150, 120, 150, 100, 150];
  columnWidths.forEach((width, index) => {
    dataSheet.setColumnWidth(index + 1, width);
  });

  // Add data validation for Priority column
  const priorityRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['High', 'Medium', 'Low'], true)
    .setAllowInvalid(false)
    .build();
  dataSheet.getRange('H2:H1000').setDataValidation(priorityRule);

  // Add data validation for Task Type column
  const taskTypeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Task', 'Milestone'], true)
    .setAllowInvalid(false)
    .build();
  dataSheet.getRange('L2:L1000').setDataValidation(taskTypeRule);

  // Add data validation for % Complete column
  const percentRule = SpreadsheetApp.newDataValidation()
    .requireNumberBetween(0, 100)
    .setAllowInvalid(false)
    .build();
  dataSheet.getRange('G2:G1000').setDataValidation(percentRule);

  // Format date columns
  dataSheet.getRange('C2:D1000').setNumberFormat('yyyy-mm-dd');

  // Format percentage column
  dataSheet.getRange('G2:G1000').setNumberFormat('0"%"');

  // Freeze header row
  dataSheet.setFrozenRows(1);

  // Create Timeline View sheet if it doesn't exist
  let timelineSheet = ss.getSheetByName(TIMELINE_SHEET_NAME);
  if (!timelineSheet) {
    timelineSheet = ss.insertSheet(TIMELINE_SHEET_NAME, 1);
    timelineSheet.getRange('A1').setValue('Timeline will be generated here. Use "Visual Gantt > Generate Timeline" to create your chart.');
    timelineSheet.getRange('A1').setFontStyle('italic');
    timelineSheet.getRange('A1').setFontColor('#666666');
  }

  // Switch to data sheet
  ss.setActiveSheet(dataSheet);

  SpreadsheetApp.getUi().alert(
    'Setup Complete',
    `Data sheet "${DATA_SHEET_NAME}" has been created.\n\nNext steps:\n1. Enter your project tasks\n2. Click "Visual Gantt > Generate Timeline"`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Creates sample project data for demonstration
 */
function createSampleData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let dataSheet = ss.getSheetByName(DATA_SHEET_NAME);

  if (!dataSheet) {
    setupDataSheet();
    dataSheet = ss.getSheetByName(DATA_SHEET_NAME);
  }

  // Check if there's existing data
  const lastRow = dataSheet.getLastRow();
  if (lastRow > 1) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Existing Data',
      'There is existing data in the sheet. Do you want to replace it with sample data?',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      return;
    }
    // Clear existing data (keep headers)
    if (lastRow > 1) {
      dataSheet.getRange(2, 1, lastRow - 1, COLUMN_HEADERS.length).clearContent();
    }
  }

  // Sample data for B1 Bidirectional Charger project
  const today = new Date();
  const baseDate = new Date(today.getFullYear(), today.getMonth(), 1); // Start of current month

  const sampleData = [
    // Mechanical Engineering Tasks
    ['MECH-001', 'Mechanical Design Phase', addDays(baseDate, 0), addDays(baseDate, 45), '', 'John Smith', 75, 'High', 'https://jira.example.com/MECH-001', '', '', 'Task', 'Mechanical'],
    ['MECH-002', 'Enclosure CAD Design', addDays(baseDate, 0), addDays(baseDate, 14), '', 'John Smith', 100, 'High', 'https://jira.example.com/MECH-002', 'MECH-001', '', 'Task', 'Mechanical'],
    ['MECH-003', 'Thermal Analysis', addDays(baseDate, 7), addDays(baseDate, 21), '', 'Sarah Johnson', 80, 'High', 'https://jira.example.com/MECH-003', 'MECH-001', 'MECH-002', 'Task', 'Mechanical'],
    ['MECH-004', 'Prototype Fabrication', addDays(baseDate, 21), addDays(baseDate, 35), '', 'John Smith', 40, 'Medium', 'https://jira.example.com/MECH-004', 'MECH-001', 'MECH-003', 'Task', 'Mechanical'],
    ['MECH-005', 'Mechanical Design Complete', addDays(baseDate, 45), addDays(baseDate, 45), '', 'John Smith', 0, 'High', 'https://jira.example.com/MECH-005', '', 'MECH-004', 'Milestone', 'Mechanical'],

    // Electrical Engineering Tasks
    ['ELEC-001', 'Electrical Design Phase', addDays(baseDate, 5), addDays(baseDate, 50), '', 'Mike Chen', 60, 'High', 'https://jira.example.com/ELEC-001', '', '', 'Task', 'Electrical'],
    ['ELEC-002', 'Power Stage Design', addDays(baseDate, 5), addDays(baseDate, 20), '', 'Mike Chen', 100, 'High', 'https://jira.example.com/ELEC-002', 'ELEC-001', '', 'Task', 'Electrical'],
    ['ELEC-003', 'PCB Layout', addDays(baseDate, 15), addDays(baseDate, 30), '', 'Lisa Wong', 70, 'High', 'https://jira.example.com/ELEC-003', 'ELEC-001', 'ELEC-002', 'Task', 'Electrical'],
    ['ELEC-004', 'EMC Pre-compliance', addDays(baseDate, 30), addDays(baseDate, 40), '', 'Mike Chen', 20, 'Medium', 'https://jira.example.com/ELEC-004', 'ELEC-001', 'ELEC-003', 'Task', 'Electrical'],
    ['ELEC-005', 'Electrical Validation Complete', addDays(baseDate, 50), addDays(baseDate, 50), '', 'Mike Chen', 0, 'High', 'https://jira.example.com/ELEC-005', '', 'ELEC-004', 'Milestone', 'Electrical'],

    // Firmware Development Tasks
    ['FW-001', 'Firmware Development', addDays(baseDate, 10), addDays(baseDate, 55), '', 'Alex Rivera', 45, 'High', 'https://jira.example.com/FW-001', '', '', 'Task', 'Firmware'],
    ['FW-002', 'CAN Protocol Implementation', addDays(baseDate, 10), addDays(baseDate, 25), '', 'Alex Rivera', 100, 'High', 'https://jira.example.com/FW-002', 'FW-001', '', 'Task', 'Firmware'],
    ['FW-003', 'Bidirectional Control Logic', addDays(baseDate, 20), addDays(baseDate, 40), '', 'Alex Rivera', 50, 'High', 'https://jira.example.com/FW-003', 'FW-001', 'FW-002', 'Task', 'Firmware'],
    ['FW-004', 'Safety State Machine', addDays(baseDate, 35), addDays(baseDate, 50), '', 'David Park', 15, 'High', 'https://jira.example.com/FW-004', 'FW-001', 'FW-003', 'Task', 'Firmware'],
    ['FW-005', 'Firmware Release v1.0', addDays(baseDate, 55), addDays(baseDate, 55), '', 'Alex Rivera', 0, 'High', 'https://jira.example.com/FW-005', '', 'FW-004', 'Milestone', 'Firmware'],

    // Test Engineering Tasks
    ['TEST-001', 'Test Engineering Phase', addDays(baseDate, 25), addDays(baseDate, 65), '', 'Emily Taylor', 30, 'Medium', 'https://jira.example.com/TEST-001', '', '', 'Task', 'Test'],
    ['TEST-002', 'Test Plan Development', addDays(baseDate, 25), addDays(baseDate, 35), '', 'Emily Taylor', 100, 'Medium', 'https://jira.example.com/TEST-002', 'TEST-001', '', 'Task', 'Test'],
    ['TEST-003', 'DVT Execution', addDays(baseDate, 40), addDays(baseDate, 55), '', 'Emily Taylor', 25, 'High', 'https://jira.example.com/TEST-003', 'TEST-001', 'TEST-002,ELEC-004', 'Task', 'Test'],
    ['TEST-004', 'Certification Testing', addDays(baseDate, 55), addDays(baseDate, 65), '', 'Emily Taylor', 0, 'High', 'https://jira.example.com/TEST-004', 'TEST-001', 'TEST-003,FW-005', 'Task', 'Test'],
    ['TEST-005', 'Product Certification Complete', addDays(baseDate, 65), addDays(baseDate, 65), '', 'Emily Taylor', 0, 'High', 'https://jira.example.com/TEST-005', '', 'TEST-004', 'Milestone', 'Test'],

    // Program Management
    ['PM-001', 'Customer Review - Honda', addDays(baseDate, 30), addDays(baseDate, 30), '', 'Program Manager', 0, 'High', 'https://jira.example.com/PM-001', '', 'MECH-004,ELEC-003', 'Milestone', 'Program'],
    ['PM-002', 'Customer Review - Mazda', addDays(baseDate, 45), addDays(baseDate, 45), '', 'Program Manager', 0, 'High', 'https://jira.example.com/PM-002', '', 'FW-004', 'Milestone', 'Program'],
    ['PM-003', 'Production Release', addDays(baseDate, 70), addDays(baseDate, 70), '', 'Program Manager', 0, 'High', 'https://jira.example.com/PM-003', '', 'TEST-005', 'Milestone', 'Program']
  ];

  // Write sample data
  if (sampleData.length > 0) {
    dataSheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
  }

  SpreadsheetApp.getUi().alert(
    'Sample Data Created',
    'Sample B1 Bidirectional Charger project data has been added.\n\nClick "Visual Gantt > Generate Timeline" to create your Gantt chart.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Helper function to add days to a date
 */
function addDays(date, days) {
  const result = new Date(date);
  result.setDate(result.getDate() + days);
  return result;
}

// ============================================================================
// DATA READING & VALIDATION
// ============================================================================

/**
 * Reads task data from the data sheet
 * @returns {Object} Object containing tasks array and any validation errors
 */
function readTaskData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(DATA_SHEET_NAME);

  if (!dataSheet) {
    return {
      tasks: [],
      errors: [`Data sheet "${DATA_SHEET_NAME}" not found. Please run "Setup Data Sheet" first.`]
    };
  }

  const lastRow = dataSheet.getLastRow();
  const lastCol = dataSheet.getLastColumn();

  if (lastRow < 2) {
    return {
      tasks: [],
      errors: ['No task data found. Please add tasks to the data sheet.']
    };
  }

  const dataRange = dataSheet.getRange(2, 1, lastRow - 1, Math.max(lastCol, COLUMN_HEADERS.length));
  const data = dataRange.getValues();

  const tasks = [];
  const errors = [];

  data.forEach((row, index) => {
    const rowNum = index + 2;

    // Skip empty rows
    if (!row[0] && !row[1]) {
      return;
    }

    const task = {
      id: row[0] ? String(row[0]).trim() : `TASK-${rowNum}`,
      name: row[1] ? String(row[1]).trim() : '',
      startDate: parseDate(row[2]),
      endDate: parseDate(row[3]),
      duration: row[4] ? Number(row[4]) : null,
      owner: row[5] ? String(row[5]).trim() : '',
      percentComplete: row[6] !== '' && row[6] !== null ? Number(row[6]) : 0,
      priority: row[7] ? String(row[7]).trim() : 'Medium',
      jiraTicket: row[8] ? String(row[8]).trim() : '',
      parentTask: row[9] ? String(row[9]).trim() : '',
      dependencies: row[10] ? String(row[10]).split(',').map(d => d.trim()).filter(d => d) : [],
      taskType: row[11] ? String(row[11]).trim() : 'Task',
      swimlane: row[12] ? String(row[12]).trim() : 'Default',
      rowNumber: rowNum
    };

    // Validate required fields
    if (!task.name) {
      errors.push(`Row ${rowNum}: Task name is required`);
      return;
    }

    if (!task.startDate) {
      errors.push(`Row ${rowNum}: Start date is required for "${task.name}"`);
      return;
    }

    // Calculate end date from duration if not provided
    if (!task.endDate && task.duration) {
      task.endDate = addDays(task.startDate, task.duration - 1);
    }

    // Calculate duration from dates if not provided
    if (task.startDate && task.endDate && !task.duration) {
      task.duration = Math.ceil((task.endDate - task.startDate) / (1000 * 60 * 60 * 24)) + 1;
    }

    // For milestones, end date equals start date
    if (task.taskType === 'Milestone') {
      task.endDate = task.startDate;
      task.duration = 1;
    }

    // Validate end date
    if (!task.endDate) {
      errors.push(`Row ${rowNum}: End date or duration is required for "${task.name}"`);
      return;
    }

    // Validate date order
    if (task.startDate > task.endDate) {
      errors.push(`Row ${rowNum}: End date must be after start date for "${task.name}"`);
      return;
    }

    // Validate percent complete
    if (task.percentComplete < 0 || task.percentComplete > 100) {
      errors.push(`Row ${rowNum}: Percent complete must be between 0 and 100 for "${task.name}"`);
      task.percentComplete = Math.max(0, Math.min(100, task.percentComplete));
    }

    tasks.push(task);
  });

  // Build task ID lookup map
  const taskMap = new Map();
  tasks.forEach(task => taskMap.set(task.id, task));

  // Validate dependencies exist
  tasks.forEach(task => {
    task.dependencies = task.dependencies.filter(depId => {
      if (!taskMap.has(depId)) {
        errors.push(`Row ${task.rowNumber}: Dependency "${depId}" not found for "${task.name}"`);
        return false;
      }
      return true;
    });
  });

  // Check for circular dependencies
  const circularErrors = detectCircularDependencies(tasks, taskMap);
  errors.push(...circularErrors);

  // Build parent-child relationships
  tasks.forEach(task => {
    if (task.parentTask && taskMap.has(task.parentTask)) {
      task.parent = taskMap.get(task.parentTask);
      task.isSubtask = true;
    } else {
      task.isSubtask = false;
    }
  });

  return { tasks, errors, taskMap };
}

/**
 * Parses a date value from various formats
 */
function parseDate(value) {
  if (!value) return null;

  if (value instanceof Date) {
    return isNaN(value.getTime()) ? null : value;
  }

  if (typeof value === 'string') {
    const parsed = new Date(value);
    return isNaN(parsed.getTime()) ? null : parsed;
  }

  if (typeof value === 'number') {
    // Excel/Sheets serial date number
    const parsed = new Date((value - 25569) * 86400 * 1000);
    return isNaN(parsed.getTime()) ? null : parsed;
  }

  return null;
}

/**
 * Detects circular dependencies in task relationships
 */
function detectCircularDependencies(tasks, taskMap) {
  const errors = [];
  const visited = new Set();
  const recursionStack = new Set();

  function hasCycle(taskId, path) {
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
    if (task && task.dependencies) {
      for (const depId of task.dependencies) {
        if (hasCycle(depId, [...path, taskId])) {
          return true;
        }
      }
    }

    recursionStack.delete(taskId);
    return false;
  }

  tasks.forEach(task => {
    if (!visited.has(task.id)) {
      hasCycle(task.id, []);
    }
  });

  return errors;
}

// ============================================================================
// CONFIGURATION
// ============================================================================

/**
 * Gets the current configuration
 */
function getConfig() {
  const userProperties = PropertiesService.getUserProperties();
  const savedConfig = userProperties.getProperty(CONFIG_KEY);

  if (savedConfig) {
    try {
      return { ...DEFAULT_CONFIG, ...JSON.parse(savedConfig) };
    } catch (e) {
      return { ...DEFAULT_CONFIG };
    }
  }

  return { ...DEFAULT_CONFIG };
}

/**
 * Saves configuration
 */
function saveConfig(config) {
  const userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty(CONFIG_KEY, JSON.stringify(config));
}

/**
 * Shows the configuration dialog
 */
function showConfigDialog() {
  const html = HtmlService.createHtmlOutputFromFile('ConfigDialog')
    .setWidth(500)
    .setHeight(600)
    .setTitle('Visual Gantt Settings');

  SpreadsheetApp.getUi().showModalDialog(html, 'Visual Gantt Settings');
}

/**
 * Gets configuration for the dialog
 */
function getConfigForDialog() {
  return getConfig();
}

/**
 * Saves configuration from the dialog
 */
function saveConfigFromDialog(config) {
  saveConfig(config);
  return { success: true };
}

// ============================================================================
// TIMELINE GENERATION
// ============================================================================

/**
 * Main function to generate the timeline
 */
function generateTimeline() {
  const ui = SpreadsheetApp.getUi();

  // Read task data
  const { tasks, errors, taskMap } = readTaskData();

  // Show validation errors if any critical ones
  const criticalErrors = errors.filter(e => !e.includes('not found'));
  if (criticalErrors.length > 0 && tasks.length === 0) {
    ui.alert('Validation Errors', criticalErrors.join('\n'), ui.ButtonSet.OK);
    return;
  }

  if (tasks.length === 0) {
    ui.alert('No Data', 'No valid tasks found. Please add tasks to the data sheet.', ui.ButtonSet.OK);
    return;
  }

  // Get configuration
  const config = getConfig();

  // Calculate date range if not specified
  let startDate = config.startDate ? new Date(config.startDate) : null;
  let endDate = config.endDate ? new Date(config.endDate) : null;

  if (!startDate || !endDate) {
    const dates = tasks.flatMap(t => [t.startDate, t.endDate]).filter(d => d);
    if (!startDate) {
      startDate = new Date(Math.min(...dates));
      startDate.setDate(startDate.getDate() - 7); // Add padding
    }
    if (!endDate) {
      endDate = new Date(Math.max(...dates));
      endDate.setDate(endDate.getDate() + 7); // Add padding
    }
  }

  // Prepare data for chart
  const chartData = {
    tasks: prepareTasksForChart(tasks, config),
    config: config,
    startDate: startDate.toISOString(),
    endDate: endDate.toISOString(),
    today: new Date().toISOString()
  };

  // Generate chart using HTML template
  const template = HtmlService.createTemplateFromFile('ChartRenderer');
  template.chartData = JSON.stringify(chartData);

  const html = template.evaluate()
    .setWidth(config.chartWidth + 100)
    .setHeight(config.chartHeight + 100);

  // Show chart generation dialog
  ui.showModalDialog(html, 'Generating Timeline...');
}

/**
 * Prepares tasks for chart rendering
 */
function prepareTasksForChart(tasks, config) {
  // Group by swimlane if configured
  let sortedTasks = [...tasks];

  if (config.swimlaneGrouping !== 'None') {
    const groupField = config.swimlaneGrouping === 'Owner' ? 'owner' : 'swimlane';

    // Sort by swimlane/group, then by start date
    sortedTasks.sort((a, b) => {
      const groupA = a[groupField] || 'Other';
      const groupB = b[groupField] || 'Other';

      if (groupA !== groupB) {
        return groupA.localeCompare(groupB);
      }

      // Within same group, parent tasks first, then by start date
      if (a.isSubtask !== b.isSubtask) {
        return a.isSubtask ? 1 : -1;
      }

      return a.startDate - b.startDate;
    });
  } else {
    // Sort by start date only
    sortedTasks.sort((a, b) => a.startDate - b.startDate);
  }

  // Assign colors based on color scheme
  const swimlaneColorMap = new Map();
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

    // Convert dates to ISO strings for JSON
    task.startDateStr = task.startDate.toISOString();
    task.endDateStr = task.endDate.toISOString();
  });

  return sortedTasks;
}

/**
 * Refreshes the timeline (alias for generateTimeline)
 */
function refreshTimeline() {
  generateTimeline();
}

/**
 * Inserts the chart image into the timeline sheet
 */
function insertChartImage(imageBlob) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let timelineSheet = ss.getSheetByName(TIMELINE_SHEET_NAME);

  if (!timelineSheet) {
    timelineSheet = ss.insertSheet(TIMELINE_SHEET_NAME);
  }

  // Clear existing content
  timelineSheet.clear();

  // Insert image
  const image = timelineSheet.insertImage(imageBlob, 1, 1);

  // Add timestamp
  const config = getConfig();
  const lastRow = Math.ceil(config.chartHeight / 21) + 2;
  timelineSheet.getRange(lastRow, 1).setValue('Generated: ' + new Date().toLocaleString());
  timelineSheet.getRange(lastRow, 1).setFontStyle('italic');
  timelineSheet.getRange(lastRow, 1).setFontColor('#666666');

  return true;
}

/**
 * Receives image data from client and inserts into sheet
 */
function saveChartToSheet(base64Data) {
  try {
    // Remove data URL prefix if present
    const base64Content = base64Data.replace(/^data:image\/png;base64,/, '');

    // Create blob from base64
    const blob = Utilities.newBlob(
      Utilities.base64Decode(base64Content),
      'image/png',
      'gantt_chart.png'
    );

    // Insert into sheet
    insertChartImage(blob);

    return { success: true, message: 'Timeline saved to sheet!' };
  } catch (e) {
    return { success: false, message: 'Error saving chart: ' + e.message };
  }
}

// ============================================================================
// EXPORT FUNCTIONS
// ============================================================================

/**
 * Exports timeline as PNG
 */
function exportAsPng() {
  const html = HtmlService.createHtmlOutputFromFile('ExportDialog')
    .setWidth(400)
    .setHeight(200)
    .setTitle('Export as PNG');

  SpreadsheetApp.getUi().showModalDialog(html, 'Export as PNG');
}

/**
 * Exports timeline as PDF
 */
function exportAsPdf() {
  const html = HtmlService.createHtmlOutputFromFile('ExportPdfDialog')
    .setWidth(400)
    .setHeight(200)
    .setTitle('Export as PDF');

  SpreadsheetApp.getUi().showModalDialog(html, 'Export as PDF');
}

/**
 * Saves export file to Drive and returns download URL
 */
function saveExportFile(base64Data, filename, mimeType) {
  try {
    const base64Content = base64Data.replace(/^data:[^;]+;base64,/, '');
    const blob = Utilities.newBlob(
      Utilities.base64Decode(base64Content),
      mimeType,
      filename
    );

    const file = DriveApp.createFile(blob);
    const url = file.getDownloadUrl();

    return { success: true, url: url, filename: filename };
  } catch (e) {
    return { success: false, message: 'Error exporting: ' + e.message };
  }
}

/**
 * Gets export data for the export dialogs
 */
function getExportData() {
  const { tasks, errors } = readTaskData();
  const config = getConfig();

  if (tasks.length === 0) {
    return { tasks: [], errors: errors };
  }

  // Calculate date range
  let startDate = config.startDate ? new Date(config.startDate) : null;
  let endDate = config.endDate ? new Date(config.endDate) : null;

  if (!startDate || !endDate) {
    const dates = tasks.flatMap(t => [t.startDate, t.endDate]).filter(d => d);
    if (!startDate) {
      startDate = new Date(Math.min(...dates));
      startDate.setDate(startDate.getDate() - 7);
    }
    if (!endDate) {
      endDate = new Date(Math.max(...dates));
      endDate.setDate(endDate.getDate() + 7);
    }
  }

  return {
    tasks: prepareTasksForChart(tasks, config),
    config: config,
    startDate: startDate.toISOString(),
    endDate: endDate.toISOString(),
    today: new Date().toISOString()
  };
}

/**
 * Exports the timeline as a PNG file
 */
function exportTimelineAsPng() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const timelineSheet = ss.getSheetByName(TIMELINE_SHEET_NAME);

    if (!timelineSheet) {
      return { success: false, message: 'No timeline found. Please generate a timeline first.' };
    }

    // Get all images in the sheet
    const images = timelineSheet.getImages();

    if (images.length === 0) {
      return { success: false, message: 'No timeline image found. Please generate a timeline first.' };
    }

    // Get the first image (the timeline)
    const timelineImage = images[0];
    const blob = timelineImage.getBlob();

    // Generate filename with timestamp
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmmss');
    const filename = `Visual_Gantt_Timeline_${timestamp}.png`;

    // Save to Drive
    const file = DriveApp.createFile(blob.setName(filename));

    return {
      success: true,
      url: file.getDownloadUrl(),
      filename: filename
    };
  } catch (e) {
    return { success: false, message: 'Error exporting PNG: ' + e.message };
  }
}

/**
 * Exports the timeline as a PDF file
 */
function exportTimelineAsPdf() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const timelineSheet = ss.getSheetByName(TIMELINE_SHEET_NAME);

    if (!timelineSheet) {
      return { success: false, message: 'No timeline found. Please generate a timeline first.' };
    }

    // Get spreadsheet ID and sheet ID
    const ssId = ss.getId();
    const sheetId = timelineSheet.getSheetId();

    // Generate PDF export URL
    const url = `https://docs.google.com/spreadsheets/d/${ssId}/export?` +
      `format=pdf&` +
      `gid=${sheetId}&` +
      `size=A4&` +
      `portrait=false&` +
      `fitw=true&` +
      `gridlines=false&` +
      `printtitle=false&` +
      `sheetnames=false&` +
      `pagenum=false&` +
      `fzr=false`;

    // Fetch the PDF
    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(url, {
      headers: {
        'Authorization': 'Bearer ' + token
      }
    });

    // Generate filename
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmmss');
    const filename = `Visual_Gantt_Timeline_${timestamp}.pdf`;

    // Save to Drive
    const blob = response.getBlob().setName(filename);
    const file = DriveApp.createFile(blob);

    return {
      success: true,
      url: file.getDownloadUrl(),
      filename: filename
    };
  } catch (e) {
    return { success: false, message: 'Error exporting PDF: ' + e.message };
  }
}

// ============================================================================
// HELP & ABOUT
// ============================================================================

/**
 * Shows help dialog
 */
function showHelp() {
  const html = HtmlService.createHtmlOutputFromFile('HelpDialog')
    .setWidth(600)
    .setHeight(500)
    .setTitle('Visual Gantt Help');

  SpreadsheetApp.getUi().showModalDialog(html, 'Visual Gantt Help');
}

/**
 * Shows about dialog
 */
function showAbout() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'About Visual Gantt',
    'Visual Gantt v1.0.0\n\n' +
    'A professional Gantt chart timeline add-on for Google Sheets.\n\n' +
    'Features:\n' +
    '- Professional timeline visualization\n' +
    '- Swimlane grouping by team/category\n' +
    '- Milestone markers\n' +
    '- Progress tracking\n' +
    '- Dependency arrows\n' +
    '- PNG/PDF export\n\n' +
    'For support, please contact your administrator.',
    ui.ButtonSet.OK
  );
}

/**
 * Include HTML file content (for templates)
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
