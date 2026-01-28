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
  'Swimlane/Category',
  // New Smartsheet-like columns
  'Status',              // Not Started, In Progress, Complete, On Hold, At Risk
  'Health',              // Green, Yellow, Red
  'Baseline Start',      // Original planned start date
  'Baseline End',        // Original planned end date
  'Dependency Type',     // FS, SS, FF, SF (Finish-Start, Start-Start, Finish-Finish, Start-Finish)
  'Lag/Lead (Days)',     // Positive = lag, Negative = lead
  'Resource %',          // Resource allocation percentage (0-100)
  'Budget Cost',         // Planned cost
  'Actual Cost',         // Actual cost incurred
  'Work Days Only',      // TRUE/FALSE - exclude weekends
  'Notes'                // Comments/notes for the task
];

// Column indices for easy reference (0-based)
const COL = {
  TASK_ID: 0,
  TASK_NAME: 1,
  START_DATE: 2,
  END_DATE: 3,
  DURATION: 4,
  OWNER: 5,
  PERCENT_COMPLETE: 6,
  PRIORITY: 7,
  JIRA_TICKET: 8,
  PARENT_TASK: 9,
  DEPENDENCIES: 10,
  TASK_TYPE: 11,
  SWIMLANE: 12,
  STATUS: 13,
  HEALTH: 14,
  BASELINE_START: 15,
  BASELINE_END: 16,
  DEPENDENCY_TYPE: 17,
  LAG_LEAD: 18,
  RESOURCE_PERCENT: 19,
  BUDGET_COST: 20,
  ACTUAL_COST: 21,
  WORK_DAYS_ONLY: 22,
  NOTES: 23
};

const DEFAULT_CONFIG = {
  startDate: null,
  endDate: null,
  swimlaneGrouping: 'Category', // 'Owner', 'Category', 'None'
  colorScheme: 'priority', // 'priority', 'taskType', 'swimlane', 'status', 'health'
  showDependencies: true,
  showPercentComplete: true,
  showMilestones: true,
  showTodayMarker: true,
  barHeight: 24,
  barSpacing: 8,
  fontSize: 12,
  chartWidth: 1200,
  chartHeight: 800,
  // New Smartsheet-like settings
  showCriticalPath: true,
  showBaseline: true,
  showHealthIndicators: true,
  showCostTracking: false,
  excludeWeekends: true,
  holidays: [], // Array of holiday dates to exclude
  autoSchedule: true,
  showResourceWorkload: false,
  // Filter settings
  activeFilter: 'all', // 'all', 'myTasks', 'overdue', 'atRisk', 'custom'
  customFilterOwner: '',
  customFilterStatus: '',
  customFilterPriority: '',
  // Conditional formatting rules
  conditionalFormatting: {
    overdueRed: true,
    atRiskYellow: true,
    completedGreen: true
  },
  // Print settings
  printDateRange: 'all', // 'all', 'custom'
  printStartDate: null,
  printEndDate: null,
  printScale: 'fit', // 'fit', '100%', '75%', '50%'
  printOrientation: 'landscape'
};

// Status values
const STATUS_VALUES = ['Not Started', 'In Progress', 'Complete', 'On Hold', 'At Risk'];

// Health indicator values
const HEALTH_VALUES = ['Green', 'Yellow', 'Red'];

// Dependency type values
const DEPENDENCY_TYPES = ['FS', 'SS', 'FF', 'SF']; // Finish-Start, Start-Start, Finish-Finish, Start-Finish

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

// Status colors for color scheme
const STATUS_COLORS = {
  'Not Started': '#9E9E9E',
  'In Progress': '#2196F3',
  'Complete': '#4CAF50',
  'On Hold': '#FF9800',
  'At Risk': '#F44336',
  'default': '#6C757D'
};

// Health indicator colors
const HEALTH_COLORS = {
  'Green': '#4CAF50',
  'Yellow': '#FFC107',
  'Red': '#F44336',
  'default': '#9E9E9E'
};

// Critical path color
const CRITICAL_PATH_COLOR = '#D32F2F';

// Baseline bar color (semi-transparent gray)
const BASELINE_COLOR = 'rgba(128, 128, 128, 0.4)';

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
    .addSubMenu(ui.createMenu('Scheduling')
      .addItem('Auto-Schedule Tasks', 'autoScheduleTasks')
      .addItem('Calculate Critical Path', 'calculateAndHighlightCriticalPath')
      .addItem('Recalculate Rollups', 'recalculateRollups'))
    .addSubMenu(ui.createMenu('Baseline')
      .addItem('Save Current as Baseline', 'saveBaseline')
      .addItem('Clear Baseline', 'clearBaseline')
      .addItem('Compare to Baseline', 'showBaselineComparison'))
    .addSubMenu(ui.createMenu('Views & Filters')
      .addItem('Show All Tasks', 'filterShowAll')
      .addItem('My Tasks', 'filterMyTasks')
      .addItem('Overdue Tasks', 'filterOverdue')
      .addItem('At Risk Tasks', 'filterAtRisk')
      .addItem('Critical Path Only', 'filterCriticalPath')
      .addSeparator()
      .addItem('Resource Workload View', 'showResourceWorkload')
      .addItem('Cost Summary View', 'showCostSummary'))
    .addSeparator()
    .addItem('Configure Settings', 'showConfigDialog')
    .addItem('Manage Holidays', 'showHolidayManager')
    .addSeparator()
    .addSubMenu(ui.createMenu('Export')
      .addItem('Export as PNG', 'exportAsPng')
      .addItem('Export as PDF', 'exportAsPdf')
      .addItem('Print Settings', 'showPrintSettings'))
    .addSeparator()
    .addItem('Setup Data Sheet', 'setupDataSheet')
    .addItem('Create Sample Data', 'createSampleData')
    .addItem('Apply Conditional Formatting', 'applyConditionalFormatting')
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

  // Set column widths (expanded for new columns)
  const columnWidths = [
    80,   // Task ID
    200,  // Task Name
    100,  // Start Date
    100,  // End Date
    80,   // Duration
    120,  // Owner
    80,   // % Complete
    80,   // Priority
    150,  // Jira Ticket
    100,  // Parent Task
    150,  // Dependencies
    80,   // Task Type
    120,  // Swimlane
    100,  // Status
    70,   // Health
    100,  // Baseline Start
    100,  // Baseline End
    80,   // Dependency Type
    80,   // Lag/Lead
    80,   // Resource %
    100,  // Budget Cost
    100,  // Actual Cost
    90,   // Work Days Only
    200   // Notes
  ];
  columnWidths.forEach((width, index) => {
    dataSheet.setColumnWidth(index + 1, width);
  });

  // Add data validation for Priority column (H)
  const priorityRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['High', 'Medium', 'Low'], true)
    .setAllowInvalid(false)
    .build();
  dataSheet.getRange('H2:H1000').setDataValidation(priorityRule);

  // Add data validation for Task Type column (L)
  const taskTypeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Task', 'Milestone'], true)
    .setAllowInvalid(false)
    .build();
  dataSheet.getRange('L2:L1000').setDataValidation(taskTypeRule);

  // Add data validation for % Complete column (G)
  const percentRule = SpreadsheetApp.newDataValidation()
    .requireNumberBetween(0, 100)
    .setAllowInvalid(false)
    .build();
  dataSheet.getRange('G2:G1000').setDataValidation(percentRule);

  // Add data validation for Status column (N)
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(STATUS_VALUES, true)
    .setAllowInvalid(false)
    .build();
  dataSheet.getRange('N2:N1000').setDataValidation(statusRule);

  // Add data validation for Health column (O)
  const healthRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(HEALTH_VALUES, true)
    .setAllowInvalid(false)
    .build();
  dataSheet.getRange('O2:O1000').setDataValidation(healthRule);

  // Add data validation for Dependency Type column (R)
  const depTypeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(DEPENDENCY_TYPES, true)
    .setAllowInvalid(true) // Allow empty for default FS
    .build();
  dataSheet.getRange('R2:R1000').setDataValidation(depTypeRule);

  // Add data validation for Resource % column (T)
  const resourceRule = SpreadsheetApp.newDataValidation()
    .requireNumberBetween(0, 100)
    .setAllowInvalid(true)
    .build();
  dataSheet.getRange('T2:T1000').setDataValidation(resourceRule);

  // Add data validation for Work Days Only column (W)
  const workDaysRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['TRUE', 'FALSE'], true)
    .setAllowInvalid(true)
    .build();
  dataSheet.getRange('W2:W1000').setDataValidation(workDaysRule);

  // Format date columns (C, D, P, Q)
  dataSheet.getRange('C2:D1000').setNumberFormat('yyyy-mm-dd');
  dataSheet.getRange('P2:Q1000').setNumberFormat('yyyy-mm-dd');

  // Format percentage columns (G, T)
  dataSheet.getRange('G2:G1000').setNumberFormat('0"%"');
  dataSheet.getRange('T2:T1000').setNumberFormat('0"%"');

  // Format cost columns (U, V)
  dataSheet.getRange('U2:V1000').setNumberFormat('$#,##0.00');

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
    `Data sheet "${DATA_SHEET_NAME}" has been created with Smartsheet-like columns.\n\nNew features include:\n- Status & Health tracking\n- Baseline date comparison\n- Multiple dependency types (FS, SS, FF, SF)\n- Cost tracking\n- Resource allocation %\n- Work days calculation\n\nNext steps:\n1. Enter your project tasks\n2. Click "Visual Gantt > Generate Timeline"`,
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

  // Extended sample data with all new Smartsheet-like columns
  // Columns: Task ID, Task Name, Start Date, End Date, Duration, Owner, % Complete, Priority, Jira Ticket,
  //          Parent Task, Dependencies, Task Type, Swimlane, Status, Health, Baseline Start, Baseline End,
  //          Dependency Type, Lag/Lead, Resource %, Budget Cost, Actual Cost, Work Days Only, Notes
  const sampleData = [
    // Mechanical Engineering Tasks
    ['MECH-001', 'Mechanical Design Phase', addDays(baseDate, 0), addDays(baseDate, 45), '', 'John Smith', 75, 'High', 'https://jira.example.com/MECH-001', '', '', 'Task', 'Mechanical', 'In Progress', 'Green', addDays(baseDate, 0), addDays(baseDate, 42), 'FS', 0, 100, 50000, 38000, 'TRUE', 'Parent task for all mechanical work'],
    ['MECH-002', 'Enclosure CAD Design', addDays(baseDate, 0), addDays(baseDate, 14), '', 'John Smith', 100, 'High', 'https://jira.example.com/MECH-002', 'MECH-001', '', 'Task', 'Mechanical', 'Complete', 'Green', addDays(baseDate, 0), addDays(baseDate, 14), 'FS', 0, 100, 15000, 14500, 'TRUE', 'Initial enclosure design completed on schedule'],
    ['MECH-003', 'Thermal Analysis', addDays(baseDate, 7), addDays(baseDate, 21), '', 'Sarah Johnson', 80, 'High', 'https://jira.example.com/MECH-003', 'MECH-001', 'MECH-002', 'Task', 'Mechanical', 'In Progress', 'Yellow', addDays(baseDate, 5), addDays(baseDate, 18), 'FS', 2, 75, 12000, 10000, 'TRUE', 'Thermal simulation showing higher temps than expected'],
    ['MECH-004', 'Prototype Fabrication', addDays(baseDate, 21), addDays(baseDate, 35), '', 'John Smith', 40, 'Medium', 'https://jira.example.com/MECH-004', 'MECH-001', 'MECH-003', 'Task', 'Mechanical', 'In Progress', 'Green', addDays(baseDate, 20), addDays(baseDate, 34), 'FS', 0, 80, 20000, 8500, 'TRUE', 'CNC machining in progress'],
    ['MECH-005', 'Mechanical Design Complete', addDays(baseDate, 45), addDays(baseDate, 45), '', 'John Smith', 0, 'High', 'https://jira.example.com/MECH-005', '', 'MECH-004', 'Milestone', 'Mechanical', 'Not Started', 'Green', addDays(baseDate, 42), addDays(baseDate, 42), 'FS', 0, 0, 0, 0, 'FALSE', 'Key milestone for mechanical sign-off'],

    // Electrical Engineering Tasks
    ['ELEC-001', 'Electrical Design Phase', addDays(baseDate, 5), addDays(baseDate, 50), '', 'Mike Chen', 60, 'High', 'https://jira.example.com/ELEC-001', '', '', 'Task', 'Electrical', 'In Progress', 'Yellow', addDays(baseDate, 3), addDays(baseDate, 45), 'FS', 0, 100, 75000, 48000, 'TRUE', 'Power electronics design phase'],
    ['ELEC-002', 'Power Stage Design', addDays(baseDate, 5), addDays(baseDate, 20), '', 'Mike Chen', 100, 'High', 'https://jira.example.com/ELEC-002', 'ELEC-001', '', 'Task', 'Electrical', 'Complete', 'Green', addDays(baseDate, 3), addDays(baseDate, 18), 'FS', 0, 100, 25000, 24000, 'TRUE', 'GaN-based power stage finalized'],
    ['ELEC-003', 'PCB Layout', addDays(baseDate, 15), addDays(baseDate, 30), '', 'Lisa Wong', 70, 'High', 'https://jira.example.com/ELEC-003', 'ELEC-001', 'ELEC-002', 'Task', 'Electrical', 'In Progress', 'Green', addDays(baseDate, 15), addDays(baseDate, 28), 'SS', 5, 100, 18000, 13000, 'TRUE', '6-layer PCB, started after power stage review'],
    ['ELEC-004', 'EMC Pre-compliance', addDays(baseDate, 30), addDays(baseDate, 40), '', 'Mike Chen', 20, 'Medium', 'https://jira.example.com/ELEC-004', 'ELEC-001', 'ELEC-003', 'Task', 'Electrical', 'In Progress', 'Red', addDays(baseDate, 28), addDays(baseDate, 35), 'FF', 0, 50, 15000, 4000, 'TRUE', 'EMC chamber booked, some conducted emissions issues'],
    ['ELEC-005', 'Electrical Validation Complete', addDays(baseDate, 50), addDays(baseDate, 50), '', 'Mike Chen', 0, 'High', 'https://jira.example.com/ELEC-005', '', 'ELEC-004', 'Milestone', 'Electrical', 'Not Started', 'Yellow', addDays(baseDate, 45), addDays(baseDate, 45), 'FS', 0, 0, 0, 0, 'FALSE', 'Depends on EMC resolution'],

    // Firmware Development Tasks
    ['FW-001', 'Firmware Development', addDays(baseDate, 10), addDays(baseDate, 55), '', 'Alex Rivera', 45, 'High', 'https://jira.example.com/FW-001', '', '', 'Task', 'Firmware', 'In Progress', 'Green', addDays(baseDate, 10), addDays(baseDate, 52), 'FS', 0, 100, 60000, 28000, 'TRUE', 'Embedded firmware for bidirectional control'],
    ['FW-002', 'CAN Protocol Implementation', addDays(baseDate, 10), addDays(baseDate, 25), '', 'Alex Rivera', 100, 'High', 'https://jira.example.com/FW-002', 'FW-001', '', 'Task', 'Firmware', 'Complete', 'Green', addDays(baseDate, 10), addDays(baseDate, 25), 'FS', 0, 100, 18000, 17500, 'TRUE', 'CAN 2.0B + J1939 support complete'],
    ['FW-003', 'Bidirectional Control Logic', addDays(baseDate, 20), addDays(baseDate, 40), '', 'Alex Rivera', 50, 'High', 'https://jira.example.com/FW-003', 'FW-001', 'FW-002', 'Task', 'Firmware', 'In Progress', 'Green', addDays(baseDate, 20), addDays(baseDate, 38), 'FS', -2, 100, 22000, 11000, 'TRUE', 'V2G and G2V modes implemented'],
    ['FW-004', 'Safety State Machine', addDays(baseDate, 35), addDays(baseDate, 50), '', 'David Park', 15, 'High', 'https://jira.example.com/FW-004', 'FW-001', 'FW-003', 'Task', 'Firmware', 'In Progress', 'Yellow', addDays(baseDate, 33), addDays(baseDate, 47), 'SF', 0, 75, 20000, 3500, 'TRUE', 'ISO 26262 ASIL-B compliance work'],
    ['FW-005', 'Firmware Release v1.0', addDays(baseDate, 55), addDays(baseDate, 55), '', 'Alex Rivera', 0, 'High', 'https://jira.example.com/FW-005', '', 'FW-004', 'Milestone', 'Firmware', 'Not Started', 'Green', addDays(baseDate, 52), addDays(baseDate, 52), 'FS', 0, 0, 0, 0, 'FALSE', 'Code freeze for DVT builds'],

    // Test Engineering Tasks
    ['TEST-001', 'Test Engineering Phase', addDays(baseDate, 25), addDays(baseDate, 65), '', 'Emily Taylor', 30, 'Medium', 'https://jira.example.com/TEST-001', '', '', 'Task', 'Test', 'In Progress', 'Green', addDays(baseDate, 25), addDays(baseDate, 62), 'FS', 0, 100, 40000, 12000, 'TRUE', 'DVT and certification testing'],
    ['TEST-002', 'Test Plan Development', addDays(baseDate, 25), addDays(baseDate, 35), '', 'Emily Taylor', 100, 'Medium', 'https://jira.example.com/TEST-002', 'TEST-001', '', 'Task', 'Test', 'Complete', 'Green', addDays(baseDate, 25), addDays(baseDate, 33), 'FS', 0, 100, 8000, 7500, 'TRUE', 'Test procedures documented and reviewed'],
    ['TEST-003', 'DVT Execution', addDays(baseDate, 40), addDays(baseDate, 55), '', 'Emily Taylor', 25, 'High', 'https://jira.example.com/TEST-003', 'TEST-001', 'TEST-002,ELEC-004', 'Task', 'Test', 'In Progress', 'Yellow', addDays(baseDate, 38), addDays(baseDate, 52), 'FS', 3, 100, 20000, 5500, 'TRUE', 'Waiting on EMC-compliant boards'],
    ['TEST-004', 'Certification Testing', addDays(baseDate, 55), addDays(baseDate, 65), '', 'Emily Taylor', 0, 'High', 'https://jira.example.com/TEST-004', 'TEST-001', 'TEST-003,FW-005', 'Task', 'Test', 'Not Started', 'Green', addDays(baseDate, 52), addDays(baseDate, 62), 'FS', 0, 50, 12000, 0, 'TRUE', 'UL/CE certification at external lab'],
    ['TEST-005', 'Product Certification Complete', addDays(baseDate, 65), addDays(baseDate, 65), '', 'Emily Taylor', 0, 'High', 'https://jira.example.com/TEST-005', '', 'TEST-004', 'Milestone', 'Test', 'Not Started', 'Green', addDays(baseDate, 62), addDays(baseDate, 62), 'FS', 0, 0, 0, 0, 'FALSE', 'Final certification milestone'],

    // Program Management
    ['PM-001', 'Customer Review - Honda', addDays(baseDate, 30), addDays(baseDate, 30), '', 'Program Manager', 0, 'High', 'https://jira.example.com/PM-001', '', 'MECH-004,ELEC-003', 'Milestone', 'Program', 'Not Started', 'Green', addDays(baseDate, 28), addDays(baseDate, 28), 'FS', 0, 0, 5000, 0, 'FALSE', 'Design review with Honda engineering'],
    ['PM-002', 'Customer Review - Mazda', addDays(baseDate, 45), addDays(baseDate, 45), '', 'Program Manager', 0, 'High', 'https://jira.example.com/PM-002', '', 'FW-004', 'Milestone', 'Program', 'Not Started', 'Yellow', addDays(baseDate, 42), addDays(baseDate, 42), 'FS', 0, 0, 5000, 0, 'FALSE', 'Firmware demo for Mazda'],
    ['PM-003', 'Production Release', addDays(baseDate, 70), addDays(baseDate, 70), '', 'Program Manager', 0, 'High', 'https://jira.example.com/PM-003', '', 'TEST-005', 'Milestone', 'Program', 'Not Started', 'Green', addDays(baseDate, 65), addDays(baseDate, 65), 'FS', 0, 0, 0, 0, 'FALSE', 'Final production release gate']
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
      // Original columns
      id: row[COL.TASK_ID] ? String(row[COL.TASK_ID]).trim() : `TASK-${rowNum}`,
      name: row[COL.TASK_NAME] ? String(row[COL.TASK_NAME]).trim() : '',
      startDate: parseDate(row[COL.START_DATE]),
      endDate: parseDate(row[COL.END_DATE]),
      duration: row[COL.DURATION] ? Number(row[COL.DURATION]) : null,
      owner: row[COL.OWNER] ? String(row[COL.OWNER]).trim() : '',
      percentComplete: row[COL.PERCENT_COMPLETE] !== '' && row[COL.PERCENT_COMPLETE] !== null ? Number(row[COL.PERCENT_COMPLETE]) : 0,
      priority: row[COL.PRIORITY] ? String(row[COL.PRIORITY]).trim() : 'Medium',
      jiraTicket: row[COL.JIRA_TICKET] ? String(row[COL.JIRA_TICKET]).trim() : '',
      parentTask: row[COL.PARENT_TASK] ? String(row[COL.PARENT_TASK]).trim() : '',
      dependencies: row[COL.DEPENDENCIES] ? String(row[COL.DEPENDENCIES]).split(',').map(d => d.trim()).filter(d => d) : [],
      taskType: row[COL.TASK_TYPE] ? String(row[COL.TASK_TYPE]).trim() : 'Task',
      swimlane: row[COL.SWIMLANE] ? String(row[COL.SWIMLANE]).trim() : 'Default',

      // New Smartsheet-like columns
      status: row[COL.STATUS] ? String(row[COL.STATUS]).trim() : 'Not Started',
      health: row[COL.HEALTH] ? String(row[COL.HEALTH]).trim() : 'Green',
      baselineStart: parseDate(row[COL.BASELINE_START]),
      baselineEnd: parseDate(row[COL.BASELINE_END]),
      dependencyType: row[COL.DEPENDENCY_TYPE] ? String(row[COL.DEPENDENCY_TYPE]).trim() : 'FS',
      lagLead: row[COL.LAG_LEAD] ? Number(row[COL.LAG_LEAD]) : 0,
      resourcePercent: row[COL.RESOURCE_PERCENT] !== '' && row[COL.RESOURCE_PERCENT] !== null ? Number(row[COL.RESOURCE_PERCENT]) : 100,
      budgetCost: row[COL.BUDGET_COST] ? Number(row[COL.BUDGET_COST]) : 0,
      actualCost: row[COL.ACTUAL_COST] ? Number(row[COL.ACTUAL_COST]) : 0,
      workDaysOnly: row[COL.WORK_DAYS_ONLY] === true || row[COL.WORK_DAYS_ONLY] === 'TRUE' || row[COL.WORK_DAYS_ONLY] === 'true',
      notes: row[COL.NOTES] ? String(row[COL.NOTES]).trim() : '',

      rowNumber: rowNum,

      // Calculated fields (set later)
      isOnCriticalPath: false,
      slippage: 0, // Days behind baseline
      costVariance: 0 // Budget - Actual
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

  // Calculate slippage and cost variance for each task
  tasks.forEach(task => {
    // Slippage: how many days behind baseline
    if (task.baselineEnd && task.endDate) {
      task.slippage = Math.ceil((task.endDate - task.baselineEnd) / (1000 * 60 * 60 * 24));
    }
    // Cost variance: Budget - Actual (positive = under budget)
    task.costVariance = task.budgetCost - task.actualCost;
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

  // Calculate critical path if enabled
  if (config.showCriticalPath) {
    calculateCriticalPath(tasks, taskMap);
  }

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
      case 'status':
        task.color = STATUS_COLORS[task.status] || STATUS_COLORS.default;
        break;
      case 'health':
        task.color = HEALTH_COLORS[task.health] || HEALTH_COLORS.default;
        break;
      default:
        task.color = '#2196F3';
    }

    // Override color for critical path tasks if enabled
    if (config.showCriticalPath && task.isOnCriticalPath) {
      task.criticalPathColor = CRITICAL_PATH_COLOR;
    }

    // Convert dates to ISO strings for JSON
    task.startDateStr = task.startDate.toISOString();
    task.endDateStr = task.endDate.toISOString();

    // Convert baseline dates if they exist
    if (task.baselineStart) {
      task.baselineStartStr = task.baselineStart.toISOString();
    }
    if (task.baselineEnd) {
      task.baselineEndStr = task.baselineEnd.toISOString();
    }

    // Add health indicator for rendering
    task.healthColor = HEALTH_COLORS[task.health] || HEALTH_COLORS.default;
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
    'Visual Gantt v2.0.0 - Smartsheet Edition\n\n' +
    'A professional Gantt chart timeline add-on for Google Sheets with advanced project management features.\n\n' +
    'Core Features:\n' +
    '- Professional timeline visualization\n' +
    '- Swimlane grouping by team/category\n' +
    '- Milestone markers\n' +
    '- Progress tracking\n' +
    '- Dependency arrows\n' +
    '- PNG/PDF export\n\n' +
    'NEW Smartsheet-like Features:\n' +
    '- Critical Path calculation & highlighting\n' +
    '- Baseline tracking & comparison\n' +
    '- Multiple dependency types (FS, SS, FF, SF)\n' +
    '- Lag/Lead time support\n' +
    '- Auto-scheduling based on dependencies\n' +
    '- Work days calendar (exclude weekends/holidays)\n' +
    '- Resource workload analysis\n' +
    '- Status & Health indicators\n' +
    '- Conditional formatting\n' +
    '- Parent task rollup summaries\n' +
    '- Task filters & views\n' +
    '- Cost tracking (budget vs actual)\n' +
    '- Resource allocation percentages\n\n' +
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

// ============================================================================
// CRITICAL PATH CALCULATION
// ============================================================================

/**
 * Calculates the critical path for the project
 * Uses forward and backward pass algorithm
 * @param {Array} tasks - Array of task objects
 * @param {Map} taskMap - Map of task IDs to task objects
 * @returns {Array} Array of task IDs on the critical path
 */
function calculateCriticalPath(tasks, taskMap) {
  if (!tasks || tasks.length === 0) return [];

  const config = getConfig();

  // Initialize early start/finish and late start/finish
  tasks.forEach(task => {
    task.earlyStart = null;
    task.earlyFinish = null;
    task.lateStart = null;
    task.lateFinish = null;
    task.totalFloat = null;
  });

  // Find project start date (earliest start)
  const projectStart = new Date(Math.min(...tasks.map(t => t.startDate.getTime())));

  // Forward pass - calculate early start and early finish
  const processed = new Set();
  const queue = tasks.filter(t => t.dependencies.length === 0);

  queue.forEach(task => {
    task.earlyStart = task.startDate;
    task.earlyFinish = task.endDate;
  });

  while (queue.length > 0) {
    const task = queue.shift();
    if (processed.has(task.id)) continue;
    processed.add(task.id);

    // Find tasks that depend on this one
    tasks.forEach(successor => {
      if (successor.dependencies.includes(task.id)) {
        // Calculate early start based on dependency type and lag
        const depType = successor.dependencyType || 'FS';
        const lag = successor.lagLead || 0;

        let calculatedStart;
        switch (depType) {
          case 'FS': // Finish-to-Start
            calculatedStart = addWorkDays(task.earlyFinish, 1 + lag, config.excludeWeekends, config.holidays);
            break;
          case 'SS': // Start-to-Start
            calculatedStart = addWorkDays(task.earlyStart, lag, config.excludeWeekends, config.holidays);
            break;
          case 'FF': // Finish-to-Finish
            const duration = getWorkDaysDuration(successor.startDate, successor.endDate, config.excludeWeekends, config.holidays);
            calculatedStart = addWorkDays(task.earlyFinish, lag - duration + 1, config.excludeWeekends, config.holidays);
            break;
          case 'SF': // Start-to-Finish
            calculatedStart = addWorkDays(task.earlyStart, lag - getWorkDaysDuration(successor.startDate, successor.endDate, config.excludeWeekends, config.holidays) + 1, config.excludeWeekends, config.holidays);
            break;
          default:
            calculatedStart = addWorkDays(task.earlyFinish, 1 + lag, config.excludeWeekends, config.holidays);
        }

        if (!successor.earlyStart || calculatedStart > successor.earlyStart) {
          successor.earlyStart = calculatedStart;
          const duration = getWorkDaysDuration(successor.startDate, successor.endDate, config.excludeWeekends, config.holidays);
          successor.earlyFinish = addWorkDays(successor.earlyStart, duration - 1, config.excludeWeekends, config.holidays);
        }

        // Check if all dependencies are processed
        const allDepsProcessed = successor.dependencies.every(depId => processed.has(depId));
        if (allDepsProcessed && !processed.has(successor.id)) {
          queue.push(successor);
        }
      }
    });
  }

  // Find project end date (latest early finish)
  const projectEnd = new Date(Math.max(...tasks.filter(t => t.earlyFinish).map(t => t.earlyFinish.getTime())));

  // Backward pass - calculate late start and late finish
  const processedBackward = new Set();
  const endTasks = tasks.filter(t => {
    return !tasks.some(other => other.dependencies.includes(t.id));
  });

  endTasks.forEach(task => {
    task.lateFinish = projectEnd;
    const duration = getWorkDaysDuration(task.startDate, task.endDate, config.excludeWeekends, config.holidays);
    task.lateStart = addWorkDays(task.lateFinish, -(duration - 1), config.excludeWeekends, config.holidays);
  });

  // Process in reverse order
  const reverseQueue = [...endTasks];
  while (reverseQueue.length > 0) {
    const task = reverseQueue.shift();
    if (processedBackward.has(task.id)) continue;
    processedBackward.add(task.id);

    // Process predecessors
    task.dependencies.forEach(depId => {
      const predecessor = taskMap.get(depId);
      if (!predecessor) return;

      const depType = task.dependencyType || 'FS';
      const lag = task.lagLead || 0;

      let calculatedFinish;
      switch (depType) {
        case 'FS':
          calculatedFinish = addWorkDays(task.lateStart, -(1 + lag), config.excludeWeekends, config.holidays);
          break;
        case 'SS':
          calculatedFinish = addWorkDays(task.lateStart, -lag + getWorkDaysDuration(predecessor.startDate, predecessor.endDate, config.excludeWeekends, config.holidays) - 1, config.excludeWeekends, config.holidays);
          break;
        case 'FF':
          calculatedFinish = addWorkDays(task.lateFinish, -lag, config.excludeWeekends, config.holidays);
          break;
        case 'SF':
          calculatedFinish = addWorkDays(task.lateFinish, -lag, config.excludeWeekends, config.holidays);
          break;
        default:
          calculatedFinish = addWorkDays(task.lateStart, -(1 + lag), config.excludeWeekends, config.holidays);
      }

      if (!predecessor.lateFinish || calculatedFinish < predecessor.lateFinish) {
        predecessor.lateFinish = calculatedFinish;
        const duration = getWorkDaysDuration(predecessor.startDate, predecessor.endDate, config.excludeWeekends, config.holidays);
        predecessor.lateStart = addWorkDays(predecessor.lateFinish, -(duration - 1), config.excludeWeekends, config.holidays);
      }

      if (!processedBackward.has(predecessor.id)) {
        reverseQueue.push(predecessor);
      }
    });
  }

  // Calculate total float and identify critical path
  const criticalPath = [];
  tasks.forEach(task => {
    if (task.earlyStart && task.lateStart) {
      task.totalFloat = Math.ceil((task.lateStart - task.earlyStart) / (1000 * 60 * 60 * 24));
      task.isOnCriticalPath = task.totalFloat <= 0;
      if (task.isOnCriticalPath) {
        criticalPath.push(task.id);
      }
    }
  });

  return criticalPath;
}

/**
 * Calculates and highlights the critical path in the sheet
 */
function calculateAndHighlightCriticalPath() {
  const { tasks, errors, taskMap } = readTaskData();

  if (tasks.length === 0) {
    SpreadsheetApp.getUi().alert('No tasks found. Please add tasks first.');
    return;
  }

  const criticalPath = calculateCriticalPath(tasks, taskMap);

  // Highlight critical path tasks in the sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(DATA_SHEET_NAME);

  // Clear previous highlighting
  const lastRow = dataSheet.getLastRow();
  if (lastRow > 1) {
    dataSheet.getRange(2, 1, lastRow - 1, COLUMN_HEADERS.length).setBackground(null);
  }

  // Highlight critical path tasks
  tasks.forEach(task => {
    if (task.isOnCriticalPath) {
      dataSheet.getRange(task.rowNumber, 1, 1, COLUMN_HEADERS.length).setBackground('#FFCDD2');
    }
  });

  SpreadsheetApp.getUi().alert(
    'Critical Path Calculated',
    `Found ${criticalPath.length} tasks on the critical path.\n\nCritical tasks are highlighted in light red.\n\nCritical Path Tasks:\n${criticalPath.join(', ')}`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ============================================================================
// WORK DAYS CALCULATION
// ============================================================================

/**
 * Adds work days to a date, optionally excluding weekends and holidays
 */
function addWorkDays(startDate, days, excludeWeekends, holidays) {
  if (!startDate) return null;

  const result = new Date(startDate);
  let daysToAdd = days;
  const direction = days >= 0 ? 1 : -1;
  daysToAdd = Math.abs(daysToAdd);

  while (daysToAdd > 0) {
    result.setDate(result.getDate() + direction);

    if (excludeWeekends) {
      const dayOfWeek = result.getDay();
      if (dayOfWeek === 0 || dayOfWeek === 6) continue; // Skip weekends
    }

    if (holidays && holidays.length > 0) {
      const dateStr = Utilities.formatDate(result, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      if (holidays.includes(dateStr)) continue; // Skip holidays
    }

    daysToAdd--;
  }

  return result;
}

/**
 * Gets the number of work days between two dates
 */
function getWorkDaysDuration(startDate, endDate, excludeWeekends, holidays) {
  if (!startDate || !endDate) return 0;

  let count = 0;
  const current = new Date(startDate);

  while (current <= endDate) {
    let isWorkDay = true;

    if (excludeWeekends) {
      const dayOfWeek = current.getDay();
      if (dayOfWeek === 0 || dayOfWeek === 6) isWorkDay = false;
    }

    if (isWorkDay && holidays && holidays.length > 0) {
      const dateStr = Utilities.formatDate(current, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      if (holidays.includes(dateStr)) isWorkDay = false;
    }

    if (isWorkDay) count++;
    current.setDate(current.getDate() + 1);
  }

  return count;
}

/**
 * Shows holiday manager dialog
 */
function showHolidayManager() {
  const html = HtmlService.createHtmlOutputFromFile('HolidayManager')
    .setWidth(500)
    .setHeight(400)
    .setTitle('Manage Holidays');

  SpreadsheetApp.getUi().showModalDialog(html, 'Manage Holidays');
}

/**
 * Gets holidays from config
 */
function getHolidays() {
  const config = getConfig();
  return config.holidays || [];
}

/**
 * Saves holidays to config
 */
function saveHolidays(holidays) {
  const config = getConfig();
  config.holidays = holidays;
  saveConfig(config);
  return { success: true };
}

// ============================================================================
// BASELINE MANAGEMENT
// ============================================================================

/**
 * Saves current dates as baseline
 */
function saveBaseline() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(DATA_SHEET_NAME);

  if (!dataSheet) {
    SpreadsheetApp.getUi().alert('Data sheet not found.');
    return;
  }

  const lastRow = dataSheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('No tasks found.');
    return;
  }

  // Copy Start Date to Baseline Start, End Date to Baseline End
  const startDates = dataSheet.getRange(2, COL.START_DATE + 1, lastRow - 1, 1).getValues();
  const endDates = dataSheet.getRange(2, COL.END_DATE + 1, lastRow - 1, 1).getValues();

  dataSheet.getRange(2, COL.BASELINE_START + 1, lastRow - 1, 1).setValues(startDates);
  dataSheet.getRange(2, COL.BASELINE_END + 1, lastRow - 1, 1).setValues(endDates);

  SpreadsheetApp.getUi().alert(
    'Baseline Saved',
    'Current schedule has been saved as baseline.\n\nYou can now track schedule changes against this baseline.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Clears baseline dates
 */
function clearBaseline() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Clear Baseline',
    'Are you sure you want to clear the baseline? This cannot be undone.',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(DATA_SHEET_NAME);

  if (!dataSheet) return;

  const lastRow = dataSheet.getLastRow();
  if (lastRow > 1) {
    dataSheet.getRange(2, COL.BASELINE_START + 1, lastRow - 1, 2).clearContent();
  }

  ui.alert('Baseline cleared.');
}

/**
 * Shows baseline comparison report
 */
function showBaselineComparison() {
  const { tasks } = readTaskData();

  if (tasks.length === 0) {
    SpreadsheetApp.getUi().alert('No tasks found.');
    return;
  }

  let report = 'BASELINE COMPARISON REPORT\n';
  report += '=' .repeat(40) + '\n\n';

  let totalSlippage = 0;
  let tasksWithSlippage = 0;

  tasks.forEach(task => {
    if (task.baselineEnd && task.endDate) {
      const slippage = task.slippage;
      if (slippage !== 0) {
        tasksWithSlippage++;
        totalSlippage += slippage;
        const status = slippage > 0 ? `${slippage} days LATE` : `${Math.abs(slippage)} days EARLY`;
        report += `${task.id}: ${task.name}\n  ${status}\n\n`;
      }
    }
  });

  if (tasksWithSlippage === 0) {
    report += 'All tasks are on schedule!\n';
  } else {
    report += `\nSUMMARY:\n`;
    report += `Tasks with variance: ${tasksWithSlippage}\n`;
    report += `Average slippage: ${(totalSlippage / tasksWithSlippage).toFixed(1)} days\n`;
  }

  SpreadsheetApp.getUi().alert('Baseline Comparison', report, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ============================================================================
// AUTO-SCHEDULING
// ============================================================================

/**
 * Auto-schedules tasks based on dependencies
 * Cascades date changes through the dependency chain
 */
function autoScheduleTasks() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Auto-Schedule Tasks',
    'This will recalculate task dates based on dependencies.\n\nTasks will be scheduled as early as possible based on their predecessors.\n\nContinue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  const { tasks, taskMap } = readTaskData();
  const config = getConfig();

  if (tasks.length === 0) {
    ui.alert('No tasks found.');
    return;
  }

  // Topological sort for dependency order
  const sorted = topologicalSort(tasks, taskMap);

  // Schedule each task
  sorted.forEach(task => {
    if (task.dependencies.length === 0) {
      // No dependencies, keep original dates
      return;
    }

    let earliestStart = null;

    task.dependencies.forEach(depId => {
      const predecessor = taskMap.get(depId);
      if (!predecessor) return;

      const depType = task.dependencyType || 'FS';
      const lag = task.lagLead || 0;

      let calculatedStart;
      switch (depType) {
        case 'FS':
          calculatedStart = addWorkDays(predecessor.endDate, 1 + lag, task.workDaysOnly && config.excludeWeekends, config.holidays);
          break;
        case 'SS':
          calculatedStart = addWorkDays(predecessor.startDate, lag, task.workDaysOnly && config.excludeWeekends, config.holidays);
          break;
        case 'FF':
          const duration = task.duration || getWorkDaysDuration(task.startDate, task.endDate, task.workDaysOnly && config.excludeWeekends, config.holidays);
          calculatedStart = addWorkDays(predecessor.endDate, lag - duration + 1, task.workDaysOnly && config.excludeWeekends, config.holidays);
          break;
        case 'SF':
          const dur = task.duration || getWorkDaysDuration(task.startDate, task.endDate, task.workDaysOnly && config.excludeWeekends, config.holidays);
          calculatedStart = addWorkDays(predecessor.startDate, lag - dur + 1, task.workDaysOnly && config.excludeWeekends, config.holidays);
          break;
        default:
          calculatedStart = addWorkDays(predecessor.endDate, 1 + lag, task.workDaysOnly && config.excludeWeekends, config.holidays);
      }

      if (!earliestStart || calculatedStart > earliestStart) {
        earliestStart = calculatedStart;
      }
    });

    if (earliestStart) {
      const duration = task.duration || getWorkDaysDuration(task.startDate, task.endDate, task.workDaysOnly && config.excludeWeekends, config.holidays);
      task.startDate = earliestStart;
      task.endDate = addWorkDays(earliestStart, duration - 1, task.workDaysOnly && config.excludeWeekends, config.holidays);
    }
  });

  // Write updated dates back to sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(DATA_SHEET_NAME);

  sorted.forEach(task => {
    dataSheet.getRange(task.rowNumber, COL.START_DATE + 1).setValue(task.startDate);
    dataSheet.getRange(task.rowNumber, COL.END_DATE + 1).setValue(task.endDate);
  });

  ui.alert('Auto-Schedule Complete', `${sorted.length} tasks have been rescheduled based on dependencies.`, ui.ButtonSet.OK);
}

/**
 * Topological sort for dependency ordering
 */
function topologicalSort(tasks, taskMap) {
  const sorted = [];
  const visited = new Set();
  const temp = new Set();

  function visit(task) {
    if (temp.has(task.id)) return; // Circular dependency, skip
    if (visited.has(task.id)) return;

    temp.add(task.id);

    task.dependencies.forEach(depId => {
      const dep = taskMap.get(depId);
      if (dep) visit(dep);
    });

    temp.delete(task.id);
    visited.add(task.id);
    sorted.push(task);
  }

  tasks.forEach(task => {
    if (!visited.has(task.id)) {
      visit(task);
    }
  });

  return sorted;
}

// ============================================================================
// ROLLUP SUMMARIES
// ============================================================================

/**
 * Recalculates parent task rollups from children
 */
function recalculateRollups() {
  const { tasks, taskMap } = readTaskData();

  if (tasks.length === 0) {
    SpreadsheetApp.getUi().alert('No tasks found.');
    return;
  }

  // Group children by parent
  const childrenByParent = new Map();
  tasks.forEach(task => {
    if (task.parentTask) {
      if (!childrenByParent.has(task.parentTask)) {
        childrenByParent.set(task.parentTask, []);
      }
      childrenByParent.get(task.parentTask).push(task);
    }
  });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(DATA_SHEET_NAME);
  let updatedCount = 0;

  // Calculate rollups for each parent
  childrenByParent.forEach((children, parentId) => {
    const parent = taskMap.get(parentId);
    if (!parent) return;

    // Calculate rolled-up values
    const startDates = children.map(c => c.startDate).filter(d => d);
    const endDates = children.map(c => c.endDate).filter(d => d);
    const percentages = children.map(c => c.percentComplete);
    const budgets = children.map(c => c.budgetCost);
    const actuals = children.map(c => c.actualCost);

    if (startDates.length > 0) {
      const rollupStart = new Date(Math.min(...startDates.map(d => d.getTime())));
      const rollupEnd = new Date(Math.max(...endDates.map(d => d.getTime())));
      const rollupPercent = Math.round(percentages.reduce((a, b) => a + b, 0) / percentages.length);
      const rollupBudget = budgets.reduce((a, b) => a + b, 0);
      const rollupActual = actuals.reduce((a, b) => a + b, 0);

      // Update parent task
      dataSheet.getRange(parent.rowNumber, COL.START_DATE + 1).setValue(rollupStart);
      dataSheet.getRange(parent.rowNumber, COL.END_DATE + 1).setValue(rollupEnd);
      dataSheet.getRange(parent.rowNumber, COL.PERCENT_COMPLETE + 1).setValue(rollupPercent);
      dataSheet.getRange(parent.rowNumber, COL.BUDGET_COST + 1).setValue(rollupBudget);
      dataSheet.getRange(parent.rowNumber, COL.ACTUAL_COST + 1).setValue(rollupActual);

      updatedCount++;
    }
  });

  SpreadsheetApp.getUi().alert('Rollup Complete', `Updated ${updatedCount} parent tasks with rolled-up values from children.`, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ============================================================================
// FILTERS & VIEWS
// ============================================================================

/**
 * Shows all tasks (removes filter)
 */
function filterShowAll() {
  const config = getConfig();
  config.activeFilter = 'all';
  saveConfig(config);
  showFilteredView([]);
}

/**
 * Filters to show only current user's tasks
 */
function filterMyTasks() {
  const email = Session.getActiveUser().getEmail();
  const config = getConfig();
  config.activeFilter = 'myTasks';
  saveConfig(config);

  const { tasks } = readTaskData();
  const myTasks = tasks.filter(t =>
    t.owner.toLowerCase().includes(email.toLowerCase()) ||
    t.owner.toLowerCase().includes(email.split('@')[0].toLowerCase())
  );

  showFilteredView(myTasks.map(t => t.rowNumber));
}

/**
 * Filters to show only overdue tasks
 */
function filterOverdue() {
  const config = getConfig();
  config.activeFilter = 'overdue';
  saveConfig(config);

  const { tasks } = readTaskData();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const overdue = tasks.filter(t =>
    t.endDate < today && t.percentComplete < 100
  );

  showFilteredView(overdue.map(t => t.rowNumber));

  if (overdue.length === 0) {
    SpreadsheetApp.getUi().alert('No overdue tasks found!');
  }
}

/**
 * Filters to show only at-risk tasks
 */
function filterAtRisk() {
  const config = getConfig();
  config.activeFilter = 'atRisk';
  saveConfig(config);

  const { tasks } = readTaskData();
  const atRisk = tasks.filter(t =>
    t.status === 'At Risk' || t.health === 'Red' || t.health === 'Yellow'
  );

  showFilteredView(atRisk.map(t => t.rowNumber));

  if (atRisk.length === 0) {
    SpreadsheetApp.getUi().alert('No at-risk tasks found!');
  }
}

/**
 * Filters to show only critical path tasks
 */
function filterCriticalPath() {
  const config = getConfig();
  config.activeFilter = 'criticalPath';
  saveConfig(config);

  const { tasks, taskMap } = readTaskData();
  calculateCriticalPath(tasks, taskMap);

  const criticalTasks = tasks.filter(t => t.isOnCriticalPath);
  showFilteredView(criticalTasks.map(t => t.rowNumber));

  if (criticalTasks.length === 0) {
    SpreadsheetApp.getUi().alert('No critical path tasks found.');
  }
}

/**
 * Helper to show filtered rows
 */
function showFilteredView(visibleRows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(DATA_SHEET_NAME);

  if (!dataSheet) return;

  const lastRow = dataSheet.getLastRow();
  if (lastRow < 2) return;

  // Show all rows first
  dataSheet.showRows(2, lastRow - 1);

  // If no filter (show all), we're done
  if (visibleRows.length === 0) return;

  // Hide rows not in visible list
  for (let row = 2; row <= lastRow; row++) {
    if (!visibleRows.includes(row)) {
      dataSheet.hideRows(row);
    }
  }
}

// ============================================================================
// RESOURCE WORKLOAD VIEW
// ============================================================================

/**
 * Shows resource workload analysis
 */
function showResourceWorkload() {
  const { tasks } = readTaskData();

  if (tasks.length === 0) {
    SpreadsheetApp.getUi().alert('No tasks found.');
    return;
  }

  // Group tasks by owner
  const workloadByOwner = new Map();

  tasks.forEach(task => {
    const owner = task.owner || 'Unassigned';
    if (!workloadByOwner.has(owner)) {
      workloadByOwner.set(owner, {
        tasks: [],
        totalDays: 0,
        totalAllocation: 0
      });
    }

    const data = workloadByOwner.get(owner);
    data.tasks.push(task);
    data.totalDays += task.duration || 0;
    data.totalAllocation += (task.resourcePercent || 100) * (task.duration || 0) / 100;
  });

  // Build report
  let report = 'RESOURCE WORKLOAD REPORT\n';
  report += '=' .repeat(40) + '\n\n';

  workloadByOwner.forEach((data, owner) => {
    const overallocated = data.totalAllocation > 40; // More than 8 weeks equivalent

    report += `${owner}\n`;
    report += `-`.repeat(30) + '\n';
    report += `  Tasks assigned: ${data.tasks.length}\n`;
    report += `  Total work days: ${data.totalDays}\n`;
    report += `  Effective days (with allocation): ${data.totalAllocation.toFixed(1)}\n`;

    if (overallocated) {
      report += `  ⚠️ POTENTIALLY OVERALLOCATED\n`;
    }

    report += '\n';
  });

  SpreadsheetApp.getUi().alert('Resource Workload', report, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ============================================================================
// COST TRACKING
// ============================================================================

/**
 * Shows cost summary report
 */
function showCostSummary() {
  const { tasks } = readTaskData();

  if (tasks.length === 0) {
    SpreadsheetApp.getUi().alert('No tasks found.');
    return;
  }

  let totalBudget = 0;
  let totalActual = 0;
  let tasksBehindBudget = [];
  let tasksUnderBudget = [];

  tasks.forEach(task => {
    totalBudget += task.budgetCost || 0;
    totalActual += task.actualCost || 0;

    if (task.budgetCost > 0) {
      const variance = task.costVariance;
      if (variance < 0) {
        tasksBehindBudget.push({ task, variance });
      } else if (variance > task.budgetCost * 0.1) { // More than 10% under
        tasksUnderBudget.push({ task, variance });
      }
    }
  });

  const totalVariance = totalBudget - totalActual;
  const variancePercent = totalBudget > 0 ? (totalVariance / totalBudget * 100).toFixed(1) : 0;

  let report = 'COST SUMMARY REPORT\n';
  report += '=' .repeat(40) + '\n\n';
  report += `Total Budget: $${totalBudget.toLocaleString()}\n`;
  report += `Total Actual: $${totalActual.toLocaleString()}\n`;
  report += `Variance: $${totalVariance.toLocaleString()} (${variancePercent}%)\n`;
  report += `Status: ${totalVariance >= 0 ? '✓ UNDER BUDGET' : '⚠️ OVER BUDGET'}\n\n`;

  if (tasksBehindBudget.length > 0) {
    report += 'OVER BUDGET TASKS:\n';
    tasksBehindBudget.forEach(({ task, variance }) => {
      report += `  ${task.id}: $${Math.abs(variance).toLocaleString()} over\n`;
    });
    report += '\n';
  }

  SpreadsheetApp.getUi().alert('Cost Summary', report, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ============================================================================
// CONDITIONAL FORMATTING
// ============================================================================

/**
 * Applies conditional formatting rules to the data sheet
 */
function applyConditionalFormatting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(DATA_SHEET_NAME);

  if (!dataSheet) {
    SpreadsheetApp.getUi().alert('Data sheet not found.');
    return;
  }

  const lastRow = dataSheet.getLastRow();
  if (lastRow < 2) return;

  const config = getConfig();
  const rules = [];

  // Get existing rules and clear them
  dataSheet.clearConditionalFormatRules();

  const dataRange = dataSheet.getRange(2, 1, lastRow - 1, COLUMN_HEADERS.length);

  if (config.conditionalFormatting.overdueRed) {
    // Overdue tasks (end date < today AND % complete < 100)
    const overdueRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND($D2<TODAY(), $G2<100)`)
      .setBackground('#FFCDD2')
      .setRanges([dataRange])
      .build();
    rules.push(overdueRule);
  }

  if (config.conditionalFormatting.atRiskYellow) {
    // At Risk status or Yellow/Red health
    const atRiskRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=OR($N2="At Risk", $O2="Yellow", $O2="Red")`)
      .setBackground('#FFF9C4')
      .setRanges([dataRange])
      .build();
    rules.push(atRiskRule);
  }

  if (config.conditionalFormatting.completedGreen) {
    // Completed tasks
    const completedRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=OR($G2=100, $N2="Complete")`)
      .setBackground('#C8E6C9')
      .setRanges([dataRange])
      .build();
    rules.push(completedRule);
  }

  // Health indicator colors in Health column
  const healthRange = dataSheet.getRange(2, COL.HEALTH + 1, lastRow - 1, 1);

  const greenHealthRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Green')
    .setBackground('#C8E6C9')
    .setRanges([healthRange])
    .build();
  rules.push(greenHealthRule);

  const yellowHealthRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Yellow')
    .setBackground('#FFF9C4')
    .setRanges([healthRange])
    .build();
  rules.push(yellowHealthRule);

  const redHealthRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Red')
    .setBackground('#FFCDD2')
    .setRanges([healthRange])
    .build();
  rules.push(redHealthRule);

  // Status column colors
  const statusRange = dataSheet.getRange(2, COL.STATUS + 1, lastRow - 1, 1);

  const completeStatusRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Complete')
    .setBackground('#C8E6C9')
    .setRanges([statusRange])
    .build();
  rules.push(completeStatusRule);

  const inProgressRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('In Progress')
    .setBackground('#BBDEFB')
    .setRanges([statusRange])
    .build();
  rules.push(inProgressRule);

  const atRiskStatusRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('At Risk')
    .setBackground('#FFCDD2')
    .setRanges([statusRange])
    .build();
  rules.push(atRiskStatusRule);

  const onHoldRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('On Hold')
    .setBackground('#FFE0B2')
    .setRanges([statusRange])
    .build();
  rules.push(onHoldRule);

  dataSheet.setConditionalFormatRules(rules);

  SpreadsheetApp.getUi().alert('Conditional Formatting Applied', 'Formatting rules have been applied:\n\n- Overdue tasks: Light red\n- At Risk/Yellow/Red health: Light yellow\n- Completed tasks: Light green\n- Health & Status columns color-coded', SpreadsheetApp.getUi().ButtonSet.OK);
}

// ============================================================================
// PRINT SETTINGS
// ============================================================================

/**
 * Shows print settings dialog
 */
function showPrintSettings() {
  const html = HtmlService.createHtmlOutputFromFile('PrintSettings')
    .setWidth(450)
    .setHeight(350)
    .setTitle('Print Settings');

  SpreadsheetApp.getUi().showModalDialog(html, 'Print Settings');
}

/**
 * Gets print settings
 */
function getPrintSettings() {
  const config = getConfig();
  return {
    dateRange: config.printDateRange,
    startDate: config.printStartDate,
    endDate: config.printEndDate,
    scale: config.printScale,
    orientation: config.printOrientation
  };
}

/**
 * Saves print settings
 */
function savePrintSettings(settings) {
  const config = getConfig();
  config.printDateRange = settings.dateRange;
  config.printStartDate = settings.startDate;
  config.printEndDate = settings.endDate;
  config.printScale = settings.scale;
  config.printOrientation = settings.orientation;
  saveConfig(config);
  return { success: true };
}
