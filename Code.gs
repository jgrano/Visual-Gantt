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
const JIRA_CONFIG_KEY = 'visualGanttJiraConfig';
const SMARTSHEET_CONFIG_KEY = 'visualGanttSmartsheetConfig';
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
  'Modified'  // Tracks local changes for Jira sync
];

// Column indices (0-based) for easy reference
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
  MODIFIED: 13
};

// Default Jira configuration
const DEFAULT_JIRA_CONFIG = {
  baseUrl: '',           // e.g., 'yourcompany.atlassian.net'
  email: '',             // User's Atlassian email
  apiToken: '',          // API token (stored securely)
  pullMethod: 'jql',     // 'jql', 'project', or 'filter'
  jqlQuery: '',          // JQL query string
  projectKey: '',        // Project key for project method
  filterId: '',          // Saved filter ID
  startDateField: '',    // Custom field ID for start date (if used)
  storyPointsField: '',  // Custom field ID for story points (optional)
  lastSyncTime: null     // Timestamp of last sync
};

// Jira status category to percent complete mapping
const JIRA_STATUS_MAPPING = {
  'new': 0,
  'indeterminate': 50,
  'done': 100
};

// Jira priority mapping
const JIRA_PRIORITY_MAPPING = {
  'Highest': 'High',
  'High': 'High',
  'Medium': 'Medium',
  'Low': 'Low',
  'Lowest': 'Low'
};

// Default Smartsheet configuration
const DEFAULT_SMARTSHEET_CONFIG = {
  apiToken: '',           // Smartsheet API access token (stored securely)
  sheetId: '',            // Selected sheet ID
  sheetName: '',          // Selected sheet name
  includeChildRows: true, // Include hierarchical child rows
  syncDependencies: true, // Sync predecessor dependencies
  columnMapping: {        // Maps Visual Gantt fields to Smartsheet column IDs
    colTaskId: '',
    colTaskName: '',
    colStartDate: '',
    colEndDate: '',
    colDuration: '',
    colOwner: '',
    colPercentComplete: '',
    colPriority: '',
    colPredecessors: '',
    colCategory: ''
  },
  lastSyncTime: null      // Timestamp of last sync
};

// Smartsheet API base URL
const SMARTSHEET_API_BASE = 'https://api.smartsheet.com/2.0';

// Smartsheet priority mapping
const SMARTSHEET_PRIORITY_MAPPING = {
  'Critical': 'High',
  'High': 'High',
  'Medium': 'Medium',
  'Normal': 'Medium',
  'Low': 'Low'
};

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
// AUTHORIZATION UTILITIES
// ============================================================================

/**
 * Checks the current authorization status for the add-on
 * @returns {Object} Authorization status with isAuthorized flag and details
 */
function checkAuthorizationStatus() {
  try {
    const authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
    const authStatus = authInfo.getAuthorizationStatus();

    return {
      isAuthorized: authStatus === ScriptApp.AuthorizationStatus.NOT_REQUIRED,
      status: authStatus,
      authUrl: authInfo.getAuthorizationUrl(),
      needsAuth: authStatus === ScriptApp.AuthorizationStatus.REQUIRED
    };
  } catch (e) {
    Logger.log('Error checking authorization: ' + e.message);
    return {
      isAuthorized: false,
      status: null,
      authUrl: null,
      needsAuth: true,
      error: e.message
    };
  }
}

/**
 * Checks if the add-on has the required authorization to perform operations
 * This should be called before any operation that requires spreadsheet access
 * @returns {boolean} True if authorized, false otherwise
 */
function hasFileAccess() {
  try {
    // Try a minimal operation that requires file access
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      return false;
    }
    // Try to get spreadsheet name - this requires file scope
    ss.getName();
    return true;
  } catch (e) {
    Logger.log('File access check failed: ' + e.message);
    return false;
  }
}

/**
 * Requires authorization before proceeding with an operation
 * Shows an authorization dialog if not authorized
 * @param {string} operationName - Name of the operation for error messages
 * @returns {boolean} True if authorized and can proceed
 */
function requireAuthorization(operationName) {
  if (!hasFileAccess()) {
    const ui = SpreadsheetApp.getUi();
    const authStatus = checkAuthorizationStatus();

    let message = 'This add-on needs permission to access this spreadsheet to ' + operationName + '.\n\n';

    if (authStatus.authUrl) {
      message += 'Please use the "Visual Gantt > Request Authorization" menu item to grant access, ';
      message += 'or visit this URL:\n' + authStatus.authUrl;
    } else {
      message += 'Please:\n';
      message += '1. Go to Extensions > Visual Gantt > Request Authorization\n';
      message += '2. Follow the authorization prompts\n';
      message += '3. Try this operation again';
    }

    ui.alert('Authorization Required', message, ui.ButtonSet.OK);
    return false;
  }
  return true;
}

/**
 * Triggers the authorization flow by attempting an operation that requires auth
 * This is called from the menu to let users proactively authorize
 */
function requestAuthorization() {
  const ui = SpreadsheetApp.getUi();

  try {
    // Check current status
    const authStatus = checkAuthorizationStatus();

    if (authStatus.isAuthorized) {
      // Already authorized - verify by testing actual access
      if (hasFileAccess()) {
        ui.alert('Authorization Status',
          'Visual Gantt is fully authorized for this spreadsheet!\n\n' +
          'You can now use all features of the add-on.',
          ui.ButtonSet.OK);
        return;
      }
    }

    // Try to trigger authorization by accessing the spreadsheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const name = ss.getName();

    // If we get here, authorization was successful
    ui.alert('Authorization Successful',
      'Visual Gantt has been authorized for: ' + name + '\n\n' +
      'You can now use all features of the add-on.',
      ui.ButtonSet.OK);

  } catch (e) {
    Logger.log('Authorization request failed: ' + e.message);

    // Check if we can get an auth URL
    const authStatus = checkAuthorizationStatus();

    let message = 'Authorization is required to use Visual Gantt.\n\n';

    if (e.message.includes('PERMISSION_DENIED') || e.message.includes('Authorization')) {
      message += 'It appears the add-on does not have permission to access this spreadsheet.\n\n';
      message += 'Please try:\n';
      message += '1. Refresh the page\n';
      message += '2. Click the authorization prompt when it appears\n';
      message += '3. If no prompt appears, try removing and re-adding the add-on\n\n';

      if (authStatus.authUrl) {
        message += 'Or visit: ' + authStatus.authUrl;
      }
    } else {
      message += 'Error: ' + e.message;
    }

    ui.alert('Authorization Required', message, ui.ButtonSet.OK);
  }
}

/**
 * Gets detailed authorization diagnostics for troubleshooting
 */
function getAuthorizationDiagnostics() {
  const diagnostics = {
    timestamp: new Date().toISOString(),
    authStatus: null,
    fileAccess: false,
    propertiesAccess: {
      document: false,
      script: false,
      user: false
    },
    spreadsheetAccess: {
      getActive: false,
      getName: false,
      getSheets: false
    },
    errors: []
  };

  // Check authorization status
  try {
    diagnostics.authStatus = checkAuthorizationStatus();
  } catch (e) {
    diagnostics.errors.push('Auth status check: ' + e.message);
  }

  // Check PropertiesService access
  try {
    PropertiesService.getDocumentProperties();
    diagnostics.propertiesAccess.document = true;
  } catch (e) {
    diagnostics.errors.push('DocumentProperties: ' + e.message);
  }

  try {
    PropertiesService.getScriptProperties();
    diagnostics.propertiesAccess.script = true;
  } catch (e) {
    diagnostics.errors.push('ScriptProperties: ' + e.message);
  }

  try {
    PropertiesService.getUserProperties();
    diagnostics.propertiesAccess.user = true;
  } catch (e) {
    diagnostics.errors.push('UserProperties: ' + e.message);
  }

  // Check Spreadsheet access
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    diagnostics.spreadsheetAccess.getActive = !!ss;

    if (ss) {
      try {
        ss.getName();
        diagnostics.spreadsheetAccess.getName = true;
      } catch (e) {
        diagnostics.errors.push('Spreadsheet.getName: ' + e.message);
      }

      try {
        ss.getSheets();
        diagnostics.spreadsheetAccess.getSheets = true;
      } catch (e) {
        diagnostics.errors.push('Spreadsheet.getSheets: ' + e.message);
      }
    }
  } catch (e) {
    diagnostics.errors.push('getActiveSpreadsheet: ' + e.message);
  }

  diagnostics.fileAccess = diagnostics.spreadsheetAccess.getName &&
                           diagnostics.spreadsheetAccess.getSheets;

  return diagnostics;
}

/**
 * Shows authorization diagnostics dialog for troubleshooting
 */
function showAuthorizationDiagnostics() {
  const ui = SpreadsheetApp.getUi();
  const diagnostics = getAuthorizationDiagnostics();

  let message = 'Authorization Diagnostics\n';
  message += '========================\n\n';
  message += 'File Access: ' + (diagnostics.fileAccess ? 'YES' : 'NO') + '\n\n';

  message += 'Properties Service:\n';
  message += '  - Document: ' + (diagnostics.propertiesAccess.document ? 'OK' : 'FAILED') + '\n';
  message += '  - Script: ' + (diagnostics.propertiesAccess.script ? 'OK' : 'FAILED') + '\n';
  message += '  - User: ' + (diagnostics.propertiesAccess.user ? 'OK' : 'FAILED') + '\n\n';

  message += 'Spreadsheet Access:\n';
  message += '  - Get Active: ' + (diagnostics.spreadsheetAccess.getActive ? 'OK' : 'FAILED') + '\n';
  message += '  - Get Name: ' + (diagnostics.spreadsheetAccess.getName ? 'OK' : 'FAILED') + '\n';
  message += '  - Get Sheets: ' + (diagnostics.spreadsheetAccess.getSheets ? 'OK' : 'FAILED') + '\n\n';

  if (diagnostics.errors.length > 0) {
    message += 'Errors:\n';
    diagnostics.errors.forEach(err => {
      message += '  - ' + err + '\n';
    });
  }

  ui.alert('Authorization Diagnostics', message, ui.ButtonSet.OK);
}

// ============================================================================
// PROPERTIES SERVICE STORAGE
// ============================================================================

// Note: Configuration is stored using PropertiesService.getDocumentProperties()
// DocumentProperties requires file-scope authorization for add-ons.
// ScriptProperties is used as a fallback when document access is not available.
// UserProperties stores user-specific settings across all documents.

// ============================================================================
// MENU & TRIGGERS
// ============================================================================

/**
 * Creates the add-on menu when the spreadsheet opens
 * Handles both full and limited authorization modes
 */
function onOpen(e) {
  const ui = SpreadsheetApp.getUi();

  // Determine auth mode from event or default to FULL
  // In limited auth mode (e.g., when add-on first loads), some operations may not be available
  const authMode = e && e.authMode ? e.authMode : ScriptApp.AuthMode.FULL;

  // Create Jira submenu
  const jiraMenu = ui.createMenu('Jira')
    .addItem('Configure Connection', 'showJiraConfigDialog')
    .addSeparator()
    .addItem('Pull from Jira', 'pullFromJira')
    .addItem('Pull Selected Rows', 'pullSelectedFromJira')
    .addSeparator()
    .addItem('Push to Jira', 'pushToJira')
    .addItem('Push Selected Rows', 'pushSelectedToJira')
    .addSeparator()
    .addItem('View Sync Status', 'showSyncStatus');

  // Create Smartsheet submenu
  const smartsheetMenu = ui.createMenu('Smartsheet')
    .addItem('Configure Connection', 'showSmartsheetConfigDialog')
    .addSeparator()
    .addItem('Pull from Smartsheet', 'pullFromSmartsheet')
    .addItem('Pull Selected Rows', 'pullSelectedFromSmartsheet')
    .addSeparator()
    .addItem('Push to Smartsheet', 'pushToSmartsheet')
    .addItem('Push Selected Rows', 'pushSelectedToSmartsheet')
    .addSeparator()
    .addItem('View Sync Status', 'showSmartsheetSyncStatus');

  // Create Authorization submenu for troubleshooting
  const authMenu = ui.createMenu('Authorization')
    .addItem('Request Authorization', 'requestAuthorization')
    .addItem('Check Authorization Status', 'showAuthorizationDiagnostics')
    .addSeparator()
    .addItem('Storage Diagnostics', 'showStorageDiagnostics')
    .addItem('Migrate Settings', 'showMigrationDialog');

  ui.createMenu('Visual Gantt')
    .addItem('Generate Timeline', 'generateTimeline')
    .addItem('Refresh Timeline', 'refreshTimeline')
    .addSeparator()
    .addItem('Configure Settings', 'showConfigDialog')
    .addSeparator()
    .addSubMenu(jiraMenu)
    .addSubMenu(smartsheetMenu)
    .addSeparator()
    .addItem('Export as PNG', 'exportAsPng')
    .addItem('Export as PDF', 'exportAsPdf')
    .addSeparator()
    .addItem('Setup Data Sheet', 'setupDataSheet')
    .addItem('Create Sample Data', 'createSampleData')
    .addSeparator()
    .addSubMenu(authMenu)
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
 * This is triggered by Google Sheets when the user authorizes the add-on for this file
 */
function onFileScopeGranted(e) {
  // Log successful file scope grant
  Logger.log('File scope granted - initializing add-on');
  onOpen(e);
}

/**
 * Creates the homepage card for the Workspace Add-on sidebar
 * This displays before the user grants file scope access
 * @param {Object} e - Event object
 * @returns {CardService.Card} The homepage card
 */
function onHomepage(e) {
  const builder = CardService.newCardBuilder();

  // Header
  const header = CardService.newCardHeader()
    .setTitle('Visual Gantt')
    .setSubtitle('Project Timeline Add-on')
    .setImageUrl('https://www.gstatic.com/images/branding/product/2x/apps_script_48dp.png')
    .setImageStyle(CardService.ImageStyle.CIRCLE);

  builder.setHeader(header);

  // Check authorization status
  const hasAccess = hasFileAccess();

  if (hasAccess) {
    // User has granted file scope - show main functionality
    const authorizedSection = CardService.newCardSection()
      .setHeader('Ready to Use')
      .addWidget(CardService.newTextParagraph()
        .setText('Visual Gantt is authorized for this spreadsheet. Use the menu at Extensions > Visual Gantt to:'))
      .addWidget(CardService.newTextParagraph()
        .setText('• Generate Timeline - Create Gantt charts\n• Configure Settings - Customize appearance\n• Jira/Smartsheet - Sync with external tools\n• Export - Save as PNG or PDF'));

    const menuButton = CardService.newTextButton()
      .setText('Open Extensions Menu')
      .setOnClickAction(CardService.newAction().setFunctionName('showMenuInstructions'));

    authorizedSection.addWidget(menuButton);
    builder.addSection(authorizedSection);

  } else {
    // User needs to grant file scope
    const welcomeSection = CardService.newCardSection()
      .setHeader('Welcome!')
      .addWidget(CardService.newTextParagraph()
        .setText('Visual Gantt creates professional Gantt chart timelines for your project management needs.'))
      .addWidget(CardService.newTextParagraph()
        .setText('Features include:\n• Visual timeline generation\n• Jira & Smartsheet integration\n• PNG/PDF export\n• Customizable styling'));

    builder.addSection(welcomeSection);

    const authSection = CardService.newCardSection()
      .setHeader('Authorization Required')
      .addWidget(CardService.newTextParagraph()
        .setText('To use Visual Gantt with this spreadsheet, please grant access by clicking the button below.'))
      .addWidget(CardService.newTextParagraph()
        .setText('This allows the add-on to:\n• Read your project data\n• Create timeline sheets\n• Save your settings'));

    const authButton = CardService.newTextButton()
      .setText('Grant Access')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOnClickAction(CardService.newAction().setFunctionName('requestFileAccess'));

    authSection.addWidget(authButton);
    builder.addSection(authSection);
  }

  // Help section
  const helpSection = CardService.newCardSection()
    .setCollapsible(true)
    .setHeader('Help & Support');

  const helpButton = CardService.newTextButton()
    .setText('View Help')
    .setOnClickAction(CardService.newAction().setFunctionName('showHelp'));

  const diagButton = CardService.newTextButton()
    .setText('Check Authorization Status')
    .setOnClickAction(CardService.newAction().setFunctionName('showAuthDiagCard'));

  helpSection.addWidget(helpButton);
  helpSection.addWidget(diagButton);
  builder.addSection(helpSection);

  return builder.build();
}

/**
 * Requests file access - triggers the authorization flow
 * @returns {CardService.ActionResponse} Response to update the card
 */
function requestFileAccess() {
  try {
    // Attempt to access the spreadsheet, which will trigger auth if needed
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) {
      ss.getName(); // This requires file scope
    }

    // If we get here, we have access - refresh the homepage
    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().updateCard(onHomepage({})))
      .setNotification(CardService.newNotification()
        .setText('Authorization successful! You can now use Visual Gantt.'))
      .build();

  } catch (e) {
    Logger.log('File access request failed: ' + e.message);

    // Show error notification
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText('Please follow the authorization prompt to grant access.'))
      .build();
  }
}

/**
 * Shows instructions about using the menu
 * @returns {CardService.ActionResponse} Response with notification
 */
function showMenuInstructions() {
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification()
      .setText('Use the menu: Extensions > Visual Gantt to access all features'))
    .build();
}

/**
 * Shows authorization diagnostics as a card
 * @returns {CardService.ActionResponse} Response with diagnostics card
 */
function showAuthDiagCard() {
  const diagnostics = getAuthorizationDiagnostics();

  const builder = CardService.newCardBuilder();

  builder.setHeader(CardService.newCardHeader()
    .setTitle('Authorization Diagnostics')
    .setSubtitle(diagnostics.fileAccess ? 'Fully Authorized' : 'Authorization Needed'));

  const statusSection = CardService.newCardSection()
    .setHeader('Status');

  statusSection.addWidget(CardService.newDecoratedText()
    .setText('File Access')
    .setBottomLabel(diagnostics.fileAccess ? '✓ Granted' : '✗ Not Granted'));

  statusSection.addWidget(CardService.newDecoratedText()
    .setText('Document Properties')
    .setBottomLabel(diagnostics.propertiesAccess.document ? '✓ Available' : '✗ Not Available'));

  statusSection.addWidget(CardService.newDecoratedText()
    .setText('Script Properties')
    .setBottomLabel(diagnostics.propertiesAccess.script ? '✓ Available' : '✗ Not Available'));

  builder.addSection(statusSection);

  if (diagnostics.errors.length > 0) {
    const errorSection = CardService.newCardSection()
      .setHeader('Errors')
      .setCollapsible(true);

    diagnostics.errors.forEach(err => {
      errorSection.addWidget(CardService.newTextParagraph().setText('• ' + err));
    });

    builder.addSection(errorSection);
  }

  const backButton = CardService.newTextButton()
    .setText('Back to Home')
    .setOnClickAction(CardService.newAction().setFunctionName('goToHomepage'));

  builder.addSection(CardService.newCardSection().addWidget(backButton));

  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().pushCard(builder.build()))
    .build();
}

/**
 * Navigates back to the homepage card
 * @returns {CardService.ActionResponse} Response to show homepage
 */
function goToHomepage() {
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().popToRoot().updateCard(onHomepage({})))
    .build();
}

// ============================================================================
// SETUP FUNCTIONS
// ============================================================================

/**
 * Creates the data sheet template with proper column headers
 */
function setupDataSheet() {
  // Check authorization before proceeding
  if (!requireAuthorization('create the data sheet')) {
    return;
  }

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

  // Set column widths (includes Modified column)
  const columnWidths = [80, 200, 100, 100, 100, 120, 80, 80, 150, 120, 150, 100, 150, 70];
  columnWidths.forEach((width, index) => {
    dataSheet.setColumnWidth(index + 1, width);
  });

  // Add data validation for Modified column (Yes or empty)
  const modifiedRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Yes', ''], true)
    .setAllowInvalid(true)
    .build();
  dataSheet.getRange('N2:N1000').setDataValidation(modifiedRule);

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
  // Check authorization before proceeding
  if (!requireAuthorization('create sample data')) {
    return;
  }

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
    ['MECH-001', 'Mechanical Design Phase', addDays(baseDate, 0), addDays(baseDate, 45), '', 'John Smith', 75, 'High', 'https://jira.example.com/MECH-001', '', '', 'Task', 'Mechanical', ''],
    ['MECH-002', 'Enclosure CAD Design', addDays(baseDate, 0), addDays(baseDate, 14), '', 'John Smith', 100, 'High', 'https://jira.example.com/MECH-002', 'MECH-001', '', 'Task', 'Mechanical', ''],
    ['MECH-003', 'Thermal Analysis', addDays(baseDate, 7), addDays(baseDate, 21), '', 'Sarah Johnson', 80, 'High', 'https://jira.example.com/MECH-003', 'MECH-001', 'MECH-002', 'Task', 'Mechanical', ''],
    ['MECH-004', 'Prototype Fabrication', addDays(baseDate, 21), addDays(baseDate, 35), '', 'John Smith', 40, 'Medium', 'https://jira.example.com/MECH-004', 'MECH-001', 'MECH-003', 'Task', 'Mechanical', ''],
    ['MECH-005', 'Mechanical Design Complete', addDays(baseDate, 45), addDays(baseDate, 45), '', 'John Smith', 0, 'High', 'https://jira.example.com/MECH-005', '', 'MECH-004', 'Milestone', 'Mechanical', ''],

    // Electrical Engineering Tasks
    ['ELEC-001', 'Electrical Design Phase', addDays(baseDate, 5), addDays(baseDate, 50), '', 'Mike Chen', 60, 'High', 'https://jira.example.com/ELEC-001', '', '', 'Task', 'Electrical', ''],
    ['ELEC-002', 'Power Stage Design', addDays(baseDate, 5), addDays(baseDate, 20), '', 'Mike Chen', 100, 'High', 'https://jira.example.com/ELEC-002', 'ELEC-001', '', 'Task', 'Electrical', ''],
    ['ELEC-003', 'PCB Layout', addDays(baseDate, 15), addDays(baseDate, 30), '', 'Lisa Wong', 70, 'High', 'https://jira.example.com/ELEC-003', 'ELEC-001', 'ELEC-002', 'Task', 'Electrical', ''],
    ['ELEC-004', 'EMC Pre-compliance', addDays(baseDate, 30), addDays(baseDate, 40), '', 'Mike Chen', 20, 'Medium', 'https://jira.example.com/ELEC-004', 'ELEC-001', 'ELEC-003', 'Task', 'Electrical', ''],
    ['ELEC-005', 'Electrical Validation Complete', addDays(baseDate, 50), addDays(baseDate, 50), '', 'Mike Chen', 0, 'High', 'https://jira.example.com/ELEC-005', '', 'ELEC-004', 'Milestone', 'Electrical', ''],

    // Firmware Development Tasks
    ['FW-001', 'Firmware Development', addDays(baseDate, 10), addDays(baseDate, 55), '', 'Alex Rivera', 45, 'High', 'https://jira.example.com/FW-001', '', '', 'Task', 'Firmware', ''],
    ['FW-002', 'CAN Protocol Implementation', addDays(baseDate, 10), addDays(baseDate, 25), '', 'Alex Rivera', 100, 'High', 'https://jira.example.com/FW-002', 'FW-001', '', 'Task', 'Firmware', ''],
    ['FW-003', 'Bidirectional Control Logic', addDays(baseDate, 20), addDays(baseDate, 40), '', 'Alex Rivera', 50, 'High', 'https://jira.example.com/FW-003', 'FW-001', 'FW-002', 'Task', 'Firmware', ''],
    ['FW-004', 'Safety State Machine', addDays(baseDate, 35), addDays(baseDate, 50), '', 'David Park', 15, 'High', 'https://jira.example.com/FW-004', 'FW-001', 'FW-003', 'Task', 'Firmware', ''],
    ['FW-005', 'Firmware Release v1.0', addDays(baseDate, 55), addDays(baseDate, 55), '', 'Alex Rivera', 0, 'High', 'https://jira.example.com/FW-005', '', 'FW-004', 'Milestone', 'Firmware', ''],

    // Test Engineering Tasks
    ['TEST-001', 'Test Engineering Phase', addDays(baseDate, 25), addDays(baseDate, 65), '', 'Emily Taylor', 30, 'Medium', 'https://jira.example.com/TEST-001', '', '', 'Task', 'Test', ''],
    ['TEST-002', 'Test Plan Development', addDays(baseDate, 25), addDays(baseDate, 35), '', 'Emily Taylor', 100, 'Medium', 'https://jira.example.com/TEST-002', 'TEST-001', '', 'Task', 'Test', ''],
    ['TEST-003', 'DVT Execution', addDays(baseDate, 40), addDays(baseDate, 55), '', 'Emily Taylor', 25, 'High', 'https://jira.example.com/TEST-003', 'TEST-001', 'TEST-002,ELEC-004', 'Task', 'Test', ''],
    ['TEST-004', 'Certification Testing', addDays(baseDate, 55), addDays(baseDate, 65), '', 'Emily Taylor', 0, 'High', 'https://jira.example.com/TEST-004', 'TEST-001', 'TEST-003,FW-005', 'Task', 'Test', ''],
    ['TEST-005', 'Product Certification Complete', addDays(baseDate, 65), addDays(baseDate, 65), '', 'Emily Taylor', 0, 'High', 'https://jira.example.com/TEST-005', '', 'TEST-004', 'Milestone', 'Test', ''],

    // Program Management
    ['PM-001', 'Customer Review - Honda', addDays(baseDate, 30), addDays(baseDate, 30), '', 'Program Manager', 0, 'High', 'https://jira.example.com/PM-001', '', 'MECH-004,ELEC-003', 'Milestone', 'Program', ''],
    ['PM-002', 'Customer Review - Mazda', addDays(baseDate, 45), addDays(baseDate, 45), '', 'Program Manager', 0, 'High', 'https://jira.example.com/PM-002', '', 'FW-004', 'Milestone', 'Program', ''],
    ['PM-003', 'Production Release', addDays(baseDate, 70), addDays(baseDate, 70), '', 'Program Manager', 0, 'High', 'https://jira.example.com/PM-003', '', 'TEST-005', 'Milestone', 'Program', '']
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
// STORAGE FUNCTIONS (HIDDEN SHEET + PROPERTIES SERVICE FALLBACK)
// ============================================================================

// Hidden sheet name for storing configuration
const CONFIG_SHEET_NAME = '__VisualGanttConfig__';

/**
 * Gets or creates the hidden configuration sheet
 * @returns {GoogleAppsScript.Spreadsheet.Sheet|null} The config sheet or null if unavailable
 */
function getConfigSheet_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      Logger.log('No active spreadsheet available');
      return null;
    }

    let configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);

    if (!configSheet) {
      // Create the config sheet
      configSheet = ss.insertSheet(CONFIG_SHEET_NAME);

      // Set up headers
      configSheet.getRange('A1').setValue('__key__');
      configSheet.getRange('B1').setValue('__value__');

      // Hide the sheet immediately
      configSheet.hideSheet();

      Logger.log('Created and hid config sheet: ' + CONFIG_SHEET_NAME);
    }

    return configSheet;
  } catch (e) {
    Logger.log('Error getting/creating config sheet: ' + e.message);
    return null;
  }
}

/**
 * Reads a setting value from the hidden config sheet
 * @param {string} key - The setting key to read
 * @returns {string|null} The setting value or null if not found
 */
function readFromConfigSheet_(key) {
  try {
    const configSheet = getConfigSheet_();
    if (!configSheet) {
      return null;
    }

    // Get all data from the sheet
    const lastRow = configSheet.getLastRow();
    if (lastRow <= 1) {
      // Only header row exists
      return null;
    }

    const data = configSheet.getRange(2, 1, lastRow - 1, 2).getValues();

    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === key) {
        const value = data[i][1];
        return value ? String(value) : null;
      }
    }

    return null;
  } catch (e) {
    Logger.log('Error reading from config sheet: ' + e.message);
    return null;
  }
}

/**
 * Writes a setting value to the hidden config sheet
 * @param {string} key - The setting key
 * @param {string} value - The setting value
 * @returns {boolean} True if successfully written
 */
function writeToConfigSheet_(key, value) {
  try {
    const configSheet = getConfigSheet_();
    if (!configSheet) {
      return false;
    }

    // Find existing row with this key
    const lastRow = configSheet.getLastRow();
    let existingRow = -1;

    if (lastRow > 1) {
      const keys = configSheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < keys.length; i++) {
        if (keys[i][0] === key) {
          existingRow = i + 2; // +2 for 1-based index and header row
          break;
        }
      }
    }

    if (existingRow > 0) {
      // Update existing row
      configSheet.getRange(existingRow, 2).setValue(value);
    } else {
      // Add new row
      const newRow = lastRow + 1;
      configSheet.getRange(newRow, 1).setValue(key);
      configSheet.getRange(newRow, 2).setValue(value);
    }

    Logger.log('Saved setting to config sheet: ' + key);
    return true;
  } catch (e) {
    Logger.log('Error writing to config sheet: ' + e.message);
    return false;
  }
}

/**
 * Deletes a setting from the hidden config sheet
 * @param {string} key - The setting key to delete
 * @returns {boolean} True if successfully deleted
 */
function deleteFromConfigSheet_(key) {
  try {
    const configSheet = getConfigSheet_();
    if (!configSheet) {
      return false;
    }

    const lastRow = configSheet.getLastRow();
    if (lastRow <= 1) {
      return true; // Nothing to delete
    }

    const keys = configSheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < keys.length; i++) {
      if (keys[i][0] === key) {
        configSheet.deleteRow(i + 2); // +2 for 1-based index and header row
        Logger.log('Deleted setting from config sheet: ' + key);
        return true;
      }
    }

    return true; // Key not found, consider it deleted
  } catch (e) {
    Logger.log('Error deleting from config sheet: ' + e.message);
    return false;
  }
}

/**
 * Reads a setting from PropertiesService (fallback)
 * Attempts: DocumentProperties -> UserProperties -> ScriptProperties
 * @param {string} key - The setting key to read
 * @returns {string|null} The setting value or null if not found
 */
function readFromProperties_(key) {
  // Try DocumentProperties first
  try {
    const props = PropertiesService.getDocumentProperties();
    const value = props.getProperty(key);
    if (value) return value;
  } catch (e) {
    Logger.log('DocumentProperties read failed: ' + e.message);
  }

  // Try UserProperties
  try {
    const userProps = PropertiesService.getUserProperties();
    const value = userProps.getProperty(key);
    if (value) return value;
  } catch (e) {
    Logger.log('UserProperties read failed: ' + e.message);
  }

  // Try ScriptProperties
  try {
    const scriptProps = PropertiesService.getScriptProperties();
    const value = scriptProps.getProperty(key);
    return value || null;
  } catch (e) {
    Logger.log('ScriptProperties read failed: ' + e.message);
    return null;
  }
}

/**
 * Writes a setting to PropertiesService (fallback)
 * @param {string} key - The setting key
 * @param {string} value - The setting value
 * @returns {boolean} True if successfully written
 */
function writeToProperties_(key, value) {
  // Try DocumentProperties first
  try {
    const props = PropertiesService.getDocumentProperties();
    props.setProperty(key, value);
    Logger.log('Saved to DocumentProperties');
    return true;
  } catch (e) {
    Logger.log('DocumentProperties write failed: ' + e.message);
  }

  // Try UserProperties
  try {
    const userProps = PropertiesService.getUserProperties();
    userProps.setProperty(key, value);
    Logger.log('Saved to UserProperties');
    return true;
  } catch (e) {
    Logger.log('UserProperties write failed: ' + e.message);
  }

  // Try ScriptProperties
  try {
    const scriptProps = PropertiesService.getScriptProperties();
    scriptProps.setProperty(key, value);
    Logger.log('Saved to ScriptProperties');
    return true;
  } catch (e) {
    Logger.log('ScriptProperties write failed: ' + e.message);
  }

  return false;
}

/**
 * Deletes a setting from all PropertiesService stores
 * @param {string} key - The setting key to delete
 */
function deleteFromProperties_(key) {
  try {
    PropertiesService.getDocumentProperties().deleteProperty(key);
  } catch (e) {
    Logger.log('DocumentProperties delete failed: ' + e.message);
  }

  try {
    PropertiesService.getUserProperties().deleteProperty(key);
  } catch (e) {
    Logger.log('UserProperties delete failed: ' + e.message);
  }

  try {
    PropertiesService.getScriptProperties().deleteProperty(key);
  } catch (e) {
    Logger.log('ScriptProperties delete failed: ' + e.message);
  }
}

/**
 * Reads a setting value using multiple storage backends
 * Priority: Hidden Config Sheet -> PropertiesService
 * @param {string} key - The setting key to read
 * @returns {string|null} The setting value or null if not found
 */
function readSetting_(key) {
  // Try hidden config sheet first (works with spreadsheets.currentonly scope)
  const sheetValue = readFromConfigSheet_(key);
  if (sheetValue !== null) {
    Logger.log('Read setting from config sheet: ' + key);
    return sheetValue;
  }

  // Fall back to PropertiesService
  const propsValue = readFromProperties_(key);
  if (propsValue !== null) {
    Logger.log('Read setting from PropertiesService: ' + key);
    // Migrate to config sheet for future reads
    try {
      writeToConfigSheet_(key, propsValue);
      Logger.log('Migrated setting to config sheet: ' + key);
    } catch (e) {
      Logger.log('Migration to config sheet failed: ' + e.message);
    }
    return propsValue;
  }

  return null;
}

/**
 * Writes a setting value using multiple storage backends
 * Attempts: Hidden Config Sheet -> PropertiesService
 * @param {string} key - The setting key
 * @param {string} value - The setting value (JSON string)
 */
function writeSetting_(key, value) {
  let saved = false;

  // Try hidden config sheet first (works with spreadsheets.currentonly scope)
  if (writeToConfigSheet_(key, value)) {
    saved = true;
    Logger.log('Successfully saved setting to config sheet: ' + key);
  }

  // Also try to save to PropertiesService as backup (may fail in some environments)
  if (!saved) {
    if (writeToProperties_(key, value)) {
      saved = true;
      Logger.log('Successfully saved setting to PropertiesService: ' + key);
    }
  }

  if (!saved) {
    throw new Error(
      'Unable to save settings. This may be caused by permission restrictions.\n\n' +
      'Please try:\n' +
      '1. Go to Extensions > Visual Gantt > Authorization > Request Authorization\n' +
      '2. Follow the prompts to grant access\n' +
      '3. Try saving again\n\n' +
      'If the problem persists, your organization may have restrictions on add-on permissions.'
    );
  }
}

/**
 * Deletes a setting from all storage backends
 * @param {string} key - The setting key to delete
 */
function deleteSetting_(key) {
  // Delete from config sheet
  deleteFromConfigSheet_(key);

  // Delete from PropertiesService
  deleteFromProperties_(key);

  Logger.log('Deleted setting from all stores: ' + key);
}

/**
 * Migrates settings from PropertiesService to hidden config sheet
 * Call this to move existing settings to the new storage mechanism
 */
function migrateSettingsToConfigSheet() {
  const keysToMigrate = [CONFIG_KEY, JIRA_CONFIG_KEY, SMARTSHEET_CONFIG_KEY];
  let migrated = 0;

  for (const key of keysToMigrate) {
    const value = readFromProperties_(key);
    if (value !== null) {
      if (writeToConfigSheet_(key, value)) {
        Logger.log('Migrated setting: ' + key);
        migrated++;
      }
    }
  }

  Logger.log('Migration complete. Migrated ' + migrated + ' settings.');
  return { migrated: migrated, total: keysToMigrate.length };
}

/**
 * Diagnostic function to check storage capabilities
 * @returns {Object} Object with status of each storage method
 */
function diagnoseStorageCapabilities() {
  const testKey = '__visualGantt_storage_test__';
  const testValue = 'test_' + Date.now();
  const results = {
    configSheet: { read: false, write: false, delete: false, error: null },
    documentProperties: { read: false, write: false, delete: false, error: null },
    userProperties: { read: false, write: false, delete: false, error: null },
    scriptProperties: { read: false, write: false, delete: false, error: null }
  };

  // Test Config Sheet
  try {
    if (writeToConfigSheet_(testKey, testValue)) {
      results.configSheet.write = true;
      const readValue = readFromConfigSheet_(testKey);
      results.configSheet.read = (readValue === testValue);
      if (deleteFromConfigSheet_(testKey)) {
        results.configSheet.delete = true;
      }
    }
  } catch (e) {
    results.configSheet.error = e.message;
  }

  // Test DocumentProperties
  try {
    const props = PropertiesService.getDocumentProperties();
    props.setProperty(testKey, testValue);
    results.documentProperties.write = true;
    const readValue = props.getProperty(testKey);
    results.documentProperties.read = (readValue === testValue);
    props.deleteProperty(testKey);
    results.documentProperties.delete = true;
  } catch (e) {
    results.documentProperties.error = e.message;
  }

  // Test UserProperties
  try {
    const props = PropertiesService.getUserProperties();
    props.setProperty(testKey, testValue);
    results.userProperties.write = true;
    const readValue = props.getProperty(testKey);
    results.userProperties.read = (readValue === testValue);
    props.deleteProperty(testKey);
    results.userProperties.delete = true;
  } catch (e) {
    results.userProperties.error = e.message;
  }

  // Test ScriptProperties
  try {
    const props = PropertiesService.getScriptProperties();
    props.setProperty(testKey, testValue);
    results.scriptProperties.write = true;
    const readValue = props.getProperty(testKey);
    results.scriptProperties.read = (readValue === testValue);
    props.deleteProperty(testKey);
    results.scriptProperties.delete = true;
  } catch (e) {
    results.scriptProperties.error = e.message;
  }

  return results;
}

/**
 * Shows storage diagnostic results in a dialog
 */
function showStorageDiagnostics() {
  const results = diagnoseStorageCapabilities();

  let html = '<html><head><style>';
  html += 'body { font-family: Arial, sans-serif; padding: 20px; }';
  html += '.success { color: green; }';
  html += '.failure { color: red; }';
  html += 'table { border-collapse: collapse; width: 100%; }';
  html += 'th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }';
  html += 'th { background-color: #4285f4; color: white; }';
  html += '</style></head><body>';
  html += '<h2>Storage Diagnostics</h2>';
  html += '<table><tr><th>Storage Method</th><th>Write</th><th>Read</th><th>Delete</th><th>Error</th></tr>';

  for (const [method, status] of Object.entries(results)) {
    const formatMethod = method.replace(/([A-Z])/g, ' $1').trim();
    html += '<tr>';
    html += '<td>' + formatMethod + '</td>';
    html += '<td class="' + (status.write ? 'success' : 'failure') + '">' + (status.write ? '✓' : '✗') + '</td>';
    html += '<td class="' + (status.read ? 'success' : 'failure') + '">' + (status.read ? '✓' : '✗') + '</td>';
    html += '<td class="' + (status.delete ? 'success' : 'failure') + '">' + (status.delete ? '✓' : '✗') + '</td>';
    html += '<td>' + (status.error || '-') + '</td>';
    html += '</tr>';
  }

  html += '</table>';
  html += '<p style="margin-top: 20px;"><strong>Recommended:</strong> Config Sheet storage is the primary method and works with the spreadsheets.currentonly scope.</p>';
  html += '</body></html>';

  const output = HtmlService.createHtmlOutput(html)
    .setWidth(600)
    .setHeight(400)
    .setTitle('Storage Diagnostics');

  SpreadsheetApp.getUi().showModalDialog(output, 'Storage Diagnostics');
}

/**
 * Shows a dialog to migrate settings from PropertiesService to config sheet
 */
function showMigrationDialog() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Migrate Settings',
    'This will migrate any existing settings from PropertiesService to the hidden config sheet.\n\n' +
    'This is useful if you are experiencing PERMISSION_DENIED errors with PropertiesService.\n\n' +
    'Do you want to proceed?',
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.YES) {
    try {
      const result = migrateSettingsToConfigSheet();
      ui.alert(
        'Migration Complete',
        'Successfully migrated ' + result.migrated + ' of ' + result.total + ' settings to the config sheet.\n\n' +
        'Your settings should now work correctly.',
        ui.ButtonSet.OK
      );
    } catch (e) {
      ui.alert(
        'Migration Failed',
        'An error occurred during migration: ' + e.message,
        ui.ButtonSet.OK
      );
    }
  }
}


// ============================================================================
// CONFIGURATION
// ============================================================================

/**
 * Gets the current configuration from Document Properties
 */
function getConfig() {
  try {
    const savedConfig = readSetting_(CONFIG_KEY);
    if (savedConfig) {
      try {
        return { ...DEFAULT_CONFIG, ...JSON.parse(savedConfig) };
      } catch (e) {
        Logger.log('Invalid JSON in config: ' + e.message);
        return { ...DEFAULT_CONFIG };
      }
    }
  } catch (e) {
    Logger.log('Error reading config: ' + e.message);
  }

  return { ...DEFAULT_CONFIG };
}

/**
 * Saves configuration to Document Properties
 */
function saveConfig(config) {
  // Check file access before saving
  if (!hasFileAccess()) {
    throw new Error('Authorization required to save settings. Please use Visual Gantt > Authorization > Request Authorization.');
  }

  const configJson = JSON.stringify(config);

  try {
    writeSetting_(CONFIG_KEY, configJson);
  } catch (e) {
    Logger.log('Error saving config: ' + e.message);
    throw new Error('Unable to save settings: ' + e.message);
  }
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
  // Check authorization before proceeding
  if (!requireAuthorization('generate the timeline')) {
    return;
  }

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
    'Visual Gantt v1.1.0\n\n' +
    'A professional Gantt chart timeline add-on for Google Sheets.\n\n' +
    'Features:\n' +
    '- Professional timeline visualization\n' +
    '- Swimlane grouping by team/category\n' +
    '- Milestone markers\n' +
    '- Progress tracking\n' +
    '- Dependency arrows\n' +
    '- PNG/PDF export\n' +
    '- Jira Cloud integration\n' +
    '- Smartsheet integration\n\n' +
    'For support, please contact your administrator.',
    ui.ButtonSet.OK
  );
}

// ============================================================================
// JIRA INTEGRATION
// ============================================================================

/**
 * Gets Jira configuration from Document Properties
 */
function getJiraConfig() {
  try {
    const savedConfig = readSetting_(JIRA_CONFIG_KEY);
    if (savedConfig) {
      try {
        return { ...DEFAULT_JIRA_CONFIG, ...JSON.parse(savedConfig) };
      } catch (e) {
        Logger.log('Invalid JSON in Jira config: ' + e.message);
        return { ...DEFAULT_JIRA_CONFIG };
      }
    }
  } catch (e) {
    Logger.log('Error reading Jira config: ' + e.message);
  }

  return { ...DEFAULT_JIRA_CONFIG };
}

/**
 * Saves Jira configuration to Document Properties
 */
function saveJiraConfig(config) {
  // Check file access before saving
  if (!hasFileAccess()) {
    throw new Error('Authorization required to save Jira settings. Please use Visual Gantt > Authorization > Request Authorization.');
  }

  const configJson = JSON.stringify(config);

  try {
    writeSetting_(JIRA_CONFIG_KEY, configJson);
  } catch (e) {
    Logger.log('Error saving Jira config: ' + e.message);
    throw new Error('Unable to save Jira settings: ' + e.message);
  }
}

/**
 * Shows Jira configuration dialog
 */
function showJiraConfigDialog() {
  const html = HtmlService.createHtmlOutputFromFile('JiraConfigDialog')
    .setWidth(550)
    .setHeight(650)
    .setTitle('Jira Configuration');

  SpreadsheetApp.getUi().showModalDialog(html, 'Jira Configuration');
}

/**
 * Gets Jira config for dialog
 */
function getJiraConfigForDialog() {
  const config = getJiraConfig();
  // Don't send the full API token to client, just indicate if it's set
  return {
    ...config,
    apiToken: config.apiToken ? '••••••••' : ''
  };
}

/**
 * Saves Jira config from dialog
 */
function saveJiraConfigFromDialog(config) {
  const existingConfig = getJiraConfig();

  // If apiToken is the masked value, keep the existing token
  if (config.apiToken === '••••••••') {
    config.apiToken = existingConfig.apiToken;
  }

  saveJiraConfig(config);
  return { success: true };
}

/**
 * Tests Jira connection
 */
function testJiraConnection() {
  const config = getJiraConfig();

  if (!config.baseUrl || !config.email || !config.apiToken) {
    return { success: false, message: 'Please configure Jira URL, email, and API token first.' };
  }

  try {
    const response = jiraApiRequest('/rest/api/3/myself');
    return {
      success: true,
      message: `Connected as ${response.displayName} (${response.emailAddress})`
    };
  } catch (e) {
    return { success: false, message: 'Connection failed: ' + e.message };
  }
}

/**
 * Makes an authenticated request to Jira API
 */
function jiraApiRequest(endpoint, method, payload) {
  const config = getJiraConfig();

  if (!config.baseUrl || !config.email || !config.apiToken) {
    throw new Error('Jira not configured. Please configure connection first.');
  }

  const baseUrl = config.baseUrl.replace(/\/$/, '');
  const url = baseUrl.startsWith('http')
    ? `${baseUrl}${endpoint}`
    : `https://${baseUrl}${endpoint}`;

  const options = {
    method: method || 'GET',
    headers: {
      'Authorization': 'Basic ' + Utilities.base64Encode(config.email + ':' + config.apiToken),
      'Content-Type': 'application/json',
      'Accept': 'application/json'
    },
    muteHttpExceptions: true
  };

  if (payload) {
    options.payload = JSON.stringify(payload);
  }

  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();

  if (responseCode >= 400) {
    let errorMessage = `HTTP ${responseCode}`;
    try {
      const errorJson = JSON.parse(responseText);
      errorMessage = errorJson.errorMessages?.join(', ') || errorJson.message || errorMessage;
    } catch (e) {
      errorMessage = responseText || errorMessage;
    }
    throw new Error(errorMessage);
  }

  return responseText ? JSON.parse(responseText) : null;
}

/**
 * Fetches issues from Jira based on configured method
 */
function fetchJiraIssues() {
  const config = getJiraConfig();
  let jql = '';

  switch (config.pullMethod) {
    case 'jql':
      jql = config.jqlQuery;
      break;
    case 'project':
      jql = `project = "${config.projectKey}" ORDER BY key ASC`;
      break;
    case 'filter':
      // Get filter JQL
      const filter = jiraApiRequest(`/rest/api/3/filter/${config.filterId}`);
      jql = filter.jql;
      break;
    default:
      throw new Error('Invalid pull method configured');
  }

  if (!jql) {
    throw new Error('No JQL query, project key, or filter ID configured');
  }

  // Fetch issues with pagination
  const allIssues = [];
  let startAt = 0;
  const maxResults = 100;

  // Fields to fetch
  const fields = [
    'key', 'summary', 'status', 'priority', 'assignee', 'duedate',
    'issuetype', 'parent', 'issuelinks', 'labels', 'components',
    'timeoriginalestimate', 'timespent', 'timeestimate'
  ];

  // Add custom fields if configured
  if (config.startDateField) {
    fields.push(config.startDateField);
  }

  do {
    const response = jiraApiRequest(
      `/rest/api/3/search?jql=${encodeURIComponent(jql)}&startAt=${startAt}&maxResults=${maxResults}&fields=${fields.join(',')}&expand=names`
    );

    allIssues.push(...response.issues);
    startAt += response.issues.length;

    if (startAt >= response.total) {
      break;
    }
  } while (true);

  return allIssues;
}

/**
 * Converts a Jira issue to a sheet row format
 */
function jiraIssueToRow(issue, config) {
  const fields = issue.fields;

  // Extract key as Task ID
  const taskId = issue.key;

  // Extract name
  const taskName = fields.summary || '';

  // Extract dates
  const startDate = config.startDateField && fields[config.startDateField]
    ? new Date(fields[config.startDateField])
    : null;
  const endDate = fields.duedate ? new Date(fields.duedate) : null;

  // Calculate duration
  let duration = '';
  if (startDate && endDate) {
    duration = Math.ceil((endDate - startDate) / (1000 * 60 * 60 * 24)) + 1;
  }

  // Extract owner (assignee)
  const owner = fields.assignee?.displayName || '';

  // Map status to percent complete
  const statusCategory = fields.status?.statusCategory?.key || 'new';
  const percentComplete = JIRA_STATUS_MAPPING[statusCategory] ?? 0;

  // Map priority
  const jiraPriority = fields.priority?.name || 'Medium';
  const priority = JIRA_PRIORITY_MAPPING[jiraPriority] || 'Medium';

  // Build Jira ticket URL
  const baseUrl = getJiraConfig().baseUrl.replace(/\/$/, '');
  const jiraUrl = baseUrl.startsWith('http')
    ? `${baseUrl}/browse/${taskId}`
    : `https://${baseUrl}/browse/${taskId}`;

  // Extract parent task
  const parentTask = fields.parent?.key || '';

  // Extract dependencies from issue links
  const dependencies = [];
  if (fields.issuelinks) {
    fields.issuelinks.forEach(link => {
      // "is blocked by" or "depends on" links
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

  // Extract swimlane (use first component or first label)
  let swimlane = 'Default';
  if (fields.components && fields.components.length > 0) {
    swimlane = fields.components[0].name;
  } else if (fields.labels && fields.labels.length > 0) {
    swimlane = fields.labels[0];
  }

  return {
    taskId,
    taskName,
    startDate,
    endDate,
    duration,
    owner,
    percentComplete,
    priority,
    jiraUrl,
    parentTask,
    dependencies: dependencies.join(','),
    taskType,
    swimlane,
    modified: '',  // Not modified when pulled from Jira
    // Store original Jira data for conflict detection
    _jiraStatus: fields.status?.name,
    _jiraStatusCategory: statusCategory
  };
}

/**
 * Gets current sheet data as a map keyed by Task ID
 */
function getSheetDataMap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(DATA_SHEET_NAME);

  if (!dataSheet) {
    return new Map();
  }

  const lastRow = dataSheet.getLastRow();
  if (lastRow < 2) {
    return new Map();
  }

  const dataRange = dataSheet.getRange(2, 1, lastRow - 1, COLUMN_HEADERS.length);
  const data = dataRange.getValues();

  const map = new Map();
  data.forEach((row, index) => {
    const taskId = row[COL.TASK_ID];
    if (taskId) {
      map.set(String(taskId).trim(), {
        rowIndex: index + 2, // 1-based row number
        data: row
      });
    }
  });

  return map;
}

/**
 * Detects conflicts between Jira data and sheet data
 */
function detectConflicts(jiraRows, sheetDataMap) {
  const conflicts = [];

  jiraRows.forEach(jiraRow => {
    const sheetEntry = sheetDataMap.get(jiraRow.taskId);

    if (sheetEntry) {
      const sheetRow = sheetEntry.data;
      const isModified = String(sheetRow[COL.MODIFIED]).toLowerCase() === 'yes';

      if (isModified) {
        // Check for actual differences
        const diffs = [];

        // Compare dates
        const sheetStartDate = parseDate(sheetRow[COL.START_DATE]);
        const sheetEndDate = parseDate(sheetRow[COL.END_DATE]);

        if (jiraRow.startDate && sheetStartDate &&
            jiraRow.startDate.getTime() !== sheetStartDate.getTime()) {
          diffs.push({
            field: 'Start Date',
            jiraValue: formatDateForDisplay(jiraRow.startDate),
            sheetValue: formatDateForDisplay(sheetStartDate)
          });
        }

        if (jiraRow.endDate && sheetEndDate &&
            jiraRow.endDate.getTime() !== sheetEndDate.getTime()) {
          diffs.push({
            field: 'End Date',
            jiraValue: formatDateForDisplay(jiraRow.endDate),
            sheetValue: formatDateForDisplay(sheetEndDate)
          });
        }

        // Compare owner
        if (jiraRow.owner !== String(sheetRow[COL.OWNER]).trim()) {
          diffs.push({
            field: 'Owner',
            jiraValue: jiraRow.owner || '(none)',
            sheetValue: String(sheetRow[COL.OWNER]).trim() || '(none)'
          });
        }

        // Compare priority
        if (jiraRow.priority !== String(sheetRow[COL.PRIORITY]).trim()) {
          diffs.push({
            field: 'Priority',
            jiraValue: jiraRow.priority,
            sheetValue: String(sheetRow[COL.PRIORITY]).trim()
          });
        }

        if (diffs.length > 0) {
          conflicts.push({
            taskId: jiraRow.taskId,
            taskName: jiraRow.taskName,
            rowIndex: sheetEntry.rowIndex,
            diffs: diffs,
            jiraData: jiraRow,
            sheetData: sheetRow
          });
        }
      }
    }
  });

  return conflicts;
}

/**
 * Formats a date for display
 */
function formatDateForDisplay(date) {
  if (!date) return '';
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/**
 * Main pull function - fetches from Jira and updates sheet
 */
function pullFromJira() {
  // Check authorization before proceeding
  if (!requireAuthorization('pull from Jira')) {
    return;
  }

  const ui = SpreadsheetApp.getUi();

  try {
    // Fetch issues from Jira
    ui.alert('Pulling from Jira', 'Fetching issues from Jira...', ui.ButtonSet.OK);

    const issues = fetchJiraIssues();

    if (issues.length === 0) {
      ui.alert('No Issues', 'No issues found matching your query.', ui.ButtonSet.OK);
      return;
    }

    // Convert to row format
    const config = getJiraConfig();
    const jiraRows = issues.map(issue => jiraIssueToRow(issue, config));

    // Get current sheet data
    const sheetDataMap = getSheetDataMap();

    // Detect conflicts
    const conflicts = detectConflicts(jiraRows, sheetDataMap);

    if (conflicts.length > 0) {
      // Show conflict resolution dialog
      showConflictDialog(jiraRows, conflicts);
    } else {
      // No conflicts, apply updates directly
      applyJiraPull(jiraRows, []);
      ui.alert('Pull Complete', `Successfully pulled ${issues.length} issues from Jira.`, ui.ButtonSet.OK);
    }

    // Update last sync time
    config.lastSyncTime = new Date().toISOString();
    saveJiraConfig(config);

  } catch (e) {
    ui.alert('Error', 'Failed to pull from Jira: ' + e.message, ui.ButtonSet.OK);
  }
}

/**
 * Shows conflict resolution dialog
 */
function showConflictDialog(jiraRows, conflicts) {
  const template = HtmlService.createTemplateFromFile('ConflictDialog');
  template.conflicts = JSON.stringify(conflicts);
  template.jiraRowsJson = JSON.stringify(jiraRows);

  const html = template.evaluate()
    .setWidth(700)
    .setHeight(600)
    .setTitle('Resolve Conflicts');

  SpreadsheetApp.getUi().showModalDialog(html, 'Resolve Conflicts');
}

/**
 * Applies Jira pull with conflict resolutions
 * @param {Array} jiraRows - All rows from Jira
 * @param {Array} keepLocalIds - Task IDs where local changes should be kept
 */
function applyJiraPull(jiraRows, keepLocalIds) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let dataSheet = ss.getSheetByName(DATA_SHEET_NAME);

  if (!dataSheet) {
    setupDataSheet();
    dataSheet = ss.getSheetByName(DATA_SHEET_NAME);
  }

  const keepLocalSet = new Set(keepLocalIds);
  const sheetDataMap = getSheetDataMap();

  // Prepare updates and inserts
  const updates = [];
  const inserts = [];

  jiraRows.forEach(jiraRow => {
    const sheetEntry = sheetDataMap.get(jiraRow.taskId);

    if (sheetEntry) {
      // Update existing row (unless keeping local)
      if (!keepLocalSet.has(jiraRow.taskId)) {
        updates.push({
          rowIndex: sheetEntry.rowIndex,
          data: jiraRowToSheetRow(jiraRow)
        });
      }
    } else {
      // Insert new row
      inserts.push(jiraRowToSheetRow(jiraRow));
    }
  });

  // Apply updates
  updates.forEach(update => {
    dataSheet.getRange(update.rowIndex, 1, 1, COLUMN_HEADERS.length)
      .setValues([update.data]);
  });

  // Apply inserts
  if (inserts.length > 0) {
    const lastRow = dataSheet.getLastRow();
    dataSheet.getRange(lastRow + 1, 1, inserts.length, COLUMN_HEADERS.length)
      .setValues(inserts);
  }

  return {
    success: true,
    updated: updates.length,
    inserted: inserts.length
  };
}

/**
 * Converts a jiraRow object to sheet row array
 */
function jiraRowToSheetRow(jiraRow) {
  return [
    jiraRow.taskId,
    jiraRow.taskName,
    jiraRow.startDate,
    jiraRow.endDate,
    jiraRow.duration,
    jiraRow.owner,
    jiraRow.percentComplete,
    jiraRow.priority,
    jiraRow.jiraUrl,
    jiraRow.parentTask,
    jiraRow.dependencies,
    jiraRow.taskType,
    jiraRow.swimlane,
    ''  // Modified = empty (fresh from Jira)
  ];
}

/**
 * Pulls only selected rows from Jira
 */
function pullSelectedFromJira() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  if (sheet.getName() !== DATA_SHEET_NAME) {
    ui.alert('Wrong Sheet', `Please select rows in the "${DATA_SHEET_NAME}" sheet.`, ui.ButtonSet.OK);
    return;
  }

  const selection = sheet.getActiveRange();
  const startRow = selection.getRow();
  const numRows = selection.getNumRows();

  if (startRow < 2) {
    ui.alert('Invalid Selection', 'Please select data rows (not the header).', ui.ButtonSet.OK);
    return;
  }

  // Get task IDs from selected rows
  const taskIds = [];
  for (let i = 0; i < numRows; i++) {
    const taskId = sheet.getRange(startRow + i, 1).getValue();
    if (taskId) {
      taskIds.push(String(taskId).trim());
    }
  }

  if (taskIds.length === 0) {
    ui.alert('No Tasks', 'No task IDs found in selection.', ui.ButtonSet.OK);
    return;
  }

  try {
    // Build JQL for specific issues
    const jql = `key in (${taskIds.join(',')})`;
    const config = getJiraConfig();

    // Fetch specific issues
    const response = jiraApiRequest(
      `/rest/api/3/search?jql=${encodeURIComponent(jql)}&maxResults=${taskIds.length}`
    );

    if (response.issues.length === 0) {
      ui.alert('No Issues', 'Selected issues not found in Jira.', ui.ButtonSet.OK);
      return;
    }

    const jiraRows = response.issues.map(issue => jiraIssueToRow(issue, config));
    const sheetDataMap = getSheetDataMap();
    const conflicts = detectConflicts(jiraRows, sheetDataMap);

    if (conflicts.length > 0) {
      showConflictDialog(jiraRows, conflicts);
    } else {
      applyJiraPull(jiraRows, []);
      ui.alert('Pull Complete', `Successfully pulled ${response.issues.length} issues.`, ui.ButtonSet.OK);
    }

  } catch (e) {
    ui.alert('Error', 'Failed to pull from Jira: ' + e.message, ui.ButtonSet.OK);
  }
}

/**
 * Pushes changes to Jira
 */
function pushToJira() {
  // Check authorization before proceeding
  if (!requireAuthorization('push to Jira')) {
    return;
  }

  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(DATA_SHEET_NAME);

  if (!dataSheet) {
    ui.alert('No Data', 'Data sheet not found.', ui.ButtonSet.OK);
    return;
  }

  // Find all modified rows
  const lastRow = dataSheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('No Data', 'No tasks to push.', ui.ButtonSet.OK);
    return;
  }

  const data = dataSheet.getRange(2, 1, lastRow - 1, COLUMN_HEADERS.length).getValues();
  const modifiedRows = [];

  data.forEach((row, index) => {
    if (String(row[COL.MODIFIED]).toLowerCase() === 'yes') {
      modifiedRows.push({
        rowIndex: index + 2,
        data: row
      });
    }
  });

  if (modifiedRows.length === 0) {
    ui.alert('No Changes', 'No modified rows to push. Mark rows as "Yes" in the Modified column to push changes.', ui.ButtonSet.OK);
    return;
  }

  const response = ui.alert(
    'Push to Jira',
    `Push ${modifiedRows.length} modified row(s) to Jira?\n\nThis will update: Due Date, Start Date (if configured), Assignee, Priority, and Status.`,
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  const results = pushRowsToJira(modifiedRows);

  // Clear modified flag for successful pushes
  results.successful.forEach(taskId => {
    const sheetDataMap = getSheetDataMap();
    const entry = sheetDataMap.get(taskId);
    if (entry) {
      dataSheet.getRange(entry.rowIndex, COL.MODIFIED + 1).setValue('');
    }
  });

  let message = `Successfully pushed ${results.successful.length} task(s).`;
  if (results.failed.length > 0) {
    message += `\n\nFailed to push ${results.failed.length} task(s):\n${results.failed.map(f => `${f.taskId}: ${f.error}`).join('\n')}`;
  }

  ui.alert('Push Complete', message, ui.ButtonSet.OK);
}

/**
 * Pushes specific rows to Jira
 */
function pushRowsToJira(rows) {
  const config = getJiraConfig();
  const successful = [];
  const failed = [];

  rows.forEach(row => {
    const taskId = String(row.data[COL.TASK_ID]).trim();

    try {
      const updatePayload = {
        fields: {}
      };

      // Due date (End Date)
      const endDate = parseDate(row.data[COL.END_DATE]);
      if (endDate) {
        updatePayload.fields.duedate = formatDateForJira(endDate);
      }

      // Start date (custom field)
      if (config.startDateField) {
        const startDate = parseDate(row.data[COL.START_DATE]);
        if (startDate) {
          updatePayload.fields[config.startDateField] = formatDateForJira(startDate);
        }
      }

      // Priority
      const priority = String(row.data[COL.PRIORITY]).trim();
      if (priority) {
        const jiraPriority = Object.keys(JIRA_PRIORITY_MAPPING).find(
          k => JIRA_PRIORITY_MAPPING[k] === priority
        ) || priority;
        updatePayload.fields.priority = { name: jiraPriority };
      }

      // Update the issue
      jiraApiRequest(`/rest/api/3/issue/${taskId}`, 'PUT', updatePayload);

      // Handle assignee separately (different API)
      const owner = String(row.data[COL.OWNER]).trim();
      if (owner) {
        try {
          // Search for user by display name
          const users = jiraApiRequest(`/rest/api/3/user/search?query=${encodeURIComponent(owner)}`);
          if (users && users.length > 0) {
            jiraApiRequest(`/rest/api/3/issue/${taskId}/assignee`, 'PUT', { accountId: users[0].accountId });
          }
        } catch (assigneeError) {
          // Assignee update failed, but main update succeeded
          console.log(`Failed to update assignee for ${taskId}: ${assigneeError.message}`);
        }
      }

      // Handle status transition
      const percentComplete = Number(row.data[COL.PERCENT_COMPLETE]) || 0;
      try {
        updateIssueStatus(taskId, percentComplete);
      } catch (statusError) {
        console.log(`Failed to update status for ${taskId}: ${statusError.message}`);
      }

      successful.push(taskId);

    } catch (e) {
      failed.push({ taskId, error: e.message });
    }
  });

  return { successful, failed };
}

/**
 * Updates issue status based on percent complete
 */
function updateIssueStatus(taskId, percentComplete) {
  // Get available transitions
  const transitionsResponse = jiraApiRequest(`/rest/api/3/issue/${taskId}/transitions`);
  const transitions = transitionsResponse.transitions || [];

  // Determine target status category
  let targetCategory;
  if (percentComplete === 0) {
    targetCategory = 'new';
  } else if (percentComplete === 100) {
    targetCategory = 'done';
  } else {
    targetCategory = 'indeterminate';
  }

  // Find a transition that leads to the target category
  const transition = transitions.find(t =>
    t.to?.statusCategory?.key === targetCategory
  );

  if (transition) {
    jiraApiRequest(`/rest/api/3/issue/${taskId}/transitions`, 'POST', {
      transition: { id: transition.id }
    });
  }
}

/**
 * Formats a date for Jira API
 */
function formatDateForJira(date) {
  if (!date) return null;
  return Utilities.formatDate(date, 'UTC', 'yyyy-MM-dd');
}

/**
 * Pushes only selected rows to Jira
 */
function pushSelectedToJira() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  if (sheet.getName() !== DATA_SHEET_NAME) {
    ui.alert('Wrong Sheet', `Please select rows in the "${DATA_SHEET_NAME}" sheet.`, ui.ButtonSet.OK);
    return;
  }

  const selection = sheet.getActiveRange();
  const startRow = selection.getRow();
  const numRows = selection.getNumRows();

  if (startRow < 2) {
    ui.alert('Invalid Selection', 'Please select data rows (not the header).', ui.ButtonSet.OK);
    return;
  }

  // Get selected rows
  const selectedRows = [];
  for (let i = 0; i < numRows; i++) {
    const rowData = sheet.getRange(startRow + i, 1, 1, COLUMN_HEADERS.length).getValues()[0];
    if (rowData[COL.TASK_ID]) {
      selectedRows.push({
        rowIndex: startRow + i,
        data: rowData
      });
    }
  }

  if (selectedRows.length === 0) {
    ui.alert('No Tasks', 'No tasks found in selection.', ui.ButtonSet.OK);
    return;
  }

  const response = ui.alert(
    'Push to Jira',
    `Push ${selectedRows.length} selected row(s) to Jira?`,
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  const results = pushRowsToJira(selectedRows);

  // Clear modified flag for successful pushes
  const dataSheet = ss.getSheetByName(DATA_SHEET_NAME);
  results.successful.forEach(taskId => {
    selectedRows.forEach(row => {
      if (String(row.data[COL.TASK_ID]).trim() === taskId) {
        dataSheet.getRange(row.rowIndex, COL.MODIFIED + 1).setValue('');
      }
    });
  });

  let message = `Successfully pushed ${results.successful.length} task(s).`;
  if (results.failed.length > 0) {
    message += `\n\nFailed: ${results.failed.map(f => `${f.taskId}: ${f.error}`).join('\n')}`;
  }

  ui.alert('Push Complete', message, ui.ButtonSet.OK);
}

/**
 * Shows sync status dialog
 */
function showSyncStatus() {
  const config = getJiraConfig();
  const sheetDataMap = getSheetDataMap();

  let modifiedCount = 0;
  sheetDataMap.forEach(entry => {
    if (String(entry.data[COL.MODIFIED]).toLowerCase() === 'yes') {
      modifiedCount++;
    }
  });

  const lastSync = config.lastSyncTime
    ? new Date(config.lastSyncTime).toLocaleString()
    : 'Never';

  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Jira Sync Status',
    `Connection: ${config.baseUrl || 'Not configured'}\n\n` +
    `Last sync: ${lastSync}\n` +
    `Total tasks in sheet: ${sheetDataMap.size}\n` +
    `Modified (pending push): ${modifiedCount}`,
    ui.ButtonSet.OK
  );
}

/**
 * Marks a row as modified when edited
 * Can be set up as an onEdit trigger
 */
function markRowAsModified(e) {
  const sheet = e.source.getActiveSheet();

  if (sheet.getName() !== DATA_SHEET_NAME) {
    return;
  }

  const range = e.range;
  const row = range.getRow();

  // Skip header row
  if (row < 2) {
    return;
  }

  // Skip if editing the Modified column itself
  const col = range.getColumn();
  if (col === COL.MODIFIED + 1) {
    return;
  }

  // Check if this row has a Jira ticket (indicating it came from Jira)
  const jiraTicket = sheet.getRange(row, COL.JIRA_TICKET + 1).getValue();
  if (jiraTicket) {
    // Mark as modified
    sheet.getRange(row, COL.MODIFIED + 1).setValue('Yes');
  }
}

/**
 * Include HTML file content (for templates)
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================================================
// SMARTSHEET INTEGRATION
// ============================================================================

/**
 * Gets Smartsheet configuration from Document Properties
 */
function getSmartsheetConfig() {
  // Helper function to parse and merge config
  function parseConfig(savedConfig) {
    const parsed = JSON.parse(savedConfig);
    return {
      ...DEFAULT_SMARTSHEET_CONFIG,
      ...parsed,
      columnMapping: { ...DEFAULT_SMARTSHEET_CONFIG.columnMapping, ...(parsed.columnMapping || {}) }
    };
  }

  try {
    const savedConfig = readSetting_(SMARTSHEET_CONFIG_KEY);
    if (savedConfig) {
      try {
        return parseConfig(savedConfig);
      } catch (e) {
        Logger.log('Invalid JSON in Smartsheet config: ' + e.message);
        return { ...DEFAULT_SMARTSHEET_CONFIG };
      }
    }
  } catch (e) {
    Logger.log('Error reading Smartsheet config: ' + e.message);
  }

  return { ...DEFAULT_SMARTSHEET_CONFIG };
}

/**
 * Saves Smartsheet configuration to Document Properties
 */
function saveSmartsheetConfig(config) {
  // Check file access before saving
  if (!hasFileAccess()) {
    throw new Error('Authorization required to save Smartsheet settings. Please use Visual Gantt > Authorization > Request Authorization.');
  }

  const configJson = JSON.stringify(config);

  try {
    writeSetting_(SMARTSHEET_CONFIG_KEY, configJson);
  } catch (e) {
    Logger.log('Error saving Smartsheet config: ' + e.message);
    throw new Error('Unable to save Smartsheet settings: ' + e.message);
  }
}

/**
 * Shows Smartsheet configuration dialog
 */
function showSmartsheetConfigDialog() {
  const html = HtmlService.createHtmlOutputFromFile('SmartsheetConfigDialog')
    .setWidth(600)
    .setHeight(750)
    .setTitle('Smartsheet Configuration');

  SpreadsheetApp.getUi().showModalDialog(html, 'Smartsheet Configuration');
}

/**
 * Gets Smartsheet config for dialog
 */
function getSmartsheetConfigForDialog() {
  const config = getSmartsheetConfig();
  // Don't send the full API token to client, just indicate if it's set
  return {
    ...config,
    apiToken: config.apiToken ? '••••••••' : ''
  };
}

/**
 * Saves Smartsheet config from dialog
 */
function saveSmartsheetConfigFromDialog(config) {
  const existingConfig = getSmartsheetConfig();

  // If apiToken is the masked value, keep the existing token
  if (config.apiToken === '••••••••') {
    config.apiToken = existingConfig.apiToken;
  }

  saveSmartsheetConfig(config);
  return { success: true };
}

/**
 * Makes an authenticated request to Smartsheet API
 */
function smartsheetApiRequest(endpoint, method, payload, token) {
  const config = getSmartsheetConfig();
  const apiToken = token || config.apiToken;

  if (!apiToken) {
    throw new Error('Smartsheet not configured. Please configure connection first.');
  }

  const url = `${SMARTSHEET_API_BASE}${endpoint}`;

  const options = {
    method: method || 'GET',
    headers: {
      'Authorization': 'Bearer ' + apiToken,
      'Content-Type': 'application/json',
      'Accept': 'application/json'
    },
    muteHttpExceptions: true
  };

  if (payload) {
    options.payload = JSON.stringify(payload);
  }

  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();

  if (responseCode >= 400) {
    let errorMessage = `HTTP ${responseCode}`;
    try {
      const errorJson = JSON.parse(responseText);
      errorMessage = errorJson.message || errorJson.errorCode || errorMessage;
    } catch (e) {
      errorMessage = responseText || errorMessage;
    }
    throw new Error(errorMessage);
  }

  return responseText ? JSON.parse(responseText) : null;
}

/**
 * Tests Smartsheet connection
 */
function testSmartsheetConnection() {
  const config = getSmartsheetConfig();
  return testSmartsheetConnectionWithToken(config.apiToken);
}

/**
 * Tests Smartsheet connection with a specific token
 */
function testSmartsheetConnectionWithToken(token) {
  if (!token || token === '••••••••') {
    return { success: false, message: 'Please enter an API token.' };
  }

  try {
    const response = smartsheetApiRequest('/users/me', 'GET', null, token);
    return {
      success: true,
      message: `Connected as ${response.firstName} ${response.lastName} (${response.email})`
    };
  } catch (e) {
    return { success: false, message: 'Connection failed: ' + e.message };
  }
}

/**
 * Loads list of available Smartsheet sheets
 */
function loadSmartsheetList() {
  const config = getSmartsheetConfig();
  return loadSmartsheetListWithToken(config.apiToken);
}

/**
 * Loads list of available Smartsheet sheets with a specific token
 */
function loadSmartsheetListWithToken(token) {
  if (!token || token === '••••••••') {
    return { success: false, message: 'Please enter an API token.' };
  }

  try {
    const response = smartsheetApiRequest('/sheets?includeAll=true', 'GET', null, token);
    const sheets = response.data.map(sheet => ({
      id: sheet.id,
      name: sheet.name,
      accessLevel: sheet.accessLevel
    }));

    // Sort by name
    sheets.sort((a, b) => a.name.localeCompare(b.name));

    return { success: true, sheets: sheets };
  } catch (e) {
    return { success: false, message: 'Failed to load sheets: ' + e.message };
  }
}

/**
 * Loads columns for a specific Smartsheet
 */
function loadSmartsheetColumns(sheetId) {
  const config = getSmartsheetConfig();
  return loadSmartsheetColumnsWithToken(config.apiToken, sheetId);
}

/**
 * Loads columns for a specific Smartsheet with a specific token
 */
function loadSmartsheetColumnsWithToken(token, sheetId) {
  if (!token || token === '••••••••') {
    return { success: false, message: 'Please enter an API token.' };
  }

  if (!sheetId) {
    return { success: false, message: 'Please select a sheet.' };
  }

  try {
    const response = smartsheetApiRequest(`/sheets/${sheetId}?include=columns`, 'GET', null, token);
    const columns = response.columns.map(col => ({
      id: col.id,
      title: col.title,
      type: col.type,
      primary: col.primary || false,
      index: col.index
    }));

    return { success: true, columns: columns };
  } catch (e) {
    return { success: false, message: 'Failed to load columns: ' + e.message };
  }
}

/**
 * Fetches rows from Smartsheet
 */
function fetchSmartsheetRows() {
  const config = getSmartsheetConfig();

  if (!config.sheetId) {
    throw new Error('No Smartsheet selected. Please configure connection first.');
  }

  // Fetch the sheet with all data
  const response = smartsheetApiRequest(
    `/sheets/${config.sheetId}?include=attachments,discussions,rowPermalink&level=2`
  );

  return {
    sheet: response,
    columns: response.columns,
    rows: response.rows
  };
}

/**
 * Converts a Smartsheet row to a sheet row format
 */
function smartsheetRowToLocalRow(row, columns, config) {
  const columnMap = {};
  columns.forEach(col => {
    columnMap[col.id] = col;
  });

  // Get cell values by column ID
  const cellMap = {};
  row.cells.forEach(cell => {
    cellMap[cell.columnId] = cell;
  });

  // Helper to get cell value
  function getCellValue(mappingKey) {
    const colId = config.columnMapping[mappingKey];
    if (!colId) return null;

    const cell = cellMap[colId];
    if (!cell) return null;

    // Handle different value types
    if (cell.displayValue !== undefined) {
      return cell.displayValue;
    }
    if (cell.value !== undefined) {
      return cell.value;
    }
    if (cell.objectValue) {
      // Handle contact lists
      if (cell.objectValue.objectType === 'MULTI_CONTACT') {
        return cell.objectValue.values.map(v => v.name || v.email).join(', ');
      }
      if (cell.objectValue.name) {
        return cell.objectValue.name;
      }
    }
    return null;
  }

  // Helper to parse date
  function parseSmartsheetDate(value) {
    if (!value) return null;
    const parsed = new Date(value);
    return isNaN(parsed.getTime()) ? null : parsed;
  }

  // Extract values using column mapping
  const taskId = getCellValue('colTaskId') || `SS-${row.id}`;
  const taskName = getCellValue('colTaskName') || '';
  const startDateRaw = getCellValue('colStartDate');
  const endDateRaw = getCellValue('colEndDate');
  const durationRaw = getCellValue('colDuration');
  const owner = getCellValue('colOwner') || '';
  const percentCompleteRaw = getCellValue('colPercentComplete');
  const priorityRaw = getCellValue('colPriority');
  const predecessorsRaw = getCellValue('colPredecessors');
  const category = getCellValue('colCategory') || 'Default';

  // Parse dates
  const startDate = parseSmartsheetDate(startDateRaw);
  const endDate = parseSmartsheetDate(endDateRaw);

  // Parse duration
  let duration = null;
  if (durationRaw) {
    // Duration might be in format "5d" or just "5"
    const durationMatch = String(durationRaw).match(/(\d+)/);
    if (durationMatch) {
      duration = parseInt(durationMatch[1], 10);
    }
  }

  // Parse percent complete
  let percentComplete = 0;
  if (percentCompleteRaw !== null && percentCompleteRaw !== undefined) {
    if (typeof percentCompleteRaw === 'number') {
      // Might be 0-1 or 0-100
      percentComplete = percentCompleteRaw > 1 ? percentCompleteRaw : percentCompleteRaw * 100;
    } else {
      const percentMatch = String(percentCompleteRaw).match(/(\d+)/);
      if (percentMatch) {
        percentComplete = parseInt(percentMatch[1], 10);
      }
    }
  }

  // Map priority
  const priority = SMARTSHEET_PRIORITY_MAPPING[priorityRaw] || 'Medium';

  // Parse predecessors
  let dependencies = [];
  if (predecessorsRaw && config.syncDependencies) {
    // Predecessors in Smartsheet are typically in format "1, 2FS+1d" or just row numbers
    // We'll convert these to task IDs later
    const predParts = String(predecessorsRaw).split(',');
    predParts.forEach(pred => {
      const trimmed = pred.trim();
      // Extract just the row number (before any FS, SS, FF, SF and lag)
      const match = trimmed.match(/^(\d+)/);
      if (match) {
        dependencies.push(`SS-ROW-${match[1]}`); // Temporary placeholder
      }
    });
  }

  // Build Smartsheet URL
  const smartsheetUrl = row.permalink || `https://app.smartsheet.com/sheets/${config.sheetId}?rowId=${row.id}`;

  // Determine parent task (for hierarchical rows)
  const parentTask = row.parentId ? `SS-${row.parentId}` : '';

  // Determine task type (based on duration or if it has no children)
  let taskType = 'Task';
  if (duration === 0 || (startDate && endDate && startDate.getTime() === endDate.getTime())) {
    taskType = 'Milestone';
  }

  return {
    taskId: String(taskId).trim(),
    taskName: String(taskName).trim(),
    startDate,
    endDate,
    duration,
    owner: String(owner).trim(),
    percentComplete: Math.round(percentComplete),
    priority,
    smartsheetUrl,
    parentTask,
    dependencies: dependencies.join(','),
    taskType,
    swimlane: String(category).trim(),
    modified: '',
    // Store original Smartsheet data for sync
    _smartsheetRowId: row.id,
    _smartsheetRowNumber: row.rowNumber,
    _smartsheetParentId: row.parentId,
    _smartsheetIndent: row.indent || 0
  };
}

/**
 * Converts Smartsheet rows to local format, resolving dependencies
 */
function convertSmartsheetRowsToLocal(rows, columns, config) {
  // First pass: convert all rows
  const localRows = rows.map(row => smartsheetRowToLocalRow(row, columns, config));

  // Build row number to task ID mapping for dependency resolution
  const rowNumberToTaskId = {};
  localRows.forEach(row => {
    if (row._smartsheetRowNumber) {
      rowNumberToTaskId[row._smartsheetRowNumber] = row.taskId;
    }
  });

  // Second pass: resolve dependency references
  localRows.forEach(row => {
    if (row.dependencies) {
      const deps = row.dependencies.split(',').filter(d => d);
      const resolvedDeps = deps.map(dep => {
        const match = dep.match(/SS-ROW-(\d+)/);
        if (match) {
          const rowNum = parseInt(match[1], 10);
          return rowNumberToTaskId[rowNum] || dep;
        }
        return dep;
      }).filter(d => d && !d.startsWith('SS-ROW-'));
      row.dependencies = resolvedDeps.join(',');
    }
  });

  return localRows;
}

/**
 * Converts a local row to Smartsheet row array
 */
function localRowToSheetRow(localRow) {
  return [
    localRow.taskId,
    localRow.taskName,
    localRow.startDate,
    localRow.endDate,
    localRow.duration || '',
    localRow.owner,
    localRow.percentComplete,
    localRow.priority,
    localRow.smartsheetUrl || '',
    localRow.parentTask,
    localRow.dependencies,
    localRow.taskType,
    localRow.swimlane,
    ''  // Modified = empty (fresh from Smartsheet)
  ];
}

/**
 * Detects conflicts between Smartsheet data and sheet data
 */
function detectSmartsheetConflicts(smartsheetRows, sheetDataMap) {
  const conflicts = [];

  smartsheetRows.forEach(ssRow => {
    const sheetEntry = sheetDataMap.get(ssRow.taskId);

    if (sheetEntry) {
      const sheetRow = sheetEntry.data;
      const isModified = String(sheetRow[COL.MODIFIED]).toLowerCase() === 'yes';

      if (isModified) {
        // Check for actual differences
        const diffs = [];

        // Compare dates
        const sheetStartDate = parseDate(sheetRow[COL.START_DATE]);
        const sheetEndDate = parseDate(sheetRow[COL.END_DATE]);

        if (ssRow.startDate && sheetStartDate &&
            ssRow.startDate.getTime() !== sheetStartDate.getTime()) {
          diffs.push({
            field: 'Start Date',
            smartsheetValue: formatDateForDisplay(ssRow.startDate),
            sheetValue: formatDateForDisplay(sheetStartDate)
          });
        }

        if (ssRow.endDate && sheetEndDate &&
            ssRow.endDate.getTime() !== sheetEndDate.getTime()) {
          diffs.push({
            field: 'End Date',
            smartsheetValue: formatDateForDisplay(ssRow.endDate),
            sheetValue: formatDateForDisplay(sheetEndDate)
          });
        }

        // Compare owner
        if (ssRow.owner !== String(sheetRow[COL.OWNER]).trim()) {
          diffs.push({
            field: 'Owner',
            smartsheetValue: ssRow.owner || '(none)',
            sheetValue: String(sheetRow[COL.OWNER]).trim() || '(none)'
          });
        }

        // Compare percent complete
        const sheetPercent = Number(sheetRow[COL.PERCENT_COMPLETE]) || 0;
        if (ssRow.percentComplete !== sheetPercent) {
          diffs.push({
            field: '% Complete',
            smartsheetValue: ssRow.percentComplete + '%',
            sheetValue: sheetPercent + '%'
          });
        }

        // Compare priority
        if (ssRow.priority !== String(sheetRow[COL.PRIORITY]).trim()) {
          diffs.push({
            field: 'Priority',
            smartsheetValue: ssRow.priority,
            sheetValue: String(sheetRow[COL.PRIORITY]).trim()
          });
        }

        if (diffs.length > 0) {
          conflicts.push({
            taskId: ssRow.taskId,
            taskName: ssRow.taskName,
            rowIndex: sheetEntry.rowIndex,
            diffs: diffs,
            smartsheetData: ssRow,
            sheetData: sheetRow
          });
        }
      }
    }
  });

  return conflicts;
}

/**
 * Main pull function - fetches from Smartsheet and updates sheet
 */
function pullFromSmartsheet() {
  // Check authorization before proceeding
  if (!requireAuthorization('pull from Smartsheet')) {
    return;
  }

  const ui = SpreadsheetApp.getUi();

  try {
    const config = getSmartsheetConfig();

    if (!config.apiToken || !config.sheetId) {
      ui.alert('Not Configured', 'Please configure Smartsheet connection first.', ui.ButtonSet.OK);
      return;
    }

    // Fetch data from Smartsheet
    ui.alert('Pulling from Smartsheet', 'Fetching data from Smartsheet...', ui.ButtonSet.OK);

    const { sheet, columns, rows } = fetchSmartsheetRows();

    if (rows.length === 0) {
      ui.alert('No Data', 'No rows found in the Smartsheet.', ui.ButtonSet.OK);
      return;
    }

    // Convert to local format
    const localRows = convertSmartsheetRowsToLocal(rows, columns, config);

    // Filter out rows without task names
    const validRows = localRows.filter(row => row.taskName);

    if (validRows.length === 0) {
      ui.alert('No Valid Data', 'No rows with task names found.', ui.ButtonSet.OK);
      return;
    }

    // Get current sheet data
    const sheetDataMap = getSheetDataMap();

    // Detect conflicts
    const conflicts = detectSmartsheetConflicts(validRows, sheetDataMap);

    if (conflicts.length > 0) {
      // Show conflict resolution dialog
      showSmartsheetConflictDialog(validRows, conflicts);
    } else {
      // No conflicts, apply updates directly
      applySmartsheetPull(validRows, []);
      ui.alert('Pull Complete', `Successfully pulled ${validRows.length} rows from Smartsheet.`, ui.ButtonSet.OK);
    }

    // Update last sync time
    config.lastSyncTime = new Date().toISOString();
    saveSmartsheetConfig(config);

  } catch (e) {
    ui.alert('Error', 'Failed to pull from Smartsheet: ' + e.message, ui.ButtonSet.OK);
  }
}

/**
 * Shows Smartsheet conflict resolution dialog
 */
function showSmartsheetConflictDialog(smartsheetRows, conflicts) {
  // Reuse the same conflict dialog with Smartsheet-specific data
  const template = HtmlService.createTemplateFromFile('SmartsheetConflictDialog');
  template.conflicts = JSON.stringify(conflicts);
  template.smartsheetRowsJson = JSON.stringify(smartsheetRows);

  const html = template.evaluate()
    .setWidth(700)
    .setHeight(600)
    .setTitle('Resolve Smartsheet Conflicts');

  SpreadsheetApp.getUi().showModalDialog(html, 'Resolve Smartsheet Conflicts');
}

/**
 * Applies Smartsheet pull with conflict resolutions
 * @param {Array} smartsheetRows - All rows from Smartsheet
 * @param {Array} keepLocalIds - Task IDs where local changes should be kept
 */
function applySmartsheetPull(smartsheetRows, keepLocalIds) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let dataSheet = ss.getSheetByName(DATA_SHEET_NAME);

  if (!dataSheet) {
    setupDataSheet();
    dataSheet = ss.getSheetByName(DATA_SHEET_NAME);
  }

  const keepLocalSet = new Set(keepLocalIds);
  const sheetDataMap = getSheetDataMap();

  // Prepare updates and inserts
  const updates = [];
  const inserts = [];

  smartsheetRows.forEach(ssRow => {
    const sheetEntry = sheetDataMap.get(ssRow.taskId);

    if (sheetEntry) {
      // Update existing row (unless keeping local)
      if (!keepLocalSet.has(ssRow.taskId)) {
        updates.push({
          rowIndex: sheetEntry.rowIndex,
          data: localRowToSheetRow(ssRow)
        });
      }
    } else {
      // Insert new row
      inserts.push(localRowToSheetRow(ssRow));
    }
  });

  // Apply updates
  updates.forEach(update => {
    dataSheet.getRange(update.rowIndex, 1, 1, COLUMN_HEADERS.length)
      .setValues([update.data]);
  });

  // Apply inserts
  if (inserts.length > 0) {
    const lastRow = dataSheet.getLastRow();
    dataSheet.getRange(lastRow + 1, 1, inserts.length, COLUMN_HEADERS.length)
      .setValues(inserts);
  }

  return {
    success: true,
    updated: updates.length,
    inserted: inserts.length
  };
}

/**
 * Pulls only selected rows from Smartsheet
 */
function pullSelectedFromSmartsheet() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  if (sheet.getName() !== DATA_SHEET_NAME) {
    ui.alert('Wrong Sheet', `Please select rows in the "${DATA_SHEET_NAME}" sheet.`, ui.ButtonSet.OK);
    return;
  }

  const selection = sheet.getActiveRange();
  const startRow = selection.getRow();
  const numRows = selection.getNumRows();

  if (startRow < 2) {
    ui.alert('Invalid Selection', 'Please select data rows (not the header).', ui.ButtonSet.OK);
    return;
  }

  // Get task IDs from selected rows
  const taskIds = [];
  for (let i = 0; i < numRows; i++) {
    const taskId = sheet.getRange(startRow + i, 1).getValue();
    if (taskId) {
      taskIds.push(String(taskId).trim());
    }
  }

  if (taskIds.length === 0) {
    ui.alert('No Tasks', 'No task IDs found in selection.', ui.ButtonSet.OK);
    return;
  }

  try {
    const config = getSmartsheetConfig();

    // Fetch all data from Smartsheet
    const { sheet: ssSheet, columns, rows } = fetchSmartsheetRows();

    // Convert to local format
    const localRows = convertSmartsheetRowsToLocal(rows, columns, config);

    // Filter to only selected task IDs
    const selectedRows = localRows.filter(row => taskIds.includes(row.taskId));

    if (selectedRows.length === 0) {
      ui.alert('No Matches', 'Selected tasks not found in Smartsheet.', ui.ButtonSet.OK);
      return;
    }

    const sheetDataMap = getSheetDataMap();
    const conflicts = detectSmartsheetConflicts(selectedRows, sheetDataMap);

    if (conflicts.length > 0) {
      showSmartsheetConflictDialog(selectedRows, conflicts);
    } else {
      applySmartsheetPull(selectedRows, []);
      ui.alert('Pull Complete', `Successfully pulled ${selectedRows.length} rows.`, ui.ButtonSet.OK);
    }

  } catch (e) {
    ui.alert('Error', 'Failed to pull from Smartsheet: ' + e.message, ui.ButtonSet.OK);
  }
}

/**
 * Pushes changes to Smartsheet
 */
function pushToSmartsheet() {
  // Check authorization before proceeding
  if (!requireAuthorization('push to Smartsheet')) {
    return;
  }

  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(DATA_SHEET_NAME);

  if (!dataSheet) {
    ui.alert('No Data', 'Data sheet not found.', ui.ButtonSet.OK);
    return;
  }

  const config = getSmartsheetConfig();
  if (!config.apiToken || !config.sheetId) {
    ui.alert('Not Configured', 'Please configure Smartsheet connection first.', ui.ButtonSet.OK);
    return;
  }

  // Find all modified rows
  const lastRow = dataSheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('No Data', 'No tasks to push.', ui.ButtonSet.OK);
    return;
  }

  const data = dataSheet.getRange(2, 1, lastRow - 1, COLUMN_HEADERS.length).getValues();
  const modifiedRows = [];

  data.forEach((row, index) => {
    if (String(row[COL.MODIFIED]).toLowerCase() === 'yes') {
      modifiedRows.push({
        rowIndex: index + 2,
        data: row
      });
    }
  });

  if (modifiedRows.length === 0) {
    ui.alert('No Changes', 'No modified rows to push. Mark rows as "Yes" in the Modified column to push changes.', ui.ButtonSet.OK);
    return;
  }

  const response = ui.alert(
    'Push to Smartsheet',
    `Push ${modifiedRows.length} modified row(s) to Smartsheet?\n\nThis will update: Start Date, End Date, Duration, Owner, % Complete, and Priority.`,
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  const results = pushRowsToSmartsheet(modifiedRows);

  // Clear modified flag for successful pushes
  results.successful.forEach(taskId => {
    const sheetDataMap = getSheetDataMap();
    const entry = sheetDataMap.get(taskId);
    if (entry) {
      dataSheet.getRange(entry.rowIndex, COL.MODIFIED + 1).setValue('');
    }
  });

  let message = `Successfully pushed ${results.successful.length} task(s).`;
  if (results.failed.length > 0) {
    message += `\n\nFailed to push ${results.failed.length} task(s):\n${results.failed.map(f => `${f.taskId}: ${f.error}`).join('\n')}`;
  }

  ui.alert('Push Complete', message, ui.ButtonSet.OK);
}

/**
 * Pushes specific rows to Smartsheet
 */
function pushRowsToSmartsheet(rows) {
  const config = getSmartsheetConfig();
  const successful = [];
  const failed = [];

  // First, we need to map task IDs to Smartsheet row IDs
  // Fetch current Smartsheet data
  let smartsheetRowMap = {};
  try {
    const { rows: ssRows, columns } = fetchSmartsheetRows();
    const localRows = convertSmartsheetRowsToLocal(ssRows, columns, config);
    localRows.forEach(row => {
      smartsheetRowMap[row.taskId] = row._smartsheetRowId;
    });
  } catch (e) {
    return { successful: [], failed: rows.map(r => ({ taskId: String(r.data[COL.TASK_ID]), error: 'Failed to fetch Smartsheet data: ' + e.message })) };
  }

  // Build update payload
  const rowUpdates = [];

  rows.forEach(row => {
    const taskId = String(row.data[COL.TASK_ID]).trim();
    const ssRowId = smartsheetRowMap[taskId];

    if (!ssRowId) {
      failed.push({ taskId, error: 'Task not found in Smartsheet' });
      return;
    }

    try {
      const cells = [];

      // Build cells array based on column mapping
      const mapping = config.columnMapping;

      // Start Date
      if (mapping.colStartDate) {
        const startDate = parseDate(row.data[COL.START_DATE]);
        if (startDate) {
          cells.push({
            columnId: Number(mapping.colStartDate),
            value: formatDateForSmartsheet(startDate)
          });
        }
      }

      // End Date
      if (mapping.colEndDate) {
        const endDate = parseDate(row.data[COL.END_DATE]);
        if (endDate) {
          cells.push({
            columnId: Number(mapping.colEndDate),
            value: formatDateForSmartsheet(endDate)
          });
        }
      }

      // Duration
      if (mapping.colDuration) {
        const duration = row.data[COL.DURATION];
        if (duration) {
          cells.push({
            columnId: Number(mapping.colDuration),
            value: duration
          });
        }
      }

      // Owner
      if (mapping.colOwner) {
        const owner = String(row.data[COL.OWNER]).trim();
        if (owner) {
          cells.push({
            columnId: Number(mapping.colOwner),
            value: owner
          });
        }
      }

      // % Complete
      if (mapping.colPercentComplete) {
        const percent = Number(row.data[COL.PERCENT_COMPLETE]) || 0;
        cells.push({
          columnId: Number(mapping.colPercentComplete),
          value: percent / 100  // Smartsheet expects 0-1 for percentages
        });
      }

      // Priority
      if (mapping.colPriority) {
        const priority = String(row.data[COL.PRIORITY]).trim();
        if (priority) {
          // Map back to Smartsheet priority values
          const ssPriority = Object.keys(SMARTSHEET_PRIORITY_MAPPING).find(
            k => SMARTSHEET_PRIORITY_MAPPING[k] === priority
          ) || priority;
          cells.push({
            columnId: Number(mapping.colPriority),
            value: ssPriority
          });
        }
      }

      if (cells.length > 0) {
        rowUpdates.push({
          id: ssRowId,
          cells: cells,
          _taskId: taskId
        });
      } else {
        successful.push(taskId); // No cells to update, but not a failure
      }

    } catch (e) {
      failed.push({ taskId, error: e.message });
    }
  });

  // Send batch update to Smartsheet
  if (rowUpdates.length > 0) {
    try {
      // Smartsheet allows up to 500 rows per request
      const batchSize = 500;
      for (let i = 0; i < rowUpdates.length; i += batchSize) {
        const batch = rowUpdates.slice(i, i + batchSize);
        const payload = batch.map(r => ({ id: r.id, cells: r.cells }));

        smartsheetApiRequest(`/sheets/${config.sheetId}/rows`, 'PUT', payload);

        // Mark as successful
        batch.forEach(r => {
          successful.push(r._taskId);
        });
      }
    } catch (e) {
      // Mark remaining as failed
      rowUpdates.forEach(r => {
        if (!successful.includes(r._taskId)) {
          failed.push({ taskId: r._taskId, error: e.message });
        }
      });
    }
  }

  return { successful, failed };
}

/**
 * Formats a date for Smartsheet API
 */
function formatDateForSmartsheet(date) {
  if (!date) return null;
  return Utilities.formatDate(date, 'UTC', 'yyyy-MM-dd');
}

/**
 * Pushes only selected rows to Smartsheet
 */
function pushSelectedToSmartsheet() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  if (sheet.getName() !== DATA_SHEET_NAME) {
    ui.alert('Wrong Sheet', `Please select rows in the "${DATA_SHEET_NAME}" sheet.`, ui.ButtonSet.OK);
    return;
  }

  const config = getSmartsheetConfig();
  if (!config.apiToken || !config.sheetId) {
    ui.alert('Not Configured', 'Please configure Smartsheet connection first.', ui.ButtonSet.OK);
    return;
  }

  const selection = sheet.getActiveRange();
  const startRow = selection.getRow();
  const numRows = selection.getNumRows();

  if (startRow < 2) {
    ui.alert('Invalid Selection', 'Please select data rows (not the header).', ui.ButtonSet.OK);
    return;
  }

  // Get selected rows
  const selectedRows = [];
  for (let i = 0; i < numRows; i++) {
    const rowData = sheet.getRange(startRow + i, 1, 1, COLUMN_HEADERS.length).getValues()[0];
    if (rowData[COL.TASK_ID]) {
      selectedRows.push({
        rowIndex: startRow + i,
        data: rowData
      });
    }
  }

  if (selectedRows.length === 0) {
    ui.alert('No Tasks', 'No tasks found in selection.', ui.ButtonSet.OK);
    return;
  }

  const response = ui.alert(
    'Push to Smartsheet',
    `Push ${selectedRows.length} selected row(s) to Smartsheet?`,
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  const results = pushRowsToSmartsheet(selectedRows);

  // Clear modified flag for successful pushes
  const dataSheet = ss.getSheetByName(DATA_SHEET_NAME);
  results.successful.forEach(taskId => {
    selectedRows.forEach(row => {
      if (String(row.data[COL.TASK_ID]).trim() === taskId) {
        dataSheet.getRange(row.rowIndex, COL.MODIFIED + 1).setValue('');
      }
    });
  });

  let message = `Successfully pushed ${results.successful.length} task(s).`;
  if (results.failed.length > 0) {
    message += `\n\nFailed: ${results.failed.map(f => `${f.taskId}: ${f.error}`).join('\n')}`;
  }

  ui.alert('Push Complete', message, ui.ButtonSet.OK);
}

/**
 * Shows Smartsheet sync status
 */
function showSmartsheetSyncStatus() {
  const config = getSmartsheetConfig();
  const sheetDataMap = getSheetDataMap();

  let modifiedCount = 0;
  sheetDataMap.forEach(entry => {
    if (String(entry.data[COL.MODIFIED]).toLowerCase() === 'yes') {
      modifiedCount++;
    }
  });

  const lastSync = config.lastSyncTime
    ? new Date(config.lastSyncTime).toLocaleString()
    : 'Never';

  const sheetName = config.sheetName || config.sheetId || 'Not configured';

  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Smartsheet Sync Status',
    `Connected Sheet: ${sheetName}\n\n` +
    `Last sync: ${lastSync}\n` +
    `Total tasks in sheet: ${sheetDataMap.size}\n` +
    `Modified (pending push): ${modifiedCount}`,
    ui.ButtonSet.OK
  );
}
