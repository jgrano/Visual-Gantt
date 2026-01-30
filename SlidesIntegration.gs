/**
 * SlidesIntegration.gs - Google Slides Add-on Integration
 * Enables Visual Gantt charts to be inserted and managed in Google Slides
 *
 * Supports 3 approaches:
 * 1. Image insertion (static, high-quality)
 * 2. Linked Sheets chart (auto-updating from source)
 * 3. Native shapes (fully interactive, animatable)
 */

// ============================================================================
// SLIDES ADD-ON ENTRY POINTS
// ============================================================================

/**
 * Creates the sidebar when the add-on is opened in Slides
 * Default: Shows unified sidebar (self-contained, no external dependencies)
 * @returns {GoogleAppsScript.Card_Service.Card}
 */
function onSlidesHomepage() {
  // Show unified sidebar by default
  showUnifiedSlidesSidebar();
  return null; // Return null since we're showing HTML sidebar
}

/**
 * Shows the Slides sidebar UI (Drive-based workflow)
 */
function showSlidesSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('SlidesSidebar')
    .setTitle('Visual Gantt')
    .setWidth(300);

  SlidesApp.getUi().showSidebar(html);
}

/**
 * Creates the card-based sidebar for Slides add-on
 * @returns {GoogleAppsScript.Card_Service.Card}
 */
function createSlidesSidebar() {
  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle('Visual Gantt')
      .setSubtitle('Insert project timelines'))
    .addSection(createProjectSelectionSection())
    .addSection(createInsertOptionsSection())
    .addSection(createHelpSection())
    .build();

  return card;
}

/**
 * Creates the project selection section
 * @returns {GoogleAppsScript.Card_Service.CardSection}
 */
function createProjectSelectionSection() {
  const section = CardService.newCardSection()
    .setHeader('Select Project');

  // Button to browse Drive projects
  section.addWidget(CardService.newTextButton()
    .setText('Browse Projects')
    .setOnClickAction(CardService.newAction()
      .setFunctionName('showProjectPicker')));

  // Button to connect to Sheets
  section.addWidget(CardService.newTextButton()
    .setText('Connect to Sheet')
    .setOnClickAction(CardService.newAction()
      .setFunctionName('showSheetPicker')));

  return section;
}

/**
 * Creates the insert options section
 * @returns {GoogleAppsScript.Card_Service.CardSection}
 */
function createInsertOptionsSection() {
  const section = CardService.newCardSection()
    .setHeader('Insert Timeline');

  // Approach 1: Insert as image
  section.addWidget(CardService.newTextButton()
    .setText('Insert as Image')
    .setOnClickAction(CardService.newAction()
      .setFunctionName('insertTimelineAsImage')));

  // Approach 2: Link to Sheet
  section.addWidget(CardService.newTextButton()
    .setText('Link to Sheet Chart')
    .setOnClickAction(CardService.newAction()
      .setFunctionName('insertLinkedSheetChart')));

  // Approach 3: Insert as shapes
  section.addWidget(CardService.newTextButton()
    .setText('Insert as Shapes')
    .setOnClickAction(CardService.newAction()
      .setFunctionName('insertTimelineAsShapes')));

  return section;
}

/**
 * Creates the help section
 * @returns {GoogleAppsScript.Card_Service.CardSection}
 */
function createHelpSection() {
  const section = CardService.newCardSection();

  section.addWidget(CardService.newTextButton()
    .setText('Open Projects Folder')
    .setOnClickAction(CardService.newAction()
      .setFunctionName('openProjectsFolder')));

  section.addWidget(CardService.newTextButton()
    .setText('Help & Documentation')
    .setOnClickAction(CardService.newAction()
      .setFunctionName('showSlidesHelp')));

  return section;
}

// ============================================================================
// APPROACH 1: IMAGE INSERTION
// ============================================================================

/**
 * Inserts a Gantt chart as a static image on the current slide
 * This is the highest quality, presentation-ready approach
 * @returns {Object} Result of the insertion
 */
function insertTimelineAsImage() {
  try {
    const presentation = SlidesApp.getActivePresentation();
    const selection = presentation.getSelection();
    const currentSlide = selection.getCurrentPage();

    if (!currentSlide || currentSlide.getPageType() !== SlidesApp.PageType.SLIDE) {
      return {
        success: false,
        message: 'Please select a slide first'
      };
    }

    // Get selected project
    const projectId = PropertiesService.getUserProperties().getProperty('selectedProjectId');
    if (!projectId) {
      return {
        success: false,
        message: 'Please select a project first'
      };
    }

    // Load project data
    const projectData = loadProjectFromDrive(projectId);
    if (!projectData) {
      return {
        success: false,
        message: 'Could not load project data'
      };
    }

    // Generate chart image
    const chartBlob = generateChartBlobFromProject(projectData);
    if (!chartBlob) {
      return {
        success: false,
        message: 'Failed to generate chart'
      };
    }

    // Insert image on slide
    const image = currentSlide.insertImage(chartBlob);

    // Position and size the image
    const slideWidth = currentSlide.getPageWidth();
    const slideHeight = currentSlide.getPageHeight();

    // Center the image with margins
    const margin = 20;
    const maxWidth = slideWidth - (2 * margin);
    const maxHeight = slideHeight - (2 * margin);

    // Scale to fit
    const imageWidth = image.getWidth();
    const imageHeight = image.getHeight();
    const scale = Math.min(maxWidth / imageWidth, maxHeight / imageHeight, 1);

    image.setWidth(imageWidth * scale);
    image.setHeight(imageHeight * scale);

    // Center on slide
    image.setLeft((slideWidth - (imageWidth * scale)) / 2);
    image.setTop((slideHeight - (imageHeight * scale)) / 2);

    // Add alt text with project info
    image.setDescription('Visual Gantt: ' + projectData.name);

    return {
      success: true,
      message: 'Timeline inserted successfully',
      projectName: projectData.name
    };
  } catch (e) {
    Logger.log('Error inserting timeline as image: ' + e.message);
    return {
      success: false,
      message: 'Failed to insert timeline: ' + e.message
    };
  }
}

/**
 * Generates a chart blob from project data (reuses existing renderer)
 * @param {Object} projectData - The project data
 * @returns {GoogleAppsScript.Base.Blob} The chart image blob
 */
function generateChartBlobFromProject(projectData) {
  try {
    // Use existing chart generation logic
    const config = projectData.config || getConfig();
    const tasks = projectData.tasks || [];

    // Generate chart data (reuse existing logic from Code.gs)
    const chartData = {
      tasks: tasks,
      config: config
    };

    // This would call the existing chart renderer
    // For now, return a placeholder - will integrate with ChartRenderer.html
    return null; // TODO: Integrate with existing chart generation
  } catch (e) {
    Logger.log('Error generating chart blob: ' + e.message);
    return null;
  }
}

// ============================================================================
// APPROACH 2: LINKED SHEET CHART
// ============================================================================

/**
 * Inserts a linked chart from Google Sheets
 * The chart auto-updates when the source sheet changes
 * @returns {Object} Result of the insertion
 */
function insertLinkedSheetChart() {
  try {
    const presentation = SlidesApp.getActivePresentation();
    const selection = presentation.getSelection();
    const currentSlide = selection.getCurrentPage();

    if (!currentSlide || currentSlide.getPageType() !== SlidesApp.PageType.SLIDE) {
      return {
        success: false,
        message: 'Please select a slide first'
      };
    }

    // Get linked sheet info
    const sheetId = PropertiesService.getUserProperties().getProperty('linkedSheetId');
    if (!sheetId) {
      return {
        success: false,
        message: 'Please connect to a Google Sheet first'
      };
    }

    // Open the sheet and get the timeline image
    const ss = SpreadsheetApp.openById(sheetId);
    const timelineSheet = ss.getSheetByName(TIMELINE_SHEET_NAME);

    if (!timelineSheet) {
      return {
        success: false,
        message: 'Timeline View sheet not found. Generate timeline in Sheets first.'
      };
    }

    // Get all images from the timeline sheet
    const images = timelineSheet.getImages();
    if (images.length === 0) {
      return {
        success: false,
        message: 'No timeline image found. Generate timeline in Sheets first.'
      };
    }

    // Use the first (or most recent) image
    const sheetImage = images[images.length - 1];
    const imageBlob = sheetImage.getBlob();

    // Insert the image on the slide
    const slideImage = currentSlide.insertImage(imageBlob);

    // Position and size
    const slideWidth = currentSlide.getPageWidth();
    const slideHeight = currentSlide.getPageHeight();
    const margin = 20;

    slideImage.setWidth(slideWidth - (2 * margin));
    slideImage.setLeft(margin);
    slideImage.setTop(margin);

    // Set link back to sheet
    slideImage.setLinkUrl(ss.getUrl());
    slideImage.setDescription('Linked to: ' + ss.getName());

    return {
      success: true,
      message: 'Linked timeline inserted successfully',
      sheetName: ss.getName(),
      sheetUrl: ss.getUrl()
    };
  } catch (e) {
    Logger.log('Error inserting linked sheet chart: ' + e.message);
    return {
      success: false,
      message: 'Failed to insert linked chart: ' + e.message
    };
  }
}

/**
 * Refreshes all linked sheet charts on the current slide
 * @returns {Object} Result of the refresh
 */
function refreshLinkedCharts() {
  try {
    const presentation = SlidesApp.getActivePresentation();
    const selection = presentation.getSelection();
    const currentSlide = selection.getCurrentPage();

    if (!currentSlide || currentSlide.getPageType() !== SlidesApp.PageType.SLIDE) {
      return {
        success: false,
        message: 'Please select a slide first'
      };
    }

    // Get all images on the slide
    const images = currentSlide.getImages();
    let refreshedCount = 0;

    for (const image of images) {
      const linkUrl = image.getLink() ? image.getLink().getUrl() : null;

      // Check if it's a Sheets link
      if (linkUrl && linkUrl.includes('docs.google.com/spreadsheets')) {
        try {
          // Extract sheet ID from URL
          const sheetIdMatch = linkUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);
          if (sheetIdMatch) {
            const sheetId = sheetIdMatch[1];
            const ss = SpreadsheetApp.openById(sheetId);
            const timelineSheet = ss.getSheetByName(TIMELINE_SHEET_NAME);

            if (timelineSheet) {
              const sheetImages = timelineSheet.getImages();
              if (sheetImages.length > 0) {
                const newImageBlob = sheetImages[sheetImages.length - 1].getBlob();

                // Replace the image while preserving position and size
                const left = image.getLeft();
                const top = image.getTop();
                const width = image.getWidth();
                const height = image.getHeight();

                image.replace(newImageBlob);
                image.setLeft(left);
                image.setTop(top);
                image.setWidth(width);
                image.setHeight(height);

                refreshedCount++;
              }
            }
          }
        } catch (e) {
          Logger.log('Error refreshing image: ' + e.message);
        }
      }
    }

    return {
      success: true,
      message: 'Refreshed ' + refreshedCount + ' linked chart(s)',
      count: refreshedCount
    };
  } catch (e) {
    Logger.log('Error refreshing linked charts: ' + e.message);
    return {
      success: false,
      message: 'Failed to refresh charts: ' + e.message
    };
  }
}

// ============================================================================
// APPROACH 3: NATIVE SHAPES (calls SlidesShapesRenderer.gs)
// ============================================================================

/**
 * Inserts timeline as native Slides shapes (fully interactive)
 * See SlidesShapesRenderer.gs for implementation
 * @returns {Object} Result of the insertion
 */
function insertTimelineAsShapes() {
  // This calls the dedicated shapes renderer
  return insertTimelineAsNativeShapes();
}

// ============================================================================
// PROJECT & SHEET PICKERS
// ============================================================================

/**
 * Shows the project picker dialog
 */
function showProjectPicker() {
  const html = HtmlService.createHtmlOutputFromFile('ProjectSelector')
    .setWidth(400)
    .setHeight(500)
    .setTitle('Select Project');

  SlidesApp.getUi().showModalDialog(html, 'Select Project');
}

/**
 * Shows the sheet picker (for linking)
 */
function showSheetPicker() {
  const html = HtmlService.createHtmlOutput(`
    <html>
      <body>
        <h3>Connect to Google Sheet</h3>
        <p>Enter the Google Sheet URL or ID:</p>
        <input type="text" id="sheetUrl" style="width: 100%; padding: 8px;" placeholder="https://docs.google.com/spreadsheets/d/...">
        <br><br>
        <button onclick="connectSheet()">Connect</button>
        <div id="status"></div>

        <script>
          function connectSheet() {
            const url = document.getElementById('sheetUrl').value;
            google.script.run
              .withSuccessHandler(onSuccess)
              .withFailureHandler(onError)
              .connectToSheet(url);
          }

          function onSuccess(result) {
            if (result.success) {
              document.getElementById('status').innerHTML = '<p style="color: green;">✓ Connected to: ' + result.sheetName + '</p>';
              setTimeout(() => google.script.host.close(), 2000);
            } else {
              document.getElementById('status').innerHTML = '<p style="color: red;">Error: ' + result.message + '</p>';
            }
          }

          function onError(error) {
            document.getElementById('status').innerHTML = '<p style="color: red;">Error: ' + error.message + '</p>';
          }
        </script>
      </body>
    </html>
  `)
    .setWidth(400)
    .setHeight(200);

  SlidesApp.getUi().showModalDialog(html, 'Connect to Sheet');
}

/**
 * Connects to a Google Sheet for linking
 * @param {string} sheetUrl - The sheet URL or ID
 * @returns {Object} Result of the connection
 */
function connectToSheet(sheetUrl) {
  try {
    // Extract sheet ID from URL
    let sheetId = sheetUrl;
    const urlMatch = sheetUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (urlMatch) {
      sheetId = urlMatch[1];
    }

    // Try to open the sheet
    const ss = SpreadsheetApp.openById(sheetId);
    const sheetName = ss.getName();

    // Store the connection
    PropertiesService.getUserProperties().setProperty('linkedSheetId', sheetId);
    PropertiesService.getUserProperties().setProperty('linkedSheetName', sheetName);

    return {
      success: true,
      sheetId: sheetId,
      sheetName: sheetName
    };
  } catch (e) {
    return {
      success: false,
      message: 'Could not access sheet: ' + e.message
    };
  }
}

/**
 * Gets the currently selected project ID
 * @returns {string} The project ID
 */
function getSelectedProjectId() {
  return PropertiesService.getUserProperties().getProperty('selectedProjectId');
}

/**
 * Sets the selected project ID
 * @param {string} projectId - The project ID to select
 */
function setSelectedProjectId(projectId) {
  PropertiesService.getUserProperties().setProperty('selectedProjectId', projectId);
}

// ============================================================================
// HELP & UTILITIES
// ============================================================================

/**
 * Shows help documentation for Slides integration
 */
function showSlidesHelp() {
  const html = HtmlService.createHtmlOutput(`
    <html>
      <head>
        <style>
          body { font-family: Arial, sans-serif; padding: 20px; }
          h2 { color: #4285F4; }
          .approach { background: #f5f5f5; padding: 15px; margin: 10px 0; border-radius: 5px; }
          .approach h3 { margin-top: 0; }
        </style>
      </head>
      <body>
        <h2>Visual Gantt for Slides</h2>
        <p>Three ways to insert timelines:</p>

        <div class="approach">
          <h3>1. Insert as Image</h3>
          <p><strong>Best for:</strong> Presentations, static snapshots</p>
          <p><strong>Features:</strong> High quality, no dependencies</p>
          <p><strong>Updates:</strong> Manual (re-insert to update)</p>
        </div>

        <div class="approach">
          <h3>2. Link to Sheet Chart</h3>
          <p><strong>Best for:</strong> Live data, auto-updating</p>
          <p><strong>Features:</strong> Links back to source Sheet</p>
          <p><strong>Updates:</strong> Click "Refresh" to update</p>
        </div>

        <div class="approach">
          <h3>3. Insert as Shapes</h3>
          <p><strong>Best for:</strong> Customization, animations</p>
          <p><strong>Features:</strong> Each task is editable, animatable</p>
          <p><strong>Updates:</strong> Edit directly in Slides</p>
        </div>

        <h3>Data Storage</h3>
        <p>All project data is stored in Google Drive at:<br>
        <code>Google Drive / Visual Gantt Projects</code></p>

        <h3>Getting Started</h3>
        <ol>
          <li>Create project in Sheets or browse Drive projects</li>
          <li>Select insertion method</li>
          <li>Click insert button</li>
        </ol>
      </body>
    </html>
  `)
    .setWidth(500)
    .setHeight(600);

  SlidesApp.getUi().showModalDialog(html, 'Help - Visual Gantt for Slides');
}

// ============================================================================
// UNIFIED WORKFLOW (SINGLE-FILE APPROACH)
// ============================================================================

/**
 * Shows the unified sidebar (self-contained, no external dependencies)
 */
function showUnifiedSlidesSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('UnifiedSlidesSidebar')
    .setTitle('Visual Gantt')
    .setWidth(300);

  SlidesApp.getUi().showSidebar(html);
}

/**
 * Shows the task editor dialog
 */
function showTaskEditor() {
  const html = HtmlService.createHtmlOutputFromFile('TaskEditor')
    .setWidth(500)
    .setHeight(600)
    .setTitle('Task Manager');

  SlidesApp.getUi().showModalDialog(html, 'Task Manager');
}

/**
 * Shows import dialog for CSV/JSON
 */
function showImportDialog() {
  const html = HtmlService.createHtmlOutput(`
    <html>
      <head>
        <style>
          body { font-family: Arial, sans-serif; padding: 20px; }
          textarea { width: 100%; height: 300px; font-family: monospace; font-size: 12px; }
          .button { padding: 10px 20px; margin: 5px; cursor: pointer; }
          .primary { background: #1a73e8; color: white; border: none; border-radius: 4px; }
          .secondary { background: #f1f3f4; color: #202124; border: none; border-radius: 4px; }
          #status { margin-top: 10px; padding: 10px; border-radius: 4px; display: none; }
          .success { background: #e6f4ea; color: #1e8e3e; display: block; }
          .error { background: #fce8e6; color: #c5221f; display: block; }
        </style>
      </head>
      <body>
        <h2>Import Data</h2>
        <p>Paste JSON or CSV data below:</p>
        <textarea id="importData" placeholder="Paste JSON or CSV here..."></textarea>
        <div id="status"></div>
        <br>
        <button class="button primary" onclick="importData()">Import</button>
        <button class="button secondary" onclick="google.script.host.close()">Cancel</button>

        <script>
          function importData() {
            const data = document.getElementById('importData').value;
            if (!data.trim()) {
              showStatus('Please paste data first', 'error');
              return;
            }

            // Try JSON first
            try {
              JSON.parse(data);
              google.script.run
                .withSuccessHandler(onSuccess)
                .withFailureHandler(onError)
                .importProjectJSON(data);
              return;
            } catch (e) {
              // Not JSON, try CSV
            }

            // Try CSV
            if (data.includes(',')) {
              google.script.run
                .withSuccessHandler(onSuccess)
                .withFailureHandler(onError)
                .importTasksFromCSV(data);
            } else {
              showStatus('Invalid data format. Please paste JSON or CSV.', 'error');
            }
          }

          function onSuccess(result) {
            if (result.success) {
              showStatus('✓ Imported ' + result.taskCount + ' tasks', 'success');
              setTimeout(() => google.script.host.close(), 2000);
            } else {
              showStatus('Error: ' + result.message, 'error');
            }
          }

          function onError(error) {
            showStatus('Error: ' + error.message, 'error');
          }

          function showStatus(message, type) {
            const status = document.getElementById('status');
            status.textContent = message;
            status.className = type;
          }
        </script>
      </body>
    </html>
  `)
    .setWidth(600)
    .setHeight(500);

  SlidesApp.getUi().showModalDialog(html, 'Import Data');
}

/**
 * Inserts timeline as image using embedded presentation data
 * @returns {Object} Result of the insertion
 */
function insertTimelineAsImageFromPresentation() {
  try {
    const projectData = loadProjectFromPresentation();

    if (!projectData || !projectData.tasks || projectData.tasks.length === 0) {
      return {
        success: false,
        message: 'No tasks found. Add tasks first.'
      };
    }

    const presentation = SlidesApp.getActivePresentation();
    const selection = presentation.getSelection();
    const currentSlide = selection.getCurrentPage();

    if (!currentSlide || currentSlide.getPageType() !== SlidesApp.PageType.SLIDE) {
      return {
        success: false,
        message: 'Please select a slide first'
      };
    }

    // Generate chart blob (reuses existing logic)
    const chartBlob = generateChartBlobFromProject(projectData);
    if (!chartBlob) {
      return {
        success: false,
        message: 'Failed to generate chart'
      };
    }

    // Insert image
    const image = currentSlide.insertImage(chartBlob);

    // Position and size
    const slideWidth = currentSlide.getPageWidth();
    const slideHeight = currentSlide.getPageHeight();
    const margin = 20;
    const maxWidth = slideWidth - (2 * margin);
    const maxHeight = slideHeight - (2 * margin);

    const imageWidth = image.getWidth();
    const imageHeight = image.getHeight();
    const scale = Math.min(maxWidth / imageWidth, maxHeight / imageHeight, 1);

    image.setWidth(imageWidth * scale);
    image.setHeight(imageHeight * scale);
    image.setLeft((slideWidth - (imageWidth * scale)) / 2);
    image.setTop((slideHeight - (imageHeight * scale)) / 2);

    image.setDescription('Visual Gantt: ' + (projectData.name || 'Timeline'));

    return {
      success: true,
      message: 'Timeline inserted with ' + projectData.tasks.length + ' tasks',
      taskCount: projectData.tasks.length
    };
  } catch (e) {
    Logger.log('Error inserting timeline as image: ' + e.message);
    return {
      success: false,
      message: 'Failed to insert: ' + e.message
    };
  }
}

/**
 * Inserts timeline as native shapes using embedded presentation data
 * @returns {Object} Result of the insertion
 */
function insertTimelineAsShapesFromPresentation() {
  try {
    const projectData = loadProjectFromPresentation();

    if (!projectData || !projectData.tasks || projectData.tasks.length === 0) {
      return {
        success: false,
        message: 'No tasks found. Add tasks first.'
      };
    }

    const presentation = SlidesApp.getActivePresentation();
    const selection = presentation.getSelection();
    const currentSlide = selection.getCurrentPage();

    if (!currentSlide || currentSlide.getPageType() !== SlidesApp.PageType.SLIDE) {
      return {
        success: false,
        message: 'Please select a slide first'
      };
    }

    // Render as shapes (reuses SlidesShapesRenderer.gs)
    const result = renderGanttAsShapes(currentSlide, projectData);

    return result;
  } catch (e) {
    Logger.log('Error inserting timeline as shapes: ' + e.message);
    return {
      success: false,
      message: 'Failed to insert: ' + e.message
    };
  }
}

/**
 * Refreshes the timeline on the current slide (regenerates from latest data)
 * @returns {Object} Result of the refresh
 */
function refreshTimelineOnSlide() {
  try {
    const projectData = loadProjectFromPresentation();

    if (!projectData || !projectData.tasks || projectData.tasks.length === 0) {
      return {
        success: false,
        message: 'No tasks found'
      };
    }

    const presentation = SlidesApp.getActivePresentation();
    const selection = presentation.getSelection();
    const currentSlide = selection.getCurrentPage();

    if (!currentSlide || currentSlide.getPageType() !== SlidesApp.PageType.SLIDE) {
      return {
        success: false,
        message: 'Please select a slide first'
      };
    }

    // Find and remove existing timeline images/shapes
    const images = currentSlide.getImages();
    images.forEach(image => {
      const desc = image.getDescription();
      if (desc && desc.includes('Visual Gantt')) {
        image.remove();
      }
    });

    // Re-insert
    return insertTimelineAsImageFromPresentation();
  } catch (e) {
    Logger.log('Error refreshing timeline: ' + e.message);
    return {
      success: false,
      message: 'Failed to refresh: ' + e.message
    };
  }
}
