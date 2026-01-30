/**
 * SlidesShapesRenderer.gs - Native Slides Shapes Renderer
 * Renders Gantt charts as individual shapes in Google Slides
 *
 * Benefits:
 * - Each task is a separate shape (fully editable)
 * - Can apply Slides animations to individual tasks
 * - Shapes can be rearranged, colored, and styled
 * - No external dependencies (no image files)
 *
 * Usage: Called by SlidesIntegration.gs
 */

// ============================================================================
// MAIN RENDERING FUNCTION
// ============================================================================

/**
 * Inserts a Gantt chart as native Slides shapes
 * @returns {Object} Result of the insertion
 */
function insertTimelineAsNativeShapes() {
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
    const projectId = getSelectedProjectId();
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

    // Render as shapes
    const result = renderGanttAsShapes(currentSlide, projectData);

    return result;
  } catch (e) {
    Logger.log('Error inserting timeline as shapes: ' + e.message);
    return {
      success: false,
      message: 'Failed to insert timeline: ' + e.message
    };
  }
}

/**
 * Renders a Gantt chart as native Slides shapes
 * @param {GoogleAppsScript.Slides.Slide} slide - The slide to render on
 * @param {Object} projectData - The project data
 * @returns {Object} Result of the rendering
 */
function renderGanttAsShapes(slide, projectData) {
  try {
    const config = projectData.config || getConfig();
    const tasks = projectData.tasks || [];

    if (tasks.length === 0) {
      return {
        success: false,
        message: 'No tasks to render'
      };
    }

    // Calculate layout dimensions
    const slideWidth = slide.getPageWidth();
    const slideHeight = slide.getPageHeight();

    const layout = calculateLayout(tasks, config, slideWidth, slideHeight);

    // Create a group to hold all chart elements
    const chartElements = [];

    // 1. Draw background grid
    drawGrid(slide, layout, chartElements);

    // 2. Draw swimlane headers
    drawSwimlaneHeaders(slide, layout, chartElements);

    // 3. Draw date headers
    drawDateHeaders(slide, layout, chartElements);

    // 4. Draw task bars
    drawTaskBars(slide, tasks, layout, config, chartElements);

    // 5. Draw today marker
    if (config.showTodayMarker) {
      drawTodayMarker(slide, layout, chartElements);
    }

    // 6. Draw dependencies
    if (config.showDependencies) {
      drawDependencies(slide, tasks, layout, chartElements);
    }

    // Group all elements (optional - allows moving chart as one unit)
    // Note: Apps Script doesn't support shape grouping via API yet
    // Elements are independent for maximum flexibility

    return {
      success: true,
      message: 'Timeline rendered successfully with ' + chartElements.length + ' shapes',
      shapeCount: chartElements.length,
      taskCount: tasks.length
    };
  } catch (e) {
    Logger.log('Error rendering Gantt as shapes: ' + e.message);
    return {
      success: false,
      message: 'Rendering failed: ' + e.message
    };
  }
}

// ============================================================================
// LAYOUT CALCULATION
// ============================================================================

/**
 * Calculates the layout for the Gantt chart
 * @param {Array} tasks - The tasks to render
 * @param {Object} config - The configuration
 * @param {number} slideWidth - Slide width in points
 * @param {number} slideHeight - Slide height in points
 * @returns {Object} Layout information
 */
function calculateLayout(tasks, config, slideWidth, slideHeight) {
  const margin = 20;
  const headerHeight = 50;
  const swimlaneHeaderWidth = 120;
  const barHeight = config.barHeight || 24;
  const barSpacing = config.barSpacing || 8;

  // Calculate date range
  let minDate = null;
  let maxDate = null;

  tasks.forEach(task => {
    const startDate = new Date(task.startDate);
    const endDate = new Date(task.endDate);

    if (!minDate || startDate < minDate) minDate = startDate;
    if (!maxDate || endDate > maxDate) maxDate = endDate;
  });

  // Add padding to date range
  if (config.startDate) minDate = new Date(config.startDate);
  if (config.endDate) maxDate = new Date(config.endDate);

  const daysPadding = 5;
  minDate = new Date(minDate.getTime() - (daysPadding * 24 * 60 * 60 * 1000));
  maxDate = new Date(maxDate.getTime() + (daysPadding * 24 * 60 * 60 * 1000));

  const totalDays = Math.ceil((maxDate - minDate) / (1000 * 60 * 60 * 24));

  // Calculate chart area
  const chartLeft = margin + swimlaneHeaderWidth;
  const chartTop = margin + headerHeight;
  const chartWidth = slideWidth - chartLeft - margin;
  const chartHeight = slideHeight - chartTop - margin;

  const pixelsPerDay = chartWidth / totalDays;

  // Group tasks by swimlane
  const swimlanes = groupTasksBySwimlane(tasks, config);

  return {
    margin,
    headerHeight,
    swimlaneHeaderWidth,
    barHeight,
    barSpacing,
    chartLeft,
    chartTop,
    chartWidth,
    chartHeight,
    minDate,
    maxDate,
    totalDays,
    pixelsPerDay,
    swimlanes,
    slideWidth,
    slideHeight
  };
}

/**
 * Groups tasks by swimlane
 * @param {Array} tasks - The tasks
 * @param {Object} config - The configuration
 * @returns {Array} Array of swimlane objects
 */
function groupTasksBySwimlane(tasks, config) {
  const groupBy = config.swimlaneGrouping || 'none';

  if (groupBy === 'none') {
    return [{
      name: 'All Tasks',
      tasks: tasks
    }];
  }

  const swimlaneMap = {};

  tasks.forEach(task => {
    let key = 'Ungrouped';

    if (groupBy === 'swimlane' || groupBy === 'category') {
      key = task.swimlane || 'Ungrouped';
    } else if (groupBy === 'owner') {
      key = task.owner || 'Unassigned';
    } else if (groupBy === 'project') {
      key = task.project || 'No Project';
    }

    if (!swimlaneMap[key]) {
      swimlaneMap[key] = [];
    }
    swimlaneMap[key].push(task);
  });

  return Object.keys(swimlaneMap).map(name => ({
    name: name,
    tasks: swimlaneMap[name]
  }));
}

// ============================================================================
// DRAWING FUNCTIONS
// ============================================================================

/**
 * Draws the background grid
 * @param {GoogleAppsScript.Slides.Slide} slide
 * @param {Object} layout
 * @param {Array} elements - Array to collect created shapes
 */
function drawGrid(slide, layout, elements) {
  // Draw vertical grid lines for each week
  const weeksInChart = Math.ceil(layout.totalDays / 7);

  for (let i = 0; i <= weeksInChart; i++) {
    const x = layout.chartLeft + (i * 7 * layout.pixelsPerDay);

    if (x > layout.chartLeft + layout.chartWidth) break;

    const line = slide.insertLine(
      SlidesApp.LineCategory.STRAIGHT,
      x, layout.chartTop,
      x, layout.chartTop + layout.chartHeight
    );

    line.getLineFill().setSolidFill('#E0E0E0');
    line.setWeight(0.5);

    elements.push(line);
  }

  // Draw horizontal grid lines for swimlanes
  let currentY = layout.chartTop;

  layout.swimlanes.forEach(swimlane => {
    const swimlaneHeight = swimlane.tasks.length * (layout.barHeight + layout.barSpacing);

    const line = slide.insertLine(
      SlidesApp.LineCategory.STRAIGHT,
      layout.chartLeft, currentY,
      layout.chartLeft + layout.chartWidth, currentY
    );

    line.getLineFill().setSolidFill('#E0E0E0');
    line.setWeight(1);

    elements.push(line);

    currentY += swimlaneHeight;
  });
}

/**
 * Draws swimlane headers
 * @param {GoogleAppsScript.Slides.Slide} slide
 * @param {Object} layout
 * @param {Array} elements
 */
function drawSwimlaneHeaders(slide, layout, elements) {
  let currentY = layout.chartTop;

  layout.swimlanes.forEach(swimlane => {
    const swimlaneHeight = swimlane.tasks.length * (layout.barHeight + layout.barSpacing);

    const shape = slide.insertShape(
      SlidesApp.ShapeType.RECTANGLE,
      layout.margin,
      currentY,
      layout.swimlaneHeaderWidth,
      swimlaneHeight
    );

    shape.getFill().setSolidFill('#F5F5F5');
    shape.getBorder().setWeight(1);
    shape.getBorder().getLineFill().setSolidFill('#CCCCCC');

    const textRange = shape.getText();
    textRange.setText(swimlane.name);
    textRange.getTextStyle().setFontSize(12);
    textRange.getTextStyle().setBold(true);
    textRange.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    elements.push(shape);

    currentY += swimlaneHeight;
  });
}

/**
 * Draws date headers (month/week labels)
 * @param {GoogleAppsScript.Slides.Slide} slide
 * @param {Object} layout
 * @param {Array} elements
 */
function drawDateHeaders(slide, layout, elements) {
  // Draw month headers
  let currentDate = new Date(layout.minDate);
  let currentMonth = currentDate.getMonth();
  let monthStartX = layout.chartLeft;

  while (currentDate <= layout.maxDate) {
    if (currentDate.getMonth() !== currentMonth) {
      // Draw previous month
      const monthEndX = layout.chartLeft + ((currentDate - layout.minDate) / (1000 * 60 * 60 * 24)) * layout.pixelsPerDay;
      const monthWidth = monthEndX - monthStartX;

      if (monthWidth > 10) {
        const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
        const monthLabel = monthNames[currentMonth] + ' ' + currentDate.getFullYear();

        const shape = slide.insertShape(
          SlidesApp.ShapeType.RECTANGLE,
          monthStartX,
          layout.margin,
          monthWidth,
          layout.headerHeight / 2
        );

        shape.getFill().setSolidFill('#4285F4');
        shape.getBorder().setWeight(0);

        const textRange = shape.getText();
        textRange.setText(monthLabel);
        textRange.getTextStyle().setFontSize(11);
        textRange.getTextStyle().setBold(true);
        textRange.getTextStyle().setForegroundColor('#FFFFFF');
        textRange.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

        elements.push(shape);
      }

      currentMonth = currentDate.getMonth();
      monthStartX = monthEndX;
    }

    currentDate.setDate(currentDate.getDate() + 1);
  }
}

/**
 * Draws task bars
 * @param {GoogleAppsScript.Slides.Slide} slide
 * @param {Array} tasks
 * @param {Object} layout
 * @param {Object} config
 * @param {Array} elements
 */
function drawTaskBars(slide, tasks, layout, config, elements) {
  let currentY = layout.chartTop;

  layout.swimlanes.forEach(swimlane => {
    swimlane.tasks.forEach(task => {
      const startDate = new Date(task.startDate);
      const endDate = new Date(task.endDate);

      const startX = layout.chartLeft + ((startDate - layout.minDate) / (1000 * 60 * 60 * 24)) * layout.pixelsPerDay;
      const endX = layout.chartLeft + ((endDate - layout.minDate) / (1000 * 60 * 60 * 24)) * layout.pixelsPerDay;
      const barWidth = Math.max(endX - startX, 5); // Minimum 5pt width

      // Determine bar color
      const color = getTaskColor(task, config);

      // Draw task bar
      let shape;
      if (task.taskType === 'Milestone') {
        // Draw diamond for milestones
        shape = slide.insertShape(
          SlidesApp.ShapeType.DIAMOND,
          startX - 10,
          currentY,
          20,
          layout.barHeight
        );
      } else {
        // Draw rectangle for regular tasks
        shape = slide.insertShape(
          SlidesApp.ShapeType.RECTANGLE,
          startX,
          currentY,
          barWidth,
          layout.barHeight
        );
      }

      shape.getFill().setSolidFill(color);
      shape.getBorder().setWeight(1);
      shape.getBorder().getLineFill().setSolidFill('#333333');

      // Add task name
      const textRange = shape.getText();
      textRange.setText(task.taskName || '');
      textRange.getTextStyle().setFontSize(10);
      textRange.getTextStyle().setForegroundColor('#FFFFFF');

      // Add progress bar if applicable
      if (config.showProgress && task.percentComplete > 0 && task.taskType !== 'Milestone') {
        const progressWidth = barWidth * (task.percentComplete / 100);
        const progressBar = slide.insertShape(
          SlidesApp.ShapeType.RECTANGLE,
          startX,
          currentY + layout.barHeight - 4,
          progressWidth,
          4
        );
        progressBar.getFill().setSolidFill('#4CAF50');
        progressBar.getBorder().setWeight(0);

        elements.push(progressBar);
      }

      elements.push(shape);

      currentY += layout.barHeight + layout.barSpacing;
    });
  });
}

/**
 * Draws the today marker
 * @param {GoogleAppsScript.Slides.Slide} slide
 * @param {Object} layout
 * @param {Array} elements
 */
function drawTodayMarker(slide, layout, elements) {
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  if (today >= layout.minDate && today <= layout.maxDate) {
    const x = layout.chartLeft + ((today - layout.minDate) / (1000 * 60 * 60 * 24)) * layout.pixelsPerDay;

    const line = slide.insertLine(
      SlidesApp.LineCategory.STRAIGHT,
      x, layout.chartTop,
      x, layout.chartTop + layout.chartHeight
    );

    line.getLineFill().setSolidFill('#FF0000');
    line.setWeight(2);
    line.setDashStyle(SlidesApp.DashStyle.DASH);

    elements.push(line);
  }
}

/**
 * Draws dependency arrows
 * @param {GoogleAppsScript.Slides.Slide} slide
 * @param {Array} tasks
 * @param {Object} layout
 * @param {Array} elements
 */
function drawDependencies(slide, tasks, layout, elements) {
  // Build task position map
  const taskPositions = {};
  let currentY = layout.chartTop;

  layout.swimlanes.forEach(swimlane => {
    swimlane.tasks.forEach(task => {
      const startDate = new Date(task.startDate);
      const endDate = new Date(task.endDate);

      const startX = layout.chartLeft + ((startDate - layout.minDate) / (1000 * 60 * 60 * 24)) * layout.pixelsPerDay;
      const endX = layout.chartLeft + ((endDate - layout.minDate) / (1000 * 60 * 60 * 24)) * layout.pixelsPerDay;

      taskPositions[task.taskId] = {
        startX: startX,
        endX: endX,
        y: currentY + (layout.barHeight / 2)
      };

      currentY += layout.barHeight + layout.barSpacing;
    });
  });

  // Draw arrows
  tasks.forEach(task => {
    if (!task.dependencies) return;

    const deps = task.dependencies.split(',').map(d => d.trim());
    const taskPos = taskPositions[task.taskId];

    if (!taskPos) return;

    deps.forEach(depId => {
      const depPos = taskPositions[depId];
      if (!depPos) return;

      // Draw arrow from end of dependency to start of task
      const line = slide.insertLine(
        SlidesApp.LineCategory.BENT,
        depPos.endX,
        depPos.y,
        taskPos.startX,
        taskPos.y
      );

      line.getLineFill().setSolidFill('#999999');
      line.setWeight(2);
      line.setEndArrow(SlidesApp.ArrowStyle.FILL_ARROW);

      elements.push(line);
    });
  });
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Gets the color for a task based on configuration
 * @param {Object} task - The task
 * @param {Object} config - The configuration
 * @returns {string} Hex color code
 */
function getTaskColor(task, config) {
  const colorBy = config.colorBy || 'priority';

  if (colorBy === 'priority') {
    const priorityColors = {
      'High': '#D32F2F',
      'Medium': '#FFA000',
      'Low': '#388E3C'
    };
    return priorityColors[task.priority] || '#757575';
  } else if (colorBy === 'taskType') {
    return task.taskType === 'Milestone' ? '#9C27B0' : '#1976D2';
  } else if (colorBy === 'swimlane') {
    // Generate color from swimlane name
    return stringToColor(task.swimlane || '');
  }

  return '#1976D2'; // Default blue
}

/**
 * Generates a color from a string (for consistent swimlane colors)
 * @param {string} str - The string
 * @returns {string} Hex color code
 */
function stringToColor(str) {
  let hash = 0;
  for (let i = 0; i < str.length; i++) {
    hash = str.charCodeAt(i) + ((hash << 5) - hash);
  }

  const hue = hash % 360;
  return 'hsl(' + hue + ', 70%, 50%)';
}
