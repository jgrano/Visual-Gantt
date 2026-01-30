# Visual Gantt for Google Slides

Complete integration guide for inserting Gantt charts into Google Slides presentations with Google Drive storage.

## Overview

Visual Gantt now supports **Google Slides** with three distinct insertion methods:

1. **Insert as Image** - High-quality static images
2. **Link to Sheet Chart** - Auto-updating from source Google Sheets
3. **Native Shapes** - Fully interactive, editable, and animatable

All project data is stored on **Google Drive**, enabling seamless access from both Sheets and Slides.

---

## Architecture

### Data Storage (Google Drive)

```
Google Drive/
├── Visual Gantt Settings/
│   └── config.json              # Global settings
└── Visual Gantt Projects/
    ├── project-001.json         # Project data + metadata
    ├── project-002.json
    └── ...
```

**Benefits:**
- ✅ Data decoupled from Sheets/Slides
- ✅ Projects accessible from any presentation
- ✅ Version history (Drive maintains revisions)
- ✅ Shareable project files
- ✅ No permission issues (unlike PropertiesService in Workspace)

### New Files

| File | Purpose |
|------|---------|
| `DriveStorage.gs` | Google Drive project data management |
| `SlidesIntegration.gs` | Slides add-on entry points & core logic |
| `SlidesShapesRenderer.gs` | Native shapes rendering engine |
| `SlidesSidebar.html` | Slides add-on sidebar UI |
| `ProjectSelector.html` | Drive project picker dialog |
| `appsscript.json` | Updated with Slides scopes |

---

## Installation

### 1. Update Existing Add-on

If you already have Visual Gantt for Sheets:

1. Open your Apps Script project
2. Add the new files:
   - `DriveStorage.gs`
   - `SlidesIntegration.gs`
   - `SlidesShapesRenderer.gs`
   - `SlidesSidebar.html`
   - `ProjectSelector.html`
3. Replace `appsscript.json` with the updated version
4. Save and deploy

### 2. Grant New Permissions

The add-on now requires:

```json
"oauthScopes": [
  "https://www.googleapis.com/auth/spreadsheets.currentonly",
  "https://www.googleapis.com/auth/presentations",  // NEW
  "https://www.googleapis.com/auth/script.container.ui",
  "https://www.googleapis.com/auth/drive.file",
  "https://www.googleapis.com/auth/script.external_request"
]
```

Users will be prompted to authorize Slides access on first use.

---

## Usage

### From Google Sheets

#### Export Project to Drive

1. Open your project in Google Sheets
2. Go to **Visual Gantt > Export to Drive**
3. Project is saved to `Google Drive/Visual Gantt Projects/`

Or use the API:

```javascript
function exportMyProject() {
  const result = exportSheetToProject();
  Logger.log(result.projectId); // "project-abc123"
}
```

### From Google Slides

#### 1. Open the Add-on

1. Open Google Slides
2. Go to **Extensions > Visual Gantt**
3. Sidebar appears

#### 2. Select a Project

**Option A: Browse Drive Projects**
- Click "Browse Projects"
- Select from list of existing projects
- Click "Select Project"

**Option B: Connect to Sheet**
- Click "Connect to Sheet"
- Enter Sheet URL or ID
- Timeline must be generated in Sheet first

#### 3. Choose Insertion Method

Click one of three approaches:

---

## Approach 1: Insert as Image

### Best For
- Final presentations
- Static snapshots
- Maximum quality
- No dependencies

### Features
- ✅ High-resolution image
- ✅ Presentation-ready
- ✅ No external links
- ✅ Fast rendering

### How It Works

1. Select "📸 Insert as Image"
2. Click "Insert Timeline"
3. Image is inserted on current slide
4. Auto-sized and centered

**Updates:** Manual (re-insert to update)

### API Usage

```javascript
function insertImageExample() {
  // From Slides
  const result = insertTimelineAsImage();

  if (result.success) {
    Logger.log('Inserted: ' + result.projectName);
  }
}
```

### Under the Hood

```
Project Data (Drive)
  ↓
Chart Renderer (Canvas/HTML5)
  ↓
Image Blob (PNG)
  ↓
Slide.insertImage()
```

---

## Approach 2: Link to Sheet Chart

### Best For
- Live data presentations
- Auto-updating timelines
- Collaboration
- Source traceability

### Features
- ✅ Links to source Sheet
- ✅ Click to refresh
- ✅ Shows latest data
- ✅ Visual indication of source

### How It Works

1. Connect to a Google Sheet (must have generated timeline)
2. Select "🔗 Link to Sheet"
3. Click "Insert Timeline"
4. Image from Sheet's "Timeline View" is inserted
5. Click image → Links to source Sheet

**Updates:** Click "Refresh Linked Charts" in sidebar

### API Usage

```javascript
// Connect to Sheet
function connectExample() {
  const sheetUrl = 'https://docs.google.com/spreadsheets/d/abc123...';
  const result = connectToSheet(sheetUrl);

  Logger.log('Connected to: ' + result.sheetName);
}

// Insert linked chart
function insertLinkedExample() {
  const result = insertLinkedSheetChart();

  if (result.success) {
    Logger.log('Inserted from: ' + result.sheetUrl);
  }
}

// Refresh all linked charts on current slide
function refreshExample() {
  const result = refreshLinkedCharts();

  Logger.log('Refreshed ' + result.count + ' charts');
}
```

### Refresh Workflow

```
User clicks "Refresh"
  ↓
Find all linked images on slide
  ↓
Extract Sheet ID from link URL
  ↓
Fetch latest timeline image from Sheet
  ↓
Replace image (preserve position/size)
```

---

## Approach 3: Native Shapes

### Best For
- Interactive presentations
- Custom animations
- Manual adjustments
- Creative flexibility

### Features
- ✅ Each task is a separate shape
- ✅ Editable directly in Slides
- ✅ Apply animations to individual tasks
- ✅ Rearrange, recolor, resize
- ✅ No external dependencies

### How It Works

1. Select "✨ Native Shapes"
2. Click "Insert Timeline"
3. Chart is rendered as ~50-200+ individual shapes:
   - Task bars (rectangles)
   - Milestones (diamonds)
   - Headers (text boxes)
   - Grid lines
   - Dependency arrows
   - Progress bars

**Updates:** Edit shapes directly in Slides

### API Usage

```javascript
function insertShapesExample() {
  const result = insertTimelineAsShapes();

  if (result.success) {
    Logger.log('Created ' + result.shapeCount + ' shapes');
    Logger.log('Tasks: ' + result.taskCount);
  }
}
```

### Animation Example

After inserting as shapes:

1. Select a task bar shape
2. **Insert > Animation**
3. Choose animation (Fade in, Fly in, etc.)
4. Set trigger (On click, After previous, etc.)

**Result:** Build your timeline progressively during presentation!

### Shape Types Created

| Shape | Type | Purpose |
|-------|------|---------|
| Task bars | Rectangle | Individual tasks |
| Milestones | Diamond | Key events |
| Progress bars | Small rectangle | % completion |
| Swimlane headers | Rectangle + text | Category labels |
| Date headers | Rectangle + text | Month labels |
| Grid lines | Lines | Week separators |
| Today marker | Dashed line | Current date |
| Dependencies | Bent arrows | Task relationships |

### Customization Examples

```javascript
// After insertion, access shapes via Slides UI or API

// Change task color
function colorizeTask() {
  const slide = SlidesApp.getActivePresentation().getSelection().getCurrentPage();
  const shapes = slide.getShapes();

  shapes.forEach(shape => {
    const text = shape.getText().asString();
    if (text.includes('Critical Task')) {
      shape.getFill().setSolidFill('#FF0000'); // Make it red
    }
  });
}

// Animate by swimlane
function animateBySwimlane() {
  // 1. Select all shapes in "Engineering" swimlane
  // 2. Insert > Animation > Fade In
  // 3. Set trigger: "After previous"
  // Result: Engineering tasks appear together
}
```

---

## Comparison Table

| Feature | Image | Linked Sheet | Native Shapes |
|---------|-------|--------------|---------------|
| **Quality** | ★★★★★ | ★★★★☆ | ★★★☆☆ |
| **Updates** | Manual | Click refresh | Edit in Slides |
| **Animations** | No | No | Yes ★★★★★ |
| **Editability** | No | No | Yes ★★★★★ |
| **File Size** | Small | Small | Large (many shapes) |
| **Dependencies** | None | Needs source Sheet | None |
| **Best Use** | Final decks | Live dashboards | Interactive demos |

---

## Workflows

### Workflow 1: Sheets → Slides Presentation

**Use Case:** Project manager presents weekly status

1. **In Sheets:**
   - Update project data (dates, progress, etc.)
   - Visual Gantt > Export to Drive

2. **In Slides:**
   - Open presentation
   - Visual Gantt sidebar > Browse Projects
   - Select updated project
   - "Insert as Image"

**Result:** Latest timeline in presentation

---

### Workflow 2: Auto-Updating Dashboard

**Use Case:** Executive dashboard that stays current

1. **In Sheets:**
   - Maintain project data (updated daily)
   - Generate timeline (manual or triggered)

2. **In Slides:**
   - Create dashboard presentation
   - "Link to Sheet Chart"
   - Share presentation

3. **Before Presenting:**
   - Click "Refresh Linked Charts"
   - Dashboard shows latest data

**Result:** Always-current dashboard without manual updates

---

### Workflow 3: Animated Reveal

**Use Case:** Reveal project plan step-by-step

1. **In Slides:**
   - "Insert as Shapes"
   - Select swimlane 1 shapes → Add "Fade In" animation
   - Select swimlane 2 shapes → Add "Fade In" after previous
   - Continue for all swimlanes

2. **During Presentation:**
   - Click through to reveal each team's tasks
   - Narrate as each appears

**Result:** Engaging, progressive reveal of complex timeline

---

## Drive Storage API

### Save Project

```javascript
function saveExample() {
  const projectData = {
    name: 'Q4 Product Launch',
    tasks: [
      {
        taskId: 'TASK-001',
        taskName: 'Design Phase',
        startDate: '2025-10-01',
        endDate: '2025-10-15',
        owner: 'Alice',
        percentComplete: 50,
        priority: 'High'
      },
      // ... more tasks
    ],
    config: getConfig() // Optional
  };

  const result = saveProjectToDrive(projectData);
  Logger.log('Project ID: ' + result.projectId);
  Logger.log('Drive URL: ' + result.url);
}
```

### Load Project

```javascript
function loadExample() {
  const projectData = loadProjectFromDrive('project-abc123');

  if (projectData) {
    Logger.log('Loaded: ' + projectData.name);
    Logger.log('Tasks: ' + projectData.tasks.length);
  }
}
```

### List Projects

```javascript
function listExample() {
  const projects = listProjectsInDrive();

  projects.forEach(project => {
    Logger.log(project.name + ' (' + project.taskCount + ' tasks)');
  });
}
```

### Import to Sheet

```javascript
function importExample() {
  const result = importProjectToSheet('project-abc123');

  if (result.success) {
    Logger.log('Imported ' + result.taskCount + ' tasks');
  }
}
```

---

## Advanced Customization

### Custom Layout (Shapes Renderer)

Edit `SlidesShapesRenderer.gs` → `calculateLayout()`:

```javascript
// Increase bar height
const barHeight = 32; // Default: 24

// Wider swimlane headers
const swimlaneHeaderWidth = 150; // Default: 120

// More padding
const margin = 30; // Default: 20
```

### Custom Colors

Edit `SlidesShapesRenderer.gs` → `getTaskColor()`:

```javascript
function getTaskColor(task, config) {
  // Custom color scheme
  if (task.project === 'C1') return '#FF5722'; // Orange
  if (task.project === 'C2') return '#2196F3'; // Blue
  if (task.project === 'Vue') return '#4CAF50'; // Green

  return '#757575'; // Default gray
}
```

### Filter Tasks

Before rendering, filter tasks:

```javascript
function insertFilteredTimeline() {
  const projectData = loadProjectFromDrive('project-abc123');

  // Only show high-priority tasks
  projectData.tasks = projectData.tasks.filter(t => t.priority === 'High');

  const slide = SlidesApp.getActivePresentation().getSelection().getCurrentPage();
  renderGanttAsShapes(slide, projectData);
}
```

---

## Troubleshooting

### "Please select a project first"

**Cause:** No project selected in sidebar
**Fix:** Click "Browse Projects" and select one

### "Could not load project data"

**Cause:** Project ID is invalid or file deleted
**Fix:** Browse projects again and re-select

### "Timeline View sheet not found"

**Cause:** Linked Sheet doesn't have generated timeline
**Fix:** Open Sheet → Visual Gantt > Generate Timeline

### Shapes rendering slowly

**Cause:** Too many tasks (>100)
**Fix:** Filter tasks or increase `barHeight` to reduce vertical space

### Refresh not updating

**Cause:** Sheet hasn't been regenerated
**Fix:** Open source Sheet → Generate new timeline → Refresh in Slides

### Permission denied

**Cause:** Missing OAuth scopes
**Fix:** Visual Gantt > Authorization > Request Authorization

---

## Performance

| Approach | Tasks | Render Time | File Size Impact |
|----------|-------|-------------|------------------|
| Image | 50 | <2s | +100 KB |
| Image | 200 | <5s | +200 KB |
| Linked | 50 | <2s | +100 KB |
| Linked | 200 | <5s | +200 KB |
| Shapes | 50 | ~10s | +500 KB (50 shapes) |
| Shapes | 200 | ~40s | +2 MB (200+ shapes) |

**Recommendation:** Use shapes for <100 tasks; use images for larger projects.

---

## Migration Guide

### Existing Users (Sheets-only)

1. **Backup** your current spreadsheet
2. **Update** Apps Script with new files
3. **Reauthorize** when prompted (new Slides scope)
4. **Export** existing project: Visual Gantt > Export to Drive
5. **Test** in Slides: Open any presentation → Extensions > Visual Gantt

Your existing Sheets workflow remains unchanged!

---

## FAQ

### Q: Can I use Slides without Sheets?

**A:** Yes! Projects stored on Drive can be accessed from Slides without opening Sheets. However, you need Sheets to initially create projects.

### Q: Do linked charts update automatically?

**A:** No. Click "Refresh Linked Charts" to fetch the latest data. This is intentional to avoid breaking presentations mid-edit.

### Q: Can I edit shapes after insertion?

**A:** Yes! Shapes approach gives you full control. Right-click any shape → Format options.

### Q: Can I insert multiple projects on one slide?

**A:** Yes! Insert each project separately. Position and size them manually.

### Q: Does this work with Google Workspace?

**A:** Yes! Drive storage solves permission issues common with PropertiesService in Workspace domains.

### Q: Can I share projects between users?

**A:** Yes! Share the project JSON file from `Visual Gantt Projects` folder. Recipient imports via "Browse Projects" dialog.

---

## Roadmap

Potential future enhancements:

- [ ] Auto-refresh linked charts (on presentation open)
- [ ] Export animations template
- [ ] Shape grouping (when Apps Script API supports it)
- [ ] Template gallery (pre-built layouts)
- [ ] Real-time collaboration on project files
- [ ] Slides → Drive → Sheets sync (bidirectional)

---

## Support

- **Issues:** GitHub repository
- **Documentation:** This file + `README.md`
- **API Reference:** Inline JSDoc comments in `.gs` files

---

## License

MIT License (same as main project)

---

Built with ❤️ for project managers and engineers
