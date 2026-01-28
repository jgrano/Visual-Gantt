# Visual Gantt - Google Sheets Timeline Add-on

A professional Gantt chart timeline add-on for Google Sheets that creates presentation-ready project timelines with swimlanes, milestones, dependencies, and progress tracking.

## Features

- **Professional Timeline Visualization**: Clean, presentation-ready Gantt charts
- **Swimlane Grouping**: Group tasks by team, category, or owner
- **Milestone Markers**: Diamond-shaped markers for key project events
- **Progress Tracking**: Visual progress bars showing % completion
- **Dependency Arrows**: Show task relationships and prerequisites
- **Today Marker**: Vertical line indicating the current date
- **Color Coding**: Color by priority, task type, or swimlane
- **Hierarchical Tasks**: Parent/child task relationships with indentation
- **Export Options**: Save to sheet, export as PNG or PDF
- **Jira Integration**: Clickable Jira ticket links

## Installation

### Method 1: Copy to Your Google Sheet

1. Open a new or existing Google Spreadsheet
2. Go to **Extensions > Apps Script**
3. Delete any existing code in `Code.gs`
4. Copy the contents of each file from this repository:
   - `Code.gs` - Main script
   - `ChartRenderer.html` - Timeline rendering
   - `ConfigDialog.html` - Settings dialog
   - `ExportDialog.html` - PNG export
   - `ExportPdfDialog.html` - PDF export
   - `HelpDialog.html` - Help documentation
5. Copy `appsscript.json` to the manifest (View > Show manifest file)
6. Save the project (Ctrl+S)
7. Refresh your spreadsheet
8. The **Visual Gantt** menu will appear

### Method 2: Deploy as Add-on (For Organizations)

1. Create a new Apps Script project
2. Copy all files from this repository
3. Click **Deploy > Test deployments** for testing
4. Click **Deploy > New deployment** for production
5. Select **Add-on** as deployment type
6. Configure scopes and submit

## Quick Start

1. **Setup Data Sheet**: Click `Visual Gantt > Setup Data Sheet`
2. **Add Sample Data** (optional): Click `Visual Gantt > Create Sample Data`
3. **Enter Your Tasks**: Fill in the Project Data sheet
4. **Generate Timeline**: Click `Visual Gantt > Generate Timeline`
5. **Save to Sheet**: Click "Save to Sheet" in the preview

## Data Sheet Format

| Column | Required | Description |
|--------|----------|-------------|
| Task ID | Yes | Unique identifier (e.g., MECH-001) |
| Task Name | Yes | Descriptive task name |
| Start Date | Yes | Task start date |
| End Date | * | Task end date (or provide Duration) |
| Duration (Days) | * | Length in days (or provide End Date) |
| Owner | No | Person responsible |
| % Complete | No | Progress (0-100) |
| Priority | No | High, Medium, or Low |
| Jira Ticket | No | URL to Jira issue |
| Parent Task | No | Task ID of parent for subtasks |
| Dependencies | No | Comma-separated prerequisite Task IDs |
| Task Type | No | Task or Milestone |
| Swimlane/Category | No | Team/category for grouping |

## Configuration Options

Access via `Visual Gantt > Configure Settings`:

### Date Range
- **Start/End Date**: Custom date range or auto-calculate from tasks

### Grouping & Colors
- **Swimlane Grouping**: By Category, Owner, or None
- **Color Scheme**: By Priority, Task Type, or Swimlane

### Display Options
- Show/hide dependencies
- Show/hide progress bars
- Show/hide milestones
- Show/hide today marker

### Layout
- Bar height (16-48px)
- Bar spacing (4-20px)
- Chart width (800-3000px)
- Chart height (400-2000px)
- Font size (10-16px)

## Example Use Case

**B1 Bidirectional Charger Development Project**

Track development across multiple engineering disciplines:
- Mechanical design and prototyping
- Electrical validation and PCB layout
- Firmware development and safety systems
- Test engineering and certification

Swimlanes group tasks by discipline, dependencies show critical path, and milestones mark customer reviews and production release.

## Export Options

1. **Save to Sheet**: Embed the timeline image in the Timeline View sheet
2. **Export as PNG**: Download high-quality image for presentations
3. **Export as PDF**: Create PDF document for printing/sharing

## Performance

- Optimized for projects with 100+ tasks
- Fast refresh (<5 seconds for typical 50-task project)
- High-resolution canvas rendering with DPI awareness

## File Structure

```
Visual-Gantt/
├── appsscript.json      # Manifest with OAuth scopes
├── Code.gs              # Main Apps Script code
├── ChartRenderer.html   # Canvas-based chart rendering
├── ConfigDialog.html    # Settings configuration UI
├── ExportDialog.html    # PNG export dialog
├── ExportPdfDialog.html # PDF export dialog
├── HelpDialog.html      # Help documentation
└── README.md            # This file
```

## OAuth Scopes

The add-on requires these permissions:
- `spreadsheets.currentonly` - Read/write current spreadsheet
- `script.container.ui` - Show dialogs and menus
- `drive.file` - Save exports to Drive

## Troubleshooting

### Timeline not generating
- Ensure data sheet is named "Project Data"
- Check that Start Date column has valid dates
- Verify End Date or Duration is provided

### Dependencies not showing
- Verify dependency Task IDs exist in the sheet
- Check for typos in Task ID references
- Enable "Show Dependencies" in settings

### Export fails
- Generate and save timeline first
- Check Drive permissions
- Try refreshing the spreadsheet

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test in Google Sheets
5. Submit a pull request

## License

MIT License - See LICENSE file for details

## Support

For issues and feature requests, please open a GitHub issue.

---

Built for professional project management in hardware engineering teams.
