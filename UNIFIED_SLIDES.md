# Visual Gantt: Unified Single-File Approach

**The simplest way to use Visual Gantt in Google Slides** - everything in one file, no external dependencies.

## Problem with Multi-File Approach

The Drive-based workflow (SLIDES_INTEGRATION.md) requires:
- Data stored in Drive
- Connecting to Sheets
- Browsing and selecting projects
- Multiple files to manage

**This felt disjointed.**

## Solution: Self-Contained Presentations

With the unified approach, **everything lives in the presentation file itself**:

```
Your Presentation.pptx
├── Slides (your content)
├── Speaker notes
└── Embedded Project Data ← All tasks stored here!
```

**Benefits:**
- ✅ **Single file** - Share the .pptx, share everything
- ✅ **No Drive/Sheets** - Completely self-contained
- ✅ **Built-in editor** - Add/edit tasks directly in Slides
- ✅ **Instant sync** - Data and charts in same file
- ✅ **Offline-ready** - No external dependencies

---

## Quick Start

### 1. Open Slides Add-on

```
Google Slides → Extensions → Visual Gantt
```

**Unified sidebar appears with 3 tabs:**
- **Tasks** - Manage project data
- **Insert** - Add timeline to slide
- **Settings** - Configure appearance

### 2. Add Tasks

**Tasks Tab:**
1. Click "+ Add Task"
2. Fill in:
   - Task Name
   - Start/End dates
   - Owner, Priority, etc.
3. Save

Repeat for all tasks. Data is saved automatically to the presentation.

### 3. Insert Timeline

**Insert Tab:**
1. Choose method:
   - 📸 **Static Image** - High-quality presentation image
   - ✨ **Interactive Shapes** - Editable, animatable tasks
2. Click "Insert Timeline"

Done! Timeline appears on your current slide.

---

## Complete Workflow

### Scenario: Project Status Presentation

**Step 1: Create Presentation**
```
File → New → Presentation
Title: "Q4 Product Launch Status"
```

**Step 2: Open Visual Gantt**
```
Extensions → Visual Gantt
```

Unified sidebar appears.

**Step 3: Add Tasks (Tasks Tab)**

Click "+ Add Task" and create:

| Task | Start | End | Owner | Priority |
|------|-------|-----|-------|----------|
| Design Phase | 2025-03-01 | 2025-03-15 | Alice | High |
| Development | 2025-03-16 | 2025-04-30 | Bob | High |
| Testing | 2025-05-01 | 2025-05-15 | Carol | Medium |
| Launch | 2025-05-20 | 2025-05-20 | Team | High |

**Tip:** Use Import CSV/JSON for bulk adds.

**Step 4: Insert Timeline (Insert Tab)**
```
Select: ✨ Interactive Shapes
Click: Insert Timeline
```

**Step 5: Animate (Optional)**
```
Select task bars for "Design Phase"
Insert → Animation → Fade In (On Click)

Select "Development" bars
Insert → Animation → Fly In (After Previous)

...repeat for all phases
```

**Step 6: Present!**
```
File → Present (Cmd/Ctrl + F5)
Click through to reveal phases progressively
```

---

## Data Storage

### Where Is Data Stored?

**Primary:** `DocumentProperties` (presentation metadata)
**Fallback:** Hidden slide with JSON in speaker notes

### Viewing Raw Data

**Method 1: Export JSON**
```
Settings Tab → Export as JSON
```

Downloads `project.json` with all task data.

**Method 2: Check Hidden Slide**

If using fallback storage:
```
View → Grid View
Look for: "[Visual Gantt Data - Do Not Delete]"
```

**Warning:** Don't delete this slide - it contains your project data!

### Data Format

```json
{
  "name": "Q4 Product Launch",
  "tasks": [
    {
      "taskId": "TASK-001",
      "taskName": "Design Phase",
      "startDate": "2025-03-01",
      "endDate": "2025-03-15",
      "owner": "Alice",
      "percentComplete": 100,
      "priority": "High",
      "taskType": "Task",
      "swimlane": "Engineering",
      "dependencies": ""
    }
  ],
  "config": {
    "swimlaneGrouping": "swimlane",
    "colorBy": "priority",
    "barHeight": 24
  },
  "metadata": {
    "created": "2025-03-01T10:00:00Z",
    "modified": "2025-03-10T15:30:00Z"
  }
}
```

---

## Features

### Task Management

**Add Task:**
```
Tasks Tab → + Add Task
```

**Edit Task:**
```
Tasks Tab → Click task → Edit in Task Manager
```

**Delete Task:**
```
Task Manager → Delete button
```

**Import Tasks:**
```
Tasks Tab → Import CSV/JSON
```

Paste CSV:
```
taskName,startDate,endDate,owner,priority
Design Phase,2025-03-01,2025-03-15,Alice,High
Development,2025-03-16,2025-04-30,Bob,High
```

Or JSON:
```json
{
  "tasks": [
    {"taskName": "Design", "startDate": "2025-03-01", "endDate": "2025-03-15"}
  ]
}
```

### Timeline Insertion

#### Option 1: Static Image 📸

**Best for:** Final presentations, print-ready decks

**Pros:**
- High quality
- No shape clutter
- Lightweight file size

**Cons:**
- Not editable after insertion
- Must re-insert to update

**How:**
```
Insert Tab → Select "Static Image" → Insert Timeline
```

#### Option 2: Interactive Shapes ✨

**Best for:** Animations, customization, live editing

**Pros:**
- Each task = separate shape
- Fully editable (move, resize, recolor)
- Apply Slides animations
- No external dependencies

**Cons:**
- Larger file size (many shapes)
- Slower rendering (50+ tasks)

**How:**
```
Insert Tab → Select "Interactive Shapes" → Insert Timeline
```

**Then animate:**
```
Select task shapes → Insert → Animation
```

### Settings

**Configure appearance:**
```
Settings Tab
```

Options:
- **Group By:** Swimlane, Owner, Priority, None
- **Color By:** Priority, Task Type, Swimlane
- **Bar Height:** Small (20px), Medium (24px), Large (32px)

Changes apply to next insertion.

---

## Comparison: Unified vs Drive-Based

| Feature | Unified (Single-File) | Drive-Based (Multi-File) |
|---------|----------------------|--------------------------|
| **Data Location** | In presentation | Google Drive |
| **Dependencies** | None | Drive + optional Sheets |
| **Sharing** | Share .pptx = done | Share Drive folder |
| **Offline** | Yes | No |
| **Task Editing** | Built-in sidebar | Sheets or Drive JSON |
| **Best For** | Self-contained decks | Multi-presentation projects |

**Use Unified when:**
- Single standalone presentation
- Sharing with external parties
- No access to Sheets
- Prefer simplicity

**Use Drive-Based when:**
- Multiple presentations use same data
- Team collaboration on project data
- Already using Sheets for planning
- Want centralized project management

---

## Advanced Usage

### Animations

**Progressive Reveal:**

1. Insert as shapes
2. Select all "Phase 1" task bars:
   ```
   Click first task bar
   Hold Shift, click other Phase 1 bars
   ```
3. Insert → Animation → Fade In
4. Set: On Click
5. Repeat for Phase 2, 3, etc.

**Result:** Click to reveal each phase!

**Critical Path Highlight:**

1. Insert as shapes
2. Find critical path tasks
3. Right-click → Format options
4. Fill color → Red
5. Add "Pulse" animation

**Result:** Critical tasks stand out!

### Duplicate Presentations

**Problem:** Want same project in multiple decks?

**Solution 1: Copy/Paste**
```
Presentation A → Extensions → Visual Gantt → Settings → Export JSON
Presentation B → Extensions → Visual Gantt → Tasks → Import JSON
```

**Solution 2: Template**
```
Create master.pptx with all tasks
File → Make a copy (for each new presentation)
```

Each copy is self-contained!

### Bulk Updates

**Update multiple tasks:**

1. Export as JSON:
   ```
   Settings Tab → Export as JSON
   ```

2. Edit in text editor:
   ```json
   // Change all dates, owners, etc.
   ```

3. Re-import:
   ```
   Tasks Tab → Import CSV/JSON → Paste edited JSON
   ```

---

## Troubleshooting

### "No tasks found"

**Cause:** No tasks added yet
**Fix:** Go to Tasks tab → + Add Task

### Timeline not appearing

**Cause:** Slide not selected
**Fix:** Click a slide in the sidebar first, then insert

### Changes not saving

**Cause:** Permission issue with DocumentProperties
**Fix:** Data automatically falls back to hidden slide. Check for slide named "[Visual Gantt Data - Do Not Delete]"

### Too many shapes (slow)

**Cause:** >100 tasks with Interactive Shapes
**Fix:** Use Static Image instead, or filter tasks

### Can't animate timeline

**Cause:** Inserted as Static Image
**Fix:** Use Interactive Shapes instead

### Lost data after sharing

**Cause:** Presentation was copied without data
**Fix:** Always use File → Make a copy (not download/re-upload)

---

## API Reference

### Task Manager Functions

```javascript
// Get all tasks
const tasks = getTasksFromPresentation();

// Add task
const result = addTaskToPresentation({
  taskName: 'New Task',
  startDate: '2025-03-01',
  endDate: '2025-03-15',
  owner: 'Alice',
  priority: 'High'
});

// Update task
updateTaskInPresentation('TASK-001', {
  percentComplete: 50,
  owner: 'Bob'
});

// Delete task
deleteTaskFromPresentation('TASK-001');
```

### Data Management

```javascript
// Save project
saveProjectToPresentation(projectData);

// Load project
const project = loadProjectFromPresentation();

// Check if data exists
const hasData = hasEmbeddedData();

// Delete all data
deleteProjectFromPresentation();
```

### Import/Export

```javascript
// Export as JSON
const json = exportProjectJSON();

// Import JSON
importProjectJSON(jsonString);

// Import CSV
importTasksFromCSV(csvString);
```

### Configuration

```javascript
// Get config
const config = getConfigFromPresentation();

// Update config
updateConfigInPresentation({
  swimlaneGrouping: 'owner',
  colorBy: 'priority',
  barHeight: 32
});
```

---

## Migration Guide

### From Sheets to Unified Slides

**You have:** Project in Google Sheets
**You want:** Self-contained Slides presentation

**Steps:**

1. **In Sheets:**
   ```
   Visual Gantt → Projects → Export to Drive
   ```

2. **Get JSON:**
   ```
   Open Drive → Visual Gantt Projects folder
   Open project-xxx.json
   Copy all content
   ```

3. **In Slides:**
   ```
   Extensions → Visual Gantt
   Tasks Tab → Import CSV/JSON
   Paste JSON → Import
   ```

4. **Insert:**
   ```
   Insert Tab → Choose method → Insert Timeline
   ```

Done! Now self-contained.

### From Drive-Based to Unified

**You have:** Drive-stored projects
**You want:** Data embedded in presentation

**Steps:**

1. **Export from Drive:**
   ```
   Already have project-xxx.json
   ```

2. **Import to Presentation:**
   ```
   Extensions → Visual Gantt (unified sidebar)
   Tasks Tab → Import CSV/JSON
   Paste JSON content
   ```

3. **(Optional) Delete from Drive:**
   ```
   No longer needed - data now in .pptx
   ```

---

## Best Practices

### File Naming

Include status in filename:
```
✓ Q4_Launch_Status_2025-03-10.pptx
✓ Engineering_Timeline_v3.pptx
✗ Presentation1.pptx
```

### Task IDs

Use consistent prefixes:
```
MECH-001, MECH-002  (Mechanical)
ELEC-001, ELEC-002  (Electrical)
FW-001, FW-002      (Firmware)
```

Makes dependencies clearer.

### Swimlanes

Group logically:
```
✓ Engineering, Design, QA, Management
✓ Frontend, Backend, DevOps, Security
✗ Random, Other, Misc
```

### Backups

**Presentations are not version-controlled!**

Backup strategies:
1. **Drive auto-save:** Already enabled
2. **Manual snapshots:**
   ```
   File → Make a copy → Add date to filename
   ```
3. **Export JSON regularly:**
   ```
   Settings Tab → Export as JSON (save externally)
   ```

---

## FAQ

### Q: Can I use both unified and Drive-based?

**A:** Yes! They're independent:
- Unified: Data in presentation
- Drive-based: Data in Drive files

Use whichever fits your workflow.

### Q: How do I switch between them?

**A:** In SlidesIntegration.gs, call:
```javascript
showUnifiedSlidesSidebar()  // Single-file approach
// OR
showSlidesSidebar()  // Drive-based approach
```

### Q: What's the file size impact?

**Embedded data:** ~5-50 KB (negligible)
**Static Image:** ~100-200 KB per timeline
**Interactive Shapes:** ~500 KB - 2 MB (many shapes)

### Q: Can I edit shapes after insertion?

**A:** Only if inserted as "Interactive Shapes". Static images are not editable.

### Q: How do I share with non-Google users?

**A:**
```
File → Download → Microsoft PowerPoint (.pptx)
```

Data is preserved! They can view timeline but not edit tasks (no add-on).

### Q: Can I use this without internet?

**A:** Yes! Once presentation is downloaded, everything works offline. Add-on requires internet for sidebar UI.

### Q: Is data secure?

**A:** Data is stored in your presentation file with same permissions as the .pptx. Only people with access to the presentation can see the data.

---

## Roadmap

Potential enhancements:
- [ ] Drag-and-drop task reordering
- [ ] Gantt chart templates gallery
- [ ] Real-time collaboration (multiple editors)
- [ ] Undo/redo for task edits
- [ ] Task filtering in sidebar
- [ ] Export to MS Project format

---

## Support

- **Issues:** GitHub repository
- **Documentation:** This file + SLIDES_INTEGRATION.md
- **Quick Help:** Settings Tab → About

---

## Summary

**Unified Single-File Approach:**

```
1. Open Slides presentation
2. Extensions → Visual Gantt
3. Tasks Tab → Add tasks
4. Insert Tab → Choose method → Insert
5. Done!
```

**Everything in one file. No external dependencies. Simple.**

---

Built for simplicity, powered by Google Apps Script.
