# Visual Gantt - GCP Edition

A professional project timeline visualization tool built on Google Cloud Platform.

## Architecture

```
gcp/
├── backend/                 # Cloud Functions backend
│   ├── src/
│   │   ├── functions/      # API endpoint handlers
│   │   │   ├── projects.ts # Project CRUD operations
│   │   │   ├── tasks.ts    # Task management
│   │   │   ├── jira.ts     # Jira integration
│   │   │   ├── smartsheet.ts # Smartsheet integration
│   │   │   └── chart.ts    # Chart data & export
│   │   ├── services/       # Shared services
│   │   │   └── auth.ts     # Firebase Auth middleware
│   │   └── index.ts        # Express app & Cloud Functions entry
│   ├── package.json
│   └── tsconfig.json
├── frontend/               # React web application
│   ├── src/
│   │   ├── components/     # React components
│   │   ├── hooks/          # Custom React hooks
│   │   ├── services/       # API client & Firebase
│   │   └── utils/          # Utility functions
│   ├── package.json
│   ├── vite.config.ts
│   └── tailwind.config.js
├── shared/                 # Shared TypeScript types
│   └── types.ts
├── firebase.json           # Firebase configuration
├── firestore.rules        # Firestore security rules
├── firestore.indexes.json # Firestore indexes
└── storage.rules          # Cloud Storage security rules
```

## Features

- **Project Management**: Create and manage multiple projects
- **Task Management**: Full CRUD for tasks with dependencies
- **Gantt Chart Visualization**: Canvas-based rendering with zoom/pan
- **Jira Integration**: Bidirectional sync with Jira Cloud
- **Smartsheet Integration**: Bidirectional sync with Smartsheet
- **Export Options**: PNG and PDF export
- **Team Collaboration**: Share projects with team members

## Prerequisites

- Node.js 20+
- Firebase CLI (`npm install -g firebase-tools`)
- Google Cloud Platform account with billing enabled
- Firebase project with:
  - Authentication (Google provider)
  - Firestore
  - Cloud Storage
  - Cloud Functions

## Setup

### 1. Create Firebase Project

```bash
# Login to Firebase
firebase login

# Create a new project (or use existing)
firebase projects:create your-project-id

# Initialize Firebase in this directory
firebase init
```

### 2. Configure Environment Variables

**Backend** (`backend/.env`):
```bash
# No additional env vars needed - uses Firebase Admin SDK
```

**Frontend** (`frontend/.env`):
```bash
# Copy from frontend/.env.example and fill in values
cp frontend/.env.example frontend/.env
```

Get Firebase config from: Firebase Console > Project Settings > Your apps > Web app

### 3. Install Dependencies

```bash
# Backend
cd backend
npm install

# Frontend
cd ../frontend
npm install
```

### 4. Deploy Firestore Rules & Indexes

```bash
firebase deploy --only firestore:rules
firebase deploy --only firestore:indexes
firebase deploy --only storage:rules
```

## Development

### Run Locally with Emulators

```bash
# Start all emulators
firebase emulators:start

# In another terminal, start frontend dev server
cd frontend
npm run dev
```

The frontend will be available at `http://localhost:3000`

### Build for Production

```bash
# Build backend
cd backend
npm run build

# Build frontend
cd ../frontend
npm run build
```

## Deployment

### Deploy Everything

```bash
firebase deploy
```

### Deploy Individual Components

```bash
# Deploy only Cloud Functions
firebase deploy --only functions

# Deploy only Hosting (frontend)
firebase deploy --only hosting

# Deploy only Firestore rules
firebase deploy --only firestore:rules
```

## API Endpoints

### Projects
- `GET /projects` - List all projects
- `GET /projects/:projectId` - Get project
- `POST /projects` - Create project
- `PUT /projects/:projectId` - Update project
- `DELETE /projects/:projectId` - Delete project
- `PUT /projects/:projectId/config` - Update project config

### Tasks
- `GET /projects/:projectId/tasks` - List tasks
- `GET /projects/:projectId/tasks/:taskId` - Get task
- `POST /projects/:projectId/tasks` - Create task
- `POST /projects/:projectId/tasks/bulk` - Bulk create tasks
- `PUT /projects/:projectId/tasks/:taskId` - Update task
- `DELETE /projects/:projectId/tasks/:taskId` - Delete task
- `POST /projects/:projectId/tasks/validate` - Validate tasks

### Jira Integration
- `POST /projects/:projectId/jira/test` - Test connection
- `GET /projects/:projectId/jira/config` - Get config
- `PUT /projects/:projectId/jira/config` - Save config
- `POST /projects/:projectId/jira/pull` - Pull from Jira
- `POST /projects/:projectId/jira/push` - Push to Jira
- `POST /projects/:projectId/jira/conflicts` - Detect conflicts
- `POST /projects/:projectId/jira/resolve` - Resolve conflicts

### Smartsheet Integration
- `POST /projects/:projectId/smartsheet/test` - Test connection
- `GET /projects/:projectId/smartsheet/config` - Get config
- `PUT /projects/:projectId/smartsheet/config` - Save config
- `GET /projects/:projectId/smartsheet/sheets` - List sheets
- `GET /projects/:projectId/smartsheet/sheets/:sheetId/columns` - Get columns
- `POST /projects/:projectId/smartsheet/pull` - Pull from Smartsheet
- `POST /projects/:projectId/smartsheet/push` - Push to Smartsheet

### Chart
- `GET /projects/:projectId/chart` - Get chart data
- `POST /projects/:projectId/chart/export/png` - Export as PNG
- `POST /projects/:projectId/chart/export/pdf` - Export as PDF

## Data Model

### Project
```typescript
{
  id: string;
  name: string;
  description?: string;
  ownerId: string;
  members: ProjectMember[];
  createdAt: string;
  updatedAt: string;
  config: ProjectConfig;
  jiraConfig?: JiraConfig;
  smartsheetConfig?: SmartsheetConfig;
}
```

### Task
```typescript
{
  id: string;
  name: string;
  startDate: string;
  endDate: string;
  duration?: number;
  owner?: string;
  percentComplete: number;
  priority: 'High' | 'Medium' | 'Low';
  taskType: 'Task' | 'Milestone';
  swimlane?: string;
  dependencies: string[];
  parentTask?: string;
  jiraTicket?: string;
  smartsheetRowId?: string;
}
```

## Migration from Google Sheets Add-on

To migrate existing data from the Google Sheets add-on:

1. Export your task data from Google Sheets as CSV
2. Create a new project in Visual Gantt GCP
3. Use the bulk import API to import tasks:

```javascript
// Example bulk import
const tasks = csvData.map(row => ({
  id: row.taskId,
  name: row.taskName,
  startDate: row.startDate,
  endDate: row.endDate,
  owner: row.owner,
  percentComplete: row.percentComplete,
  priority: row.priority,
  taskType: row.taskType,
  swimlane: row.swimlane,
  dependencies: row.dependencies?.split(',') || [],
}));

await fetch(`/api/projects/${projectId}/tasks/bulk`, {
  method: 'POST',
  headers: {
    'Content-Type': 'application/json',
    'Authorization': `Bearer ${idToken}`,
  },
  body: JSON.stringify({ tasks, replaceAll: true }),
});
```

## License

MIT
