import { useState } from 'react';
import { useParams, Link, useNavigate } from 'react-router-dom';
import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query';
import { projectApi, taskApi } from '../services/api';
import { LoadingSpinner } from './LoadingSpinner';
import {
  BarChart3,
  Plus,
  Trash2,
  Edit,
  Settings,
  ArrowLeft,
  Calendar,
  User,
  CheckCircle2,
  AlertCircle,
  Download,
} from 'lucide-react';
import { format } from 'date-fns';
import type { Task, Priority, TaskType } from '@shared/types';

export function ProjectDetailPage() {
  const { projectId } = useParams<{ projectId: string }>();
  const navigate = useNavigate();
  const queryClient = useQueryClient();
  const [showTaskModal, setShowTaskModal] = useState(false);
  const [editingTask, setEditingTask] = useState<Task | null>(null);

  const { data: project, isLoading: projectLoading } = useQuery({
    queryKey: ['project', projectId],
    queryFn: () => projectApi.get(projectId!),
    enabled: !!projectId,
  });

  const { data: tasks, isLoading: tasksLoading } = useQuery({
    queryKey: ['tasks', projectId],
    queryFn: () => taskApi.list(projectId!),
    enabled: !!projectId,
  });

  const createTaskMutation = useMutation({
    mutationFn: (task: Partial<Task>) => taskApi.create(projectId!, task),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['tasks', projectId] });
      setShowTaskModal(false);
      setEditingTask(null);
    },
  });

  const updateTaskMutation = useMutation({
    mutationFn: ({ taskId, task }: { taskId: string; task: Partial<Task> }) =>
      taskApi.update(projectId!, taskId, task),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['tasks', projectId] });
      setShowTaskModal(false);
      setEditingTask(null);
    },
  });

  const deleteTaskMutation = useMutation({
    mutationFn: (taskId: string) => taskApi.delete(projectId!, taskId),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['tasks', projectId] });
    },
  });

  if (projectLoading || tasksLoading) {
    return (
      <div className="flex items-center justify-center py-12">
        <LoadingSpinner size="lg" />
      </div>
    );
  }

  if (!project) {
    return (
      <div className="text-center py-12">
        <AlertCircle className="w-16 h-16 text-gray-300 mx-auto mb-4" />
        <h3 className="text-lg font-medium text-gray-900 mb-2">Project not found</h3>
        <Link to="/projects" className="text-primary-600 hover:underline">
          Return to projects
        </Link>
      </div>
    );
  }

  const priorityColors: Record<Priority, string> = {
    High: 'bg-error-100 text-error-700',
    Medium: 'bg-warning-100 text-warning-700',
    Low: 'bg-success-100 text-success-700',
  };

  return (
    <div>
      {/* Header */}
      <div className="flex items-center gap-4 mb-6">
        <Link
          to="/projects"
          className="p-2 text-gray-500 hover:text-gray-700 hover:bg-gray-100 rounded-lg transition-colors"
        >
          <ArrowLeft className="w-5 h-5" />
        </Link>
        <div className="flex-1">
          <h1 className="text-2xl font-bold text-gray-900">{project.name}</h1>
          {project.description && (
            <p className="text-gray-600 mt-1">{project.description}</p>
          )}
        </div>
        <Link
          to={`/projects/${projectId}/chart`}
          className="btn btn-primary"
        >
          <BarChart3 className="w-5 h-5 mr-2" />
          View Gantt Chart
        </Link>
      </div>

      {/* Stats */}
      <div className="grid grid-cols-1 md:grid-cols-4 gap-4 mb-8">
        <StatCard
          label="Total Tasks"
          value={tasks?.length || 0}
          icon={<CheckCircle2 className="w-5 h-5" />}
        />
        <StatCard
          label="Completed"
          value={tasks?.filter((t) => t.percentComplete === 100).length || 0}
          icon={<CheckCircle2 className="w-5 h-5 text-success-500" />}
        />
        <StatCard
          label="In Progress"
          value={tasks?.filter((t) => t.percentComplete > 0 && t.percentComplete < 100).length || 0}
          icon={<CheckCircle2 className="w-5 h-5 text-warning-500" />}
        />
        <StatCard
          label="Not Started"
          value={tasks?.filter((t) => t.percentComplete === 0).length || 0}
          icon={<CheckCircle2 className="w-5 h-5 text-gray-400" />}
        />
      </div>

      {/* Tasks section */}
      <div className="card">
        <div className="flex items-center justify-between mb-4">
          <h2 className="text-lg font-semibold text-gray-900">Tasks</h2>
          <button
            onClick={() => {
              setEditingTask(null);
              setShowTaskModal(true);
            }}
            className="btn btn-primary btn-sm"
          >
            <Plus className="w-4 h-4 mr-1" />
            Add Task
          </button>
        </div>

        {tasks && tasks.length > 0 ? (
          <div className="overflow-x-auto">
            <table className="w-full">
              <thead>
                <tr className="border-b border-gray-200">
                  <th className="text-left py-3 px-4 text-sm font-medium text-gray-500">Task</th>
                  <th className="text-left py-3 px-4 text-sm font-medium text-gray-500">Dates</th>
                  <th className="text-left py-3 px-4 text-sm font-medium text-gray-500">Owner</th>
                  <th className="text-left py-3 px-4 text-sm font-medium text-gray-500">Priority</th>
                  <th className="text-left py-3 px-4 text-sm font-medium text-gray-500">Progress</th>
                  <th className="text-right py-3 px-4 text-sm font-medium text-gray-500">Actions</th>
                </tr>
              </thead>
              <tbody>
                {tasks.map((task) => (
                  <tr key={task.id} className="border-b border-gray-100 hover:bg-gray-50">
                    <td className="py-3 px-4">
                      <div>
                        <div className="font-medium text-gray-900">{task.name}</div>
                        <div className="text-sm text-gray-500">{task.id}</div>
                      </div>
                    </td>
                    <td className="py-3 px-4">
                      <div className="text-sm">
                        <div>{format(new Date(task.startDate), 'MMM d, yyyy')}</div>
                        <div className="text-gray-500">
                          to {format(new Date(task.endDate), 'MMM d, yyyy')}
                        </div>
                      </div>
                    </td>
                    <td className="py-3 px-4">
                      <span className="text-sm text-gray-700">{task.owner || '-'}</span>
                    </td>
                    <td className="py-3 px-4">
                      <span
                        className={`inline-flex px-2 py-1 text-xs font-medium rounded ${priorityColors[task.priority]}`}
                      >
                        {task.priority}
                      </span>
                    </td>
                    <td className="py-3 px-4">
                      <div className="flex items-center gap-2">
                        <div className="w-24 h-2 bg-gray-200 rounded-full overflow-hidden">
                          <div
                            className="h-full bg-primary-500 rounded-full"
                            style={{ width: `${task.percentComplete}%` }}
                          />
                        </div>
                        <span className="text-sm text-gray-600">{task.percentComplete}%</span>
                      </div>
                    </td>
                    <td className="py-3 px-4">
                      <div className="flex items-center justify-end gap-1">
                        <button
                          onClick={() => {
                            setEditingTask(task);
                            setShowTaskModal(true);
                          }}
                          className="p-1 text-gray-500 hover:text-primary-600 hover:bg-primary-50 rounded"
                        >
                          <Edit className="w-4 h-4" />
                        </button>
                        <button
                          onClick={() => {
                            if (confirm('Delete this task?')) {
                              deleteTaskMutation.mutate(task.id);
                            }
                          }}
                          className="p-1 text-gray-500 hover:text-error-600 hover:bg-error-50 rounded"
                        >
                          <Trash2 className="w-4 h-4" />
                        </button>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        ) : (
          <div className="text-center py-8 text-gray-500">
            <p>No tasks yet. Add your first task to get started.</p>
          </div>
        )}
      </div>

      {/* Task modal */}
      {showTaskModal && (
        <TaskModal
          task={editingTask}
          onClose={() => {
            setShowTaskModal(false);
            setEditingTask(null);
          }}
          onSave={(task) => {
            if (editingTask) {
              updateTaskMutation.mutate({ taskId: editingTask.id, task });
            } else {
              createTaskMutation.mutate(task);
            }
          }}
          isLoading={createTaskMutation.isPending || updateTaskMutation.isPending}
        />
      )}
    </div>
  );
}

function StatCard({
  label,
  value,
  icon,
}: {
  label: string;
  value: number;
  icon: React.ReactNode;
}) {
  return (
    <div className="card">
      <div className="flex items-center gap-3">
        <div className="p-2 bg-gray-100 rounded-lg">{icon}</div>
        <div>
          <div className="text-2xl font-bold text-gray-900">{value}</div>
          <div className="text-sm text-gray-500">{label}</div>
        </div>
      </div>
    </div>
  );
}

function TaskModal({
  task,
  onClose,
  onSave,
  isLoading,
}: {
  task: Task | null;
  onClose: () => void;
  onSave: (task: Partial<Task>) => void;
  isLoading: boolean;
}) {
  const [formData, setFormData] = useState<Partial<Task>>({
    id: task?.id || '',
    name: task?.name || '',
    startDate: task?.startDate || new Date().toISOString().split('T')[0],
    endDate: task?.endDate || new Date().toISOString().split('T')[0],
    owner: task?.owner || '',
    percentComplete: task?.percentComplete || 0,
    priority: task?.priority || 'Medium',
    taskType: task?.taskType || 'Task',
    swimlane: task?.swimlane || '',
    dependencies: task?.dependencies || [],
  });

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    onSave(formData);
  };

  return (
    <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-4">
      <div className="bg-white rounded-xl shadow-xl max-w-lg w-full p-6 animate-slide-in max-h-[90vh] overflow-y-auto">
        <h2 className="text-xl font-semibold text-gray-900 mb-4">
          {task ? 'Edit Task' : 'Add Task'}
        </h2>
        <form onSubmit={handleSubmit}>
          <div className="space-y-4">
            <div className="grid grid-cols-2 gap-4">
              <div className="col-span-2">
                <label className="label">Task ID</label>
                <input
                  type="text"
                  className="input"
                  value={formData.id}
                  onChange={(e) => setFormData({ ...formData, id: e.target.value })}
                  placeholder="e.g., TASK-001"
                />
              </div>
              <div className="col-span-2">
                <label className="label">Task Name *</label>
                <input
                  type="text"
                  className="input"
                  value={formData.name}
                  onChange={(e) => setFormData({ ...formData, name: e.target.value })}
                  placeholder="Enter task name"
                  required
                />
              </div>
              <div>
                <label className="label">Start Date *</label>
                <input
                  type="date"
                  className="input"
                  value={formData.startDate}
                  onChange={(e) => setFormData({ ...formData, startDate: e.target.value })}
                  required
                />
              </div>
              <div>
                <label className="label">End Date *</label>
                <input
                  type="date"
                  className="input"
                  value={formData.endDate}
                  onChange={(e) => setFormData({ ...formData, endDate: e.target.value })}
                  required
                />
              </div>
              <div>
                <label className="label">Owner</label>
                <input
                  type="text"
                  className="input"
                  value={formData.owner}
                  onChange={(e) => setFormData({ ...formData, owner: e.target.value })}
                  placeholder="Assignee name"
                />
              </div>
              <div>
                <label className="label">Progress (%)</label>
                <input
                  type="number"
                  className="input"
                  min={0}
                  max={100}
                  value={formData.percentComplete}
                  onChange={(e) =>
                    setFormData({ ...formData, percentComplete: parseInt(e.target.value) || 0 })
                  }
                />
              </div>
              <div>
                <label className="label">Priority</label>
                <select
                  className="input"
                  value={formData.priority}
                  onChange={(e) =>
                    setFormData({ ...formData, priority: e.target.value as Priority })
                  }
                >
                  <option value="High">High</option>
                  <option value="Medium">Medium</option>
                  <option value="Low">Low</option>
                </select>
              </div>
              <div>
                <label className="label">Task Type</label>
                <select
                  className="input"
                  value={formData.taskType}
                  onChange={(e) =>
                    setFormData({ ...formData, taskType: e.target.value as TaskType })
                  }
                >
                  <option value="Task">Task</option>
                  <option value="Milestone">Milestone</option>
                </select>
              </div>
              <div className="col-span-2">
                <label className="label">Category/Swimlane</label>
                <input
                  type="text"
                  className="input"
                  value={formData.swimlane}
                  onChange={(e) => setFormData({ ...formData, swimlane: e.target.value })}
                  placeholder="e.g., Development, Design"
                />
              </div>
              <div className="col-span-2">
                <label className="label">Dependencies (comma-separated Task IDs)</label>
                <input
                  type="text"
                  className="input"
                  value={formData.dependencies?.join(', ')}
                  onChange={(e) =>
                    setFormData({
                      ...formData,
                      dependencies: e.target.value.split(',').map((d) => d.trim()).filter(Boolean),
                    })
                  }
                  placeholder="e.g., TASK-001, TASK-002"
                />
              </div>
            </div>
          </div>
          <div className="flex justify-end gap-3 mt-6">
            <button type="button" onClick={onClose} className="btn btn-secondary">
              Cancel
            </button>
            <button type="submit" disabled={isLoading} className="btn btn-primary">
              {isLoading ? (
                <>
                  <LoadingSpinner size="sm" className="mr-2" />
                  Saving...
                </>
              ) : (
                'Save Task'
              )}
            </button>
          </div>
        </form>
      </div>
    </div>
  );
}
