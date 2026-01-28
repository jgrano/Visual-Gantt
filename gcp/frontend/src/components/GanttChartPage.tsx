import { useRef, useEffect, useState } from 'react';
import { useParams, Link } from 'react-router-dom';
import { useQuery } from '@tanstack/react-query';
import { chartApi } from '../services/api';
import { LoadingSpinner } from './LoadingSpinner';
import {
  ArrowLeft,
  Download,
  Settings,
  RefreshCw,
  ZoomIn,
  ZoomOut,
  Maximize2,
} from 'lucide-react';
import type { ChartData, Task, ProjectConfig } from '@shared/types';
import { jsPDF } from 'jspdf';

// Color constants
const COLORS = {
  background: '#FFFFFF',
  grid: '#E0E0E0',
  gridLight: '#F5F5F5',
  text: '#333333',
  textLight: '#666666',
  today: '#FF5722',
  swimlaneHeader: '#F8F9FA',
  dependency: '#9E9E9E',
  progress: 'rgba(255,255,255,0.4)',
  milestone: '#9C27B0',
};

export function GanttChartPage() {
  const { projectId } = useParams<{ projectId: string }>();
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const containerRef = useRef<HTMLDivElement>(null);
  const [zoom, setZoom] = useState(1);
  const [exporting, setExporting] = useState(false);

  const { data: chartData, isLoading, error, refetch } = useQuery({
    queryKey: ['chartData', projectId],
    queryFn: () => chartApi.getData(projectId!),
    enabled: !!projectId,
  });

  // Render chart when data changes
  useEffect(() => {
    if (chartData && canvasRef.current) {
      renderChart(canvasRef.current, chartData, zoom);
    }
  }, [chartData, zoom]);

  const handleExportPng = async () => {
    if (!canvasRef.current) return;

    setExporting(true);
    try {
      const dataUrl = canvasRef.current.toDataURL('image/png');

      // Download locally
      const link = document.createElement('a');
      link.download = `visual_gantt_${projectId}.png`;
      link.href = dataUrl;
      link.click();
    } finally {
      setExporting(false);
    }
  };

  const handleExportPdf = async () => {
    if (!canvasRef.current) return;

    setExporting(true);
    try {
      const canvas = canvasRef.current;
      const imgData = canvas.toDataURL('image/png');

      // Create PDF with landscape orientation
      const pdf = new jsPDF({
        orientation: canvas.width > canvas.height ? 'landscape' : 'portrait',
        unit: 'px',
        format: [canvas.width, canvas.height],
      });

      pdf.addImage(imgData, 'PNG', 0, 0, canvas.width, canvas.height);
      pdf.save(`visual_gantt_${projectId}.pdf`);
    } finally {
      setExporting(false);
    }
  };

  if (isLoading) {
    return (
      <div className="flex items-center justify-center py-12">
        <LoadingSpinner size="lg" />
      </div>
    );
  }

  if (error || !chartData) {
    return (
      <div className="text-center py-12">
        <p className="text-error-600 mb-4">
          {error instanceof Error ? error.message : 'Failed to load chart data'}
        </p>
        <Link to={`/projects/${projectId}`} className="text-primary-600 hover:underline">
          Return to project
        </Link>
      </div>
    );
  }

  return (
    <div className="flex flex-col h-[calc(100vh-12rem)]">
      {/* Header */}
      <div className="flex items-center justify-between mb-4">
        <div className="flex items-center gap-4">
          <Link
            to={`/projects/${projectId}`}
            className="p-2 text-gray-500 hover:text-gray-700 hover:bg-gray-100 rounded-lg transition-colors"
          >
            <ArrowLeft className="w-5 h-5" />
          </Link>
          <h1 className="text-xl font-bold text-gray-900">Gantt Chart</h1>
          <span className="text-sm text-gray-500">
            {chartData.tasks.length} tasks
          </span>
        </div>

        <div className="flex items-center gap-2">
          {/* Zoom controls */}
          <div className="flex items-center gap-1 bg-gray-100 rounded-lg p-1">
            <button
              onClick={() => setZoom((z) => Math.max(0.5, z - 0.25))}
              className="p-1.5 hover:bg-white rounded transition-colors"
              title="Zoom out"
            >
              <ZoomOut className="w-4 h-4" />
            </button>
            <span className="px-2 text-sm font-medium min-w-[3rem] text-center">
              {Math.round(zoom * 100)}%
            </span>
            <button
              onClick={() => setZoom((z) => Math.min(2, z + 0.25))}
              className="p-1.5 hover:bg-white rounded transition-colors"
              title="Zoom in"
            >
              <ZoomIn className="w-4 h-4" />
            </button>
            <button
              onClick={() => setZoom(1)}
              className="p-1.5 hover:bg-white rounded transition-colors"
              title="Reset zoom"
            >
              <Maximize2 className="w-4 h-4" />
            </button>
          </div>

          <button
            onClick={() => refetch()}
            className="btn btn-secondary btn-sm"
            title="Refresh"
          >
            <RefreshCw className="w-4 h-4" />
          </button>

          <div className="relative group">
            <button
              className="btn btn-primary btn-sm"
              disabled={exporting}
            >
              {exporting ? (
                <LoadingSpinner size="sm" />
              ) : (
                <>
                  <Download className="w-4 h-4 mr-1" />
                  Export
                </>
              )}
            </button>
            <div className="absolute right-0 top-full mt-1 bg-white rounded-lg shadow-lg border border-gray-200 py-1 opacity-0 invisible group-hover:opacity-100 group-hover:visible transition-all z-10">
              <button
                onClick={handleExportPng}
                className="w-full px-4 py-2 text-left text-sm hover:bg-gray-100"
              >
                Export as PNG
              </button>
              <button
                onClick={handleExportPdf}
                className="w-full px-4 py-2 text-left text-sm hover:bg-gray-100"
              >
                Export as PDF
              </button>
            </div>
          </div>
        </div>
      </div>

      {/* Chart container */}
      <div
        ref={containerRef}
        className="flex-1 bg-white rounded-lg border border-gray-200 overflow-auto"
      >
        <canvas ref={canvasRef} className="block" />
      </div>

      {/* Legend */}
      <div className="flex flex-wrap items-center gap-4 mt-4 p-3 bg-gray-50 rounded-lg">
        {chartData.config.colorScheme === 'priority' && (
          <>
            <LegendItem color="#DC3545" label="High Priority" />
            <LegendItem color="#FFC107" label="Medium Priority" />
            <LegendItem color="#28A745" label="Low Priority" />
          </>
        )}
        {chartData.config.showMilestones && (
          <div className="flex items-center gap-2 text-sm text-gray-600">
            <div
              className="w-3 h-3 bg-purple-600 transform rotate-45"
              style={{ transform: 'rotate(45deg)' }}
            />
            <span>Milestone</span>
          </div>
        )}
        {chartData.config.showTodayMarker && (
          <div className="flex items-center gap-2 text-sm text-gray-600">
            <div className="w-4 h-0.5 bg-orange-500" />
            <span>Today</span>
          </div>
        )}
        {chartData.config.showPercentComplete && (
          <div className="flex items-center gap-2 text-sm text-gray-600">
            <div className="w-4 h-3 bg-gradient-to-r from-white/40 to-primary-500 rounded" />
            <span>% Complete</span>
          </div>
        )}
      </div>
    </div>
  );
}

function LegendItem({ color, label }: { color: string; label: string }) {
  return (
    <div className="flex items-center gap-2 text-sm text-gray-600">
      <div className="w-4 h-4 rounded" style={{ backgroundColor: color }} />
      <span>{label}</span>
    </div>
  );
}

// ============================================================================
// Canvas Rendering Functions
// ============================================================================

function renderChart(canvas: HTMLCanvasElement, data: ChartData, zoom: number): void {
  const ctx = canvas.getContext('2d');
  if (!ctx) return;

  const { tasks, config, startDate, endDate, today } = data;
  const start = new Date(startDate);
  const end = new Date(endDate);
  const todayDate = new Date(today);

  // Layout constants
  const MARGIN = { top: 80, right: 40, bottom: 40, left: 250 };
  const BAR_HEIGHT = (config.barHeight || 24) * zoom;
  const BAR_SPACING = (config.barSpacing || 8) * zoom;
  const ROW_HEIGHT = BAR_HEIGHT + BAR_SPACING;
  const FONT_SIZE = (config.fontSize || 12) * zoom;
  const HEADER_HEIGHT = 60 * zoom;

  // Calculate dimensions
  const totalDays = Math.ceil((end.getTime() - start.getTime()) / (1000 * 60 * 60 * 24));
  const chartWidth = Math.max((config.chartWidth || 1200) * zoom, totalDays * 15 * zoom);
  const chartHeight = Math.max(
    (config.chartHeight || 600) * zoom,
    MARGIN.top + tasks.length * ROW_HEIGHT + MARGIN.bottom + 100
  );

  // Set canvas size
  const dpr = window.devicePixelRatio || 1;
  canvas.width = (chartWidth + MARGIN.left + MARGIN.right) * dpr;
  canvas.height = chartHeight * dpr;
  canvas.style.width = `${chartWidth + MARGIN.left + MARGIN.right}px`;
  canvas.style.height = `${chartHeight}px`;
  ctx.scale(dpr, dpr);

  const canvasWidth = chartWidth + MARGIN.left + MARGIN.right;
  const canvasHeight = chartHeight;

  // Chart area
  const chartArea = {
    x: MARGIN.left,
    y: MARGIN.top,
    width: chartWidth,
    height: canvasHeight - MARGIN.top - MARGIN.bottom,
  };

  // Helper functions
  const dateToX = (date: Date | string): number => {
    const d = new Date(date);
    const daysDiff = (d.getTime() - start.getTime()) / (1000 * 60 * 60 * 24);
    return chartArea.x + (daysDiff / totalDays) * chartArea.width;
  };

  const taskToY = (index: number): number => {
    return chartArea.y + index * ROW_HEIGHT;
  };

  // Clear canvas
  ctx.fillStyle = COLORS.background;
  ctx.fillRect(0, 0, canvasWidth, canvasHeight);

  // Draw title
  ctx.fillStyle = COLORS.text;
  ctx.font = `bold ${18 * zoom}px 'Segoe UI', sans-serif`;
  ctx.textAlign = 'left';
  ctx.fillText('Project Timeline', MARGIN.left, 30 * zoom);

  ctx.font = `${12 * zoom}px 'Segoe UI', sans-serif`;
  ctx.fillStyle = COLORS.textLight;
  const dateRange = `${formatDate(start)} - ${formatDate(end)}`;
  ctx.fillText(dateRange, MARGIN.left, 50 * zoom);

  // Draw grid
  ctx.strokeStyle = COLORS.gridLight;
  ctx.lineWidth = 1;

  // Horizontal grid lines
  for (let i = 0; i <= tasks.length; i++) {
    const y = taskToY(i);
    ctx.beginPath();
    ctx.moveTo(chartArea.x, y);
    ctx.lineTo(chartArea.x + chartArea.width, y);
    ctx.stroke();
  }

  // Vertical grid lines (weekly)
  ctx.strokeStyle = COLORS.grid;
  const currentDate = new Date(start);
  while (currentDate <= end) {
    const x = dateToX(currentDate);
    ctx.beginPath();
    ctx.moveTo(x, chartArea.y);
    ctx.lineTo(x, chartArea.y + tasks.length * ROW_HEIGHT);
    ctx.stroke();
    currentDate.setDate(currentDate.getDate() + 7);
  }

  // Draw swimlanes if enabled
  if (config.swimlaneGrouping !== 'None') {
    drawSwimlanes(ctx, tasks, config, chartArea, canvasWidth, taskToY, FONT_SIZE, ROW_HEIGHT);
  }

  // Draw today marker
  if (config.showTodayMarker && todayDate >= start && todayDate <= end) {
    const x = dateToX(todayDate);
    ctx.strokeStyle = COLORS.today;
    ctx.lineWidth = 2;
    ctx.setLineDash([5, 5]);
    ctx.beginPath();
    ctx.moveTo(x, MARGIN.top - HEADER_HEIGHT);
    ctx.lineTo(x, chartArea.y + tasks.length * ROW_HEIGHT);
    ctx.stroke();
    ctx.setLineDash([]);

    ctx.fillStyle = COLORS.today;
    ctx.font = `bold ${(FONT_SIZE - 2)}px 'Segoe UI', sans-serif`;
    ctx.textAlign = 'center';
    ctx.fillText('Today', x, MARGIN.top - HEADER_HEIGHT - 5);
  }

  // Draw dependencies
  if (config.showDependencies) {
    drawDependencies(ctx, tasks, dateToX, taskToY, BAR_HEIGHT);
  }

  // Draw tasks
  tasks.forEach((task, index) => {
    const x = dateToX(task.startDate);
    const y = taskToY(index);
    const width = Math.max(dateToX(task.endDate) - x, 4 * zoom);

    if (task.taskType === 'Milestone') {
      drawMilestone(ctx, x, y + BAR_HEIGHT / 2, task.color || COLORS.milestone, BAR_HEIGHT);
    } else {
      drawTaskBar(ctx, x, y, width, task, config, BAR_HEIGHT, FONT_SIZE);
    }
  });

  // Draw date axis
  drawDateAxis(ctx, chartArea, start, end, totalDays, MARGIN, HEADER_HEIGHT, FONT_SIZE, zoom);

  // Draw task labels
  drawTaskLabels(ctx, tasks, chartArea, MARGIN, taskToY, FONT_SIZE, BAR_HEIGHT, ROW_HEIGHT);
}

function drawSwimlanes(
  ctx: CanvasRenderingContext2D,
  tasks: Task[],
  config: ProjectConfig,
  chartArea: { x: number; y: number; width: number; height: number },
  canvasWidth: number,
  taskToY: (index: number) => number,
  fontSize: number,
  rowHeight: number
): void {
  const groupField = config.swimlaneGrouping === 'Owner' ? 'owner' : 'swimlane';
  let currentGroup: string | null = null;
  let groupStartIndex = 0;
  const swimlaneGroups: { name: string; startIndex: number; endIndex: number }[] = [];

  tasks.forEach((task, index) => {
    const group = (task[groupField as keyof Task] as string) || 'Other';
    if (group !== currentGroup) {
      if (currentGroup !== null) {
        swimlaneGroups.push({
          name: currentGroup,
          startIndex: groupStartIndex,
          endIndex: index - 1,
        });
      }
      currentGroup = group;
      groupStartIndex = index;
    }
  });

  if (currentGroup !== null) {
    swimlaneGroups.push({
      name: currentGroup,
      startIndex: groupStartIndex,
      endIndex: tasks.length - 1,
    });
  }

  swimlaneGroups.forEach((group, idx) => {
    const y1 = taskToY(group.startIndex);
    const y2 = taskToY(group.endIndex + 1);
    const height = y2 - y1;

    if (idx % 2 === 0) {
      ctx.fillStyle = 'rgba(66, 133, 244, 0.03)';
      ctx.fillRect(0, y1, canvasWidth, height);
    }

    ctx.save();
    ctx.fillStyle = COLORS.textLight;
    ctx.font = `bold ${fontSize - 1}px 'Segoe UI', sans-serif`;
    ctx.textAlign = 'center';
    ctx.translate(15, y1 + height / 2);
    ctx.rotate(-Math.PI / 2);
    ctx.fillText(group.name, 0, 0);
    ctx.restore();

    if (idx > 0) {
      ctx.strokeStyle = 'rgba(66, 133, 244, 0.2)';
      ctx.lineWidth = 1;
      ctx.beginPath();
      ctx.moveTo(0, y1);
      ctx.lineTo(canvasWidth, y1);
      ctx.stroke();
    }
  });
}

function drawDependencies(
  ctx: CanvasRenderingContext2D,
  tasks: Task[],
  dateToX: (date: string) => number,
  taskToY: (index: number) => number,
  barHeight: number
): void {
  const taskIndexMap = new Map<string, number>();
  tasks.forEach((task, index) => taskIndexMap.set(task.id, index));

  ctx.strokeStyle = COLORS.dependency;
  ctx.lineWidth = 1.5;

  tasks.forEach((task, index) => {
    if (!task.dependencies || task.dependencies.length === 0) return;

    task.dependencies.forEach((depId) => {
      const depIndex = taskIndexMap.get(depId);
      if (depIndex === undefined) return;

      const depTask = tasks[depIndex];
      const fromX = dateToX(depTask.endDate) + 2;
      const fromY = taskToY(depIndex) + barHeight / 2;
      const toX = dateToX(task.startDate) - 2;
      const toY = taskToY(index) + barHeight / 2;

      ctx.beginPath();
      ctx.moveTo(fromX, fromY);

      const midX = (fromX + toX) / 2;
      if (toY === fromY) {
        ctx.lineTo(toX, toY);
      } else {
        ctx.bezierCurveTo(midX, fromY, midX, toY, toX, toY);
      }
      ctx.stroke();

      // Arrow head
      const arrowSize = 6;
      ctx.beginPath();
      ctx.moveTo(toX, toY);
      ctx.lineTo(toX - arrowSize, toY - arrowSize / 2);
      ctx.lineTo(toX - arrowSize, toY + arrowSize / 2);
      ctx.closePath();
      ctx.fillStyle = COLORS.dependency;
      ctx.fill();
    });
  });
}

function drawTaskBar(
  ctx: CanvasRenderingContext2D,
  x: number,
  y: number,
  width: number,
  task: Task,
  config: ProjectConfig,
  barHeight: number,
  fontSize: number
): void {
  const radius = 4;

  // Main bar
  ctx.fillStyle = task.color || '#2196F3';
  roundRect(ctx, x, y, width, barHeight, radius);
  ctx.fill();

  // Progress bar
  if (config.showPercentComplete && task.percentComplete > 0) {
    const progressWidth = (width * task.percentComplete) / 100;
    ctx.fillStyle = COLORS.progress;

    if (progressWidth >= width) {
      roundRect(ctx, x, y, width, barHeight, radius);
    } else {
      ctx.beginPath();
      ctx.moveTo(x + radius, y);
      ctx.lineTo(x + progressWidth, y);
      ctx.lineTo(x + progressWidth, y + barHeight);
      ctx.lineTo(x + radius, y + barHeight);
      ctx.arcTo(x, y + barHeight, x, y + barHeight - radius, radius);
      ctx.lineTo(x, y + radius);
      ctx.arcTo(x, y, x + radius, y, radius);
      ctx.closePath();
    }
    ctx.fill();

    if (width > 40) {
      ctx.fillStyle = '#FFFFFF';
      ctx.font = `bold ${fontSize - 2}px 'Segoe UI', sans-serif`;
      ctx.textAlign = 'left';
      ctx.fillText(`${task.percentComplete}%`, x + 5, y + barHeight / 2 + 4);
    }
  }

  // Border for completed tasks
  if (task.percentComplete === 100) {
    ctx.strokeStyle = 'rgba(255,255,255,0.5)';
    ctx.lineWidth = 2;
    roundRect(ctx, x, y, width, barHeight, radius);
    ctx.stroke();
  }

  // Subtask indicator
  if (task.isSubtask) {
    ctx.fillStyle = 'rgba(0,0,0,0.1)';
    ctx.fillRect(x, y, 3, barHeight);
  }
}

function drawMilestone(
  ctx: CanvasRenderingContext2D,
  x: number,
  y: number,
  color: string,
  barHeight: number
): void {
  const size = barHeight / 2;

  ctx.save();
  ctx.translate(x, y);
  ctx.rotate(Math.PI / 4);

  ctx.fillStyle = color;
  ctx.fillRect(-size / 2, -size / 2, size, size);

  ctx.strokeStyle = '#FFFFFF';
  ctx.lineWidth = 2;
  ctx.strokeRect(-size / 2, -size / 2, size, size);

  ctx.restore();
}

function drawDateAxis(
  ctx: CanvasRenderingContext2D,
  chartArea: { x: number; y: number; width: number },
  start: Date,
  end: Date,
  totalDays: number,
  margin: { top: number },
  headerHeight: number,
  fontSize: number,
  zoom: number
): void {
  ctx.fillStyle = '#F8F9FA';
  ctx.fillRect(chartArea.x, margin.top - headerHeight, chartArea.width, headerHeight);

  ctx.strokeStyle = COLORS.grid;
  ctx.lineWidth = 1;
  ctx.beginPath();
  ctx.moveTo(chartArea.x, margin.top);
  ctx.lineTo(chartArea.x + chartArea.width, margin.top);
  ctx.stroke();

  // Month headers
  ctx.font = `bold ${fontSize}px 'Segoe UI', sans-serif`;
  ctx.fillStyle = COLORS.text;
  ctx.textAlign = 'center';

  let currentMonth = new Date(start.getFullYear(), start.getMonth(), 1);
  while (currentMonth <= end) {
    const nextMonth = new Date(currentMonth.getFullYear(), currentMonth.getMonth() + 1, 1);
    const monthStart = new Date(Math.max(currentMonth.getTime(), start.getTime()));
    const monthEnd = new Date(Math.min(nextMonth.getTime() - 1, end.getTime()));

    const x1 = chartArea.x + ((monthStart.getTime() - start.getTime()) / (1000 * 60 * 60 * 24) / totalDays) * chartArea.width;
    const x2 = chartArea.x + ((monthEnd.getTime() - start.getTime()) / (1000 * 60 * 60 * 24) / totalDays) * chartArea.width;
    const centerX = (x1 + x2) / 2;

    const monthName = currentMonth.toLocaleString('default', { month: 'short', year: 'numeric' });
    ctx.fillText(monthName, centerX, margin.top - headerHeight + 20 * zoom);

    ctx.strokeStyle = COLORS.grid;
    ctx.beginPath();
    ctx.moveTo(x1, margin.top - headerHeight);
    ctx.lineTo(x1, margin.top);
    ctx.stroke();

    currentMonth = nextMonth;
  }
}

function drawTaskLabels(
  ctx: CanvasRenderingContext2D,
  tasks: Task[],
  chartArea: { x: number; y: number },
  margin: { left: number },
  taskToY: (index: number) => number,
  fontSize: number,
  barHeight: number,
  rowHeight: number
): void {
  ctx.fillStyle = '#FAFAFA';
  ctx.fillRect(0, chartArea.y, margin.left - 10, tasks.length * rowHeight);

  tasks.forEach((task, index) => {
    const y = taskToY(index) + barHeight / 2 + 4;
    const x = task.isSubtask ? 45 : 30;

    ctx.fillStyle = COLORS.text;
    ctx.font = `${task.isSubtask ? '' : 'bold '}${fontSize}px 'Segoe UI', sans-serif`;
    ctx.textAlign = 'left';

    const maxWidth = margin.left - x - 20;
    const displayName = truncateText(ctx, task.name, maxWidth, fontSize);
    ctx.fillText(displayName, x, y);

    if (task.owner && !task.isSubtask) {
      ctx.fillStyle = COLORS.textLight;
      ctx.font = `${fontSize - 2}px 'Segoe UI', sans-serif`;
      const ownerY = y + 12;
      if (ownerY < taskToY(index + 1)) {
        ctx.fillText(task.owner, x, ownerY);
      }
    }
  });

  ctx.strokeStyle = COLORS.grid;
  ctx.lineWidth = 1;
  ctx.beginPath();
  ctx.moveTo(margin.left - 10, chartArea.y);
  ctx.lineTo(margin.left - 10, chartArea.y + tasks.length * rowHeight);
  ctx.stroke();
}

// Utility functions
function roundRect(
  ctx: CanvasRenderingContext2D,
  x: number,
  y: number,
  width: number,
  height: number,
  radius: number
): void {
  ctx.beginPath();
  ctx.moveTo(x + radius, y);
  ctx.lineTo(x + width - radius, y);
  ctx.quadraticCurveTo(x + width, y, x + width, y + radius);
  ctx.lineTo(x + width, y + height - radius);
  ctx.quadraticCurveTo(x + width, y + height, x + width - radius, y + height);
  ctx.lineTo(x + radius, y + height);
  ctx.quadraticCurveTo(x, y + height, x, y + height - radius);
  ctx.lineTo(x, y + radius);
  ctx.quadraticCurveTo(x, y, x + radius, y);
  ctx.closePath();
}

function truncateText(
  ctx: CanvasRenderingContext2D,
  text: string,
  maxWidth: number,
  fontSize: number
): string {
  ctx.font = `${fontSize}px 'Segoe UI', sans-serif`;
  if (ctx.measureText(text).width <= maxWidth) {
    return text;
  }

  let truncated = text;
  while (ctx.measureText(truncated + '...').width > maxWidth && truncated.length > 0) {
    truncated = truncated.slice(0, -1);
  }
  return truncated + '...';
}

function formatDate(date: Date): string {
  return date.toLocaleDateString('en-US', {
    year: 'numeric',
    month: 'short',
    day: 'numeric',
  });
}
