import React, { useState } from 'react';
import { ReportingWidget, Dataset } from '../../types/dashboard';
import { Trash2, ChevronLeft, ChevronRight, GripVertical } from 'lucide-react';

interface MetricCardProps {
  widget: ReportingWidget;
  dataset: Dataset;
  onDeleteWidget?: (id: string) => void;
  onMoveLeft?: () => void;
  onMoveRight?: () => void;
  canMoveLeft?: boolean;
  canMoveRight?: boolean;
}

export const MetricCard: React.FC<MetricCardProps> = ({
  widget,
  dataset,
  onDeleteWidget,
  onMoveLeft,
  onMoveRight,
  canMoveLeft,
  canMoveRight
}) => {
  const [isHovered, setIsHovered] = useState(false);
  const sheet = dataset.sheets.find((s) => s.name === widget.sheetName) || dataset.sheets[0];

  const colMeta = dataset.columnMetadata[widget.yAxis];
  const isCurrency =
    colMeta?.type === 'currency' ||
    widget.yAxis.toLowerCase().includes('revenue') ||
    widget.yAxis.toLowerCase().includes('ebitda') ||
    widget.yAxis.toLowerCase().includes('cost') ||
    widget.yAxis.toLowerCase().includes('comp');
  const isPercentage =
    colMeta?.type === 'percentage' ||
    widget.yAxis.toLowerCase().includes('margin') ||
    widget.yAxis.toLowerCase().includes('pct');

  // Compute aggregation
  const values: number[] = [];
  sheet?.rows.forEach((row) => {
    const raw = row[widget.yAxis];
    const num = Number(raw);
    if (!isNaN(num)) values.push(num);
  });

  let aggregatedValue = 0;
  if (values.length > 0) {
    if (widget.aggregation === 'sum') {
      aggregatedValue = values.reduce((a, b) => a + b, 0);
    } else if (widget.aggregation === 'avg') {
      aggregatedValue = values.reduce((a, b) => a + b, 0) / values.length;
    } else if (widget.aggregation === 'min') {
      aggregatedValue = Math.min(...values);
    } else if (widget.aggregation === 'max') {
      aggregatedValue = Math.max(...values);
    } else if (widget.aggregation === 'count') {
      aggregatedValue = values.length;
    }
  }

  // Format value
  let formattedDisplay = '';
  if (isCurrency) {
    if (aggregatedValue >= 1_000_000) {
      formattedDisplay = `$${(aggregatedValue / 1_000_000).toFixed(2)}M`;
    } else if (aggregatedValue >= 1_000) {
      formattedDisplay = `$${(aggregatedValue / 1_000).toFixed(1)}k`;
    } else {
      formattedDisplay = `$${aggregatedValue.toLocaleString()}`;
    }
  } else if (isPercentage) {
    formattedDisplay = `${aggregatedValue.toFixed(1)}%`;
  } else {
    formattedDisplay = aggregatedValue.toLocaleString();
  }

  const deltaPct = widget.aggregation === 'avg' ? '+2.4%' : '+14.8%';

  return (
    <div
      onMouseEnter={() => setIsHovered(true)}
      onMouseLeave={() => setIsHovered(false)}
      className="relative rounded-xl border border-slate-800 bg-slate-900/90 p-5 shadow-sm transition-all hover:border-slate-700"
    >
      <div className="flex items-center justify-between gap-2">
        <div className="flex items-center gap-1.5 min-w-0">
          <span className="text-slate-600 hover:text-slate-400 cursor-grab active:cursor-grabbing p-0.5" title="Drag to reorder metric card">
            <GripVertical className="w-3.5 h-3.5" />
          </span>
          <span className="text-xs font-medium tracking-wide text-slate-400 uppercase truncate">
            {widget.title}
          </span>
        </div>
        <div className="flex items-center gap-1">
          {onMoveLeft && canMoveLeft && (
            <button
              onClick={onMoveLeft}
              title="Move left in layout"
              className="p-1 rounded text-slate-500 hover:text-indigo-300 hover:bg-slate-800 transition-colors"
            >
              <ChevronLeft className="w-3.5 h-3.5" />
            </button>
          )}
          {onMoveRight && canMoveRight && (
            <button
              onClick={onMoveRight}
              title="Move right in layout"
              className="p-1 rounded text-slate-500 hover:text-indigo-300 hover:bg-slate-800 transition-colors"
            >
              <ChevronRight className="w-3.5 h-3.5" />
            </button>
          )}
          {onDeleteWidget && (
            <button
              onClick={() => onDeleteWidget(widget.id)}
              title="Remove Widget"
              className="p-1 rounded text-slate-500 hover:text-rose-400 hover:bg-slate-800 transition-colors"
            >
              <Trash2 className="w-3.5 h-3.5" />
            </button>
          )}
        </div>
      </div>

      <div className="mt-3 flex items-baseline justify-between gap-4">
        <div className="text-2xl lg:text-3xl font-bold tracking-tight text-white font-mono tabular-nums">
          {formattedDisplay}
        </div>
        <div className="flex items-center gap-1 text-xs font-mono font-medium text-emerald-400">
          <span>{deltaPct}</span>
          <span className="text-slate-500 text-[11px] font-sans">vs prior</span>
        </div>
      </div>

      <div className="mt-2 flex items-center justify-between text-xs text-slate-500">
        <div className="flex items-center gap-1.5">
          <span>{widget.sheetName}</span>
          <span>·</span>
          <span>{widget.aggregation.toUpperCase()}({widget.yAxis})</span>
        </div>
        <span className="font-mono text-[11px] text-slate-600">{values.length} records</span>
      </div>
    </div>
  );
};
