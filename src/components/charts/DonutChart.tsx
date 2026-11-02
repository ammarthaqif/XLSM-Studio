import React, { useState, useMemo } from 'react';
import { ReportingWidget, Dataset, WidgetWidth } from '../../types/dashboard';
import { Trash2, ChevronLeft, ChevronRight, GripVertical } from 'lucide-react';
import { useTheme } from '../../context/ThemeContext';

interface DonutChartProps {
  widget: ReportingWidget;
  dataset: Dataset;
  onDeleteWidget?: (id: string) => void;
  onMoveLeft?: () => void;
  onMoveRight?: () => void;
  canMoveLeft?: boolean;
  canMoveRight?: boolean;
  onChangeWidth?: (width: WidgetWidth) => void;
}

const PALETTE = ['#6366f1', '#10b981', '#f59e0b', '#ec4899', '#06b6d4', '#8b5cf6'];

export const DonutChart: React.FC<DonutChartProps> = ({
  widget,
  dataset,
  onDeleteWidget,
  onMoveLeft,
  onMoveRight,
  canMoveLeft,
  canMoveRight,
  onChangeWidth
}) => {
  const { isLight } = useTheme();
  const [hoveredIdx, setHoveredIdx] = useState<number | null>(null);
  const sheet = dataset.sheets.find((s) => s.name === widget.sheetName) || dataset.sheets[0];

  const { slices, total } = useMemo(() => {
    if (!sheet) return { slices: [], total: 0 };
    const groups: Record<string, number> = {};

    sheet.rows.forEach((row) => {
      const key = String(row[widget.xAxis] || 'Other');
      const val = Number(row[widget.yAxis]) || 0;
      groups[key] = (groups[key] || 0) + val;
    });

    const totalSum = Object.values(groups).reduce((a, b) => a + b, 0);

    const items = Object.entries(groups)
      .map(([label, value], i) => ({
        label,
        value,
        color: PALETTE[i % PALETTE.length],
        percentage: totalSum > 0 ? (value / totalSum) * 100 : 0
      }))
      .sort((a, b) => b.value - a.value);

    // Compute stroke dash offsets
    const radius = 64;
    const circumference = 2 * Math.PI * radius;
    let accumulated = 0;

    const sliceData = items.map((item) => {
      const strokeLength = (item.percentage / 100) * circumference;
      const strokeDashoffset = -accumulated;
      accumulated += strokeLength;
      return {
        ...item,
        circumference,
        strokeLength,
        strokeDashoffset
      };
    });

    return { slices: sliceData, total: totalSum };
  }, [sheet, widget.xAxis, widget.yAxis]);

  const formatNumber = (num: number) => {
    if (num >= 1_000_000) return `$${(num / 1_000_000).toFixed(1)}M`;
    if (num >= 1_000) return `$${(num / 1_000).toFixed(0)}k`;
    return num.toLocaleString();
  };

  return (
    <div
      className={`rounded-xl border p-5 shadow-sm transition-all ${
        isLight
          ? 'bg-white border-slate-200 hover:border-slate-300 text-slate-900 shadow-slate-100'
          : 'border-slate-800 bg-slate-900/90 hover:border-slate-700 text-white'
      }`}
    >
      <div className="flex flex-wrap items-center justify-between gap-3 mb-3">
        <div className="flex items-center gap-2 min-w-0">
          <span
            className={`cursor-grab active:cursor-grabbing p-0.5 ${
              isLight ? 'text-slate-400 hover:text-slate-600' : 'text-slate-600 hover:text-slate-400'
            }`}
            title="Drag to reorder chart in layout"
          >
            <GripVertical className="w-4 h-4" />
          </span>
          <div>
            <h4
              className={`text-sm font-semibold tracking-tight truncate ${
                isLight ? 'text-slate-900' : 'text-white'
              }`}
            >
              {widget.title}
            </h4>
            <p className={`text-xs mt-0.5 ${isLight ? 'text-slate-500' : 'text-slate-400'}`}>
              Share by <span className={isLight ? 'text-slate-700 font-medium' : 'text-slate-300'}>{widget.xAxis}</span>
            </p>
          </div>
        </div>

        <div className="flex items-center gap-2 flex-wrap">
          {/* Width adjustment selector */}
          {onChangeWidth && (
            <select
              value={widget.width}
              onChange={(e) => onChangeWidth(e.target.value as WidgetWidth)}
              className={`border rounded-lg px-2 py-1 text-xs font-mono focus:outline-none focus:border-indigo-500 ${
                isLight
                  ? 'bg-slate-50 border-slate-300 text-slate-700'
                  : 'bg-slate-800/90 border-slate-700 text-slate-300'
              }`}
              title="Adjust widget layout width"
            >
              <option value="1/3">1/3 Width</option>
              <option value="1/2">1/2 Width</option>
              <option value="2/3">2/3 Width</option>
              <option value="full">Full Width</option>
            </select>
          )}

          {/* Reorder Arrows */}
          <div
            className={`flex items-center gap-0.5 border rounded-lg p-0.5 ${
              isLight ? 'border-slate-300 bg-slate-100' : 'border-slate-700 bg-slate-800/60'
            }`}
          >
            {onMoveLeft && canMoveLeft && (
              <button
                onClick={onMoveLeft}
                title="Move left/up in layout"
                className={`p-1 rounded transition-colors ${
                  isLight
                    ? 'text-slate-600 hover:text-indigo-600 hover:bg-slate-200'
                    : 'text-slate-400 hover:text-indigo-300 hover:bg-slate-700'
                }`}
              >
                <ChevronLeft className="w-3.5 h-3.5" />
              </button>
            )}
            {onMoveRight && canMoveRight && (
              <button
                onClick={onMoveRight}
                title="Move right/down in layout"
                className={`p-1 rounded transition-colors ${
                  isLight
                    ? 'text-slate-600 hover:text-indigo-600 hover:bg-slate-200'
                    : 'text-slate-400 hover:text-indigo-300 hover:bg-slate-700'
                }`}
              >
                <ChevronRight className="w-3.5 h-3.5" />
              </button>
            )}
          </div>

          {onDeleteWidget && (
            <button
              onClick={() => onDeleteWidget(widget.id)}
              title="Delete widget"
              className={`p-1.5 rounded-lg transition-colors ${
                isLight
                  ? 'text-slate-400 hover:text-rose-600 hover:bg-rose-50'
                  : 'text-slate-500 hover:text-rose-400 hover:bg-slate-800'
              }`}
            >
              <Trash2 className="w-4 h-4" />
            </button>
          )}
        </div>
      </div>

      <div className="flex flex-col sm:flex-row items-center justify-center gap-6 py-2">
        {/* SVG Donut */}
        <div className="relative w-36 h-36 shrink-0">
          <svg viewBox="0 0 160 160" className="w-full h-full -rotate-90">
            <circle
              cx="80"
              cy="80"
              r="64"
              fill="transparent"
              stroke={isLight ? '#e2e8f0' : '#1e293b'}
              strokeWidth="20"
            />
            {slices.map((slice, i) => (
              <circle
                key={slice.label}
                cx="80"
                cy="80"
                r="64"
                fill="transparent"
                stroke={slice.color}
                strokeWidth={hoveredIdx === i ? 23 : 20}
                strokeDasharray={`${slice.strokeLength} ${slice.circumference - slice.strokeLength}`}
                strokeDashoffset={slice.strokeDashoffset}
                onMouseEnter={() => setHoveredIdx(i)}
                onMouseLeave={() => setHoveredIdx(null)}
                className="transition-all duration-200 cursor-pointer"
              />
            ))}
          </svg>

          {/* Center text */}
          <div className="absolute inset-0 flex flex-col items-center justify-center pointer-events-none">
            <span
              className={`text-[10px] font-mono uppercase tracking-wider ${
                isLight ? 'text-slate-500' : 'text-slate-400'
              }`}
            >
              {hoveredIdx !== null ? slices[hoveredIdx]?.label : 'Total'}
            </span>
            <span
              className={`text-sm font-bold font-mono tabular-nums ${
                isLight ? 'text-slate-900' : 'text-white'
              }`}
            >
              {hoveredIdx !== null
                ? `${slices[hoveredIdx]?.percentage.toFixed(1)}%`
                : formatNumber(total)}
            </span>
          </div>
        </div>

        {/* Legend Breakdown */}
        <div className="w-full space-y-1.5">
          {slices.map((slice, i) => (
            <div
              key={slice.label}
              onMouseEnter={() => setHoveredIdx(i)}
              onMouseLeave={() => setHoveredIdx(null)}
              className={`flex items-center justify-between text-xs p-1 rounded cursor-pointer transition-colors ${
                hoveredIdx === i
                  ? isLight
                    ? 'bg-slate-100 text-slate-900 font-semibold'
                    : 'bg-slate-800 text-white'
                  : isLight
                  ? 'text-slate-700 hover:bg-slate-50'
                  : 'text-slate-300 hover:bg-slate-800/40'
              }`}
            >
              <div className="flex items-center gap-2 truncate">
                <span
                  className="w-2.5 h-2.5 rounded-sm shrink-0"
                  style={{ backgroundColor: slice.color }}
                />
                <span className="truncate">{slice.label}</span>
              </div>
              <div className="flex items-center gap-2 font-mono tabular-nums shrink-0 ml-2">
                <span className={`text-[11px] ${isLight ? 'text-slate-500' : 'text-slate-400'}`}>
                  {formatNumber(slice.value)}
                </span>
                <span
                  className={`w-12 text-right ${
                    isLight ? 'font-semibold text-slate-800' : 'font-semibold text-slate-200'
                  }`}
                >
                  {slice.percentage.toFixed(1)}%
                </span>
              </div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
};
