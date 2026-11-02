import React, { useState, useMemo } from 'react';
import { ReportingWidget, Dataset, WidgetWidth } from '../../types/dashboard';
import { Trash2, ArrowUpDown, ChevronLeft, ChevronRight, GripVertical } from 'lucide-react';
import { useTheme } from '../../context/ThemeContext';

interface BarGroupChartProps {
  widget: ReportingWidget;
  dataset: Dataset;
  onDeleteWidget?: (id: string) => void;
  onMoveLeft?: () => void;
  onMoveRight?: () => void;
  canMoveLeft?: boolean;
  canMoveRight?: boolean;
  onChangeWidth?: (width: WidgetWidth) => void;
}

export const BarGroupChart: React.FC<BarGroupChartProps> = ({
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
  const [hoverIndex, setHoverIndex] = useState<number | null>(null);
  const [sortDescending, setSortDescending] = useState(true);

  const sheet = dataset.sheets.find((s) => s.name === widget.sheetName) || dataset.sheets[0];

  // Group and aggregate data by xAxis
  const aggregatedData = useMemo(() => {
    if (!sheet) return [];
    const groups: Record<string, { sum: number; count: number }> = {};

    sheet.rows.forEach((row) => {
      const key = String(row[widget.xAxis] || 'Unassigned');
      const val = Number(row[widget.yAxis]) || 0;
      if (!groups[key]) {
        groups[key] = { sum: 0, count: 0 };
      }
      groups[key].sum += val;
      groups[key].count += 1;
    });

    const result = Object.entries(groups).map(([label, stats]) => {
      let finalVal = stats.sum;
      if (widget.aggregation === 'avg' && stats.count > 0) {
        finalVal = stats.sum / stats.count;
      } else if (widget.aggregation === 'count') {
        finalVal = stats.count;
      }
      return {
        label,
        value: finalVal,
        count: stats.count
      };
    });

    result.sort((a, b) => (sortDescending ? b.value - a.value : a.value - b.value));
    return result;
  }, [sheet, widget.xAxis, widget.yAxis, widget.aggregation, sortDescending]);

  const maxVal = useMemo(() => {
    const max = Math.max(...aggregatedData.map((d) => d.value), 1);
    return max * 1.15;
  }, [aggregatedData]);

  const formatNumber = (num: number) => {
    if (num >= 1_000_000) return `$${(num / 1_000_000).toFixed(2)}M`;
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
      <div className="flex flex-wrap items-center justify-between gap-3 mb-4">
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
              Grouped by{' '}
              <span className={`font-medium ${isLight ? 'text-slate-800' : 'text-slate-300'}`}>
                {widget.xAxis}
              </span>{' '}
              · {widget.aggregation.toUpperCase()}
            </p>
          </div>
        </div>

        <div className="flex items-center gap-2 flex-wrap">
          <button
            onClick={() => setSortDescending(!sortDescending)}
            title="Toggle sort order"
            className={`flex items-center gap-1 px-2 py-1 text-xs font-mono rounded-md border transition-colors ${
              isLight
                ? 'border-slate-300 bg-slate-100 text-slate-700 hover:text-slate-900 hover:bg-slate-200'
                : 'border-slate-700 bg-slate-800/60 text-slate-300 hover:text-white'
            }`}
          >
            <ArrowUpDown className="w-3.5 h-3.5" />
            <span>{sortDescending ? 'High-Low' : 'Low-High'}</span>
          </button>

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

      {/* Bar Chart Area */}
      <div className="space-y-3.5 mt-2">
        {aggregatedData.map((item, idx) => {
          const pct = Math.min(Math.max((item.value / maxVal) * 100, 4), 100);
          const isHovered = hoverIndex === idx;

          return (
            <div
              key={item.label}
              onMouseEnter={() => setHoverIndex(idx)}
              onMouseLeave={() => setHoverIndex(null)}
              className="group cursor-pointer"
            >
              <div className="flex items-center justify-between text-xs mb-1">
                <span
                  className={`font-medium transition-colors ${
                    isHovered
                      ? isLight
                        ? 'text-amber-700 font-bold'
                        : 'text-amber-300'
                      : isLight
                      ? 'text-slate-800'
                      : 'text-slate-300'
                  }`}
                >
                  {item.label}
                </span>
                <div className="flex items-center gap-2">
                  <span
                    className={`text-[11px] font-mono ${
                      isLight ? 'text-slate-500' : 'text-slate-500'
                    }`}
                  >
                    ({item.count} rows)
                  </span>
                  <span
                    className={`font-mono tabular-nums font-semibold ${
                      isLight ? 'text-slate-900' : 'text-slate-100'
                    }`}
                  >
                    {formatNumber(item.value)}
                  </span>
                </div>
              </div>

              {/* Bar track and fill */}
              <div
                className={`relative w-full h-3 rounded overflow-hidden ${
                  isLight ? 'bg-slate-100' : 'bg-slate-800/80'
                }`}
              >
                <div
                  className={`h-full transition-all duration-300 ease-out rounded ${
                    isHovered
                      ? isLight
                        ? 'bg-amber-500'
                        : 'bg-amber-400'
                      : isLight
                      ? 'bg-amber-500/85'
                      : 'bg-amber-500/80'
                  }`}
                  style={{ width: `${pct}%` }}
                />
              </div>
            </div>
          );
        })}
      </div>

      <div
        className={`mt-4 pt-3 border-t flex items-center justify-between text-xs font-mono ${
          isLight ? 'border-slate-200 text-slate-500' : 'border-slate-800 text-slate-500'
        }`}
      >
        <span>Total categories: {aggregatedData.length}</span>
        <span>Peak: {aggregatedData[0] ? formatNumber(aggregatedData[0].value) : 0}</span>
      </div>
    </div>
  );
};
