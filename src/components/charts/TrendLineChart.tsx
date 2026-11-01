import React, { useState, useMemo } from 'react';
import { ReportingWidget, Dataset, WidgetWidth } from '../../types/dashboard';
import { Trash2, ChevronLeft, ChevronRight, GripVertical } from 'lucide-react';

interface TrendLineChartProps {
  widget: ReportingWidget;
  dataset: Dataset;
  onDeleteWidget?: (id: string) => void;
  onMoveLeft?: () => void;
  onMoveRight?: () => void;
  canMoveLeft?: boolean;
  canMoveRight?: boolean;
  onChangeWidth?: (width: WidgetWidth) => void;
}

export const TrendLineChart: React.FC<TrendLineChartProps> = ({
  widget,
  dataset,
  onDeleteWidget,
  onMoveLeft,
  onMoveRight,
  canMoveLeft,
  canMoveRight,
  onChangeWidth
}) => {
  const [hoverIndex, setHoverIndex] = useState<number | null>(null);
  const [showSecondaryMetric, setShowSecondaryMetric] = useState(true);
  const [timeFilter, setTimeFilter] = useState<'All' | '6M' | '3M'>('All');

  const sheet = dataset.sheets.find((s) => s.name === widget.sheetName) || dataset.sheets[0];

  // Secondary metric candidate (e.g. OPEX if primary is Revenue, or Margin)
  const secondaryKey = useMemo(() => {
    if (!sheet) return null;
    const candidates = ['OPEX', 'EBITDA', 'Cost', 'SafetyStock', 'AnnualTarget'];
    return candidates.find((c) => c !== widget.yAxis && sheet.columns.includes(c)) || null;
  }, [sheet, widget.yAxis]);

  // Aggregate data by xAxis
  const chartData = useMemo(() => {
    if (!sheet) return [];
    let rows = [...sheet.rows];

    if (timeFilter === '6M') {
      rows = rows.slice(-6);
    } else if (timeFilter === '3M') {
      rows = rows.slice(-3);
    }

    return rows.map((row) => ({
      label: String(row[widget.xAxis] || 'N/A'),
      primaryVal: Number(row[widget.yAxis]) || 0,
      secondaryVal: secondaryKey ? Number(row[secondaryKey]) || 0 : 0
    }));
  }, [sheet, widget.xAxis, widget.yAxis, secondaryKey, timeFilter]);

  // Dimensions
  const width = 640;
  const height = 240;
  const paddingLeft = 55;
  const paddingRight = 20;
  const paddingTop = 20;
  const paddingBottom = 35;
  const chartWidth = width - paddingLeft - paddingRight;
  const chartHeight = height - paddingTop - paddingBottom;

  const maxVal = useMemo(() => {
    if (chartData.length === 0) return 100;
    const all = chartData.flatMap((d) => [d.primaryVal, showSecondaryMetric ? d.secondaryVal : 0]);
    const max = Math.max(...all, 1);
    return Math.ceil(max * 1.15);
  }, [chartData, showSecondaryMetric]);

  const points = useMemo(() => {
    if (chartData.length <= 1) return [];
    return chartData.map((d, i) => {
      const x = paddingLeft + (i / (chartData.length - 1)) * chartWidth;
      const yPrimary = paddingTop + chartHeight - (d.primaryVal / maxVal) * chartHeight;
      const ySecondary = paddingTop + chartHeight - (d.secondaryVal / maxVal) * chartHeight;
      return { x, yPrimary, ySecondary, ...d };
    });
  }, [chartData, maxVal, chartWidth, chartHeight]);

  // Generate SVG path for primary curve
  const primaryPath = useMemo(() => {
    if (points.length < 2) return '';
    return points.reduce((acc, p, i) => {
      if (i === 0) return `M ${p.x} ${p.yPrimary}`;
      const prev = points[i - 1];
      const cx = (prev.x + p.x) / 2;
      return `${acc} C ${cx} ${prev.yPrimary}, ${cx} ${p.yPrimary}, ${p.x} ${p.yPrimary}`;
    }, '');
  }, [points]);

  // Area path for gradient fill
  const primaryAreaPath = useMemo(() => {
    if (points.length < 2) return '';
    const first = points[0];
    const last = points[points.length - 1];
    const baseline = paddingTop + chartHeight;
    return `${primaryPath} L ${last.x} ${baseline} L ${first.x} ${baseline} Z`;
  }, [primaryPath, points, chartHeight]);

  // Secondary path
  const secondaryPath = useMemo(() => {
    if (!showSecondaryMetric || points.length < 2 || !secondaryKey) return '';
    return points.reduce((acc, p, i) => {
      if (i === 0) return `M ${p.x} ${p.ySecondary}`;
      const prev = points[i - 1];
      const cx = (prev.x + p.x) / 2;
      return `${acc} C ${cx} ${prev.ySecondary}, ${cx} ${p.ySecondary}, ${p.x} ${p.ySecondary}`;
    }, '');
  }, [points, showSecondaryMetric, secondaryKey]);

  const formatNumber = (num: number) => {
    if (num >= 1_000_000) return `$${(num / 1_000_000).toFixed(1)}M`;
    if (num >= 1_000) return `$${(num / 1_000).toFixed(0)}k`;
    return num.toLocaleString();
  };

  return (
    <div className="rounded-xl border border-slate-800 bg-slate-900/90 p-5 shadow-sm transition-all hover:border-slate-700">
      {/* Header */}
      <div className="flex flex-wrap items-center justify-between gap-3 mb-4">
        <div className="flex items-center gap-2 min-w-0">
          <span className="text-slate-600 hover:text-slate-400 cursor-grab active:cursor-grabbing p-0.5" title="Drag to reorder chart in layout">
            <GripVertical className="w-4 h-4" />
          </span>
          <div>
            <div className="flex items-center gap-2">
              <h4 className="text-sm font-semibold tracking-tight text-white truncate">{widget.title}</h4>
              <span className="text-xs text-slate-500 font-mono">({widget.xAxis} series)</span>
            </div>
            <p className="text-xs text-slate-400 mt-0.5">
              Ingested from sheet <span className="text-slate-300 font-medium">{widget.sheetName}</span>
            </p>
          </div>
        </div>

        {/* Controls */}
        <div className="flex items-center gap-2 flex-wrap">
          {/* Filter tabs */}
          <div className="flex items-center gap-1 p-0.5 bg-slate-800/80 rounded-lg border border-slate-700">
            {(['All', '6M', '3M'] as const).map((t) => (
              <button
                key={t}
                onClick={() => setTimeFilter(t)}
                className={`px-2.5 py-1 text-xs font-medium rounded-md transition-colors ${
                  timeFilter === t
                    ? 'bg-slate-700 text-white shadow-sm'
                    : 'text-slate-400 hover:text-slate-200'
                }`}
              >
                {t}
              </button>
            ))}
          </div>

          {/* Secondary toggle */}
          {secondaryKey && (
            <button
              onClick={() => setShowSecondaryMetric(!showSecondaryMetric)}
              className={`px-2 py-1 text-xs font-medium rounded-md border transition-colors ${
                showSecondaryMetric
                  ? 'border-indigo-500/40 bg-indigo-950/40 text-indigo-300'
                  : 'border-slate-700 bg-slate-800/60 text-slate-400 hover:text-slate-200'
              }`}
            >
              + {secondaryKey}
            </button>
          )}

          {/* Width adjustment selector */}
          {onChangeWidth && (
            <select
              value={widget.width}
              onChange={(e) => onChangeWidth(e.target.value as WidgetWidth)}
              className="bg-slate-800/90 border border-slate-700 rounded-lg px-2 py-1 text-xs text-slate-300 font-mono focus:outline-none focus:border-indigo-500"
              title="Adjust widget layout width"
            >
              <option value="1/3">1/3 Width</option>
              <option value="1/2">1/2 Width</option>
              <option value="2/3">2/3 Width</option>
              <option value="full">Full Width</option>
            </select>
          )}

          {/* Reorder Arrows */}
          <div className="flex items-center gap-0.5 border border-slate-700 rounded-lg bg-slate-800/60 p-0.5">
            {onMoveLeft && canMoveLeft && (
              <button
                onClick={onMoveLeft}
                title="Move left/up in layout"
                className="p-1 rounded text-slate-400 hover:text-indigo-300 hover:bg-slate-700 transition-colors"
              >
                <ChevronLeft className="w-3.5 h-3.5" />
              </button>
            )}
            {onMoveRight && canMoveRight && (
              <button
                onClick={onMoveRight}
                title="Move right/down in layout"
                className="p-1 rounded text-slate-400 hover:text-indigo-300 hover:bg-slate-700 transition-colors"
              >
                <ChevronRight className="w-3.5 h-3.5" />
              </button>
            )}
          </div>

          {onDeleteWidget && (
            <button
              onClick={() => onDeleteWidget(widget.id)}
              title="Delete widget"
              className="p-1.5 rounded-lg text-slate-500 hover:text-rose-400 hover:bg-slate-800 transition-colors"
            >
              <Trash2 className="w-4 h-4" />
            </button>
          )}
        </div>
      </div>

      {/* SVG Canvas */}
      <div className="relative w-full overflow-hidden">
        <svg
          viewBox={`0 0 ${width} ${height}`}
          className="w-full h-auto select-none"
          onMouseLeave={() => setHoverIndex(null)}
        >
          <defs>
            <linearGradient id="areaGradient" x1="0" y1="0" x2="0" y2="1">
              <stop offset="0%" stopColor="#10b981" stopOpacity="0.28" />
              <stop offset="100%" stopColor="#10b981" stopOpacity="0.0" />
            </linearGradient>
            <linearGradient id="secondaryGradient" x1="0" y1="0" x2="0" y2="1">
              <stop offset="0%" stopColor="#6366f1" stopOpacity="0.18" />
              <stop offset="100%" stopColor="#6366f1" stopOpacity="0.0" />
            </linearGradient>
          </defs>

          {/* Grid lines */}
          {[0, 0.25, 0.5, 0.75, 1].map((pct, i) => {
            const y = paddingTop + chartHeight * (1 - pct);
            const val = maxVal * pct;
            return (
              <g key={i}>
                <line
                  x1={paddingLeft}
                  y1={y}
                  x2={width - paddingRight}
                  y2={y}
                  stroke="#334155"
                  strokeDasharray="3 3"
                  strokeOpacity="0.5"
                />
                <text
                  x={paddingLeft - 8}
                  y={y + 3}
                  textAnchor="end"
                  className="text-[10px] fill-slate-500 font-mono tabular-nums"
                >
                  {formatNumber(val)}
                </text>
              </g>
            );
          })}

          {/* Area under primary line */}
          {primaryAreaPath && <path d={primaryAreaPath} fill="url(#areaGradient)" />}

          {/* Secondary line (OPEX / comparison) */}
          {secondaryPath && (
            <path
              d={secondaryPath}
              fill="none"
              stroke="#818cf8"
              strokeWidth="2"
              strokeDasharray="4 4"
            />
          )}

          {/* Primary Curve */}
          {primaryPath && (
            <path
              d={primaryPath}
              fill="none"
              stroke="#10b981"
              strokeWidth="2.5"
              strokeLinecap="round"
              strokeLinejoin="round"
            />
          )}

          {/* Interactive Hover Crosshairs & Data Points */}
          {points.map((p, idx) => {
            const isHovered = hoverIndex === idx;
            return (
              <g
                key={idx}
                onMouseEnter={() => setHoverIndex(idx)}
                className="cursor-pointer"
              >
                {/* Transparent hit target column */}
                <rect
                  x={p.x - 20}
                  y={paddingTop}
                  width="40"
                  height={chartHeight}
                  fill="transparent"
                />

                {isHovered && (
                  <line
                    x1={p.x}
                    y1={paddingTop}
                    x2={p.x}
                    y2={paddingTop + chartHeight}
                    stroke="#94a3b8"
                    strokeWidth="1"
                    strokeDasharray="2 2"
                  />
                )}

                {/* Primary dot */}
                <circle
                  cx={p.x}
                  cy={p.yPrimary}
                  r={isHovered ? 6 : 3.5}
                  fill="#0f172a"
                  stroke="#10b981"
                  strokeWidth={isHovered ? 3 : 2}
                  className="transition-all"
                />

                {/* Secondary dot */}
                {showSecondaryMetric && secondaryKey && (
                  <circle
                    cx={p.x}
                    cy={p.ySecondary}
                    r={isHovered ? 5 : 3}
                    fill="#0f172a"
                    stroke="#818cf8"
                    strokeWidth={isHovered ? 2.5 : 1.5}
                    className="transition-all"
                  />
                )}

                {/* X-axis labels */}
                <text
                  x={p.x}
                  y={height - 10}
                  textAnchor="middle"
                  className={`text-[10px] font-mono select-none ${
                    isHovered ? 'fill-slate-100 font-semibold' : 'fill-slate-500'
                  }`}
                >
                  {p.label.length > 7 ? p.label.slice(5) : p.label}
                </text>
              </g>
            );
          })}
        </svg>

        {/* Hover Tooltip Box */}
        {hoverIndex !== null && points[hoverIndex] && (
          <div
            className="pointer-events-none absolute z-20 rounded-lg border border-slate-700 bg-slate-900/95 p-2.5 text-xs shadow-xl backdrop-blur-sm"
            style={{
              left: `${(points[hoverIndex].x / width) * 100}%`,
              top: '15%',
              transform: 'translate(-50%, 0)'
            }}
          >
            <div className="font-semibold text-slate-200 border-b border-slate-800 pb-1 mb-1.5 flex items-center justify-between gap-3">
              <span>{points[hoverIndex].label}</span>
              <span className="text-[10px] text-slate-400 font-mono">Row #{hoverIndex + 1}</span>
            </div>
            <div className="space-y-1 font-mono">
              <div className="flex items-center justify-between gap-3 text-emerald-400">
                <span className="text-slate-400 font-sans">{widget.yAxis}:</span>
                <span className="font-semibold tabular-nums">{formatNumber(points[hoverIndex].primaryVal)}</span>
              </div>
              {showSecondaryMetric && secondaryKey && (
                <div className="flex items-center justify-between gap-3 text-indigo-400">
                  <span className="text-slate-400 font-sans">{secondaryKey}:</span>
                  <span className="font-semibold tabular-nums">{formatNumber(points[hoverIndex].secondaryVal)}</span>
                </div>
              )}
            </div>
          </div>
        )}
      </div>

      {/* Legend & Stats Footer */}
      <div className="mt-2 pt-2 border-t border-slate-800/80 flex items-center justify-between text-xs text-slate-400">
        <div className="flex items-center gap-4">
          <div className="flex items-center gap-1.5">
            <span className="w-2.5 h-2.5 rounded-full bg-emerald-500 inline-block"></span>
            <span className="text-slate-300 font-medium">{widget.yAxis}</span>
          </div>
          {showSecondaryMetric && secondaryKey && (
            <div className="flex items-center gap-1.5">
              <span className="w-2.5 h-0.5 bg-indigo-400 inline-block"></span>
              <span className="text-slate-300">{secondaryKey}</span>
            </div>
          )}
        </div>
        <span className="text-slate-500 text-[11px] font-mono">
          {points.length} data points · Live reactive SVG
        </span>
      </div>
    </div>
  );
};
