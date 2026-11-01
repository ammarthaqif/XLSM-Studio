import React, { useState } from 'react';
import { Dataset, ReportingWidget, WidgetType, AggregationType, WidgetWidth } from '../../types/dashboard';
import { X, Sparkles, BarChart3, LineChart, PieChart, Activity, Table } from 'lucide-react';

interface WidgetStudioModalProps {
  isOpen: boolean;
  onClose: () => void;
  dataset: Dataset;
  onSaveWidget: (widget: ReportingWidget) => void;
}

export const WidgetStudioModal: React.FC<WidgetStudioModalProps> = ({
  isOpen,
  onClose,
  dataset,
  onSaveWidget
}) => {
  const [title, setTitle] = useState('');
  const [type, setType] = useState<WidgetType>('bar');
  const [sheetName, setSheetName] = useState(dataset.activeSheetName || dataset.sheets[0]?.name || '');
  
  const currentSheet = dataset.sheets.find((s) => s.name === sheetName) || dataset.sheets[0];
  const columns = currentSheet?.columns || [];

  const [xAxis, setXAxis] = useState(columns[0] || '');
  const [yAxis, setYAxis] = useState(columns[1] || columns[0] || '');
  const [aggregation, setAggregation] = useState<AggregationType>('sum');
  const [width, setWidth] = useState<WidgetWidth>('1/2');
  const [colorTheme, setColorTheme] = useState<'emerald' | 'indigo' | 'amber' | 'cyan' | 'rose' | 'slate'>('indigo');

  if (!isOpen) return null;

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    const finalTitle = title.trim() || `${aggregation.toUpperCase()}(${yAxis}) by ${xAxis}`;
    const newWidget: ReportingWidget = {
      id: `w-${Date.now()}-${Math.random().toString(36).substring(2, 6)}`,
      title: finalTitle,
      type,
      sheetName,
      xAxis: xAxis || columns[0],
      yAxis: yAxis || columns[1] || columns[0],
      aggregation,
      width,
      colorTheme
    };
    onSaveWidget(newWidget);
    onClose();
  };

  const chartTypes: { type: WidgetType; label: string; icon: React.ReactNode }[] = [
    { type: 'metric', label: 'Metric Card', icon: <Activity className="w-4 h-4" /> },
    { type: 'line', label: 'Trend Line', icon: <LineChart className="w-4 h-4" /> },
    { type: 'bar', label: 'Bar Series', icon: <BarChart3 className="w-4 h-4" /> },
    { type: 'donut', label: 'Donut Share', icon: <PieChart className="w-4 h-4" /> },
    { type: 'table', label: 'Summary Table', icon: <Table className="w-4 h-4" /> }
  ];

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/70 backdrop-blur-sm animate-in fade-in duration-200">
      <div className="relative w-full max-w-xl rounded-2xl border border-slate-700 bg-slate-900 shadow-2xl p-6 overflow-hidden">
        {/* Header */}
        <div className="flex items-center justify-between border-b border-slate-800 pb-4 mb-5">
          <div className="flex items-center gap-2">
            <div className="p-2 rounded-lg bg-indigo-950/70 border border-indigo-500/30 text-indigo-400">
              <Sparkles className="w-4 h-4" />
            </div>
            <div>
              <h3 className="text-base font-semibold text-white">Create Custom Reporting Widget</h3>
              <p className="text-xs text-slate-400">Configure dimensional metrics and visualization parameters</p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="p-1 rounded-lg text-slate-400 hover:text-white hover:bg-slate-800 transition-colors"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Form */}
        <form onSubmit={handleSubmit} className="space-y-4">
          {/* Title */}
          <div>
            <label className="block text-xs font-medium text-slate-300 mb-1">
              Widget Title
            </label>
            <input
              type="text"
              value={title}
              onChange={(e) => setTitle(e.target.value)}
              placeholder="e.g. Operating Margin by Division"
              className="w-full bg-slate-800/90 border border-slate-700 rounded-lg px-3 py-2 text-xs text-slate-100 placeholder-slate-500 focus:outline-none focus:border-indigo-500"
            />
          </div>

          {/* Visualization Type */}
          <div>
            <label className="block text-xs font-medium text-slate-300 mb-1.5">
              Visualization Type
            </label>
            <div className="grid grid-cols-2 sm:grid-cols-5 gap-2">
              {chartTypes.map((t) => (
                <button
                  type="button"
                  key={t.type}
                  onClick={() => setType(t.type)}
                  className={`flex flex-col items-center justify-center gap-1.5 p-2.5 rounded-lg border text-xs font-medium transition-all ${
                    type === t.type
                      ? 'border-indigo-500 bg-indigo-950/60 text-white shadow-sm ring-1 ring-indigo-500'
                      : 'border-slate-800 bg-slate-800/50 text-slate-400 hover:text-slate-200 hover:bg-slate-800'
                  }`}
                >
                  {t.icon}
                  <span className="text-[11px] whitespace-nowrap">{t.label}</span>
                </button>
              ))}
            </div>
          </div>

          {/* Grid layout for sheet and dimensions */}
          <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
            {/* Sheet */}
            <div>
              <label className="block text-xs font-medium text-slate-300 mb-1">Workbook Sheet</label>
              <select
                value={sheetName}
                onChange={(e) => setSheetName(e.target.value)}
                className="w-full bg-slate-800 border border-slate-700 rounded-lg px-2.5 py-1.5 text-xs text-slate-200 focus:outline-none focus:border-indigo-500"
              >
                {dataset.sheets.map((s) => (
                  <option key={s.name} value={s.name}>{s.name}</option>
                ))}
              </select>
            </div>

            {/* X Axis */}
            <div>
              <label className="block text-xs font-medium text-slate-300 mb-1">Dimension (X Axis)</label>
              <select
                value={xAxis}
                onChange={(e) => setXAxis(e.target.value)}
                className="w-full bg-slate-800 border border-slate-700 rounded-lg px-2.5 py-1.5 text-xs text-slate-200 focus:outline-none focus:border-indigo-500"
              >
                {columns.map((c) => (
                  <option key={c} value={c}>{c}</option>
                ))}
              </select>
            </div>

            {/* Y Axis */}
            <div>
              <label className="block text-xs font-medium text-slate-300 mb-1">Metric (Y Axis)</label>
              <select
                value={yAxis}
                onChange={(e) => setYAxis(e.target.value)}
                className="w-full bg-slate-800 border border-slate-700 rounded-lg px-2.5 py-1.5 text-xs text-slate-200 focus:outline-none focus:border-indigo-500"
              >
                {columns.map((c) => (
                  <option key={c} value={c}>{c}</option>
                ))}
              </select>
            </div>
          </div>

          {/* Aggregation & Width */}
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 pt-1">
            <div>
              <label className="block text-xs font-medium text-slate-300 mb-1">Formula Aggregation</label>
              <div className="flex items-center gap-1 bg-slate-800 p-1 rounded-lg border border-slate-700">
                {(['sum', 'avg', 'count', 'min', 'max'] as const).map((agg) => (
                  <button
                    type="button"
                    key={agg}
                    onClick={() => setAggregation(agg)}
                    className={`flex-1 py-1 text-[11px] font-mono uppercase font-semibold rounded transition-colors ${
                      aggregation === agg ? 'bg-indigo-600 text-white' : 'text-slate-400 hover:text-slate-200'
                    }`}
                  >
                    {agg}
                  </button>
                ))}
              </div>
            </div>

            <div>
              <label className="block text-xs font-medium text-slate-300 mb-1">Dashboard Grid Width</label>
              <div className="flex items-center gap-1 bg-slate-800 p-1 rounded-lg border border-slate-700">
                {(['1/3', '1/2', '2/3', 'full'] as const).map((w) => (
                  <button
                    type="button"
                    key={w}
                    onClick={() => setWidth(w)}
                    className={`flex-1 py-1 text-xs font-mono rounded transition-colors ${
                      width === w ? 'bg-indigo-600 text-white font-semibold' : 'text-slate-400 hover:text-slate-200'
                    }`}
                  >
                    {w}
                  </button>
                ))}
              </div>
            </div>
          </div>

          {/* Color Theme */}
          <div>
            <label className="block text-xs font-medium text-slate-300 mb-1">Color Palette</label>
            <div className="flex items-center gap-2">
              {(
                [
                  { id: 'emerald', bg: 'bg-emerald-500' },
                  { id: 'indigo', bg: 'bg-indigo-500' },
                  { id: 'amber', bg: 'bg-amber-500' },
                  { id: 'cyan', bg: 'bg-cyan-500' },
                  { id: 'rose', bg: 'bg-rose-500' },
                  { id: 'slate', bg: 'bg-slate-500' }
                ] as const
              ).map((c) => (
                <button
                  type="button"
                  key={c.id}
                  onClick={() => setColorTheme(c.id)}
                  className={`w-7 h-7 rounded-full flex items-center justify-center transition-all ${c.bg} ${
                    colorTheme === c.id ? 'ring-2 ring-white scale-110 shadow-md' : 'opacity-60 hover:opacity-100'
                  }`}
                />
              ))}
            </div>
          </div>

          {/* Action Buttons */}
          <div className="flex items-center justify-between gap-3 pt-3 border-t border-slate-800">
            <span className="text-[11px] text-slate-500">
              Developed by <span className="text-indigo-400 font-medium">Ammar Thaqif</span>
            </span>
            <div className="flex items-center gap-3">
              <button
                type="button"
                onClick={onClose}
                className="px-4 py-2 text-xs font-medium text-slate-400 hover:text-white transition-colors"
              >
                Cancel
              </button>
              <button
                type="submit"
                className="px-4 py-2 text-xs font-semibold text-white bg-indigo-600 hover:bg-indigo-500 rounded-lg shadow-sm transition-colors"
              >
                Pin Widget to Dashboard
              </button>
            </div>
          </div>
        </form>
      </div>
    </div>
  );
};
