/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useMemo } from 'react';
import {
  Dataset,
  ReportingWidget,
  LegacyMacro,
  WidgetWidth
} from './types/dashboard';
import { SAMPLE_DATASETS, INITIAL_WIDGETS } from './data/sampleDatasets';
import {
  getLocalDatasets,
  saveLocalDatasets,
  getLocalWidgets,
  saveLocalWidgets,
  resetWidgetLayoutStorage
} from './utils/localStorage';

// Visualization components
import { MetricCard } from './components/charts/MetricCard';
import { TrendLineChart } from './components/charts/TrendLineChart';
import { BarGroupChart } from './components/charts/BarGroupChart';
import { DonutChart } from './components/charts/DonutChart';
import { PivotTable } from './components/charts/PivotTable';
import { SpreadsheetGrid } from './components/spreadsheet/SpreadsheetGrid';

// Modals & Panels
import { WidgetStudioModal } from './components/widgets/WidgetStudioModal';
import { LegacyMacroStudio } from './components/macros/LegacyMacroStudio';
import { ExcelUploadModal } from './components/upload/ExcelUploadModal';
import { ExportPdfModal } from './components/modals/ExportPdfModal';
import { useTheme } from './context/ThemeContext';

// Icons
import {
  Plus,
  Upload,
  FileSpreadsheet,
  Menu,
  X,
  Code2,
  Download,
  Check,
  RotateCcw,
  CheckCircle2,
  Save,
  LayoutGrid,
  Sun,
  Moon,
  FileText
} from 'lucide-react';
import { exportSheetToCSV, exportAllSheetsToCSVZip } from './utils/excelParser';

export default function App() {
  const { theme, toggleTheme, isLight } = useTheme();

  // 1. Core State
  const [dataset, setDataset] = useState<Dataset>(() => {
    const cached = getLocalDatasets();
    return cached && cached.length > 0 ? cached[0] : SAMPLE_DATASETS[0];
  });

  const [activeSheetName, setActiveSheetName] = useState<string>(
    dataset.activeSheetName || dataset.sheets[0]?.name || 'Sheet1'
  );

  const [widgets, setWidgets] = useState<ReportingWidget[]>(() => {
    const cached = getLocalWidgets();
    return cached && cached.length > 0 ? cached : INITIAL_WIDGETS;
  });

  const [activeTab, setActiveTab] = useState<'dashboard' | 'spreadsheet' | 'pivot' | 'macros'>('dashboard');

  // 2. Modals, Notifications & Layout State
  const [uploadModalOpen, setUploadModalOpen] = useState(false);
  const [widgetModalOpen, setWidgetModalOpen] = useState(false);
  const [exportPdfOpen, setExportPdfOpen] = useState(false);
  const [mobileMenuOpen, setMobileMenuOpen] = useState(false);
  const [csvDownloadToast, setCsvDownloadToast] = useState<{ message: string; filename: string; rows: number } | null>(null);
  const [layoutToast, setLayoutToast] = useState<string | null>(null);
  const [draggedWidgetId, setDraggedWidgetId] = useState<string | null>(null);
  const [dragOverWidgetId, setDragOverWidgetId] = useState<string | null>(null);

  // Download active spreadsheet data as CSV
  const handleDownloadCSV = (allSheets: boolean = false) => {
    try {
      if (allSheets && dataset.sheets.length > 1) {
        exportAllSheetsToCSVZip(dataset).then((res) => {
          setCsvDownloadToast({
            message: `Exported all ${res.sheetCount} sheets to CSV archive`,
            filename: res.filename,
            rows: dataset.sheets.reduce((acc, s) => acc + (s.rowCount || s.rows.length), 0)
          });
          setTimeout(() => setCsvDownloadToast(null), 4000);
        });
        return;
      }

      const sheet = dataset.sheets.find((s) => s.name === activeSheetName) || dataset.sheets[0];
      if (!sheet) return;
      const { filename, rowCount } = exportSheetToCSV(sheet, `${dataset.name}_${sheet.name}.csv`);
      setCsvDownloadToast({
        message: `Successfully extracted ${sheet.name} to CSV`,
        filename,
        rows: rowCount
      });
      setTimeout(() => setCsvDownloadToast(null), 4000);
    } catch (err: any) {
      console.error('Error downloading CSV:', err);
    }
  };

  // Ingest new dataset
  const handleDatasetLoaded = (newDataset: Dataset) => {
    setDataset(newDataset);
    setActiveSheetName(newDataset.activeSheetName || newDataset.sheets[0]?.name || 'Sheet1');
    saveLocalDatasets([newDataset]);

    // Create auto-generated widgets matching new sheet columns
    const firstSheet = newDataset.sheets[0];
    const newCols = firstSheet?.columns || [];
    const numCols = newCols.filter(
      (c) => newDataset.columnMetadata[c]?.type === 'currency' || newDataset.columnMetadata[c]?.type === 'number'
    );
    const catCols = newCols.filter((c) => newDataset.columnMetadata[c]?.type === 'text' || newDataset.columnMetadata[c]?.type === 'date');

    const autoWidgets: ReportingWidget[] = [];
    if (numCols[0]) {
      autoWidgets.push({
        id: `w-auto-metric-${Date.now()}`,
        title: `Total ${numCols[0]}`,
        type: 'metric',
        sheetName: firstSheet.name,
        xAxis: catCols[0] || newCols[0],
        yAxis: numCols[0],
        aggregation: 'sum',
        width: '1/3',
        colorTheme: 'emerald'
      });
    }
    if (numCols[1] || numCols[0]) {
      autoWidgets.push({
        id: `w-auto-trend-${Date.now()}`,
        title: `${numCols[1] || numCols[0]} Trajectory`,
        type: 'line',
        sheetName: firstSheet.name,
        xAxis: catCols[0] || newCols[0],
        yAxis: numCols[1] || numCols[0],
        aggregation: 'sum',
        width: '2/3',
        colorTheme: 'indigo'
      });
    }
    if (catCols[0] && numCols[0]) {
      autoWidgets.push({
        id: `w-auto-donut-${Date.now()}`,
        title: `Share by ${catCols[0]}`,
        type: 'donut',
        sheetName: firstSheet.name,
        xAxis: catCols[0],
        yAxis: numCols[0],
        aggregation: 'sum',
        width: '1/3',
        colorTheme: 'amber'
      });
      autoWidgets.push({
        id: `w-auto-bar-${Date.now()}`,
        title: `Distribution of ${numCols[0]}`,
        type: 'bar',
        sheetName: firstSheet.name,
        xAxis: catCols[0],
        yAxis: numCols[0],
        aggregation: 'sum',
        width: '2/3',
        colorTheme: 'slate'
      });
    }

    if (autoWidgets.length > 0) {
      setWidgets(autoWidgets);
      saveLocalWidgets(autoWidgets);
    }
  };

  // Add Custom Reporting Widget
  const handleSaveWidget = (widget: ReportingWidget) => {
    const updated = [...widgets, widget];
    setWidgets(updated);
    saveLocalWidgets(updated);
    setLayoutToast(`Added "${widget.title}" to dashboard`);
    setTimeout(() => setLayoutToast(null), 3000);
  };

  // Delete Widget
  const handleDeleteWidget = (id: string) => {
    const target = widgets.find((w) => w.id === id);
    const updated = widgets.filter((w) => w.id !== id);
    setWidgets(updated);
    saveLocalWidgets(updated);
    if (target) {
      setLayoutToast(`Removed "${target.title}" from dashboard`);
      setTimeout(() => setLayoutToast(null), 3000);
    }
  };

  // Reorder widgets and immediately persist to localStorage
  const handleReorderWidget = (sourceId: string, targetId: string) => {
    if (sourceId === targetId) return;
    const sourceIdx = widgets.findIndex((w) => w.id === sourceId);
    const targetIdx = widgets.findIndex((w) => w.id === targetId);
    if (sourceIdx === -1 || targetIdx === -1) return;

    const updated = [...widgets];
    const [moved] = updated.splice(sourceIdx, 1);
    updated.splice(targetIdx, 0, moved);

    setWidgets(updated);
    saveLocalWidgets(updated);
    setLayoutToast(`Reordered "${moved.title}" in dashboard layout`);
    setTimeout(() => setLayoutToast(null), 3000);
  };

  // Move widget one step left/up or right/down within active subset
  const handleMoveWidgetStep = (id: string, direction: 'left' | 'right', isMetric: boolean) => {
    const subset = isMetric
      ? widgets.filter((w) => w.type === 'metric' && (w.sheetName === activeSheetName || !w.sheetName))
      : widgets.filter((w) => w.type !== 'metric' && w.type !== 'table' && (w.sheetName === activeSheetName || !w.sheetName));

    const subIdx = subset.findIndex((w) => w.id === id);
    if (subIdx === -1) return;

    const targetSubIdx = direction === 'left' ? subIdx - 1 : subIdx + 1;
    if (targetSubIdx < 0 || targetSubIdx >= subset.length) return;

    const targetId = subset[targetSubIdx].id;
    handleReorderWidget(id, targetId);
  };

  // Change widget layout width (1/3, 1/2, 2/3, full) and persist to localStorage
  const handleChangeWidgetWidth = (id: string, width: WidgetWidth) => {
    const updated = widgets.map((w) => (w.id === id ? { ...w, width } : w));
    setWidgets(updated);
    saveLocalWidgets(updated);
    setLayoutToast(`Updated layout width to ${width}`);
    setTimeout(() => setLayoutToast(null), 3000);
  };

  // Reset layout to default arrangement and persist
  const handleResetLayout = () => {
    resetWidgetLayoutStorage();
    setWidgets(INITIAL_WIDGETS);
    saveLocalWidgets(INITIAL_WIDGETS);
    setLayoutToast('Restored default dashboard layout and widget sizes');
    setTimeout(() => setLayoutToast(null), 3500);
  };

  // Update macros when modified or enhanced in Macro Studio
  const handleUpdateMacros = (updatedMacros: LegacyMacro[]) => {
    setDataset((prev) => {
      const updated = {
        ...prev,
        legacyMacros: updatedMacros
      };
      saveLocalDatasets([updated]);
      return updated;
    });
  };

  // Update dataset state and persist to local storage (e.g. from macro simulation)
  const handleUpdateDataset = (updatedDataset: Dataset) => {
    setDataset(updatedDataset);
    saveLocalDatasets([updatedDataset]);
  };

  // Filter widgets by active sheet
  const activeSheetWidgets = useMemo(() => {
    return widgets.filter((w) => w.sheetName === activeSheetName || !w.sheetName);
  }, [widgets, activeSheetName]);

  const metricWidgets = useMemo(() => {
    return activeSheetWidgets.filter((w) => w.type === 'metric');
  }, [activeSheetWidgets]);

  const chartWidgets = useMemo(() => {
    return activeSheetWidgets.filter((w) => w.type !== 'metric' && w.type !== 'table');
  }, [activeSheetWidgets]);

  return (
    <div
      className={`min-h-screen flex flex-col selection:bg-indigo-500 selection:text-white transition-colors duration-200 ${
        isLight ? 'bg-slate-50 text-slate-900' : 'bg-slate-900 text-slate-100'
      }`}
    >
      {/* Top Navigation Bar */}
      <header
        className={`flex items-center justify-between gap-8 px-6 py-3.5 border-b backdrop-blur-md sticky top-0 z-40 transition-colors ${
          isLight
            ? 'border-slate-200 bg-white/90 shadow-sm'
            : 'border-slate-800 bg-slate-950/70'
        }`}
      >
        {/* Zone 1: Brand title wordmark & Author attribution */}
        <div className="flex items-center gap-3 shrink-0">
          <div className="p-1.5 rounded-lg bg-indigo-600 text-white shadow-sm">
            <FileSpreadsheet className="w-5 h-5" />
          </div>
          <span
            className={`text-base font-bold tracking-tight whitespace-nowrap shrink-0 ${
              isLight ? 'text-slate-900' : 'text-white'
            }`}
          >
            XLSM Cloud Studio
          </span>
          <span
            className={`hidden sm:inline-flex items-center gap-1.5 px-2.5 py-0.5 rounded-full text-[11px] font-medium border ${
              isLight
                ? 'bg-slate-100 border-slate-200 text-slate-600'
                : 'bg-slate-900/90 border-slate-800 text-slate-400'
            }`}
          >
            Developed by <span className="text-indigo-500 font-semibold">Ammar Thaqif</span>
          </span>
        </div>

        {/* Zone 2: Concise single-line navigation tabs */}
        <nav
          className={`hidden md:flex items-center gap-6 text-sm font-medium ${
            isLight ? 'text-slate-600' : 'text-slate-400'
          }`}
        >
          <button
            onClick={() => setActiveTab('dashboard')}
            className={`whitespace-nowrap shrink-0 transition-colors pb-0.5 ${
              activeTab === 'dashboard'
                ? isLight
                  ? 'text-indigo-600 font-semibold border-b-2 border-indigo-600'
                  : 'text-white font-semibold border-b-2 border-indigo-500'
                : isLight
                ? 'hover:text-slate-900'
                : 'hover:text-slate-200'
            }`}
          >
            Dashboard
          </button>
          <button
            onClick={() => setActiveTab('spreadsheet')}
            className={`whitespace-nowrap shrink-0 transition-colors pb-0.5 ${
              activeTab === 'spreadsheet'
                ? isLight
                  ? 'text-indigo-600 font-semibold border-b-2 border-indigo-600'
                  : 'text-white font-semibold border-b-2 border-indigo-500'
                : isLight
                ? 'hover:text-slate-900'
                : 'hover:text-slate-200'
            }`}
          >
            Spreadsheet
          </button>
          <button
            onClick={() => setActiveTab('pivot')}
            className={`whitespace-nowrap shrink-0 transition-colors pb-0.5 ${
              activeTab === 'pivot'
                ? isLight
                  ? 'text-indigo-600 font-semibold border-b-2 border-indigo-600'
                  : 'text-white font-semibold border-b-2 border-indigo-500'
                : isLight
                ? 'hover:text-slate-900'
                : 'hover:text-slate-200'
            }`}
          >
            Pivot Matrix
          </button>
          <button
            onClick={() => setActiveTab('macros')}
            className={`whitespace-nowrap shrink-0 transition-colors pb-0.5 flex items-center gap-1.5 ${
              activeTab === 'macros'
                ? isLight
                  ? 'text-indigo-600 font-semibold border-b-2 border-indigo-600'
                  : 'text-white font-semibold border-b-2 border-indigo-500'
                : isLight
                ? 'hover:text-slate-900'
                : 'hover:text-slate-200'
            }`}
          >
            <Code2 className="w-4 h-4 text-indigo-500" />
            <span>Macro Studio</span>
            <span
              className={`text-[11px] font-mono px-1.5 py-0.2 rounded border ${
                isLight
                  ? 'bg-indigo-50 text-indigo-700 border-indigo-200'
                  : 'bg-indigo-950 text-indigo-300 border-indigo-500/30'
              }`}
            >
              {dataset.legacyMacros.length}
            </span>
          </button>
        </nav>

        {/* Zone 3: Primary Action + Theme Toggle + Quick Controls */}
        <div className="flex items-center gap-2.5 shrink-0">
          {/* Theme Toggle Button */}
          <button
            onClick={toggleTheme}
            className={`flex items-center gap-1.5 px-2.5 py-1.5 text-xs font-medium rounded-lg border transition-all ${
              isLight
                ? 'bg-slate-100 hover:bg-slate-200 border-slate-300 text-slate-800'
                : 'bg-slate-800/80 hover:bg-slate-700 border-slate-700 text-slate-200'
            }`}
            title={isLight ? 'Switch to Deep Slate Dark Theme' : 'Switch to High-Contrast Light Theme'}
          >
            {isLight ? (
              <Moon className="w-3.5 h-3.5 text-indigo-600" />
            ) : (
              <Sun className="w-3.5 h-3.5 text-amber-400" />
            )}
            <span className="hidden sm:inline font-mono text-[11px]">
              {isLight ? 'Dark' : 'Light'}
            </span>
          </button>

          {/* Primary Action Button: Ingest Excel */}
          <button
            onClick={() => setUploadModalOpen(true)}
            className="flex items-center gap-1.5 px-3.5 py-2 text-xs font-semibold text-white bg-indigo-600 hover:bg-indigo-500 rounded-lg shadow-sm transition-all whitespace-nowrap shrink-0"
          >
            <Upload className="w-3.5 h-3.5" />
            <span>Ingest Excel / XLSM</span>
          </button>

          {/* Mobile menu toggle */}
          <button
            onClick={() => setMobileMenuOpen(!mobileMenuOpen)}
            className={`md:hidden p-2 rounded-lg transition-colors ${
              isLight ? 'text-slate-600 hover:bg-slate-100' : 'text-slate-400 hover:text-white'
            }`}
          >
            {mobileMenuOpen ? <X className="w-5 h-5" /> : <Menu className="w-5 h-5" />}
          </button>
        </div>
      </header>

      {/* Mobile Drawer Navigation */}
      {mobileMenuOpen && (
        <div
          className={`md:hidden border-b p-4 space-y-2 animate-in slide-in-from-top-2 ${
            isLight ? 'bg-white border-slate-200' : 'bg-slate-950 border-slate-800'
          }`}
        >
          <div className="flex items-center justify-between pb-2 mb-2 border-b border-slate-200 dark:border-slate-800">
            <span className="text-xs font-semibold text-slate-500 uppercase">Navigation</span>
            <button
              onClick={toggleTheme}
              className={`flex items-center gap-1.5 px-2.5 py-1 text-xs font-medium rounded-md border ${
                isLight ? 'bg-slate-100 border-slate-200 text-slate-800' : 'bg-slate-800 border-slate-700 text-slate-200'
              }`}
            >
              {isLight ? <Moon className="w-3.5 h-3.5 text-indigo-600" /> : <Sun className="w-3.5 h-3.5 text-amber-400" />}
              <span>{isLight ? 'Switch to Dark' : 'Switch to Light'}</span>
            </button>
          </div>
          <button
            onClick={() => {
              setActiveTab('dashboard');
              setMobileMenuOpen(false);
            }}
            className={`w-full text-left py-2 px-3 rounded-lg text-sm font-medium ${
              activeTab === 'dashboard'
                ? isLight
                  ? 'bg-indigo-50 text-indigo-700 font-semibold'
                  : 'bg-slate-800 text-white'
                : isLight
                ? 'text-slate-600 hover:bg-slate-100'
                : 'text-slate-400'
            }`}
          >
            Dashboard
          </button>
          <button
            onClick={() => {
              setActiveTab('spreadsheet');
              setMobileMenuOpen(false);
            }}
            className={`w-full text-left py-2 px-3 rounded-lg text-sm font-medium ${
              activeTab === 'spreadsheet'
                ? isLight
                  ? 'bg-indigo-50 text-indigo-700 font-semibold'
                  : 'bg-slate-800 text-white'
                : isLight
                ? 'text-slate-600 hover:bg-slate-100'
                : 'text-slate-400'
            }`}
          >
            Spreadsheet
          </button>
          <button
            onClick={() => {
              setActiveTab('pivot');
              setMobileMenuOpen(false);
            }}
            className={`w-full text-left py-2 px-3 rounded-lg text-sm font-medium ${
              activeTab === 'pivot'
                ? isLight
                  ? 'bg-indigo-50 text-indigo-700 font-semibold'
                  : 'bg-slate-800 text-white'
                : isLight
                ? 'text-slate-600 hover:bg-slate-100'
                : 'text-slate-400'
            }`}
          >
            Pivot Matrix
          </button>
          <button
            onClick={() => {
              setActiveTab('macros');
              setMobileMenuOpen(false);
            }}
            className={`w-full text-left py-2 px-3 rounded-lg text-sm font-medium ${
              activeTab === 'macros'
                ? isLight
                  ? 'bg-indigo-50 text-indigo-700 font-semibold'
                  : 'bg-slate-800 text-white'
                : isLight
                ? 'text-slate-600 hover:bg-slate-100'
                : 'text-slate-400'
            }`}
          >
            Macro Studio ({dataset.legacyMacros.length})
          </button>
          <div
            className={`pt-3 border-t px-3 text-xs ${
              isLight ? 'border-slate-200 text-slate-600' : 'border-slate-800/80 text-slate-400'
            }`}
          >
            <div>
              Developed by <span className="text-indigo-600 dark:text-indigo-400 font-semibold">Ammar Thaqif</span>
            </div>
            <a
              href="mailto:ammarthaqif.ar@gmail.com"
              className={`text-[11px] font-mono transition-colors ${
                isLight ? 'text-slate-500 hover:text-indigo-600' : 'text-slate-500 hover:text-indigo-300'
              }`}
            >
              ammarthaqif.ar@gmail.com
            </a>
          </div>
        </div>
      )}

      {/* Main Viewport Container */}
      <main className="flex-1 max-w-7xl w-full mx-auto p-4 sm:p-6 lg:p-8 space-y-6">
        {/* Workspace Context Kicker Bar */}
        <div
          className={`flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-2 border-b transition-colors ${
            isLight ? 'border-slate-200' : 'border-slate-800/80'
          }`}
        >
          <div>
            <div className="flex flex-wrap items-center gap-2">
              <h1
                className={`text-xl sm:text-2xl font-bold tracking-tight font-sans ${
                  isLight ? 'text-slate-900' : 'text-white'
                }`}
              >
                {dataset.name}
              </h1>
              <span className="text-xs font-mono text-emerald-400 bg-emerald-950/60 px-2 py-0.5 rounded border border-emerald-500/30">
                .xlsm converted
              </span>
              <button
                onClick={() => setActiveTab('macros')}
                className={`text-xs font-mono px-2.5 py-0.5 rounded border transition-colors ${
                  isLight
                    ? 'text-indigo-700 bg-indigo-50 border-indigo-200 hover:bg-indigo-100'
                    : 'text-indigo-300 hover:text-white bg-indigo-950/70 hover:bg-indigo-900/80 border-indigo-500/40'
                }`}
                title="View all extracted VBA macros from vbaProject.bin"
              >
                {dataset.legacyMacros.length} Embedded Macros Extracted
              </button>
              <span
                className={`text-xs font-mono px-2.5 py-0.5 rounded border ${
                  isLight
                    ? 'text-emerald-700 bg-emerald-50 border-emerald-200'
                    : 'text-emerald-300 bg-emerald-950/70 border-emerald-500/40'
                }`}
                title="All sheet protection passwords and workbook lock tags were stripped"
              >
                Sheets Unprotected & Unlocked
              </span>
            </div>
            <div
              className={`mt-1.5 flex flex-wrap items-center gap-2 text-xs font-mono ${
                isLight ? 'text-slate-500' : 'text-slate-400'
              }`}
            >
              <span>{dataset.filename}</span>
              <span aria-hidden="true">·</span>
              <span>{dataset.sheets.length} Sheets</span>
              <span aria-hidden="true">·</span>
              <span>Active: {activeSheetName}</span>
              <span aria-hidden="true">·</span>
              <span>Updated {new Date(dataset.uploadedAt).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}</span>
            </div>
          </div>

          {/* Sub-actions toolbar */}
          <div className="flex items-center gap-2 flex-wrap">
            {/* Sheet selector */}
            <select
              value={activeSheetName}
              onChange={(e) => setActiveSheetName(e.target.value)}
              className={`border rounded-lg px-3 py-1.5 text-xs font-medium focus:outline-none focus:border-indigo-500 ${
                isLight
                  ? 'bg-white border-slate-300 text-slate-800'
                  : 'bg-slate-800/90 border-slate-700 text-slate-200'
              }`}
            >
              {dataset.sheets.map((s) => (
                <option key={s.name} value={s.name}>
                  Sheet: {s.name} ({s.rowCount} rows)
                </option>
              ))}
            </select>

            {/* Layout Persisted Status Badge */}
            <span
              className={`hidden sm:inline-flex items-center gap-1.5 text-[11px] font-mono px-2.5 py-1.5 rounded-lg border shadow-sm ${
                isLight
                  ? 'text-emerald-700 bg-emerald-50 border-emerald-200'
                  : 'text-emerald-400 bg-emerald-950/60 border-emerald-500/30'
              }`}
              title="Widget order, widths, and views are automatically persisted to localStorage across sessions"
            >
              <Save className="w-3.5 h-3.5 text-emerald-500" />
              <span>Layout Persisted</span>
            </span>

            {/* Reset Layout Button */}
            <button
              onClick={handleResetLayout}
              className={`flex items-center gap-1.5 px-2.5 py-1.5 text-xs font-medium border rounded-lg shadow-sm transition-colors ${
                isLight
                  ? 'text-slate-700 bg-white hover:bg-slate-100 hover:text-slate-900 border-slate-300'
                  : 'text-slate-300 bg-slate-800 hover:bg-slate-700 hover:text-white border-slate-700'
              }`}
              title="Reset dashboard widgets and layout ordering to default arrangement"
            >
              <RotateCcw className={`w-3.5 h-3.5 ${isLight ? 'text-slate-500' : 'text-slate-400'}`} />
              <span>Reset Layout</span>
            </button>

            {/* Download CSV Action Button */}
            <button
              onClick={() => handleDownloadCSV(false)}
              className={`flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold rounded-lg shadow-sm transition-all border ${
                isLight
                  ? 'text-emerald-800 bg-emerald-100/80 hover:bg-emerald-200/80 border-emerald-300'
                  : 'text-emerald-300 bg-emerald-950/80 hover:bg-emerald-900 border-emerald-500/40 hover:text-white'
              }`}
              title={`Extract and download active sheet (${activeSheetName}) as standard RFC-4180 CSV`}
            >
              <Download className="w-3.5 h-3.5 text-emerald-500" />
              <span>Download CSV</span>
            </button>

            {/* Export PDF Report Action Button */}
            <button
              onClick={() => setExportPdfOpen(true)}
              className={`flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold rounded-lg shadow-sm transition-all border ${
                isLight
                  ? 'text-indigo-700 bg-indigo-50 hover:bg-indigo-100 border-indigo-200'
                  : 'text-indigo-300 bg-indigo-950/80 hover:bg-indigo-900 border-indigo-500/40 hover:text-white'
              }`}
              title="Export full executive dashboard report with all active charts & KPIs as a professional PDF"
            >
              <FileText className="w-3.5 h-3.5 text-indigo-400" />
              <span>Export PDF</span>
            </button>

            {/* Custom Widget Builder trigger */}
            <button
              onClick={() => setWidgetModalOpen(true)}
              className={`flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold rounded-lg shadow-sm transition-colors border ${
                isLight
                  ? 'text-slate-700 bg-white hover:bg-slate-100 border-slate-300'
                  : 'text-slate-200 bg-slate-800 hover:bg-slate-700 border-slate-700'
              }`}
            >
              <Plus className="w-3.5 h-3.5" />
              <span>Add Widget</span>
            </button>
          </div>
        </div>

        {/* Layout Customization Feedback Notification */}
        {layoutToast && (
          <div className="flex items-center justify-between gap-3 px-4 py-2.5 rounded-lg bg-indigo-950/90 border border-indigo-500/50 text-indigo-200 text-xs shadow-lg animate-in fade-in slide-in-from-top-1">
            <div className="flex items-center gap-2.5">
              <span className="p-1 rounded-full bg-indigo-500/20 text-indigo-400">
                <CheckCircle2 className="w-3.5 h-3.5" />
              </span>
              <div>
                <span className="font-semibold text-white">{layoutToast}</span>
                <span className="mx-2 text-indigo-400/70">·</span>
                <span className="text-slate-300">Custom view retained across sessions</span>
              </div>
            </div>
            <button
              onClick={() => setLayoutToast(null)}
              className="text-indigo-400 hover:text-white transition-colors p-1"
              title="Dismiss"
            >
              <X className="w-3.5 h-3.5" />
            </button>
          </div>
        )}

        {/* CSV Download Feedback Notification */}
        {csvDownloadToast && (
          <div className="flex items-center justify-between gap-3 px-4 py-2.5 rounded-lg bg-emerald-950/90 border border-emerald-500/50 text-emerald-200 text-xs shadow-lg animate-in fade-in slide-in-from-top-1">
            <div className="flex items-center gap-2.5">
              <span className="p-1 rounded-full bg-emerald-500/20 text-emerald-400">
                <Check className="w-3.5 h-3.5" />
              </span>
              <div>
                <span className="font-semibold text-emerald-100">{csvDownloadToast.message}</span>
                <span className="mx-2 text-emerald-500/70">·</span>
                <span className="font-mono text-emerald-300">{csvDownloadToast.filename}</span>
                <span className="ml-1.5 text-[11px] text-emerald-400 font-mono">({csvDownloadToast.rows} rows)</span>
              </div>
            </div>
            <button
              onClick={() => setCsvDownloadToast(null)}
              className="text-emerald-400 hover:text-emerald-100 transition-colors p-1"
              title="Dismiss"
            >
              <X className="w-3.5 h-3.5" />
            </button>
          </div>
        )}

        {/* Tab 1: Live Interactive Dashboard */}
        {activeTab === 'dashboard' && (
          <div id="dashboard-export-container" className="space-y-6">
            {/* KPI Metrics Row with Drag & Drop and Order Controls */}
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
              {metricWidgets.map((w, idx) => (
                <div
                  key={w.id}
                  draggable
                  onDragStart={(e) => {
                    setDraggedWidgetId(w.id);
                    e.dataTransfer.setData('text/plain', w.id);
                  }}
                  onDragOver={(e) => {
                    e.preventDefault();
                    if (dragOverWidgetId !== w.id) setDragOverWidgetId(w.id);
                  }}
                  onDragLeave={() => {
                    if (dragOverWidgetId === w.id) setDragOverWidgetId(null);
                  }}
                  onDrop={(e) => {
                    e.preventDefault();
                    setDragOverWidgetId(null);
                    if (draggedWidgetId && draggedWidgetId !== w.id) {
                      handleReorderWidget(draggedWidgetId, w.id);
                    }
                    setDraggedWidgetId(null);
                  }}
                  onDragEnd={() => {
                    setDraggedWidgetId(null);
                    setDragOverWidgetId(null);
                  }}
                  className={`transition-all duration-200 ${
                    draggedWidgetId === w.id ? 'opacity-40 scale-[0.98]' : ''
                  } ${
                    dragOverWidgetId === w.id ? 'ring-2 ring-indigo-500 rounded-xl shadow-lg bg-indigo-950/30' : ''
                  }`}
                >
                  <MetricCard
                    widget={w}
                    dataset={dataset}
                    onDeleteWidget={handleDeleteWidget}
                    onMoveLeft={() => handleMoveWidgetStep(w.id, 'left', true)}
                    onMoveRight={() => handleMoveWidgetStep(w.id, 'right', true)}
                    canMoveLeft={idx > 0}
                    canMoveRight={idx < metricWidgets.length - 1}
                  />
                </div>
              ))}
            </div>

            {/* Visual Charts Grid with Drag & Drop, Layout Width Resizing & Order Controls */}
            <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
              {chartWidgets.map((w, idx) => {
                const colSpan =
                  w.width === 'full'
                    ? 'lg:col-span-12'
                    : w.width === '2/3'
                    ? 'lg:col-span-8'
                    : w.width === '1/3'
                    ? 'lg:col-span-4'
                    : 'lg:col-span-6';

                return (
                  <div
                    key={w.id}
                    draggable
                    onDragStart={(e) => {
                      setDraggedWidgetId(w.id);
                      e.dataTransfer.setData('text/plain', w.id);
                    }}
                    onDragOver={(e) => {
                      e.preventDefault();
                      if (dragOverWidgetId !== w.id) setDragOverWidgetId(w.id);
                    }}
                    onDragLeave={() => {
                      if (dragOverWidgetId === w.id) setDragOverWidgetId(null);
                    }}
                    onDrop={(e) => {
                      e.preventDefault();
                      setDragOverWidgetId(null);
                      if (draggedWidgetId && draggedWidgetId !== w.id) {
                        handleReorderWidget(draggedWidgetId, w.id);
                      }
                      setDraggedWidgetId(null);
                    }}
                    onDragEnd={() => {
                      setDraggedWidgetId(null);
                      setDragOverWidgetId(null);
                    }}
                    className={`${colSpan} transition-all duration-200 ${
                      draggedWidgetId === w.id ? 'opacity-40 scale-[0.98]' : ''
                    } ${
                      dragOverWidgetId === w.id ? 'ring-2 ring-indigo-500 rounded-xl shadow-lg bg-indigo-950/30' : ''
                    }`}
                  >
                    {w.type === 'line' && (
                      <TrendLineChart
                        widget={w}
                        dataset={dataset}
                        onDeleteWidget={handleDeleteWidget}
                        onMoveLeft={() => handleMoveWidgetStep(w.id, 'left', false)}
                        onMoveRight={() => handleMoveWidgetStep(w.id, 'right', false)}
                        canMoveLeft={idx > 0}
                        canMoveRight={idx < chartWidgets.length - 1}
                        onChangeWidth={(newWidth) => handleChangeWidgetWidth(w.id, newWidth)}
                      />
                    )}
                    {w.type === 'bar' && (
                      <BarGroupChart
                        widget={w}
                        dataset={dataset}
                        onDeleteWidget={handleDeleteWidget}
                        onMoveLeft={() => handleMoveWidgetStep(w.id, 'left', false)}
                        onMoveRight={() => handleMoveWidgetStep(w.id, 'right', false)}
                        canMoveLeft={idx > 0}
                        canMoveRight={idx < chartWidgets.length - 1}
                        onChangeWidth={(newWidth) => handleChangeWidgetWidth(w.id, newWidth)}
                      />
                    )}
                    {w.type === 'donut' && (
                      <DonutChart
                        widget={w}
                        dataset={dataset}
                        onDeleteWidget={handleDeleteWidget}
                        onMoveLeft={() => handleMoveWidgetStep(w.id, 'left', false)}
                        onMoveRight={() => handleMoveWidgetStep(w.id, 'right', false)}
                        canMoveLeft={idx > 0}
                        canMoveRight={idx < chartWidgets.length - 1}
                        onChangeWidth={(newWidth) => handleChangeWidgetWidth(w.id, newWidth)}
                      />
                    )}
                  </div>
                );
              })}
            </div>

            {/* Quick Spreadsheet Preview */}
            <div className="space-y-2">
              <div className="flex flex-wrap items-center justify-between gap-2">
                <div className="flex items-center gap-2">
                  <h3
                    className={`text-sm font-semibold ${
                      isLight ? 'text-slate-900' : 'text-white'
                    }`}
                  >
                    Active Spreadsheet Ledger
                  </h3>
                  <span
                    className={`text-xs font-mono ${
                      isLight ? 'text-slate-500' : 'text-slate-400'
                    }`}
                  >
                    ({activeSheetName})
                  </span>
                </div>
                <div className="flex items-center gap-2.5">
                  <button
                    onClick={() => handleDownloadCSV(false)}
                    className={`flex items-center gap-1.5 px-2.5 py-1 text-xs font-medium rounded-lg transition-colors shadow-sm border ${
                      isLight
                        ? 'text-emerald-800 bg-emerald-100/90 hover:bg-emerald-200 border-emerald-300'
                        : 'text-emerald-300 bg-emerald-950/70 hover:bg-emerald-900 border-emerald-500/40'
                    }`}
                    title={`Extract and download ${activeSheetName} as CSV`}
                  >
                    <Download className="w-3.5 h-3.5 text-emerald-500" />
                    <span>Download CSV</span>
                  </button>
                  <button
                    onClick={() => setActiveTab('spreadsheet')}
                    className={`text-xs font-medium transition-colors ${
                      isLight ? 'text-indigo-600 hover:text-indigo-800' : 'text-indigo-400 hover:text-indigo-300'
                    }`}
                  >
                    Open Full Spreadsheet View →
                  </button>
                </div>
              </div>
              <div className="h-80">
                <SpreadsheetGrid
                  dataset={dataset}
                  activeSheetName={activeSheetName}
                  onSelectSheet={setActiveSheetName}
                />
              </div>
            </div>
          </div>
        )}

        {/* Tab 2: Full Spreadsheet Grid */}
        {activeTab === 'spreadsheet' && (
          <div className="h-[750px]">
            <SpreadsheetGrid
              dataset={dataset}
              activeSheetName={activeSheetName}
              onSelectSheet={setActiveSheetName}
            />
          </div>
        )}

        {/* Tab 3: Dynamic Pivot Matrix */}
        {activeTab === 'pivot' && (
          <PivotTable
            dataset={dataset}
            activeSheetName={activeSheetName}
          />
        )}

        {/* Tab 4: Comprehensive Legacy Macro Studio */}
        {activeTab === 'macros' && (
          <LegacyMacroStudio
            dataset={dataset}
            onUpdateMacros={handleUpdateMacros}
            onUpdateDataset={handleUpdateDataset}
            onNavigateToDashboard={() => setActiveTab('dashboard')}
          />
        )}
      </main>

      {/* Footer */}
      <footer
        className={`border-t py-4 px-6 text-xs flex flex-wrap items-center justify-between gap-4 transition-colors ${
          isLight
            ? 'border-slate-200 bg-white text-slate-600 shadow-sm'
            : 'border-slate-800 bg-slate-950 text-slate-500'
        }`}
      >
        <div className="flex flex-wrap items-center gap-2">
          <span className={`font-semibold ${isLight ? 'text-slate-800' : 'text-slate-300'}`}>
            XLSM Cloud Studio
          </span>
          <span>·</span>
          <span>Macro Extraction & Unprotection Engine</span>
          <span>·</span>
          <span className={isLight ? 'text-slate-600' : 'text-slate-400'}>
            Developed by{' '}
            <span className={isLight ? 'text-indigo-600 font-semibold' : 'text-indigo-400 font-semibold'}>
              Ammar Thaqif
            </span>
          </span>
        </div>
        <div className="flex flex-wrap items-center gap-4">
          <span>{dataset.legacyMacros.length} Macros Ingested</span>
          <span>·</span>
          <span>{dataset.sheets.length} Sheets Unlocked</span>
          <span>·</span>
          <a
            href="mailto:ammarthaqif.ar@gmail.com"
            className={`font-mono text-[11px] transition-colors ${
              isLight ? 'text-slate-600 hover:text-indigo-600' : 'text-slate-400 hover:text-indigo-300'
            }`}
          >
            ammarthaqif.ar@gmail.com
          </a>
        </div>
      </footer>

      {/* Modals & Drawers */}
      <ExcelUploadModal
        isOpen={uploadModalOpen}
        onClose={() => setUploadModalOpen(false)}
        onDatasetLoaded={handleDatasetLoaded}
      />

      <WidgetStudioModal
        isOpen={widgetModalOpen}
        onClose={() => setWidgetModalOpen(false)}
        dataset={dataset}
        onSaveWidget={handleSaveWidget}
      />

      <ExportPdfModal
        isOpen={exportPdfOpen}
        onClose={() => setExportPdfOpen(false)}
        dataset={dataset}
        activeSheetName={activeSheetName}
        widgets={widgets}
        targetElementId="dashboard-export-container"
        currentTheme={theme}
        onNotification={(msg) => setLayoutToast(msg)}
      />
    </div>
  );
}
