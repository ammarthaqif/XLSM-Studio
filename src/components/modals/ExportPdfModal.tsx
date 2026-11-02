import React, { useState } from 'react';
import { Dataset, ReportingWidget } from '../../types/dashboard';
import { exportElementToPdf, PdfExportOptions } from '../../utils/pdfExport';
import {
  FileText,
  X,
  Printer,
  Sparkles,
  Layout,
  Sun,
  Moon,
  Layers,
  CheckCircle2,
  Table,
  Check
} from 'lucide-react';

interface ExportPdfModalProps {
  isOpen: boolean;
  onClose: () => void;
  dataset: Dataset;
  activeSheetName: string;
  widgets: ReportingWidget[];
  targetElementId: string;
  currentTheme: 'dark' | 'light';
  onNotification?: (msg: string) => void;
}

export const ExportPdfModal: React.FC<ExportPdfModalProps> = ({
  isOpen,
  onClose,
  dataset,
  activeSheetName,
  widgets,
  targetElementId,
  currentTheme,
  onNotification
}) => {
  const [reportTitle, setReportTitle] = useState(
    `${dataset.name} · Executive Performance Report`
  );
  const [orientation, setOrientation] = useState<'landscape' | 'portrait'>('landscape');
  const [colorMode, setColorMode] = useState<'light' | 'dark'>(
    currentTheme === 'light' ? 'light' : 'light' // Default to light for crisp print visibility
  );
  const [paperSize, setPaperSize] = useState<'a4' | 'letter'>('a4');
  const [includeKpis, setIncludeKpis] = useState(true);
  const [includeCharts, setIncludeCharts] = useState(true);
  const [includeLedger, setIncludeLedger] = useState(true);
  const [isExporting, setIsExporting] = useState(false);
  const [exportComplete, setExportComplete] = useState(false);

  if (!isOpen) return null;

  const metricWidgets = widgets.filter((w) => w.type === 'metric');
  const chartWidgets = widgets.filter((w) => w.type !== 'metric' && w.type !== 'table');

  const handleExport = async () => {
    const el = document.getElementById(targetElementId);
    if (!el) {
      if (onNotification) onNotification('Dashboard element not found for PDF export.');
      return;
    }

    setIsExporting(true);
    setExportComplete(false);

    try {
      const options: PdfExportOptions = {
        reportTitle,
        orientation,
        paperSize,
        colorMode,
        datasetName: dataset.name,
        filename: dataset.filename,
        activeSheetName,
        rowCount: dataset.sheets.find((s) => s.name === activeSheetName)?.rowCount || 0,
        authorName: 'Ammar Thaqif',
        authorEmail: 'ammarthaqif.ar@gmail.com'
      };

      // Export using html2canvas & jsPDF
      const result = await exportElementToPdf(el, options);

      setIsExporting(false);
      setExportComplete(true);

      if (onNotification) {
        onNotification(`Exported professional PDF report (${result.pages} page${result.pages > 1 ? 's' : ''}): ${result.filename}`);
      }

      setTimeout(() => {
        onClose();
        setExportComplete(false);
      }, 1200);
    } catch (err: any) {
      console.error('PDF export error:', err);
      setIsExporting(false);
      if (onNotification) {
        onNotification('Error generating PDF report. Please try again.');
      }
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-950/80 backdrop-blur-sm animate-in fade-in duration-150">
      <div
        className={`w-full max-w-xl rounded-2xl border shadow-2xl overflow-hidden transition-colors ${
          currentTheme === 'light'
            ? 'bg-white border-slate-200 text-slate-900'
            : 'bg-slate-900 border-slate-800 text-white'
        }`}
      >
        {/* Header */}
        <div
          className={`flex items-center justify-between px-6 py-4 border-b ${
            currentTheme === 'light'
              ? 'bg-slate-50 border-slate-200'
              : 'bg-slate-950/80 border-slate-800'
          }`}
        >
          <div className="flex items-center gap-2.5">
            <div className="p-2 rounded-xl bg-indigo-600/10 border border-indigo-500/20 text-indigo-400">
              <Printer className="w-5 h-5" />
            </div>
            <div>
              <h3 className="text-base font-bold tracking-tight">
                Export Professional PDF Report
              </h3>
              <p
                className={`text-xs ${
                  currentTheme === 'light' ? 'text-slate-500' : 'text-slate-400'
                }`}
              >
                Capture active charts, summary metrics, and spreadsheet ledger data
              </p>
            </div>
          </div>
          <button
            onClick={onClose}
            disabled={isExporting}
            className={`p-1.5 rounded-lg transition-colors ${
              currentTheme === 'light'
                ? 'hover:bg-slate-200 text-slate-500'
                : 'hover:bg-slate-800 text-slate-400 hover:text-white'
            }`}
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Body */}
        <div className="p-6 space-y-5 max-h-[75vh] overflow-y-auto">
          {/* Report Title Input */}
          <div className="space-y-1.5">
            <label
              className={`text-xs font-semibold uppercase tracking-wider block ${
                currentTheme === 'light' ? 'text-slate-600' : 'text-slate-400'
              }`}
            >
              Document Title
            </label>
            <input
              type="text"
              value={reportTitle}
              onChange={(e) => setReportTitle(e.target.value)}
              className={`w-full px-3.5 py-2 rounded-lg text-sm border font-medium focus:outline-none focus:ring-2 focus:ring-indigo-500 transition-colors ${
                currentTheme === 'light'
                  ? 'bg-slate-50 border-slate-300 text-slate-900'
                  : 'bg-slate-950 border-slate-700 text-white'
              }`}
            />
          </div>

          {/* Color Mode / Theme for PDF */}
          <div className="space-y-2">
            <label
              className={`text-xs font-semibold uppercase tracking-wider block ${
                currentTheme === 'light' ? 'text-slate-600' : 'text-slate-400'
              }`}
            >
              Report Color Theme
            </label>
            <div className="grid grid-cols-2 gap-3">
              <button
                type="button"
                onClick={() => setColorMode('light')}
                className={`p-3 rounded-xl border text-left transition-all flex items-start gap-3 ${
                  colorMode === 'light'
                    ? 'border-indigo-500 bg-indigo-50/60 ring-2 ring-indigo-500/20 text-slate-900'
                    : currentTheme === 'light'
                    ? 'border-slate-200 bg-slate-50 hover:bg-slate-100 text-slate-700'
                    : 'border-slate-800 bg-slate-950/60 hover:bg-slate-800 text-slate-300'
                }`}
              >
                <div className="p-1.5 rounded-lg bg-amber-500/10 text-amber-500 mt-0.5">
                  <Sun className="w-4 h-4" />
                </div>
                <div>
                  <div className="text-xs font-bold flex items-center gap-1.5">
                    <span>High-Contrast Light</span>
                    <span className="text-[10px] font-mono px-1.5 py-0.2 rounded bg-emerald-500/10 text-emerald-600 font-semibold">
                      Print Ready
                    </span>
                  </div>
                  <p className="text-[11px] text-slate-500 mt-0.5 leading-snug">
                    Crisp white paper background with maximum ink contrast for printing and executive circulation.
                  </p>
                </div>
              </button>

              <button
                type="button"
                onClick={() => setColorMode('dark')}
                className={`p-3 rounded-xl border text-left transition-all flex items-start gap-3 ${
                  colorMode === 'dark'
                    ? 'border-indigo-500 bg-indigo-950/50 ring-2 ring-indigo-500/20 text-white'
                    : currentTheme === 'light'
                    ? 'border-slate-200 bg-slate-50 hover:bg-slate-100 text-slate-700'
                    : 'border-slate-800 bg-slate-950/60 hover:bg-slate-800 text-slate-300'
                }`}
              >
                <div className="p-1.5 rounded-lg bg-indigo-500/10 text-indigo-400 mt-0.5">
                  <Moon className="w-4 h-4" />
                </div>
                <div>
                  <div className="text-xs font-bold">Deep Slate Presentation</div>
                  <p className="text-[11px] text-slate-400 mt-0.5 leading-snug">
                    Premium dark cyberpunk aesthetic matching the live studio screen viewport.
                  </p>
                </div>
              </button>
            </div>
          </div>

          {/* Orientation & Paper Format */}
          <div className="grid grid-cols-2 gap-4">
            <div className="space-y-1.5">
              <label
                className={`text-xs font-semibold uppercase tracking-wider block ${
                  currentTheme === 'light' ? 'text-slate-600' : 'text-slate-400'
                }`}
              >
                Orientation
              </label>
              <div className="grid grid-cols-2 gap-2">
                <button
                  type="button"
                  onClick={() => setOrientation('landscape')}
                  className={`py-2 px-3 text-xs font-medium rounded-lg border transition-colors flex items-center justify-center gap-1.5 ${
                    orientation === 'landscape'
                      ? 'bg-indigo-600 text-white border-indigo-500 shadow-sm'
                      : currentTheme === 'light'
                      ? 'bg-slate-100 border-slate-200 text-slate-700 hover:bg-slate-200'
                      : 'bg-slate-950 border-slate-800 text-slate-300 hover:bg-slate-800'
                  }`}
                >
                  <Layout className="w-3.5 h-3.5 rotate-90" />
                  <span>Landscape</span>
                </button>
                <button
                  type="button"
                  onClick={() => setOrientation('portrait')}
                  className={`py-2 px-3 text-xs font-medium rounded-lg border transition-colors flex items-center justify-center gap-1.5 ${
                    orientation === 'portrait'
                      ? 'bg-indigo-600 text-white border-indigo-500 shadow-sm'
                      : currentTheme === 'light'
                      ? 'bg-slate-100 border-slate-200 text-slate-700 hover:bg-slate-200'
                      : 'bg-slate-950 border-slate-800 text-slate-300 hover:bg-slate-800'
                  }`}
                >
                  <Layout className="w-3.5 h-3.5" />
                  <span>Portrait</span>
                </button>
              </div>
            </div>

            <div className="space-y-1.5">
              <label
                className={`text-xs font-semibold uppercase tracking-wider block ${
                  currentTheme === 'light' ? 'text-slate-600' : 'text-slate-400'
                }`}
              >
                Paper Size
              </label>
              <select
                value={paperSize}
                onChange={(e) => setPaperSize(e.target.value as 'a4' | 'letter')}
                className={`w-full px-3 py-2 rounded-lg text-xs font-medium border focus:outline-none focus:ring-2 focus:ring-indigo-500 ${
                  currentTheme === 'light'
                    ? 'bg-slate-50 border-slate-300 text-slate-900'
                    : 'bg-slate-950 border-slate-700 text-slate-200'
                }`}
              >
                <option value="a4">Standard A4 (210 × 297 mm)</option>
                <option value="letter">US Letter (8.5 × 11 in)</option>
              </select>
            </div>
          </div>

          {/* Section Summary Checklist */}
          <div
            className={`p-4 rounded-xl border space-y-2.5 ${
              currentTheme === 'light'
                ? 'bg-slate-50 border-slate-200'
                : 'bg-slate-950/70 border-slate-800'
            }`}
          >
            <span
              className={`text-[11px] font-semibold uppercase tracking-wider block ${
                currentTheme === 'light' ? 'text-slate-500' : 'text-slate-400'
              }`}
            >
              Included Layout Sections
            </span>

            <div className="grid grid-cols-1 sm:grid-cols-3 gap-2.5 text-xs">
              <label className="flex items-center gap-2 cursor-pointer select-none">
                <input
                  type="checkbox"
                  checked={includeKpis}
                  onChange={(e) => setIncludeKpis(e.target.checked)}
                  className="rounded border-slate-700 text-indigo-600 focus:ring-indigo-500"
                />
                <span className="flex items-center gap-1">
                  <Layers className="w-3.5 h-3.5 text-emerald-400" />
                  <span>{metricWidgets.length} KPI Metric Cards</span>
                </span>
              </label>

              <label className="flex items-center gap-2 cursor-pointer select-none">
                <input
                  type="checkbox"
                  checked={includeCharts}
                  onChange={(e) => setIncludeCharts(e.target.checked)}
                  className="rounded border-slate-700 text-indigo-600 focus:ring-indigo-500"
                />
                <span className="flex items-center gap-1">
                  <FileText className="w-3.5 h-3.5 text-indigo-400" />
                  <span>{chartWidgets.length} Visual Charts</span>
                </span>
              </label>

              <label className="flex items-center gap-2 cursor-pointer select-none">
                <input
                  type="checkbox"
                  checked={includeLedger}
                  onChange={(e) => setIncludeLedger(e.target.checked)}
                  className="rounded border-slate-700 text-indigo-600 focus:ring-indigo-500"
                />
                <span className="flex items-center gap-1">
                  <Table className="w-3.5 h-3.5 text-cyan-400" />
                  <span>Spreadsheet Ledger</span>
                </span>
              </label>
            </div>
          </div>

          {/* Report Metadata Footer Info */}
          <div
            className={`p-3 rounded-lg text-xs font-mono flex items-center justify-between border ${
              currentTheme === 'light'
                ? 'bg-emerald-50/60 border-emerald-200 text-emerald-800'
                : 'bg-emerald-950/30 border-emerald-500/20 text-emerald-300'
            }`}
          >
            <div className="flex items-center gap-2">
              <CheckCircle2 className="w-4 h-4 text-emerald-500 shrink-0" />
              <span>Report Attribution: Developed by Ammar Thaqif</span>
            </div>
            <span className="text-[11px] opacity-80">2x High-DPI Capture</span>
          </div>
        </div>

        {/* Footer Actions */}
        <div
          className={`flex items-center justify-between px-6 py-4 border-t ${
            currentTheme === 'light'
              ? 'bg-slate-50 border-slate-200'
              : 'bg-slate-950/90 border-slate-800'
          }`}
        >
          <button
            type="button"
            onClick={onClose}
            disabled={isExporting}
            className={`px-4 py-2 text-xs font-medium rounded-lg transition-colors ${
              currentTheme === 'light'
                ? 'text-slate-600 hover:bg-slate-200'
                : 'text-slate-400 hover:text-white hover:bg-slate-800'
            }`}
          >
            Cancel
          </button>

          <button
            type="button"
            onClick={handleExport}
            disabled={isExporting}
            className="flex items-center gap-2 px-5 py-2.5 text-xs font-semibold text-white bg-indigo-600 hover:bg-indigo-500 disabled:opacity-60 rounded-xl shadow-lg transition-all active:scale-95"
          >
            {isExporting ? (
              <>
                <span className="w-4 h-4 border-2 border-white/30 border-t-white rounded-full animate-spin" />
                <span>Generating High-Resolution PDF...</span>
              </>
            ) : exportComplete ? (
              <>
                <Check className="w-4 h-4 text-emerald-300" />
                <span>Downloaded Successfully!</span>
              </>
            ) : (
              <>
                <Sparkles className="w-4 h-4 text-amber-300" />
                <span>Export PDF Report</span>
              </>
            )}
          </button>
        </div>
      </div>
    </div>
  );
};
