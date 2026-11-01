import React, { useState, useRef } from 'react';
import { parseExcelFile } from '../../utils/excelParser';
import { SAMPLE_DATASETS } from '../../data/sampleDatasets';
import { Dataset } from '../../types/dashboard';
import { X, Upload, FileSpreadsheet, CheckCircle2, AlertCircle } from 'lucide-react';

interface ExcelUploadModalProps {
  isOpen: boolean;
  onClose: () => void;
  onDatasetLoaded: (dataset: Dataset) => void;
}

export const ExcelUploadModal: React.FC<ExcelUploadModalProps> = ({
  isOpen,
  onClose,
  onDatasetLoaded
}) => {
  const [isDragging, setIsDragging] = useState(false);
  const [parsing, setParsing] = useState(false);
  const [errorMsg, setErrorMsg] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  if (!isOpen) return null;

  const handleProcessFile = async (file: File) => {
    setParsing(true);
    setErrorMsg(null);
    try {
      const parsed = await parseExcelFile(file);
      onDatasetLoaded(parsed);
      onClose();
    } catch (err: any) {
      setErrorMsg(err.message || 'Failed to parse Excel workbook. Please verify the file structure.');
    } finally {
      setParsing(false);
    }
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      const file = e.dataTransfer.files[0];
      handleProcessFile(file);
    }
  };

  const handleSampleSelect = (dataset: Dataset) => {
    onDatasetLoaded(dataset);
    onClose();
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/70 backdrop-blur-sm animate-in fade-in duration-200">
      <div className="relative w-full max-w-2xl rounded-2xl border border-slate-700 bg-slate-900 shadow-2xl p-6">
        {/* Header */}
        <div className="flex items-center justify-between border-b border-slate-800 pb-4 mb-5">
          <div className="flex items-center gap-2.5">
            <div className="p-2 rounded-lg bg-emerald-950/80 border border-emerald-500/30 text-emerald-400">
              <FileSpreadsheet className="w-5 h-5" />
            </div>
            <div>
              <h3 className="text-base font-semibold text-white">Ingest Excel / XLSM Workbook</h3>
              <p className="text-xs text-slate-400">
                Supports .xlsm (Macro-Enabled), .xlsx, .xls, and .csv with automatic formula extraction
              </p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="p-1 rounded-lg text-slate-400 hover:text-white hover:bg-slate-800 transition-colors"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {errorMsg && (
          <div className="mb-4 p-3 rounded-lg border border-rose-500/30 bg-rose-950/40 text-rose-300 text-xs flex items-center gap-2">
            <AlertCircle className="w-4 h-4 shrink-0" />
            <span>{errorMsg}</span>
          </div>
        )}

        {/* Drag & Drop Area */}
        <div
          onDragOver={(e) => {
            e.preventDefault();
            setIsDragging(true);
          }}
          onDragLeave={() => setIsDragging(false)}
          onDrop={handleDrop}
          onClick={() => fileInputRef.current?.click()}
          className={`relative border-2 border-dashed rounded-xl p-7 text-center cursor-pointer transition-all ${
            isDragging
              ? 'border-indigo-500 bg-indigo-950/20'
              : 'border-slate-700 bg-slate-950/60 hover:border-slate-600 hover:bg-slate-950/90'
          }`}
        >
          <input
            type="file"
            ref={fileInputRef}
            onChange={(e) => {
              if (e.target.files && e.target.files[0]) {
                handleProcessFile(e.target.files[0]);
              }
            }}
            accept=".xlsm,.xlsx,.xls,.csv"
            className="hidden"
          />

          <div className="flex flex-col items-center justify-center space-y-2">
            <div className="p-3 rounded-full bg-slate-800 text-slate-300">
              <Upload className="w-6 h-6" />
            </div>
            <div className="text-sm font-semibold text-white">
              {parsing ? 'Extracting VBA macros & unprotecting sheets...' : 'Drop your .xlsm or .xlsx file here'}
            </div>
            <p className="text-xs text-slate-400 max-w-md">
              or click to browse. Automatically extracts embedded macros from <code className="text-indigo-300 font-mono">vbaProject.bin</code> and strips password protection locks.
            </p>
          </div>

          {/* Feature Highlights */}
          <div className="mt-4 pt-3 border-t border-slate-800/80 flex flex-wrap items-center justify-center gap-4 text-[11px] text-slate-400 font-mono">
            <span className="flex items-center gap-1 text-emerald-400">
              <CheckCircle2 className="w-3.5 h-3.5" />
              <span>Full VBA Macro Extraction</span>
            </span>
            <span className="flex items-center gap-1 text-indigo-400">
              <CheckCircle2 className="w-3.5 h-3.5" />
              <span>Auto-Unprotect Protected Sheets</span>
            </span>
            <span className="flex items-center gap-1 text-amber-400">
              <CheckCircle2 className="w-3.5 h-3.5" />
              <span>Formula Lineage Preservation</span>
            </span>
          </div>
        </div>

        {/* Divider */}
        <div className="relative my-6 text-center">
          <div className="absolute inset-0 flex items-center">
            <div className="w-full border-t border-slate-800" />
          </div>
          <span className="relative bg-slate-900 px-3 text-xs uppercase tracking-wider text-slate-500 font-mono">
            Or Test With Built-in Legacy Workbooks
          </span>
        </div>

        {/* Sample Workbooks */}
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
          {SAMPLE_DATASETS.map((sample) => (
            <div
              key={sample.id}
              onClick={() => handleSampleSelect(sample)}
              className="p-3.5 rounded-xl border border-slate-800 bg-slate-950/60 hover:bg-slate-800/60 hover:border-slate-700 cursor-pointer transition-all group"
            >
              <div className="flex items-start justify-between">
                <span className="text-xs font-semibold text-slate-200 group-hover:text-indigo-400 transition-colors">
                  {sample.name}
                </span>
                <span className="text-[10px] font-mono text-emerald-400 bg-emerald-950/60 px-1.5 py-0.5 rounded border border-emerald-500/30">
                  .xlsm
                </span>
              </div>
              <p className="text-xs text-slate-400 mt-1 line-clamp-2 leading-relaxed">
                {sample.notes}
              </p>
              <div className="mt-2.5 flex items-center justify-between text-[11px] text-slate-500 font-mono">
                <span>{sample.sheets.length} sheets</span>
                <span>{sample.legacyMacros.length} legacy macros</span>
              </div>
            </div>
          ))}
        </div>

        {/* Modal Attribution Footer */}
        <div className="mt-5 pt-3 border-t border-slate-800 flex flex-wrap items-center justify-between gap-2 text-xs text-slate-500">
          <span>Client-side XLSM & VBA unprotection engine</span>
          <span>
            Developed by <span className="text-indigo-400 font-medium">Ammar Thaqif</span>
          </span>
        </div>
      </div>
    </div>
  );
};
