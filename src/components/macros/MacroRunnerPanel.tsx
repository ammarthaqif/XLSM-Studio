import React, { useState, useMemo } from 'react';
import { Dataset, LegacyMacro, SheetData } from '../../types/dashboard';
import { simulateVbaExecution, VbaSimulationReport } from '../../utils/vbaSimulator';
import { useTheme } from '../../context/ThemeContext';
import {
  Play,
  RotateCcw,
  CheckCircle2,
  TrendingUp,
  Zap,
  ArrowRight,
  FileCode,
  RefreshCw,
  BarChart2,
  Layers,
  AlertCircle,
  Sliders,
  DollarSign,
  Percent,
  Check
} from 'lucide-react';

interface MacroRunnerPanelProps {
  dataset: Dataset;
  activeMacro?: LegacyMacro;
  macros: LegacyMacro[];
  onUpdateDataset: (updatedDataset: Dataset) => void;
  onNavigateToDashboard?: () => void;
  onNotification?: (msg: string) => void;
}

export type RunnerPreset =
  | 'revenue_growth'
  | 'opex_reduction'
  | 'margin_recalc'
  | 'cap_floor'
  | 'custom_script'
  | 'active_macro';

export const MacroRunnerPanel: React.FC<MacroRunnerPanelProps> = ({
  dataset,
  activeMacro,
  macros,
  onUpdateDataset,
  onNavigateToDashboard,
  onNotification
}) => {
  const { isLight } = useTheme();

  // Selected Runner Preset Mode
  const [selectedPreset, setSelectedPreset] = useState<RunnerPreset>('revenue_growth');

  // Parameters for presets
  const [growthPercentage, setGrowthPercentage] = useState<number>(10);
  const [selectedColumn, setSelectedColumn] = useState<string>('Revenue');
  const [targetSheetScope, setTargetSheetScope] = useState<'active' | 'all'>('active');
  const [opexCutPercentage, setOpexCutPercentage] = useState<number>(8);
  const [customVbaCode, setCustomVbaCode] = useState<string>(
    `Sub ApplyCustomOptimization()\n    Dim ws As Worksheet\n    Set ws = ActiveSheet\n    Dim r As Long, lastRow As Long\n    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row\n    \n    For r = 2 To lastRow\n        ' Example: Scale revenue by 10% and recalculate EBITDA\n        ws.Cells(r, 4).Value = Round(ws.Cells(r, 4).Value * 1.10, 2)\n        ws.Cells(r, 6).Value = ws.Cells(r, 4).Value - ws.Cells(r, 5).Value\n    Next r\nEnd Sub`
  );

  // Execution state & History
  const [isExecuting, setIsExecuting] = useState(false);
  const [lastReport, setLastReport] = useState<VbaSimulationReport | null>(null);
  const [showTraceLog, setShowTraceLog] = useState(false);
  const [originalBackup, setOriginalBackup] = useState<SheetData[] | null>(null);
  const [executionCount, setExecutionCount] = useState(0);

  // Active sheet reference
  const activeSheet = useMemo(() => {
    return (
      dataset.sheets.find((s) => s.name === dataset.activeSheetName) ||
      dataset.sheets[0] || { name: 'Sheet1', columns: [], rows: [] }
    );
  }, [dataset]);

  // Identify numeric columns in active sheet for parameter selection
  const numericColumns = useMemo(() => {
    if (!activeSheet || !activeSheet.rows || activeSheet.rows.length === 0) {
      return activeSheet.columns || ['Revenue', 'OPEX', 'EBITDA', 'Orders'];
    }
    const sample = activeSheet.rows[0];
    return activeSheet.columns.filter((c) => {
      const val = sample[c];
      return typeof val === 'number' || (!isNaN(Number(val)) && val !== '' && val !== null);
    });
  }, [activeSheet]);

  // Ensure selected column is valid
  React.useEffect(() => {
    if (numericColumns.length > 0 && !numericColumns.includes(selectedColumn)) {
      if (numericColumns.includes('Revenue')) {
        setSelectedColumn('Revenue');
      } else {
        setSelectedColumn(numericColumns[0]);
      }
    }
  }, [numericColumns, selectedColumn]);

  // Calculate high-level summary metrics before & after for impact preview
  const currentMetrics = useMemo(() => {
    let totalRevenue = 0;
    let totalOpex = 0;
    let rowCount = 0;

    activeSheet.rows.forEach((row) => {
      rowCount++;
      if (row.Revenue !== undefined) totalRevenue += Number(row.Revenue) || 0;
      if (row.OPEX !== undefined) totalOpex += Number(row.OPEX) || 0;
    });

    const ebitda = totalRevenue - totalOpex;
    const margin = totalRevenue > 0 ? (ebitda / totalRevenue) * 100 : 0;

    return {
      totalRevenue,
      totalOpex,
      ebitda,
      margin,
      rowCount
    };
  }, [activeSheet]);

  // Generate synthetic or formatted VBA script based on current preset and parameters
  const generatedVbaScript = useMemo(() => {
    const colIdx = Math.max(1, activeSheet.columns.indexOf(selectedColumn) + 1);

    if (selectedPreset === 'revenue_growth') {
      const multiplier = (1 + growthPercentage / 100).toFixed(4);
      if (targetSheetScope === 'all') {
        return `' ========================================================
' Simulated Macro: Global ${growthPercentage >= 0 ? '+' : ''}${growthPercentage}% Growth Scenario
' Target: All sheets in workbook, Column "${selectedColumn}"
' ========================================================
Sub GlobalGrowthSimulation()
    Dim ws As Worksheet
    Dim r As Long, lastRow As Long
    
    For Each ws In Worksheets
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        For r = 2 To lastRow
            ws.Cells(r, ${colIdx}).Value = Round(ws.Cells(r, ${colIdx}).Value * ${multiplier}, 2)
        Next r
    Next ws
    
    MsgBox "Applied ${growthPercentage}% growth across all sheets.", vbInformation
End Sub`;
      }

      return `' ========================================================
' Simulated Macro: ${growthPercentage >= 0 ? '+' : ''}${growthPercentage}% Scale to ${selectedColumn}
' Target Sheet: "${activeSheet.name}", Column ${colIdx} (${selectedColumn})
' ========================================================
Sub ApplyGrowthScenario()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim r As Long, lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    For r = 2 To lastRow
        ws.Cells(r, ${colIdx}).Value = Round(ws.Cells(r, ${colIdx}).Value * ${multiplier}, 2)
    Next r
    
    MsgBox "Scaled ${selectedColumn} by ${growthPercentage}% across " & (lastRow - 1) & " rows.", vbInformation
End Sub`;
    }

    if (selectedPreset === 'opex_reduction') {
      const opexColIdx = Math.max(1, activeSheet.columns.indexOf('OPEX') + 1 || colIdx);
      const factor = (1 - opexCutPercentage / 100).toFixed(4);
      return `' ========================================================
' Simulated Macro: -${opexCutPercentage}% OPEX Cost Efficiency Plan
' Target Sheet: "${activeSheet.name}", Column ${opexColIdx} (OPEX)
' ========================================================
Sub ApplyOpexOptimization()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim r As Long, lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    For r = 2 To lastRow
        ws.Cells(r, ${opexColIdx}).Value = Round(ws.Cells(r, ${opexColIdx}).Value * ${factor}, 2)
    Next r
    
    MsgBox "Reduced OPEX by ${opexCutPercentage}% across " & (lastRow - 1) & " ledger rows.", vbInformation
End Sub`;
    }

    if (selectedPreset === 'margin_recalc') {
      const revCol = Math.max(1, activeSheet.columns.indexOf('Revenue') + 1);
      const opexCol = Math.max(1, activeSheet.columns.indexOf('OPEX') + 1);
      const ebitdaCol = Math.max(1, activeSheet.columns.indexOf('EBITDA') + 1);
      const marginCol = Math.max(1, activeSheet.columns.indexOf('GrossMarginPct') + 1);

      return `' ========================================================
' Simulated Macro: Recompute EBITDA & Gross Margin Percentage
' Target Sheet: "${activeSheet.name}"
' Formula: EBITDA = Revenue - OPEX; Margin = (EBITDA / Revenue) * 100
' ========================================================
Sub RecomputeFinancialMargins()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim r As Long, lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    For r = 2 To lastRow
        ws.Cells(r, ${ebitdaCol}).Value = ws.Cells(r, ${revCol}).Value - ws.Cells(r, ${opexCol}).Value
        If ws.Cells(r, ${revCol}).Value > 0 Then
            ws.Cells(r, ${marginCol}).Value = Round((ws.Cells(r, ${ebitdaCol}).Value / ws.Cells(r, ${revCol}).Value) * 100, 1)
        End If
    Next r
    
    MsgBox "Recalculated EBITDA and Gross Margin % across ledger.", vbInformation
End Sub`;
    }

    if (selectedPreset === 'cap_floor') {
      return `' ========================================================
' Simulated Macro: Outlier Threshold Floor & Cap
' Target: Ensure ${selectedColumn} is bounded between baseline thresholds
' ========================================================
Sub EnforceThresholds()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim r As Long, lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    For r = 2 To lastRow
        ' Cap excessive values and enforce safety minimums
        If ws.Cells(r, ${colIdx}).Value < 50000 Then
            ws.Cells(r, ${colIdx}).Value = 50000
        End If
    Next r
    
    MsgBox "Normalized threshold floors for ${selectedColumn}.", vbInformation
End Sub`;
    }

    if (selectedPreset === 'active_macro') {
      return (
        activeMacro?.code ||
        `Sub ActiveMacro()\n    MsgBox "No active macro selected.", vbExclamation\nEnd Sub`
      );
    }

    return customVbaCode;
  }, [
    selectedPreset,
    growthPercentage,
    selectedColumn,
    targetSheetScope,
    opexCutPercentage,
    customVbaCode,
    activeMacro,
    activeSheet
  ]);

  // Execute Macro on Browser Data
  const handleExecuteMacro = () => {
    setIsExecuting(true);

    // Save initial backup if not already preserved
    if (!originalBackup) {
      setOriginalBackup(JSON.parse(JSON.stringify(dataset.sheets)));
    }

    setTimeout(() => {
      try {
        const report = simulateVbaExecution(generatedVbaScript, dataset);
        setLastReport(report);
        setIsExecuting(false);

        if (report.success && report.mutatedSheets && report.mutatedSheets.length > 0) {
          const updatedDataset: Dataset = {
            ...dataset,
            sheets: report.mutatedSheets
          };
          onUpdateDataset(updatedDataset);
          setExecutionCount((prev) => prev + 1);

          const mutations = report.isMultiSheet
            ? `across ${report.affectedSheets?.length || 0} sheets`
            : `in "${activeSheet.name}"`;

          if (onNotification) {
            onNotification(
              `Macro Executed Successfully! Mutated ${report.sheetDiffs ? Object.values(report.sheetDiffs).reduce((a, b) => a + b.totalMutations, 0) : 0} cell(s) ${mutations}. Dashboard charts updated!`
            );
          }
        } else if (!report.success) {
          setShowTraceLog(true);
          if (onNotification) {
            onNotification(
              `Execution stopped: ${report.errors.length} error(s) in VBA simulation script.`
            );
          }
        }
      } catch (err: any) {
        setIsExecuting(false);
        console.error('Macro runner execution fault:', err);
        if (onNotification) {
          onNotification('Failed to execute simulated macro. Please review syntax.');
        }
      }
    }, 180);
  };

  // Revert / Rollback to initial spreadsheet data
  const handleRevertToOriginal = () => {
    if (!originalBackup) return;
    const revertedDataset: Dataset = {
      ...dataset,
      sheets: JSON.parse(JSON.stringify(originalBackup))
    };
    onUpdateDataset(revertedDataset);
    setLastReport(null);
    setExecutionCount(0);
    if (onNotification) {
      onNotification('Reverted all macro operations. Original spreadsheet data restored!');
    }
  };

  const formatCurrency = (val: number) => {
    if (Math.abs(val) >= 1_000_000) return `$${(val / 1_000_000).toFixed(2)}M`;
    if (Math.abs(val) >= 1_000) return `$${(val / 1_000).toFixed(0)}k`;
    return `$${val.toLocaleString()}`;
  };

  return (
    <div className="space-y-6">
      {/* Kicker Header */}
      <div
        className={`p-4 rounded-xl border flex flex-col sm:flex-row sm:items-center justify-between gap-4 transition-all ${
          isLight
            ? 'bg-gradient-to-r from-indigo-50/90 to-purple-50/70 border-indigo-200 text-slate-800'
            : 'bg-gradient-to-r from-indigo-950/60 to-purple-950/40 border-indigo-500/30 text-slate-100'
        }`}
      >
        <div className="flex items-start gap-3">
          <div className="p-2 rounded-lg bg-indigo-600 text-white shadow-sm shrink-0 mt-0.5">
            <Play className="w-5 h-5 fill-current" />
          </div>
          <div>
            <div className="flex items-center gap-2 flex-wrap">
              <h3
                className={`text-base font-bold tracking-tight ${
                  isLight ? 'text-slate-900' : 'text-white'
                }`}
              >
                VBA Macro Runner & Chart Simulator
              </h3>
              <span className="text-[10px] font-mono px-2 py-0.5 rounded-full font-semibold bg-emerald-500/20 text-emerald-400 border border-emerald-500/30">
                Live Browser Reactive
              </span>
            </div>
            <p className={`text-xs mt-1 ${isLight ? 'text-slate-600' : 'text-slate-400'}`}>
              Execute simulated Excel VBA operations directly on your spreadsheet data. Cell mutations
              instantly flow through to KPI cards, Trend curves, Bar distributions, and Donut charts!
            </p>
          </div>
        </div>

        {/* Quick Nav to Dashboard */}
        {onNavigateToDashboard && (
          <button
            onClick={onNavigateToDashboard}
            className="flex items-center gap-1.5 px-3.5 py-2 text-xs font-semibold rounded-lg shadow-sm text-white bg-indigo-600 hover:bg-indigo-500 transition-all shrink-0 whitespace-nowrap self-start sm:self-center"
            title="Switch to Dashboard tab to view charts"
          >
            <BarChart2 className="w-4 h-4" />
            <span>View Live Dashboard</span>
            <ArrowRight className="w-3.5 h-3.5 ml-0.5" />
          </button>
        )}
      </div>

      {/* Preset Operation Selector Tabs */}
      <div className="space-y-3">
        <div className="flex items-center justify-between">
          <label
            className={`text-xs font-semibold uppercase tracking-wider font-mono ${
              isLight ? 'text-slate-600' : 'text-slate-400'
            }`}
          >
            Select Operation to Execute
          </label>
          <span className="text-xs text-slate-500 font-mono">
            Active: <span className="font-semibold">{activeSheet.name}</span> ({activeSheet.rows.length} rows)
          </span>
        </div>

        <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-2">
          {/* Preset 1: Revenue Growth */}
          <button
            type="button"
            onClick={() => setSelectedPreset('revenue_growth')}
            className={`p-3 rounded-xl border text-left transition-all flex flex-col justify-between ${
              selectedPreset === 'revenue_growth'
                ? isLight
                  ? 'border-indigo-500 bg-indigo-50/80 ring-2 ring-indigo-500/30 text-indigo-950 font-semibold'
                  : 'border-indigo-500 bg-indigo-950/70 ring-2 ring-indigo-500/30 text-white font-semibold'
                : isLight
                ? 'border-slate-200 bg-white hover:border-slate-300 text-slate-700'
                : 'border-slate-800 bg-slate-900/80 hover:border-slate-700 text-slate-300'
            }`}
          >
            <div className="flex items-center justify-between mb-1">
              <TrendingUp className="w-4 h-4 text-emerald-500" />
              {selectedPreset === 'revenue_growth' && <Check className="w-3.5 h-3.5 text-indigo-500" />}
            </div>
            <span className="text-xs">Growth Forecast</span>
            <span className="text-[10px] text-slate-500 font-mono mt-0.5">Scale Revenue %</span>
          </button>

          {/* Preset 2: OPEX Cost Reduction */}
          <button
            type="button"
            onClick={() => setSelectedPreset('opex_reduction')}
            className={`p-3 rounded-xl border text-left transition-all flex flex-col justify-between ${
              selectedPreset === 'opex_reduction'
                ? isLight
                  ? 'border-indigo-500 bg-indigo-50/80 ring-2 ring-indigo-500/30 text-indigo-950 font-semibold'
                  : 'border-indigo-500 bg-indigo-950/70 ring-2 ring-indigo-500/30 text-white font-semibold'
                : isLight
                ? 'border-slate-200 bg-white hover:border-slate-300 text-slate-700'
                : 'border-slate-800 bg-slate-900/80 hover:border-slate-700 text-slate-300'
            }`}
          >
            <div className="flex items-center justify-between mb-1">
              <DollarSign className="w-4 h-4 text-amber-500" />
              {selectedPreset === 'opex_reduction' && <Check className="w-3.5 h-3.5 text-indigo-500" />}
            </div>
            <span className="text-xs">OPEX Efficiency</span>
            <span className="text-[10px] text-slate-500 font-mono mt-0.5">Expense Cuts</span>
          </button>

          {/* Preset 3: Gross Margin Recalculator */}
          <button
            type="button"
            onClick={() => setSelectedPreset('margin_recalc')}
            className={`p-3 rounded-xl border text-left transition-all flex flex-col justify-between ${
              selectedPreset === 'margin_recalc'
                ? isLight
                  ? 'border-indigo-500 bg-indigo-50/80 ring-2 ring-indigo-500/30 text-indigo-950 font-semibold'
                  : 'border-indigo-500 bg-indigo-950/70 ring-2 ring-indigo-500/30 text-white font-semibold'
                : isLight
                ? 'border-slate-200 bg-white hover:border-slate-300 text-slate-700'
                : 'border-slate-800 bg-slate-900/80 hover:border-slate-700 text-slate-300'
            }`}
          >
            <div className="flex items-center justify-between mb-1">
              <Percent className="w-4 h-4 text-indigo-500" />
              {selectedPreset === 'margin_recalc' && <Check className="w-3.5 h-3.5 text-indigo-500" />}
            </div>
            <span className="text-xs">Recalc Margins</span>
            <span className="text-[10px] text-slate-500 font-mono mt-0.5">EBITDA & %</span>
          </button>

          {/* Preset 4: Threshold Normalization */}
          <button
            type="button"
            onClick={() => setSelectedPreset('cap_floor')}
            className={`p-3 rounded-xl border text-left transition-all flex flex-col justify-between ${
              selectedPreset === 'cap_floor'
                ? isLight
                  ? 'border-indigo-500 bg-indigo-50/80 ring-2 ring-indigo-500/30 text-indigo-950 font-semibold'
                  : 'border-indigo-500 bg-indigo-950/70 ring-2 ring-indigo-500/30 text-white font-semibold'
                : isLight
                ? 'border-slate-200 bg-white hover:border-slate-300 text-slate-700'
                : 'border-slate-800 bg-slate-900/80 hover:border-slate-700 text-slate-300'
            }`}
          >
            <div className="flex items-center justify-between mb-1">
              <Sliders className="w-4 h-4 text-cyan-500" />
              {selectedPreset === 'cap_floor' && <Check className="w-3.5 h-3.5 text-indigo-500" />}
            </div>
            <span className="text-xs">Threshold Floors</span>
            <span className="text-[10px] text-slate-500 font-mono mt-0.5">Normalize Caps</span>
          </button>

          {/* Preset 5: Run Active Workbook Macro */}
          <button
            type="button"
            onClick={() => setSelectedPreset('active_macro')}
            className={`p-3 rounded-xl border text-left transition-all flex flex-col justify-between ${
              selectedPreset === 'active_macro'
                ? isLight
                  ? 'border-indigo-500 bg-indigo-50/80 ring-2 ring-indigo-500/30 text-indigo-950 font-semibold'
                  : 'border-indigo-500 bg-indigo-950/70 ring-2 ring-indigo-500/30 text-white font-semibold'
                : isLight
                ? 'border-slate-200 bg-white hover:border-slate-300 text-slate-700'
                : 'border-slate-800 bg-slate-900/80 hover:border-slate-700 text-slate-300'
            }`}
          >
            <div className="flex items-center justify-between mb-1">
              <Zap className="w-4 h-4 text-purple-400" />
              {selectedPreset === 'active_macro' && <Check className="w-3.5 h-3.5 text-indigo-500" />}
            </div>
            <span className="text-xs truncate">Active Macro</span>
            <span className="text-[10px] text-slate-500 font-mono mt-0.5 truncate">
              {activeMacro?.name.replace(/^(Sub|Function)\s+/i, '') || 'Workbook'}
            </span>
          </button>

          {/* Preset 6: Custom VBA Script */}
          <button
            type="button"
            onClick={() => setSelectedPreset('custom_script')}
            className={`p-3 rounded-xl border text-left transition-all flex flex-col justify-between ${
              selectedPreset === 'custom_script'
                ? isLight
                  ? 'border-indigo-500 bg-indigo-50/80 ring-2 ring-indigo-500/30 text-indigo-950 font-semibold'
                  : 'border-indigo-500 bg-indigo-950/70 ring-2 ring-indigo-500/30 text-white font-semibold'
                : isLight
                ? 'border-slate-200 bg-white hover:border-slate-300 text-slate-700'
                : 'border-slate-800 bg-slate-900/80 hover:border-slate-700 text-slate-300'
            }`}
          >
            <div className="flex items-center justify-between mb-1">
              <FileCode className="w-4 h-4 text-blue-400" />
              {selectedPreset === 'custom_script' && <Check className="w-3.5 h-3.5 text-indigo-500" />}
            </div>
            <span className="text-xs">Custom Script</span>
            <span className="text-[10px] text-slate-500 font-mono mt-0.5">Direct Code</span>
          </button>
        </div>
      </div>

      {/* Operation Parameters Bar */}
      <div
        className={`p-4 rounded-xl border space-y-4 ${
          isLight ? 'bg-white border-slate-200 shadow-sm' : 'bg-slate-900/90 border-slate-800'
        }`}
      >
        <div className="flex items-center justify-between border-b pb-2.5">
          <h4 className={`text-xs font-bold uppercase tracking-wider font-mono ${isLight ? 'text-slate-800' : 'text-slate-200'}`}>
            Operation Parameters & Target Scope
          </h4>
          <span className="text-xs text-slate-500 font-mono">
            {executionCount > 0 ? `${executionCount} mutation run(s) applied` : 'Ready to run'}
          </span>
        </div>

        {/* Dynamic Controls based on preset */}
        {selectedPreset === 'revenue_growth' && (
          <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
            <div>
              <label className="block text-xs font-medium text-slate-500 mb-1">
                Growth Percentage (%):
              </label>
              <div className="flex items-center gap-1.5">
                {[-10, 5, 10, 15, 25].map((pct) => (
                  <button
                    key={pct}
                    type="button"
                    onClick={() => setGrowthPercentage(pct)}
                    className={`px-2 py-1 text-xs font-mono rounded border transition-colors ${
                      growthPercentage === pct
                        ? 'bg-indigo-600 text-white border-indigo-600 font-bold'
                        : isLight
                        ? 'bg-slate-100 hover:bg-slate-200 text-slate-700 border-slate-200'
                        : 'bg-slate-800 hover:bg-slate-700 text-slate-300 border-slate-700'
                    }`}
                  >
                    {pct > 0 ? `+${pct}%` : `${pct}%`}
                  </button>
                ))}
              </div>
            </div>

            <div>
              <label className="block text-xs font-medium text-slate-500 mb-1">
                Target Metric Column:
              </label>
              <select
                value={selectedColumn}
                onChange={(e) => setSelectedColumn(e.target.value)}
                className={`w-full border rounded-lg px-3 py-1.5 text-xs font-medium focus:outline-none focus:border-indigo-500 ${
                  isLight
                    ? 'bg-slate-50 border-slate-300 text-slate-800'
                    : 'bg-slate-800 border-slate-700 text-slate-200'
                }`}
              >
                {numericColumns.map((col) => (
                  <option key={col} value={col}>
                    {col}
                  </option>
                ))}
              </select>
            </div>

            <div>
              <label className="block text-xs font-medium text-slate-500 mb-1">
                Worksheet Scope:
              </label>
              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={() => setTargetSheetScope('active')}
                  className={`flex-1 py-1.5 text-xs font-medium rounded-lg border transition-colors ${
                    targetSheetScope === 'active'
                      ? 'bg-indigo-600 text-white border-indigo-600 font-semibold'
                      : isLight
                      ? 'bg-slate-100 text-slate-700 border-slate-200'
                      : 'bg-slate-800 text-slate-300 border-slate-700'
                  }`}
                >
                  Active ({activeSheet.name})
                </button>
                <button
                  type="button"
                  onClick={() => setTargetSheetScope('all')}
                  className={`flex-1 py-1.5 text-xs font-medium rounded-lg border transition-colors ${
                    targetSheetScope === 'all'
                      ? 'bg-indigo-600 text-white border-indigo-600 font-semibold'
                      : isLight
                      ? 'bg-slate-100 text-slate-700 border-slate-200'
                      : 'bg-slate-800 text-slate-300 border-slate-700'
                  }`}
                >
                  All ({dataset.sheets.length} Sheets)
                </button>
              </div>
            </div>
          </div>
        )}

        {selectedPreset === 'opex_reduction' && (
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
            <div>
              <label className="block text-xs font-medium text-slate-500 mb-1">
                Cost Cut Factor (% reduction):
              </label>
              <div className="flex items-center gap-1.5">
                {[5, 8, 12, 15, 20].map((cut) => (
                  <button
                    key={cut}
                    type="button"
                    onClick={() => setOpexCutPercentage(cut)}
                    className={`px-2.5 py-1 text-xs font-mono rounded border transition-colors ${
                      opexCutPercentage === cut
                        ? 'bg-amber-600 text-white border-amber-600 font-bold'
                        : isLight
                        ? 'bg-slate-100 hover:bg-slate-200 text-slate-700 border-slate-200'
                        : 'bg-slate-800 hover:bg-slate-700 text-slate-300 border-slate-700'
                    }`}
                  >
                    -{cut}%
                  </button>
                ))}
              </div>
            </div>

            <div>
              <label className="block text-xs font-medium text-slate-500 mb-1">
                Target Column:
              </label>
              <span
                className={`inline-block px-3 py-1.5 text-xs font-mono font-medium rounded-lg border ${
                  isLight ? 'bg-slate-100 border-slate-300 text-slate-800' : 'bg-slate-800 border-slate-700 text-slate-300'
                }`}
              >
                OPEX (Operating Expenses) · Auto-detects in active sheet
              </span>
            </div>
          </div>
        )}

        {selectedPreset === 'custom_script' && (
          <div>
            <label className="block text-xs font-medium text-slate-500 mb-1">
              Custom VBA Script to Execute on Workbook:
            </label>
            <textarea
              value={customVbaCode}
              onChange={(e) => setCustomVbaCode(e.target.value)}
              rows={6}
              className={`w-full font-mono text-xs p-3 rounded-lg border focus:outline-none focus:border-indigo-500 ${
                isLight
                  ? 'bg-slate-50 border-slate-300 text-slate-800'
                  : 'bg-slate-950 border-slate-700 text-slate-200'
              }`}
              placeholder="Sub CustomProcedure()... End Sub"
            />
          </div>
        )}

        {selectedPreset === 'active_macro' && activeMacro && (
          <div className="flex items-center justify-between p-3 rounded-lg border border-purple-500/30 bg-purple-950/20 text-xs">
            <div>
              <span className="font-mono font-bold text-purple-300">{activeMacro.name}</span>
              <p className="text-slate-400 mt-0.5">
                Module: {activeMacro.moduleName} · {activeMacro.lineCount || activeMacro.code.split('\n').length} lines · {activeMacro.description}
              </p>
            </div>
            <span className="text-[11px] font-mono px-2 py-0.5 rounded bg-purple-950 text-purple-200 border border-purple-500/40">
              Extracted from .xlsm
            </span>
          </div>
        )}

        {/* Generated VBA Execution Preview Box */}
        <div className="space-y-1.5">
          <div className="flex items-center justify-between text-xs">
            <span className="font-mono text-slate-400 text-[11px] uppercase">
              Generated VBA Execution Script:
            </span>
            <span className="text-[10px] text-slate-500 font-mono">
              Live transpiled from parameters
            </span>
          </div>
          <pre
            className={`p-3 rounded-lg text-[11px] font-mono overflow-x-auto border max-h-44 ${
              isLight
                ? 'bg-slate-50 border-slate-200 text-slate-800'
                : 'bg-slate-950/90 border-slate-800 text-slate-300'
            }`}
          >
            {generatedVbaScript}
          </pre>
        </div>

        {/* Primary Run & Revert Actions Bar */}
        <div className="flex flex-wrap items-center justify-between gap-3 pt-2">
          <div className="flex items-center gap-2">
            <button
              onClick={handleExecuteMacro}
              disabled={isExecuting}
              className="flex items-center gap-2 px-4 py-2.5 text-xs font-bold text-white bg-indigo-600 hover:bg-indigo-500 disabled:opacity-50 rounded-lg shadow-md transition-all whitespace-nowrap"
            >
              {isExecuting ? (
                <RefreshCw className="w-4 h-4 animate-spin" />
              ) : (
                <Play className="w-4 h-4 fill-current" />
              )}
              <span>{isExecuting ? 'Simulating VBA Execution...' : 'Run Macro on Browser Data'}</span>
            </button>

            {originalBackup && (
              <button
                onClick={handleRevertToOriginal}
                className={`flex items-center gap-1.5 px-3 py-2 text-xs font-semibold rounded-lg border transition-colors ${
                  isLight
                    ? 'bg-slate-100 hover:bg-slate-200 text-slate-700 border-slate-300'
                    : 'bg-slate-800 hover:bg-slate-700 text-slate-300 border-slate-700'
                }`}
                title="Rollback all mutations and restore original Excel values"
              >
                <RotateCcw className="w-3.5 h-3.5" />
                <span>Undo / Restore Original</span>
              </button>
            )}
          </div>

          {lastReport && (
            <div className="flex items-center gap-2">
              <span
                className={`text-xs font-mono px-2.5 py-1 rounded-lg border flex items-center gap-1.5 ${
                  lastReport.success
                    ? 'bg-emerald-950/70 border-emerald-500/40 text-emerald-300'
                    : 'bg-rose-950/70 border-rose-500/40 text-rose-300'
                }`}
              >
                {lastReport.success ? (
                  <CheckCircle2 className="w-3.5 h-3.5 text-emerald-400" />
                ) : (
                  <AlertCircle className="w-3.5 h-3.5 text-rose-400" />
                )}
                <span>
                  {lastReport.success
                    ? `Executed in ${lastReport.executionTimeMs}ms (0 faults)`
                    : `Halted with ${lastReport.errors.length} error(s)`}
                </span>
              </span>

              <button
                onClick={() => setShowTraceLog(!showTraceLog)}
                className={`text-xs font-mono px-2 py-1 rounded border transition-colors ${
                  isLight ? 'bg-slate-100 border-slate-300 text-slate-700' : 'bg-slate-800 border-slate-700 text-slate-300'
                }`}
              >
                {showTraceLog ? 'Hide Trace Log' : 'View Trace Log'}
              </button>
            </div>
          )}
        </div>
      </div>

      {/* Live Impact Card: What changed in Dashboard Charts */}
      {lastReport && lastReport.success && (
        <div
          className={`p-4 rounded-xl border animate-in fade-in slide-in-from-top-2 ${
            isLight
              ? 'bg-emerald-50/60 border-emerald-200 text-slate-800 shadow-sm'
              : 'bg-emerald-950/20 border-emerald-500/40 text-slate-200'
          }`}
        >
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-3 border-b border-emerald-500/20">
            <div className="flex items-center gap-2">
              <div className="p-1 rounded-full bg-emerald-500/20 text-emerald-400">
                <CheckCircle2 className="w-4 h-4" />
              </div>
              <h4 className="text-sm font-bold text-emerald-400">
                Live Impact on Dashboard & Metrics
              </h4>
            </div>
            {onNavigateToDashboard && (
              <button
                onClick={onNavigateToDashboard}
                className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-bold rounded-lg text-white bg-emerald-600 hover:bg-emerald-500 shadow-sm transition-all self-start sm:self-center"
              >
                <span>View Resulting Charts in Dashboard</span>
                <ArrowRight className="w-3.5 h-3.5" />
              </button>
            )}
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3 mt-3">
            {/* Metric 1: Total Revenue */}
            <div
              className={`p-3 rounded-lg border ${
                isLight ? 'bg-white border-slate-200 shadow-xs' : 'bg-slate-900/80 border-slate-800'
              }`}
            >
              <span className="text-[11px] font-mono text-slate-400 uppercase">Total Revenue</span>
              <div className="text-base font-bold font-mono tabular-nums mt-0.5 text-emerald-400">
                {formatCurrency(currentMetrics.totalRevenue)}
              </div>
              <div className="text-[10px] text-slate-500 font-mono mt-0.5">
                Live reflected in TrendLineChart
              </div>
            </div>

            {/* Metric 2: Total OPEX */}
            <div
              className={`p-3 rounded-lg border ${
                isLight ? 'bg-white border-slate-200 shadow-xs' : 'bg-slate-900/80 border-slate-800'
              }`}
            >
              <span className="text-[11px] font-mono text-slate-400 uppercase">Operating Expenses</span>
              <div className="text-base font-bold font-mono tabular-nums mt-0.5 text-indigo-400">
                {formatCurrency(currentMetrics.totalOpex)}
              </div>
              <div className="text-[10px] text-slate-500 font-mono mt-0.5">
                Live reflected in BarGroupChart
              </div>
            </div>

            {/* Metric 3: EBITDA Margin */}
            <div
              className={`p-3 rounded-lg border ${
                isLight ? 'bg-white border-slate-200 shadow-xs' : 'bg-slate-900/80 border-slate-800'
              }`}
            >
              <span className="text-[11px] font-mono text-slate-400 uppercase">EBITDA Margin %</span>
              <div className="text-base font-bold font-mono tabular-nums mt-0.5 text-amber-400">
                {currentMetrics.margin.toFixed(1)}%
              </div>
              <div className="text-[10px] text-slate-500 font-mono mt-0.5">
                Live reflected in MetricCard & DonutChart
              </div>
            </div>

            {/* Metric 4: Cell Mutations */}
            <div
              className={`p-3 rounded-lg border ${
                isLight ? 'bg-white border-slate-200 shadow-xs' : 'bg-slate-900/80 border-slate-800'
              }`}
            >
              <span className="text-[11px] font-mono text-slate-400 uppercase">Mutated Records</span>
              <div className="text-base font-bold font-mono tabular-nums mt-0.5 text-slate-200">
                {lastReport.sheetDiffs
                  ? Object.values(lastReport.sheetDiffs).reduce((a, b) => a + b.totalMutations, 0)
                  : 0}{' '}
                cells
              </div>
              <div className="text-[10px] text-slate-500 font-mono mt-0.5">
                Across {lastReport.affectedSheets?.length || 1} sheet(s)
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Execution Trace Console */}
      {showTraceLog && lastReport && (
        <div
          className={`p-4 rounded-xl border space-y-2 animate-in fade-in ${
            isLight ? 'bg-slate-50 border-slate-300' : 'bg-slate-950 border-slate-800'
          }`}
        >
          <div className="flex items-center justify-between text-xs">
            <span className="font-mono font-bold text-slate-400">VBA Engine Execution Trace Log:</span>
            <span className="text-[11px] font-mono text-slate-500">
              {lastReport.executionTrace.length} events logged
            </span>
          </div>
          <div
            className={`p-3 rounded-lg font-mono text-xs max-h-56 overflow-y-auto space-y-1 border ${
              isLight ? 'bg-white border-slate-200 text-slate-800' : 'bg-slate-900 border-slate-800 text-slate-300'
            }`}
          >
            {lastReport.executionTrace.map((msg, i) => (
              <div
                key={i}
                className={
                  msg.startsWith('❌')
                    ? 'text-rose-400'
                    : msg.startsWith('⚠️')
                    ? 'text-amber-400'
                    : msg.includes('[Cell Mutation]')
                    ? 'text-emerald-400'
                    : 'text-slate-400'
                }
              >
                {msg}
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
};
