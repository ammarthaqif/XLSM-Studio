import React, { useState, useMemo } from 'react';
import { Dataset } from '../../types/dashboard';
import { exportSheetToExcel, exportSheetToCSV } from '../../utils/excelParser';
import { Search, Download, ShieldCheck, Unlock, FunctionSquare } from 'lucide-react';

interface SpreadsheetGridProps {
  dataset: Dataset;
  activeSheetName: string;
  onSelectSheet: (name: string) => void;
}

export const SpreadsheetGrid: React.FC<SpreadsheetGridProps> = ({
  dataset,
  activeSheetName,
  onSelectSheet
}) => {
  const [searchTerm, setSearchTerm] = useState('');
  const [selectedCell, setSelectedCell] = useState<{ rowIdx: number; colKey: string; cellAddress: string } | null>(null);

  const sheet = dataset.sheets.find((s) => s.name === activeSheetName) || dataset.sheets[0];

  // Helper for column letter (0 -> A, 1 -> B, etc.)
  const getColLetter = (idx: number) => {
    let letter = '';
    while (idx >= 0) {
      letter = String.fromCharCode((idx % 26) + 65) + letter;
      idx = Math.floor(idx / 26) - 1;
    }
    return letter;
  };

  // Filtered rows
  const filteredRows = useMemo(() => {
    if (!sheet) return [];
    if (!searchTerm.trim()) return sheet.rows;
    const term = searchTerm.toLowerCase();
    return sheet.rows.filter((row) =>
      Object.values(row).some((val) => String(val).toLowerCase().includes(term))
    );
  }, [sheet, searchTerm]);

  // Selected cell value and formula
  const activeFormulaOrVal = useMemo(() => {
    if (!selectedCell || !sheet) return '';
    const formula = sheet.formulas[selectedCell.cellAddress];
    if (formula) return formula;
    const row = sheet.rows[selectedCell.rowIdx];
    return row ? String(row[selectedCell.colKey] ?? '') : '';
  }, [selectedCell, sheet]);

  const handleCellClick = (rowIdx: number, colIdx: number, colKey: string) => {
    const address = `${getColLetter(colIdx)}${rowIdx + 2}`;
    setSelectedCell({ rowIdx, colKey, cellAddress: address });
  };

  const formatCellValue = (val: any, colKey: string) => {
    const meta = dataset.columnMetadata[colKey];
    if (meta?.type === 'currency' && typeof val === 'number') {
      return `$${val.toLocaleString()}`;
    }
    if (meta?.type === 'percentage' && typeof val === 'number') {
      return `${val.toFixed(1)}%`;
    }
    if (val === null || val === undefined) return '';
    return String(val);
  };

  return (
    <div className="rounded-xl border border-slate-800 bg-slate-900/90 shadow-sm flex flex-col h-full overflow-hidden">
      {/* Top Toolbar */}
      <div className="p-4 border-b border-slate-800 flex flex-wrap items-center justify-between gap-3 bg-slate-950/40">
        {/* Sheet Tabs */}
        <div className="flex items-center gap-1.5 overflow-x-auto pb-1 max-w-full">
          {dataset.sheets.map((s) => (
            <button
              key={s.name}
              onClick={() => onSelectSheet(s.name)}
              className={`px-3 py-1.5 text-xs font-medium rounded-md whitespace-nowrap transition-colors ${
                activeSheetName === s.name
                  ? 'bg-slate-800 text-white border border-slate-700 shadow-sm'
                  : 'text-slate-400 hover:text-slate-200 hover:bg-slate-800/40'
              }`}
            >
              <span>{s.name}</span>
              <span className="ml-1.5 text-[10px] text-slate-500 font-mono">({s.rowCount})</span>
            </button>
          ))}
        </div>

        {/* Actions & Search */}
        <div className="flex items-center gap-2">
          {/* Sheet Protection Status Indicator */}
          {sheet?.protectionMeta?.wasProtected ? (
            <div
              title={`Sheet password protection was automatically stripped on ingestion.\nAlgorithm: ${sheet.protectionMeta.algorithm || 'Legacy Excel Hash'}\nUnlocked for unrestricted cell editing and calculation.`}
              className="flex items-center gap-1.5 px-2.5 py-1 text-xs font-mono text-emerald-300 bg-emerald-950/70 border border-emerald-500/40 rounded-md"
            >
              <Unlock className="w-3.5 h-3.5 text-emerald-400" />
              <span>Sheet Unprotected</span>
            </div>
          ) : (
            <div
              title="Sheet has no password lock restrictions"
              className="flex items-center gap-1 px-2 py-1 text-[11px] font-mono text-slate-400 bg-slate-900 border border-slate-800 rounded-md"
            >
              <ShieldCheck className="w-3 h-3 text-slate-500" />
              <span>Unlocked</span>
            </div>
          )}

          <div className="relative">
            <Search className="w-3.5 h-3.5 absolute left-2.5 top-1/2 -translate-y-1/2 text-slate-500" />
            <input
              type="text"
              placeholder="Search spreadsheet cells..."
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              className="bg-slate-800/90 border border-slate-700 rounded-md pl-8 pr-3 py-1 text-xs text-slate-200 placeholder-slate-500 focus:outline-none focus:border-indigo-500 w-44 sm:w-56"
            />
          </div>

          <button
            onClick={() => exportSheetToCSV(sheet, `${dataset.name}_${activeSheetName}.csv`)}
            className="flex items-center gap-1.5 px-3 py-1 text-xs font-semibold text-emerald-300 bg-emerald-950/70 hover:bg-emerald-900 border border-emerald-500/40 rounded-md transition-colors shadow-sm"
            title={`Extract and download ${sheet?.name || 'sheet'} as CSV`}
          >
            <Download className="w-3.5 h-3.5 text-emerald-400" />
            <span>Download CSV</span>
          </button>

          <button
            onClick={() => exportSheetToExcel(sheet, `${dataset.name}_${activeSheetName}.xlsx`)}
            className="flex items-center gap-1.5 px-3 py-1 text-xs font-medium text-slate-300 bg-slate-800 hover:bg-slate-700 border border-slate-700 rounded-md transition-colors"
            title="Export as Excel workbook"
          >
            <Download className="w-3.5 h-3.5" />
            <span>Export .xlsx</span>
          </button>
        </div>
      </div>

      {/* Excel Formula Bar */}
      <div className="px-4 py-2 border-b border-slate-800 bg-slate-950 flex items-center gap-3 text-xs font-mono">
        <div className="flex items-center gap-1.5 text-slate-400 bg-slate-900 px-2 py-0.5 rounded border border-slate-800">
          <FunctionSquare className="w-3.5 h-3.5 text-indigo-400" />
          <span className="text-white font-semibold min-w-8 text-center">
            {selectedCell ? selectedCell.cellAddress : 'A1'}
          </span>
        </div>
        <div className="h-4 w-px bg-slate-800" />
        <span className="text-slate-500 select-none">fx</span>
        <div className="flex-1 truncate text-slate-200 font-mono text-xs">
          {activeFormulaOrVal || '<Click any cell below to inspect value or formula>'}
        </div>
      </div>

      {/* Spreadsheet Grid Table */}
      <div className="flex-1 overflow-auto max-h-[520px]">
        <table className="w-full text-left text-xs font-mono tabular-nums border-collapse select-text">
          <thead>
            <tr className="sticky top-0 z-10 bg-slate-950 border-b border-slate-800 text-slate-400">
              <th className="w-12 py-2 px-2 text-center bg-slate-950/90 border-r border-slate-800 text-slate-500 font-mono text-[11px]">
                #
              </th>
              {sheet?.columns.map((col, cIdx) => (
                <th
                  key={col}
                  className="py-2.5 px-3 border-r border-slate-800/80 font-medium text-slate-300 whitespace-nowrap bg-slate-950"
                >
                  <div className="flex items-center gap-1.5">
                    <span className="text-slate-500 font-mono text-[10px]">
                      {getColLetter(cIdx)}
                    </span>
                    <span className="font-semibold text-slate-200 font-sans">{col}</span>
                  </div>
                </th>
              ))}
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-800/60">
            {filteredRows.map((row, rIdx) => {
              const rowNumber = rIdx + 2;
              return (
                <tr key={rIdx} className="hover:bg-slate-800/30 transition-colors">
                  <td className="py-2 px-2 text-center bg-slate-950/40 border-r border-slate-800 text-slate-500 text-[11px] font-mono select-none">
                    {rowNumber}
                  </td>
                  {sheet?.columns.map((col, cIdx) => {
                    const address = `${getColLetter(cIdx)}${rowNumber}`;
                    const hasFormula = !!sheet.formulas[address];
                    const isCellSelected =
                      selectedCell?.rowIdx === rIdx && selectedCell?.colKey === col;

                    return (
                      <td
                        key={col}
                        onClick={() => handleCellClick(rIdx, cIdx, col)}
                        className={`py-2 px-3 border-r border-slate-800/50 cursor-pointer whitespace-nowrap transition-colors relative ${
                          isCellSelected
                            ? 'bg-indigo-950/60 ring-1 ring-indigo-500 text-white font-medium'
                            : 'text-slate-300 hover:bg-slate-800/60'
                        }`}
                      >
                        <div className="flex items-center justify-between gap-1.5">
                          <span>{formatCellValue(row[col], col)}</span>
                          {hasFormula && (
                            <span
                              title={`Formula: ${sheet.formulas[address]}`}
                              className="text-[9px] text-indigo-400 font-mono bg-indigo-950 px-1 rounded border border-indigo-500/20"
                            >
                              fx
                            </span>
                          )}
                        </div>
                      </td>
                    );
                  })}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      {/* Footer Info */}
      <div className="p-3 bg-slate-950 border-t border-slate-800 flex items-center justify-between text-xs text-slate-500 font-mono">
        <div className="flex items-center gap-3">
          <span>{filteredRows.length} of {sheet?.rowCount} rows shown</span>
          <span>·</span>
          <span>{sheet?.columns.length} columns</span>
        </div>
        <span className="text-emerald-400">All data unlocked</span>
      </div>
    </div>
  );
};
