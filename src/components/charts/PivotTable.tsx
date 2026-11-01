import React, { useState, useMemo } from 'react';
import { Dataset } from '../../types/dashboard';
import { TableProperties, ArrowRightLeft } from 'lucide-react';

interface PivotTableProps {
  dataset: Dataset;
  activeSheetName: string;
}

export const PivotTable: React.FC<PivotTableProps> = ({
  dataset,
  activeSheetName
}) => {
  const sheet = dataset.sheets.find((s) => s.name === activeSheetName) || dataset.sheets[0];

  const categoricalColumns = useMemo(() => {
    if (!sheet) return [];
    return sheet.columns.filter((c) => {
      const type = dataset.columnMetadata[c]?.type;
      return type === 'text' || type === 'date';
    });
  }, [sheet, dataset.columnMetadata]);

  const numericColumns = useMemo(() => {
    if (!sheet) return [];
    return sheet.columns.filter((c) => {
      const type = dataset.columnMetadata[c]?.type;
      return type === 'currency' || type === 'number' || type === 'percentage';
    });
  }, [sheet, dataset.columnMetadata]);

  // Pivot selectors
  const [rowDimension, setRowDimension] = useState<string>(
    categoricalColumns[0] || sheet?.columns[0] || ''
  );
  const [colDimension, setColDimension] = useState<string>(
    categoricalColumns[1] || sheet?.columns[1] || ''
  );
  const [metric, setMetric] = useState<string>(
    numericColumns[0] || sheet?.columns[2] || ''
  );
  const [aggregation, setAggregation] = useState<'sum' | 'avg' | 'count'>('sum');

  // Swap dimensions
  const handleSwap = () => {
    const temp = rowDimension;
    setRowDimension(colDimension);
    setColDimension(temp);
  };

  // Build matrix
  const { rowHeaders, colHeaders, matrix, rowTotals, colTotals, grandTotal } = useMemo(() => {
    if (!sheet || !rowDimension || !colDimension || !metric) {
      return {
        rowHeaders: [],
        colHeaders: [],
        matrix: {},
        rowTotals: {},
        colTotals: {},
        grandTotal: 0
      };
    }

    const rowsSet = new Set<string>();
    const colsSet = new Set<string>();
    const cellValues: Record<string, number[]> = {};

    sheet.rows.forEach((row) => {
      const rKey = String(row[rowDimension] ?? 'N/A');
      const cKey = String(row[colDimension] ?? 'N/A');
      const mVal = Number(row[metric]) || 0;

      rowsSet.add(rKey);
      colsSet.add(cKey);

      const cellKey = `${rKey}:::${cKey}`;
      if (!cellValues[cellKey]) cellValues[cellKey] = [];
      cellValues[cellKey].push(mVal);
    });

    const rowHeaders = Array.from(rowsSet).sort();
    const colHeaders = Array.from(colsSet).sort();

    const aggregateArr = (arr: number[]) => {
      if (!arr || arr.length === 0) return 0;
      if (aggregation === 'sum') return arr.reduce((a, b) => a + b, 0);
      if (aggregation === 'avg') return arr.reduce((a, b) => a + b, 0) / arr.length;
      return arr.length;
    };

    const matrix: Record<string, Record<string, number>> = {};
    const rowTotals: Record<string, number> = {};
    const colTotals: Record<string, number> = {};
    let grandTotal = 0;

    rowHeaders.forEach((r) => {
      matrix[r] = {};
      const allRowValues: number[] = [];
      colHeaders.forEach((c) => {
        const key = `${r}:::${c}`;
        const vals = cellValues[key] || [];
        matrix[r][c] = aggregateArr(vals);
        allRowValues.push(...vals);

        if (!colTotals[c]) colTotals[c] = 0;
        colTotals[c] += matrix[r][c];
      });
      rowTotals[r] = aggregateArr(allRowValues);
      grandTotal += rowTotals[r];
    });

    return { rowHeaders, colHeaders, matrix, rowTotals, colTotals, grandTotal };
  }, [sheet, rowDimension, colDimension, metric, aggregation]);

  const formatCell = (val: number) => {
    if (val >= 1_000_000) return `$${(val / 1_000_000).toFixed(2)}M`;
    if (val >= 1_000) return `$${(val / 1_000).toFixed(0)}k`;
    return val.toLocaleString();
  };

  return (
    <div className="rounded-xl border border-slate-800 bg-slate-900/90 p-5 shadow-sm">
      {/* Pivot Controls Bar */}
      <div className="flex flex-wrap items-center justify-between gap-4 pb-4 border-b border-slate-800">
        <div className="flex items-center gap-2">
          <TableProperties className="w-5 h-5 text-indigo-400" />
          <h3 className="text-sm font-semibold text-white">Dynamic XLSM Pivot Matrix</h3>
        </div>

        <div className="flex flex-wrap items-center gap-3 text-xs">
          {/* Row selector */}
          <div className="flex items-center gap-1.5">
            <span className="text-slate-400 font-medium">Rows:</span>
            <select
              value={rowDimension}
              onChange={(e) => setRowDimension(e.target.value)}
              className="bg-slate-800 border border-slate-700 rounded-md px-2.5 py-1 text-slate-200 focus:outline-none focus:border-indigo-500"
            >
              {sheet?.columns.map((c) => (
                <option key={c} value={c}>{c}</option>
              ))}
            </select>
          </div>

          <button
            onClick={handleSwap}
            title="Swap Row & Column Dimensions"
            className="p-1 rounded bg-slate-800 border border-slate-700 text-slate-400 hover:text-white"
          >
            <ArrowRightLeft className="w-3.5 h-3.5" />
          </button>

          {/* Col selector */}
          <div className="flex items-center gap-1.5">
            <span className="text-slate-400 font-medium">Columns:</span>
            <select
              value={colDimension}
              onChange={(e) => setColDimension(e.target.value)}
              className="bg-slate-800 border border-slate-700 rounded-md px-2.5 py-1 text-slate-200 focus:outline-none focus:border-indigo-500"
            >
              {sheet?.columns.map((c) => (
                <option key={c} value={c}>{c}</option>
              ))}
            </select>
          </div>

          {/* Metric */}
          <div className="flex items-center gap-1.5">
            <span className="text-slate-400 font-medium">Value:</span>
            <select
              value={metric}
              onChange={(e) => setMetric(e.target.value)}
              className="bg-slate-800 border border-slate-700 rounded-md px-2.5 py-1 text-slate-200 focus:outline-none focus:border-indigo-500"
            >
              {sheet?.columns.map((c) => (
                <option key={c} value={c}>{c}</option>
              ))}
            </select>
          </div>

          {/* Aggregation */}
          <div className="flex items-center gap-1 bg-slate-800/80 p-0.5 rounded-lg border border-slate-700">
            {(['sum', 'avg', 'count'] as const).map((agg) => (
              <button
                key={agg}
                onClick={() => setAggregation(agg)}
                className={`px-2 py-0.5 rounded text-[11px] font-mono uppercase font-semibold transition-colors ${
                  aggregation === agg ? 'bg-indigo-600 text-white' : 'text-slate-400 hover:text-slate-200'
                }`}
              >
                {agg}
              </button>
            ))}
          </div>
        </div>
      </div>

      {/* Cross-Tab Grid */}
      <div className="mt-4 overflow-x-auto">
        <table className="w-full text-left text-xs font-mono tabular-nums border-collapse">
          <thead>
            <tr className="border-b border-slate-800 bg-slate-950/60">
              <th className="py-2.5 px-3 text-slate-400 font-medium text-left">
                {rowDimension} \ {colDimension}
              </th>
              {colHeaders.map((col) => (
                <th key={col} className="py-2.5 px-3 text-slate-300 font-medium text-right">
                  {col}
                </th>
              ))}
              <th className="py-2.5 px-3 text-indigo-400 font-semibold text-right bg-indigo-950/20">
                Grand Total
              </th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-800/60">
            {rowHeaders.map((row) => (
              <tr key={row} className="hover:bg-slate-800/40 transition-colors">
                <td className="py-2 px-3 font-semibold text-slate-200 text-left font-sans">
                  {row}
                </td>
                {colHeaders.map((col) => {
                  const val = matrix[row]?.[col] || 0;
                  return (
                    <td
                      key={col}
                      className="py-2 px-3 text-right text-slate-300"
                    >
                      {val > 0 ? formatCell(val) : '—'}
                    </td>
                  );
                })}
                <td className="py-2 px-3 text-right font-bold text-white bg-indigo-950/20">
                  {formatCell(rowTotals[row] || 0)}
                </td>
              </tr>
            ))}

            {/* Column Totals Row */}
            <tr className="border-t-2 border-slate-700 bg-slate-950/80 font-bold">
              <td className="py-2.5 px-3 text-indigo-400 font-sans">Total {aggregation.toUpperCase()}</td>
              {colHeaders.map((col) => (
                <td key={col} className="py-2.5 px-3 text-right text-slate-200">
                  {formatCell(colTotals[col] || 0)}
                </td>
              ))}
              <td className="py-2.5 px-3 text-right text-emerald-400 bg-indigo-950/40 font-bold">
                {formatCell(grandTotal)}
              </td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  );
};
