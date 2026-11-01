import { Dataset, ReportingWidget } from '../types/dashboard';

export const SAMPLE_DATASETS: Dataset[] = [
  {
    id: 'ds-sample-financial-xlsm',
    name: 'Global Financial Performance Model v7',
    filename: 'Global_Financial_Performance_Model_v7.xlsm',
    fileSize: '2.4 MB',
    uploadedAt: '2026-10-23T04:30:00.000Z',
    activeSheetName: 'Consolidated P&L',
    notes: 'Converted from legacy Excel macro-enabled workbook. Contains 4 embedded VBA modules and routines.',
    hasVbaProject: true,
    detectedModules: [
      'Module1_FinancialCore',
      'Module2_StressEngine',
      'ThisWorkbook',
      'Sheet1_ConsolidatedPL'
    ],
    unprotectionLog: [
      '🔓 Unprotected sheet "Consolidated P&L" (Stripped SHA-512 password hash: 0x9FD2)',
      '🔓 Unlocked all cell ranges for unrestricted viewing and editing'
    ],
    columnMetadata: {
      Month: { type: 'date' },
      Region: { type: 'text' },
      ProductLine: { type: 'text' },
      Revenue: { type: 'currency' },
      OPEX: { type: 'currency' },
      EBITDA: { type: 'currency' },
      GrossMarginPct: { type: 'percentage' },
      ExecutiveComp: { type: 'currency' },
      Orders: { type: 'number' }
    },
    legacyMacros: [
      {
        id: 'macro-auto-open',
        name: 'Sub Workbook_Open()',
        moduleName: 'ThisWorkbook',
        moduleType: 'document',
        procedureType: 'Event',
        isAutoRun: true,
        description: 'Auto-run initialization: checks workbook state and sets default status message.',
        triggerEvent: 'Workbook Open Event (Auto-Run)',
        originalVba: `Private Sub Workbook_Open()\n    ' Legacy workbook initialization\n    Dim ws As Worksheet\n    Set ws = ThisWorkbook.Sheets("Consolidated P&L")\n    Application.StatusBar = "Financial Model v7 Loaded."\nEnd Sub`,
        code: `Private Sub Workbook_Open()\n    ' Legacy workbook initialization\n    Dim ws As Worksheet\n    Set ws = ThisWorkbook.Sheets("Consolidated P&L")\n    Application.StatusBar = "Financial Model v7 Loaded."\nEnd Sub`,
        modernizedTs: `export function initializeWorkbookState() {\n  return { status: "ONLINE", unlocked: true };\n}`,
        lineCount: 6
      },
      {
        id: 'macro-ebitda',
        name: 'Sub RecomputeEBITDA()',
        moduleName: 'Module1_FinancialCore',
        moduleType: 'standard',
        procedureType: 'Sub',
        description: 'Recalculates EBITDA, operating margin %, and highlights margins below 40%.',
        triggerEvent: 'Worksheet_Change / Cell Recalculate',
        originalVba: `Sub RecomputeEBITDA()\n    Dim ws As Worksheet\n    Set ws = ThisWorkbook.Sheets("Consolidated P&L")\n    Dim lastRow As Long\n    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row\n    Dim r As Long\n    For r = 2 To lastRow\n        ws.Cells(r, 6).Value = ws.Cells(r, 4).Value - ws.Cells(r, 5).Value ' Revenue - OPEX\n        If ws.Cells(r, 4).Value > 0 Then\n            ws.Cells(r, 7).Value = (ws.Cells(r, 6).Value / ws.Cells(r, 4).Value) * 100\n        End If\n    Next r\nEnd Sub`,
        code: `Sub RecomputeEBITDA()\n    Dim ws As Worksheet\n    Set ws = ThisWorkbook.Sheets("Consolidated P&L")\n    Dim lastRow As Long\n    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row\n    Dim r As Long\n    For r = 2 To lastRow\n        ws.Cells(r, 6).Value = ws.Cells(r, 4).Value - ws.Cells(r, 5).Value ' Revenue - OPEX\n        If ws.Cells(r, 4).Value > 0 Then\n            ws.Cells(r, 7).Value = (ws.Cells(r, 6).Value / ws.Cells(r, 4).Value) * 100\n        End If\n    Next r\nEnd Sub`,
        modernizedTs: `export function recomputeEBITDA(rows: any[]) {\n  return rows.map(r => ({\n    ...r,\n    EBITDA: (r.Revenue || 0) - (r.OPEX || 0),\n    GrossMarginPct: r.Revenue > 0 ? (((r.Revenue - r.OPEX) / r.Revenue) * 100).toFixed(1) : 0\n  }));\n}`,
        lineCount: 13
      },
      {
        id: 'macro-stress',
        name: 'Sub RunSensitivityStressTest()',
        moduleName: 'Module2_StressEngine',
        moduleType: 'standard',
        procedureType: 'Sub',
        description: 'Simulates sensitivity shocks across revenue (-8%) and expenses (+12%).',
        triggerEvent: 'Button: Run Stress Test',
        originalVba: `Sub RunSensitivityStressTest()\n    Dim ws As Worksheet\n    Set ws = ThisWorkbook.Sheets("Consolidated P&L")\n    Dim r As Range\n    Set r = ws.Range("GrossMargin")\n    If r.Value < 0.35 Then\n        MsgBox "Notice: Margin threshold below tolerance.", vbExclamation\n    End If\nEnd Sub`,
        code: `Sub RunSensitivityStressTest()\n    Dim ws As Worksheet\n    Set ws = ThisWorkbook.Sheets("Consolidated P&L")\n    Dim r As Range\n    Set r = ws.Range("GrossMargin")\n    If r.Value < 0.35 Then\n        MsgBox "Notice: Margin threshold below tolerance.", vbExclamation\n    End If\nEnd Sub`,
        modernizedTs: `export function runStressTest(rows: any[], inflation = 0.12, drop = 0.08) {\n  return rows.map(r => ({\n    ...r,\n    StressedRevenue: Math.round(r.Revenue * (1 - drop)),\n    StressedOPEX: Math.round(r.OPEX * (1 + inflation))\n  }));\n}`,
        lineCount: 9
      },
      {
        id: 'macro-tax-fn',
        name: 'Function CalculateEffectiveTax()',
        moduleName: 'Module1_FinancialCore',
        moduleType: 'standard',
        procedureType: 'Function',
        description: 'User-defined formula function calculating regional corporate tax rates.',
        triggerEvent: 'Formula Function (=CalculateEffectiveTax)',
        originalVba: `Function CalculateEffectiveTax(ebitdaVal As Double, regionCode As String) As Double\n    Select Case regionCode\n        Case "North America": CalculateEffectiveTax = ebitdaVal * 0.21\n        Case "EMEA": CalculateEffectiveTax = ebitdaVal * 0.25\n        Case "APAC": CalculateEffectiveTax = ebitdaVal * 0.17\n        Case Else: CalculateEffectiveTax = ebitdaVal * 0.20\n    End Select\nEnd Function`,
        code: `Function CalculateEffectiveTax(ebitdaVal As Double, regionCode As String) As Double\n    Select Case regionCode\n        Case "North America": CalculateEffectiveTax = ebitdaVal * 0.21\n        Case "EMEA": CalculateEffectiveTax = ebitdaVal * 0.25\n        Case "APAC": CalculateEffectiveTax = ebitdaVal * 0.17\n        Case Else: CalculateEffectiveTax = ebitdaVal * 0.20\n    End Select\nEnd Function`,
        modernizedTs: `export function calculateTax(ebitda: number, region: string): number {\n  const rates = { 'North America': 0.21, EMEA: 0.25, APAC: 0.17 };\n  return ebitda * (rates[region] || 0.20);\n}`,
        lineCount: 9
      },
      {
        id: 'macro-consolidate-regional',
        name: 'Sub ConsolidateRegionalActuals()',
        moduleName: 'Module1_FinancialCore',
        moduleType: 'standard',
        procedureType: 'Sub',
        description: 'Multi-sheet cross-consolidation: aggregates revenues from "Consolidated P&L" and calculates regional achievement in "Regional Sales".',
        triggerEvent: 'Button: Consolidate Across Sheets',
        originalVba: `Sub ConsolidateRegionalActuals()\n    ' Cross-Sheet Consolidation: Aggregates revenues from Consolidated P&L into Regional Sales\n    Dim wsSource As Worksheet\n    Dim wsDest As Worksheet\n    Set wsSource = ThisWorkbook.Sheets("Consolidated P&L")\n    Set wsDest = ThisWorkbook.Sheets("Regional Sales")\n    \n    Dim lastRow As Long\n    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row\n    \n    Dim destRow As Long\n    For destRow = 2 To 5\n        Dim regName As String\n        regName = wsDest.Cells(destRow, 1).Value\n        Dim totalRev As Double\n        totalRev = 0\n        \n        Dim r As Long\n        For r = 2 To lastRow\n            If wsSource.Cells(r, 2).Value = regName Then\n                totalRev = totalRev + wsSource.Cells(r, 4).Value\n            End If\n        Next r\n        \n        wsDest.Cells(destRow, 3).Value = totalRev\n        Dim target As Double\n        target = wsDest.Cells(destRow, 2).Value\n        If target > 0 Then\n            wsDest.Cells(destRow, 4).Value = Round((totalRev / target) * 100, 1)\n        End If\n    Next destRow\n    \n    Application.StatusBar = "Regional actuals consolidated across 2 sheets."\nEnd Sub`,
        code: `Sub ConsolidateRegionalActuals()\n    ' Cross-Sheet Consolidation: Aggregates revenues from Consolidated P&L into Regional Sales\n    Dim wsSource As Worksheet\n    Dim wsDest As Worksheet\n    Set wsSource = ThisWorkbook.Sheets("Consolidated P&L")\n    Set wsDest = ThisWorkbook.Sheets("Regional Sales")\n    \n    Dim lastRow As Long\n    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row\n    \n    Dim destRow As Long\n    For destRow = 2 To 5\n        Dim regName As String\n        regName = wsDest.Cells(destRow, 1).Value\n        Dim totalRev As Double\n        totalRev = 0\n        \n        Dim r As Long\n        For r = 2 To lastRow\n            If wsSource.Cells(r, 2).Value = regName Then\n                totalRev = totalRev + wsSource.Cells(r, 4).Value\n            End If\n        Next r\n        \n        wsDest.Cells(destRow, 3).Value = totalRev\n        Dim target As Double\n        target = wsDest.Cells(destRow, 2).Value\n        If target > 0 Then\n            wsDest.Cells(destRow, 4).Value = Round((totalRev / target) * 100, 1)\n        End If\n    Next destRow\n    \n    Application.StatusBar = "Regional actuals consolidated across 2 sheets."\nEnd Sub`,
        modernizedTs: `export function consolidateRegional(sourceRows: any[], destRows: any[]) {\n  return destRows.map(dest => {\n    const totalRev = sourceRows\n      .filter(s => s.Region === dest.Region)\n      .reduce((sum, s) => sum + (s.Revenue || 0), 0);\n    return {\n      ...dest,\n      ActualRevenue: totalRev,\n      AchievementPct: dest.AnnualTarget > 0 ? Number(((totalRev / dest.AnnualTarget) * 100).toFixed(1)) : 0\n    };\n  });\n}`,
        lineCount: 29
      }
    ],
    sheets: [
      {
        name: 'Consolidated P&L',
        rowCount: 12,
        columnCount: 9,
        isProtected: false,
        protectionMeta: {
          wasProtected: true,
          algorithm: 'SHA-512 (Password Hash: 0x9FD2)',
          hashValue: 'B9A2F8E1...',
          unlockedAt: 'Auto-unprotected on ingestion'
        },
        columns: ['Month', 'Region', 'ProductLine', 'Revenue', 'OPEX', 'EBITDA', 'GrossMarginPct', 'ExecutiveComp', 'Orders'],
        formulas: {
          'F2': '=D2-E2',
          'G2': '=(F2/D2)*100',
          'F3': '=D3-E3',
          'G3': '=(F3/D3)*100'
        },
        rows: [
          { Month: '2026-01', Region: 'North America', ProductLine: 'Cloud SaaS', Revenue: 1420000, OPEX: 680000, EBITDA: 740000, GrossMarginPct: 52.1, ExecutiveComp: 125000, Orders: 840 },
          { Month: '2026-02', Region: 'North America', ProductLine: 'Cloud SaaS', Revenue: 1560000, OPEX: 710000, EBITDA: 850000, GrossMarginPct: 54.5, ExecutiveComp: 125000, Orders: 910 },
          { Month: '2026-03', Region: 'North America', ProductLine: 'Cloud SaaS', Revenue: 1680000, OPEX: 730000, EBITDA: 950000, GrossMarginPct: 56.5, ExecutiveComp: 130000, Orders: 1040 },
          { Month: '2026-04', Region: 'EMEA', ProductLine: 'Hardware & Edge', Revenue: 1120000, OPEX: 690000, EBITDA: 430000, GrossMarginPct: 38.4, ExecutiveComp: 110000, Orders: 620 },
          { Month: '2026-05', Region: 'EMEA', ProductLine: 'Hardware & Edge', Revenue: 1190000, OPEX: 720000, EBITDA: 470000, GrossMarginPct: 39.5, ExecutiveComp: 110000, Orders: 650 },
          { Month: '2026-06', Region: 'EMEA', ProductLine: 'Enterprise Suite', Revenue: 1340000, OPEX: 740000, EBITDA: 600000, GrossMarginPct: 44.8, ExecutiveComp: 115000, Orders: 710 },
          { Month: '2026-07', Region: 'APAC', ProductLine: 'Cloud SaaS', Revenue: 1890000, OPEX: 820000, EBITDA: 1070000, GrossMarginPct: 56.6, ExecutiveComp: 140000, Orders: 1290 },
          { Month: '2026-08', Region: 'APAC', ProductLine: 'Cloud SaaS', Revenue: 2040000, OPEX: 860000, EBITDA: 1180000, GrossMarginPct: 57.8, ExecutiveComp: 140000, Orders: 1410 },
          { Month: '2026-09', Region: 'APAC', ProductLine: 'Enterprise Suite', Revenue: 2210000, OPEX: 890000, EBITDA: 1320000, GrossMarginPct: 59.7, ExecutiveComp: 145000, Orders: 1530 },
          { Month: '2026-10', Region: 'LATAM', ProductLine: 'Hardware & Edge', Revenue: 890000, OPEX: 540000, EBITDA: 350000, GrossMarginPct: 39.3, ExecutiveComp: 95000, Orders: 480 },
          { Month: '2026-11', Region: 'LATAM', ProductLine: 'Hardware & Edge', Revenue: 950000, OPEX: 560000, EBITDA: 390000, GrossMarginPct: 41.1, ExecutiveComp: 95000, Orders: 510 },
          { Month: '2026-12', Region: 'North America', ProductLine: 'Enterprise Suite', Revenue: 2100000, OPEX: 890000, EBITDA: 1210000, GrossMarginPct: 57.6, ExecutiveComp: 150000, Orders: 1380 }
        ]
      },
      {
        name: 'Regional Sales',
        rowCount: 4,
        columnCount: 5,
        columns: ['Region', 'AnnualTarget', 'ActualRevenue', 'AchievementPct', 'Headcount'],
        formulas: {
          'D2': '=(C2/B2)*100'
        },
        rows: [
          { Region: 'North America', AnnualTarget: 5000000, ActualRevenue: 5340000, AchievementPct: 106.8, Headcount: 42 },
          { Region: 'APAC', AnnualTarget: 5500000, ActualRevenue: 6140000, AchievementPct: 111.6, Headcount: 38 },
          { Region: 'EMEA', AnnualTarget: 4000000, ActualRevenue: 3650000, AchievementPct: 91.3, Headcount: 29 },
          { Region: 'LATAM', AnnualTarget: 2000000, ActualRevenue: 1840000, AchievementPct: 92.0, Headcount: 16 }
        ]
      }
    ]
  },
  {
    id: 'ds-sample-supply-chain',
    name: 'Supply Chain & Inventory Risk 2026',
    filename: 'SupplyChain_Inventory_Risk_2026.xlsm',
    fileSize: '1.8 MB',
    uploadedAt: '2026-10-22T18:20:00.000Z',
    activeSheetName: 'Stock Levels',
    notes: 'Inventory stock level model with supplier lead time risk solvers.',
    hasVbaProject: true,
    detectedModules: ['Module1_InventoryOptimization', 'Sheet1_StockLevels'],
    columnMetadata: {
      SKU: { type: 'text' },
      Category: { type: 'text' },
      CurrentStock: { type: 'number' },
      SafetyStock: { type: 'number' },
      UnitCost: { type: 'currency' },
      LeadTimeDays: { type: 'number' },
      SupplierRiskScore: { type: 'number' },
      WarehouseLocation: { type: 'text' }
    },
    legacyMacros: [
      {
        id: 'macro-reorder',
        name: 'Sub OptimizeReorderPoints()',
        moduleName: 'Module1_InventoryOptimization',
        moduleType: 'standard',
        procedureType: 'Sub',
        description: 'VBA solver routine to calculate economic order quantity (EOQ).',
        triggerEvent: 'Nightly Scheduled Macro',
        originalVba: `Sub OptimizeReorderPoints()\n    Dim ws As Worksheet\n    Set ws = ThisWorkbook.Sheets("Stock Levels")\n    Dim r As Long\n    For r = 2 To 9\n        If ws.Cells(r, 3).Value < ws.Cells(r, 4).Value Then\n            ws.Cells(r, 3).Interior.Color = RGB(255, 230, 230)\n        End If\n    Next r\nEnd Sub`,
        code: `Sub OptimizeReorderPoints()\n    Dim ws As Worksheet\n    Set ws = ThisWorkbook.Sheets("Stock Levels")\n    Dim r As Long\n    For r = 2 To 9\n        If ws.Cells(r, 3).Value < ws.Cells(r, 4).Value Then\n            ws.Cells(r, 3).Interior.Color = RGB(255, 230, 230)\n        End If\n    Next r\nEnd Sub`,
        modernizedTs: `export function calculateEOQ(demand: number, orderCost: number, holdingCost: number): number {\n  return Math.round(Math.sqrt((2 * demand * orderCost) / holdingCost));\n}`,
        lineCount: 9
      }
    ],
    sheets: [
      {
        name: 'Stock Levels',
        rowCount: 8,
        columnCount: 8,
        columns: ['SKU', 'Category', 'CurrentStock', 'SafetyStock', 'UnitCost', 'LeadTimeDays', 'SupplierRiskScore', 'WarehouseLocation'],
        formulas: {},
        rows: [
          { SKU: 'SKU-OPT-101', Category: 'Optics & Sensors', CurrentStock: 1450, SafetyStock: 1200, UnitCost: 48.5, LeadTimeDays: 14, SupplierRiskScore: 2.1, WarehouseLocation: 'Dallas Hub' },
          { SKU: 'SKU-CHIP-204', Category: 'Semiconductors', CurrentStock: 420, SafetyStock: 800, UnitCost: 185.0, LeadTimeDays: 45, SupplierRiskScore: 7.8, WarehouseLocation: 'Frankfurt Hub' },
          { SKU: 'SKU-PWR-309', Category: 'Power Supply', CurrentStock: 3200, SafetyStock: 1500, UnitCost: 22.0, LeadTimeDays: 10, SupplierRiskScore: 1.4, WarehouseLocation: 'Dallas Hub' },
          { SKU: 'SKU-DISP-412', Category: 'Displays & Panels', CurrentStock: 680, SafetyStock: 750, UnitCost: 94.0, LeadTimeDays: 28, SupplierRiskScore: 5.2, WarehouseLocation: 'Singapore Hub' },
          { SKU: 'SKU-CHAS-501', Category: 'Chassis & Enclosure', CurrentStock: 2100, SafetyStock: 1000, UnitCost: 35.0, LeadTimeDays: 12, SupplierRiskScore: 1.8, WarehouseLocation: 'Dallas Hub' },
          { SKU: 'SKU-MOD-602', Category: 'Radio Modules', CurrentStock: 890, SafetyStock: 900, UnitCost: 62.0, LeadTimeDays: 21, SupplierRiskScore: 4.1, WarehouseLocation: 'Singapore Hub' },
          { SKU: 'SKU-MEM-708', Category: 'Semiconductors', CurrentStock: 310, SafetyStock: 600, UnitCost: 110.0, LeadTimeDays: 35, SupplierRiskScore: 6.9, WarehouseLocation: 'Frankfurt Hub' },
          { SKU: 'SKU-CBL-801', Category: 'Cables & Wiring', CurrentStock: 5400, SafetyStock: 2000, UnitCost: 8.5, LeadTimeDays: 7, SupplierRiskScore: 1.1, WarehouseLocation: 'Dallas Hub' }
        ]
      }
    ]
  }
];

export const INITIAL_WIDGETS: ReportingWidget[] = [
  {
    id: 'w-metric-rev',
    title: 'Total Converted Revenue',
    type: 'metric',
    sheetName: 'Consolidated P&L',
    xAxis: 'Month',
    yAxis: 'Revenue',
    aggregation: 'sum',
    width: '1/3',
    colorTheme: 'emerald'
  },
  {
    id: 'w-metric-ebitda',
    title: 'Operating EBITDA',
    type: 'metric',
    sheetName: 'Consolidated P&L',
    xAxis: 'Month',
    yAxis: 'EBITDA',
    aggregation: 'sum',
    width: '1/3',
    colorTheme: 'indigo'
  },
  {
    id: 'w-metric-margin',
    title: 'Average Gross Margin',
    type: 'metric',
    sheetName: 'Consolidated P&L',
    xAxis: 'Month',
    yAxis: 'GrossMarginPct',
    aggregation: 'avg',
    width: '1/3',
    colorTheme: 'cyan'
  },
  {
    id: 'w-line-trends',
    title: 'Revenue vs OPEX Trajectory',
    type: 'line',
    sheetName: 'Consolidated P&L',
    xAxis: 'Month',
    yAxis: 'Revenue',
    aggregation: 'sum',
    width: '2/3',
    colorTheme: 'emerald'
  },
  {
    id: 'w-donut-regions',
    title: 'Revenue by Region Breakdown',
    type: 'donut',
    sheetName: 'Consolidated P&L',
    xAxis: 'Region',
    yAxis: 'Revenue',
    aggregation: 'sum',
    width: '1/3',
    colorTheme: 'indigo'
  },
  {
    id: 'w-bar-products',
    title: 'EBITDA Performance by Product Line',
    type: 'bar',
    sheetName: 'Consolidated P&L',
    xAxis: 'ProductLine',
    yAxis: 'EBITDA',
    aggregation: 'sum',
    width: '1/2',
    colorTheme: 'amber'
  },
  {
    id: 'w-table-summary',
    title: 'Monthly Ledger Aggregation',
    type: 'table',
    sheetName: 'Consolidated P&L',
    xAxis: 'Month',
    yAxis: 'Revenue',
    aggregation: 'sum',
    width: '1/2',
    colorTheme: 'slate'
  }
];
