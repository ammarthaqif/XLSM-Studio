import React, { useState, useMemo } from 'react';
import { Dataset, LegacyMacro } from '../../types/dashboard';
import { enhanceVbaCode, formatVbaCode, repackUpgradedXlsm } from '../../utils/vbaExtractor';
import { downloadWorkbookFile } from '../../utils/excelParser';
import { simulateVbaExecution, evaluateImmediateExpression, VbaSimulationReport } from '../../utils/vbaSimulator';
import {
  Code2,
  FolderTree,
  FileCode,
  Download,
  Zap,
  ShieldAlert,
  Clock,
  Sparkles,
  RotateCcw,
  Check,
  Plus,
  Play,
  FileText,
  ChevronRight,
  ChevronDown,
  Layers,
  ArrowRightLeft,
  X,
  Trash2,
  AlertTriangle,
  AlertCircle,
  CheckCircle2,
  Terminal,
  RefreshCw,
  Search,
  BookOpen,
  Cpu,
  CornerDownLeft,
  Copy,
  Send,
  Sliders
} from 'lucide-react';

interface LegacyMacroStudioProps {
  dataset: Dataset;
  onUpdateMacros: (updatedMacros: LegacyMacro[]) => void;
}

interface DeleteTarget {
  type: 'procedure' | 'module';
  id?: string;
  name: string;
  moduleName?: string;
  count?: number;
}

export interface VbeMemberDef {
  name: string;
  type: 'Property' | 'Method' | 'Function' | 'Sub' | 'Constant';
  signature: string;
  description: string;
  snippet: string;
}

export interface VbeClassDef {
  name: string;
  library: 'Excel' | 'VBA';
  description: string;
  members: VbeMemberDef[];
}

const VBE_CLASSES: VbeClassDef[] = [
  {
    name: 'Application',
    library: 'Excel',
    description: 'Represents the entire Microsoft Excel application instance.',
    members: [
      {
        name: 'ScreenUpdating',
        type: 'Property',
        signature: 'Application.ScreenUpdating As Boolean',
        description: 'Enables or disables screen repainting during macro execution for maximum performance.',
        snippet: 'Application.ScreenUpdating = False'
      },
      {
        name: 'Calculation',
        type: 'Property',
        signature: 'Application.Calculation As XlCalculation',
        description: 'Returns or sets calculation mode: xlCalculationManual, xlCalculationAutomatic, or xlCalculationSemiAutomatic.',
        snippet: 'Application.Calculation = xlCalculationManual'
      },
      {
        name: 'EnableEvents',
        type: 'Property',
        signature: 'Application.EnableEvents As Boolean',
        description: 'True if events are enabled; set to False to suppress recursive Worksheet_Change events.',
        snippet: 'Application.EnableEvents = False'
      },
      {
        name: 'DisplayAlerts',
        type: 'Property',
        signature: 'Application.DisplayAlerts As Boolean',
        description: 'Suppresses prompt dialog boxes and warning messages (e.g. file overwrite prompts).',
        snippet: 'Application.DisplayAlerts = False'
      },
      {
        name: 'ActiveSheet',
        type: 'Property',
        signature: 'Application.ActiveSheet As Worksheet',
        description: 'Returns a Worksheet object that represents the currently active worksheet.',
        snippet: 'Set ws = ActiveSheet'
      },
      {
        name: 'ActiveWorkbook',
        type: 'Property',
        signature: 'Application.ActiveWorkbook As Workbook',
        description: 'Returns a Workbook object that represents the active workbook.',
        snippet: 'Set wb = ActiveWorkbook'
      },
      {
        name: 'Worksheets',
        type: 'Property',
        signature: 'Application.Worksheets(Index) As Worksheet',
        description: 'Collection of all Worksheet objects in the workbook.',
        snippet: 'Set ws = Application.Worksheets("Data")'
      },
      {
        name: 'StatusBar',
        type: 'Property',
        signature: 'Application.StatusBar As Variant',
        description: 'Returns or sets the custom text shown in the Excel status bar.',
        snippet: 'Application.StatusBar = "Processing records..."'
      },
      {
        name: 'Calculate',
        type: 'Method',
        signature: 'Application.Calculate()',
        description: 'Calculates all open workbooks and worksheets.',
        snippet: 'Application.Calculate'
      },
      {
        name: 'Quit',
        type: 'Method',
        signature: 'Application.Quit()',
        description: 'Quits Microsoft Excel.',
        snippet: 'Application.Quit'
      }
    ]
  },
  {
    name: 'Worksheet',
    library: 'Excel',
    description: 'Represents a worksheet in an Excel workbook.',
    members: [
      {
        name: 'Range',
        type: 'Property',
        signature: 'Worksheet.Range(Cell1, [Cell2]) As Range',
        description: 'Returns a Range object that represents a cell or a range of cells.',
        snippet: 'ws.Range("A1:B10")'
      },
      {
        name: 'Cells',
        type: 'Property',
        signature: 'Worksheet.Cells(RowIndex, ColumnIndex) As Range',
        description: 'Returns a Range object representing a cell by 1-based row and column index.',
        snippet: 'ws.Cells(r, c).Value'
      },
      {
        name: 'Name',
        type: 'Property',
        signature: 'Worksheet.Name As String',
        description: 'Returns or sets the name of the worksheet tab.',
        snippet: 'ws.Name = "AuditLog"'
      },
      {
        name: 'Calculate',
        type: 'Method',
        signature: 'Worksheet.Calculate()',
        description: 'Calculates formulas specifically on this worksheet.',
        snippet: 'ws.Calculate'
      },
      {
        name: 'Protect',
        type: 'Method',
        signature: 'Worksheet.Protect([Password], [DrawingObjects], [Contents], [Scenarios])',
        description: 'Protects a worksheet so that it cannot be modified.',
        snippet: 'ws.Protect Password:="Admin2026"'
      },
      {
        name: 'Unprotect',
        type: 'Method',
        signature: 'Worksheet.Unprotect([Password])',
        description: 'Removes protection from a worksheet.',
        snippet: 'ws.Unprotect Password:="Admin2026"'
      },
      {
        name: 'Rows',
        type: 'Property',
        signature: 'Worksheet.Rows As Range',
        description: 'Returns a Range object representing all the rows on the worksheet.',
        snippet: 'lastRow = ws.Rows.Count'
      },
      {
        name: 'Columns',
        type: 'Property',
        signature: 'Worksheet.Columns As Range',
        description: 'Returns a Range object representing all the columns on the worksheet.',
        snippet: 'ws.Columns("A:D").AutoFit'
      },
      {
        name: 'Visible',
        type: 'Property',
        signature: 'Worksheet.Visible As XlSheetVisibility',
        description: 'Determines whether the object is visible, hidden, or very hidden (xlSheetVeryHidden).',
        snippet: 'ws.Visible = xlSheetVisible'
      }
    ]
  },
  {
    name: 'Range',
    library: 'Excel',
    description: 'Represents a cell, a row, a column, or a selection of cells containing one or more blocks of cells.',
    members: [
      {
        name: 'Value',
        type: 'Property',
        signature: 'Range.Value([RangeValueDataType]) As Variant',
        description: 'Returns or sets a Variant value that represents the value of the specified range.',
        snippet: 'Range("A1").Value = "Total"'
      },
      {
        name: 'Value2',
        type: 'Property',
        signature: 'Range.Value2 As Variant',
        description: 'High-performance value retrieval without currency or date format conversion.',
        snippet: 'v = Range("A1").Value2'
      },
      {
        name: 'Formula',
        type: 'Property',
        signature: 'Range.Formula As String',
        description: 'Returns or sets the formula for the object in A1-style notation.',
        snippet: 'Range("C2").Formula = "=A2*B2"'
      },
      {
        name: 'Address',
        type: 'Property',
        signature: 'Range.Address([RowAbsolute], [ColumnAbsolute], [ReferenceStyle]) As String',
        description: 'Returns a String value that represents the range reference in the language of the macro.',
        snippet: 'addr = Range("A1").Address'
      },
      {
        name: 'Row',
        type: 'Property',
        signature: 'Range.Row As Long',
        description: 'Returns the number of the first row in the first area of the range.',
        snippet: 'r = Target.Row'
      },
      {
        name: 'Column',
        type: 'Property',
        signature: 'Range.Column As Long',
        description: 'Returns the number of the first column in the first area of the specified range.',
        snippet: 'c = Target.Column'
      },
      {
        name: 'Count',
        type: 'Property',
        signature: 'Range.Count As Long',
        description: 'Returns a Long value that represents the number of objects in the collection.',
        snippet: 'totalCells = Range("A1:C10").Count'
      },
      {
        name: 'Select',
        type: 'Method',
        signature: 'Range.Select() As Variant',
        description: 'Selects the range object.',
        snippet: 'Range("A1").Select'
      },
      {
        name: 'Clear',
        type: 'Method',
        signature: 'Range.Clear() As Variant',
        description: 'Clears the entire object (contents and formatting).',
        snippet: 'Range("A1:Z100").Clear'
      },
      {
        name: 'ClearContents',
        type: 'Method',
        signature: 'Range.ClearContents() As Variant',
        description: 'Clears the formulas and values from the range but preserves cell formats.',
        snippet: 'Range("B2:B50").ClearContents'
      },
      {
        name: 'Copy',
        type: 'Method',
        signature: 'Range.Copy([Destination]) As Variant',
        description: 'Copies the range to the specified range or to the Clipboard.',
        snippet: 'Range("A1:A10").Copy Destination:=Range("B1")'
      },
      {
        name: 'End',
        type: 'Property',
        signature: 'Range.End(Direction As XlDirection) As Range',
        description: 'Returns a Range object representing the cell at the end of the region (xlUp, xlDown, xlToLeft, xlToRight).',
        snippet: 'lastRow = Cells(Rows.Count, "A").End(xlUp).Row'
      },
      {
        name: 'Offset',
        type: 'Property',
        signature: 'Range.Offset([RowOffset], [ColumnOffset]) As Range',
        description: 'Returns a Range object that represents a range that is offset from the specified range.',
        snippet: 'Set nextCell = Range("A1").Offset(1, 0)'
      },
      {
        name: 'Font',
        type: 'Property',
        signature: 'Range.Font As Font',
        description: 'Returns a Font object that represents the font formatting of the specified object.',
        snippet: 'Range("A1").Font.Bold = True'
      }
    ]
  },
  {
    name: 'Workbook',
    library: 'Excel',
    description: 'Represents a Microsoft Excel workbook file.',
    members: [
      {
        name: 'Sheets',
        type: 'Property',
        signature: 'Workbook.Sheets As Sheets',
        description: 'Returns a Sheets collection that represents all the sheets in the specified workbook.',
        snippet: 'wb.Sheets.Count'
      },
      {
        name: 'Save',
        type: 'Method',
        signature: 'Workbook.Save()',
        description: 'Saves changes to the specified workbook.',
        snippet: 'ActiveWorkbook.Save'
      },
      {
        name: 'Close',
        type: 'Method',
        signature: 'Workbook.Close([SaveChanges], [Filename], [RouteWorkbook])',
        description: 'Closes the object.',
        snippet: 'ActiveWorkbook.Close SaveChanges:=False'
      },
      {
        name: 'FullName',
        type: 'Property',
        signature: 'Workbook.FullName As String',
        description: 'Returns the full pathname of the workbook on disk.',
        snippet: 'filePath = ActiveWorkbook.FullName'
      }
    ]
  },
  {
    name: 'Interaction',
    library: 'VBA',
    description: 'Standard VBA interaction dialogs and runtime control routines.',
    members: [
      {
        name: 'MsgBox',
        type: 'Function',
        signature: 'MsgBox(Prompt, [Buttons As VbMsgBoxStyle], [Title]) As VbMsgBoxResult',
        description: 'Displays a message in a dialog box and waits for the user to click a button.',
        snippet: 'MsgBox "Operation completed successfully!", vbInformation, "System Notice"'
      },
      {
        name: 'InputBox',
        type: 'Function',
        signature: 'InputBox(Prompt, [Title], [Default], [XPos], [YPos]) As String',
        description: 'Displays a prompt in a dialog box and waits for user text input.',
        snippet: 'inputVal = InputBox("Please enter multiplier:", "Parameters", "1.15")'
      },
      {
        name: 'DoEvents',
        type: 'Function',
        signature: 'DoEvents() As Integer',
        description: 'Yields execution so that the operating system can process other events.',
        snippet: 'DoEvents'
      }
    ]
  },
  {
    name: 'DateTime',
    library: 'VBA',
    description: 'VBA date, time, and high-resolution timer functions.',
    members: [
      {
        name: 'Timer',
        type: 'Function',
        signature: 'Timer() As Single',
        description: 'Returns a Single representing the number of seconds elapsed since midnight (ideal for profiling).',
        snippet: 'startTime = Timer'
      },
      {
        name: 'Now',
        type: 'Function',
        signature: 'Now() As Date',
        description: 'Returns a Date specifying the current date and time according to your system.',
        snippet: 'currentStamp = Now'
      },
      {
        name: 'Date',
        type: 'Function',
        signature: 'Date() As Date',
        description: 'Returns the current system date.',
        snippet: 'todayDate = Date'
      }
    ]
  },
  {
    name: 'ErrObject',
    library: 'VBA',
    description: 'Contains information about run-time errors.',
    members: [
      {
        name: 'Number',
        type: 'Property',
        signature: 'Err.Number As Long',
        description: 'Returns or sets a numeric value specifying an error. 0 indicates no error.',
        snippet: 'If Err.Number <> 0 Then'
      },
      {
        name: 'Description',
        type: 'Property',
        signature: 'Err.Description As String',
        description: 'Returns or sets a string expression containing a descriptive string associated with an error.',
        snippet: 'MsgBox "Error: " & Err.Description, vbCritical'
      },
      {
        name: 'Clear',
        type: 'Method',
        signature: 'Err.Clear()',
        description: 'Clears all property settings of the Err object.',
        snippet: 'Err.Clear'
      },
      {
        name: 'Raise',
        type: 'Method',
        signature: 'Err.Raise(Number As Long, [Source], [Description])',
        description: 'Generates a run-time error.',
        snippet: 'Err.Raise vbObjectError + 513, "ValidationRoutine", "Threshold exceeded"'
      }
    ]
  },
  {
    name: 'Constants',
    library: 'Excel',
    description: 'Built-in Excel enumerations and constants.',
    members: [
      {
        name: 'xlCalculationManual',
        type: 'Constant',
        signature: 'Const xlCalculationManual = -4135',
        description: 'Specifies that formulas are calculated manually when requested.',
        snippet: 'xlCalculationManual'
      },
      {
        name: 'xlCalculationAutomatic',
        type: 'Constant',
        signature: 'Const xlCalculationAutomatic = -4105',
        description: 'Specifies that formulas are recalculated automatically.',
        snippet: 'xlCalculationAutomatic'
      },
      {
        name: 'xlUp',
        type: 'Constant',
        signature: 'Const xlUp = -4162',
        description: 'Direction enumeration for moving upward.',
        snippet: 'xlUp'
      },
      {
        name: 'xlDown',
        type: 'Constant',
        signature: 'Const xlDown = -4121',
        description: 'Direction enumeration for moving downward.',
        snippet: 'xlDown'
      }
    ]
  }
];

export const LegacyMacroStudio: React.FC<LegacyMacroStudioProps> = ({
  dataset,
  onUpdateMacros
}) => {
  const [macros, setMacros] = useState<LegacyMacro[]>(dataset.legacyMacros);
  const [selectedMacroId, setSelectedMacroId] = useState<string>(
    dataset.legacyMacros[0]?.id || ''
  );
  const [activeViewTab, setActiveViewTab] = useState<'editor' | 'diff'>('editor');
  const [expandedModules, setExpandedModules] = useState<Record<string, boolean>>({
    standard: true,
    document: true,
    class: true,
    userform: true
  });
  const [statusNotification, setStatusNotification] = useState<string | null>(null);
  const [downloading, setDownloading] = useState(false);
  const [simulationReport, setSimulationReport] = useState<VbaSimulationReport | null>(null);
  const [isSimulating, setIsSimulating] = useState(false);
  const [showTrace, setShowTrace] = useState(false);

  // VBE Immediate Window State (Ctrl+G)
  const [showImmediateWindow, setShowImmediateWindow] = useState(false);
  const [immediateInput, setImmediateInput] = useState('');
  const [immediateHistory, setImmediateHistory] = useState<
    Array<{ id: string; command: string; output: string; timestamp: string; isError?: boolean }>
  >([
    {
      id: 'init-1',
      command: '? ActiveSheet.Name',
      output: dataset.activeSheetName || dataset.sheets[0]?.name || 'Sheet1',
      timestamp: new Date().toLocaleTimeString()
    },
    {
      id: 'init-2',
      command: '? Application.ScreenUpdating',
      output: 'True',
      timestamp: new Date().toLocaleTimeString()
    }
  ]);

  // VBE Object Browser State (F2)
  const [showObjectBrowser, setShowObjectBrowser] = useState(false);
  const [selectedObjLibrary, setSelectedObjLibrary] = useState<'ALL' | 'Excel' | 'VBA'>('ALL');
  const [selectedObjClass, setSelectedObjClass] = useState<string>('Application');
  const [selectedObjMember, setSelectedObjMember] = useState<string>('ScreenUpdating');
  const [objSearchQuery, setObjSearchQuery] = useState('');

  // Editor Line Gutter & Highlight State
  const [highlightedLine, setHighlightedLine] = useState<number | null>(null);
  const lineGutterRef = React.useRef<HTMLDivElement>(null);
  const editorTextareaRef = React.useRef<HTMLTextAreaElement>(null);

  // Compile VBAProject Status
  const [compileStatus, setCompileStatus] = useState<'idle' | 'success' | 'error'>('idle');
  const [compileMessage, setCompileMessage] = useState<string | null>(null);

  // Delete Confirmation Modal State
  const [deleteTarget, setDeleteTarget] = useState<DeleteTarget | null>(null);

  // Add Procedure Modal State
  const [isAddModalOpen, setIsAddModalOpen] = useState(false);
  const [newProcName, setNewProcName] = useState('CalculateCustomKPI');
  const [newProcType, setNewProcType] = useState<'Sub' | 'Function'>('Sub');
  const [newProcScope, setNewProcScope] = useState<'Public' | 'Private'>('Public');
  const [selectedModuleOption, setSelectedModuleOption] = useState<string>('');
  const [customModuleName, setCustomModuleName] = useState('');
  const [newProcDescription, setNewProcDescription] = useState('Custom procedure created in XLSM Cloud Studio.');
  const [selectedTemplate, setSelectedTemplate] = useState<'basic' | 'loop' | 'fast_calc'>('basic');

  // Sync internal state if dataset changes
  React.useEffect(() => {
    setMacros(dataset.legacyMacros);
    if (!dataset.legacyMacros.find((m) => m.id === selectedMacroId) && dataset.legacyMacros[0]) {
      setSelectedMacroId(dataset.legacyMacros[0].id);
    }
  }, [dataset.legacyMacros]);

  // Reset simulation report and compile status when active macro changes
  React.useEffect(() => {
    setSimulationReport(null);
    setShowTrace(false);
    setCompileStatus('idle');
    setCompileMessage(null);
    setHighlightedLine(null);
  }, [selectedMacroId]);

  const activeMacro = useMemo(() => {
    return macros.find((m) => m.id === selectedMacroId) || macros[0];
  }, [macros, selectedMacroId]);

  // Keyboard shortcut listener: F5 (Run), F2 (Object Browser), Ctrl+G (Immediate Window)
  React.useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'F5') {
        e.preventDefault();
        handleRunSimulation();
      } else if (e.key === 'F2') {
        e.preventDefault();
        setShowObjectBrowser((prev) => !prev);
      } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'g') {
        e.preventDefault();
        setShowImmediateWindow((prev) => !prev);
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [activeMacro]);

  // Filtered classes for Object Browser
  const filteredObjectBrowserClasses = useMemo(() => {
    return VBE_CLASSES.filter((cls) => {
      if (selectedObjLibrary !== 'ALL' && cls.library !== selectedObjLibrary) return false;
      if (!objSearchQuery.trim()) return true;
      const q = objSearchQuery.toLowerCase();
      if (cls.name.toLowerCase().includes(q)) return true;
      return cls.members.some((m) => m.name.toLowerCase().includes(q) || m.description.toLowerCase().includes(q));
    });
  }, [selectedObjLibrary, objSearchQuery]);

  const currentObjClass = useMemo(() => {
    return (
      filteredObjectBrowserClasses.find((c) => c.name === selectedObjClass) ||
      filteredObjectBrowserClasses[0] ||
      VBE_CLASSES[0]
    );
  }, [filteredObjectBrowserClasses, selectedObjClass]);

  const currentObjMember = useMemo(() => {
    if (!currentObjClass) return null;
    return (
      currentObjClass.members.find((m) => m.name === selectedObjMember) ||
      currentObjClass.members[0]
    );
  }, [currentObjClass, selectedObjMember]);

  // Procedures within current module
  const currentModuleProcedures = useMemo(() => {
    if (!activeMacro) return [];
    return macros.filter((m) => m.moduleName === activeMacro.moduleName);
  }, [macros, activeMacro]);

  // Lines of code for line numbers gutter
  const codeLines = useMemo(() => {
    return (activeMacro?.code || '').split('\n');
  }, [activeMacro?.code]);

  // Set of line numbers containing diagnostic errors
  const errorLinesSet = useMemo(() => {
    const set = new Set<number>();
    if (simulationReport?.errors) {
      simulationReport.errors.forEach((err) => set.add(err.line));
    }
    return set;
  }, [simulationReport]);

  // Group macros by module type and module name
  const groupedModules = useMemo(() => {
    const map: Record<string, Record<string, LegacyMacro[]>> = {
      standard: {},
      document: {},
      class: {},
      userform: {}
    };

    macros.forEach((m) => {
      const type = m.moduleType || 'standard';
      if (!map[type]) map[type] = {};
      if (!map[type][m.moduleName]) map[type][m.moduleName] = [];
      map[type][m.moduleName].push(m);
    });

    return map;
  }, [macros]);

  const showNotification = (msg: string) => {
    setStatusNotification(msg);
    setTimeout(() => setStatusNotification(null), 3500);
  };

  // Update active macro code
  const handleCodeChange = (newCode: string) => {
    if (simulationReport) {
      setSimulationReport(null);
    }
    const updated = macros.map((m) => {
      if (m.id === activeMacro?.id) {
        return {
          ...m,
          code: newCode,
          lastModified: new Date().toISOString(),
          lineCount: newCode.split('\n').length
        };
      }
      return m;
    });
    setMacros(updated);
    onUpdateMacros(updated);
  };

  // 1-Click Enhancement: Add Performance / Speed Boost
  const handleApplySpeedBoost = () => {
    if (!activeMacro) return;
    const { enhancedCode, appliedTags } = enhanceVbaCode(activeMacro.code, {
      addSpeedBoost: true
    });
    const updated = macros.map((m) => {
      if (m.id === activeMacro.id) {
        return {
          ...m,
          code: enhancedCode,
          isEnhanced: true,
          enhancementTags: Array.from(new Set([...(m.enhancementTags || []), ...appliedTags])),
          lastModified: new Date().toISOString(),
          lineCount: enhancedCode.split('\n').length
        };
      }
      return m;
    });
    setMacros(updated);
    onUpdateMacros(updated);
    showNotification(`Enhanced "${activeMacro.name}" with ScreenUpdating & Manual Calculation boost!`);
  };

  // 1-Click Enhancement: Add Error Handling
  const handleApplyErrorHandler = () => {
    if (!activeMacro) return;
    const { enhancedCode, appliedTags } = enhanceVbaCode(activeMacro.code, {
      addErrorHandler: true
    });
    const updated = macros.map((m) => {
      if (m.id === activeMacro.id) {
        return {
          ...m,
          code: enhancedCode,
          isEnhanced: true,
          enhancementTags: Array.from(new Set([...(m.enhancementTags || []), ...appliedTags])),
          lastModified: new Date().toISOString(),
          lineCount: enhancedCode.split('\n').length
        };
      }
      return m;
    });
    setMacros(updated);
    onUpdateMacros(updated);
    showNotification(`Injected structured On Error GoTo handler into "${activeMacro.name}"`);
  };

  // 1-Click Enhancement: Add Profiler Timer
  const handleApplyProfiler = () => {
    if (!activeMacro) return;
    const { enhancedCode, appliedTags } = enhanceVbaCode(activeMacro.code, {
      addProfiler: true
    });
    const updated = macros.map((m) => {
      if (m.id === activeMacro.id) {
        return {
          ...m,
          code: enhancedCode,
          isEnhanced: true,
          enhancementTags: Array.from(new Set([...(m.enhancementTags || []), ...appliedTags])),
          lastModified: new Date().toISOString(),
          lineCount: enhancedCode.split('\n').length
        };
      }
      return m;
    });
    setMacros(updated);
    onUpdateMacros(updated);
    showNotification(`Injected precision execution profiler timer into "${activeMacro.name}"`);
  };

  // Format Indentation
  const handleAutoFormat = () => {
    if (!activeMacro) return;
    const formatted = formatVbaCode(activeMacro.code);
    handleCodeChange(formatted);
    showNotification(`Formatted indentation for "${activeMacro.name}"`);
  };

  // Revert back to original ingested VBA
  const handleRevert = () => {
    if (!activeMacro) return;
    const updated = macros.map((m) => {
      if (m.id === activeMacro.id) {
        return {
          ...m,
          code: m.originalVba,
          isEnhanced: false,
          enhancementTags: [],
          lastModified: undefined
        };
      }
      return m;
    });
    setMacros(updated);
    onUpdateMacros(updated);
    showNotification(`Reverted "${activeMacro.name}" back to original raw VBA code.`);
  };

  // Delete initiation for single procedure
  const promptDeleteProcedure = (macro: LegacyMacro, e?: React.MouseEvent) => {
    if (e) e.stopPropagation();
    setDeleteTarget({
      type: 'procedure',
      id: macro.id,
      name: macro.name,
      moduleName: macro.moduleName
    });
  };

  // Delete initiation for entire module
  const promptDeleteModule = (moduleName: string, procs: LegacyMacro[], e?: React.MouseEvent) => {
    if (e) e.stopPropagation();
    setDeleteTarget({
      type: 'module',
      name: moduleName,
      moduleName: moduleName,
      count: procs.length
    });
  };

  // Confirm and execute deletion
  const handleConfirmDelete = () => {
    if (!deleteTarget) return;

    if (deleteTarget.type === 'procedure' && deleteTarget.id) {
      const targetId = deleteTarget.id;
      const targetName = deleteTarget.name;
      const updated = macros.filter((m) => m.id !== targetId);
      setMacros(updated);
      onUpdateMacros(updated);
      if (selectedMacroId === targetId) {
        setSelectedMacroId(updated[0]?.id || '');
      }
      showNotification(`Deleted procedure "${targetName}".`);
    } else if (deleteTarget.type === 'module' && deleteTarget.name) {
      const targetModuleName = deleteTarget.name;
      const count = deleteTarget.count || 0;
      const updated = macros.filter((m) => m.moduleName !== targetModuleName);
      setMacros(updated);
      onUpdateMacros(updated);
      const remainingSelected = updated.find((m) => m.id === selectedMacroId);
      if (!remainingSelected) {
        setSelectedMacroId(updated[0]?.id || '');
      }
      showNotification(`Deleted module "${targetModuleName}" and ${count} routine(s).`);
    }

    setDeleteTarget(null);
  };

  // Existing module names for target selection
  const existingModules = useMemo(() => {
    const list: string[] = [];
    macros.forEach((m) => {
      if (m.moduleName && !list.includes(m.moduleName)) {
        list.push(m.moduleName);
      }
    });
    if (list.length === 0) list.push('Module1_FinancialCore');
    return list;
  }, [macros]);

  // Open add procedure modal
  const openAddProcedureModal = () => {
    setSelectedModuleOption(activeMacro?.moduleName || existingModules[0] || 'Module1_CoreProcedures');
    setCustomModuleName('');
    setNewProcName(`Procedure_${Math.floor(100 + Math.random() * 900)}`);
    setNewProcType('Sub');
    setNewProcScope('Public');
    setSelectedTemplate('basic');
    setNewProcDescription('Custom procedure created in XLSM Cloud Studio.');
    setIsAddModalOpen(true);
  };

  // Live preview of generated code
  const previewGeneratedCode = useMemo(() => {
    const rawClean = newProcName.trim().replace(/[^a-zA-Z0-9_]/g, '_');
    const cleanName = rawClean || 'CustomRoutine';
    const desc = newProcDescription.trim() || 'Custom procedure created in XLSM Cloud Studio.';
    const scopePrefix = newProcScope === 'Private' ? 'Private ' : '';

    if (newProcType === 'Function') {
      return `' ========================================================
' Procedure: Function ${cleanName}
' Description: ${desc}
' ========================================================
${scopePrefix}Function ${cleanName}(ByVal inputValue As Double) As Double
    On Error GoTo ErrHandler
    
    ' Perform calculation or metric evaluation
    Dim result As Double
    result = inputValue * 1.15
    
    ${cleanName} = result
    Exit Function

ErrHandler:
    MsgBox "Error in ${cleanName}: " & Err.Description, vbCritical
    ${cleanName} = 0
End Function`;
    }

    if (selectedTemplate === 'loop') {
      return `' ========================================================
' Procedure: Sub ${cleanName}
' Description: ${desc}
' ========================================================
${scopePrefix}Sub ${cleanName}()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim r As Long
    For r = 2 To lastRow
        ' Example: Evaluate row values
        If ws.Cells(r, 1).Value <> "" Then
            ' ws.Cells(r, 2).Value = ws.Cells(r, 2).Value * 1.05
        End If
    Next r
    
    MsgBox "Processed " & (lastRow - 1) & " rows in " & ws.Name, vbInformation
End Sub`;
    }

    if (selectedTemplate === 'fast_calc') {
      return `' ========================================================
' Procedure: Sub ${cleanName}
' Description: ${desc}
' ========================================================
${scopePrefix}Sub ${cleanName}()
    ' Optimization: Prevent screen flicker and delay
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Execute core data transformation here
    
CleanExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub

ErrHandler:
    MsgBox "Execution error in ${cleanName}: " & Err.Description, vbCritical
    Resume CleanExit
End Sub`;
    }

    // Default 'basic' template
    return `' ========================================================
' Procedure: Sub ${cleanName}
' Description: ${desc}
' ========================================================
${scopePrefix}Sub ${cleanName}()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    MsgBox "Workflow executed on sheet: " & ws.Name, vbInformation
End Sub`;
  }, [newProcName, newProcType, newProcScope, selectedTemplate, newProcDescription]);

  // Submit procedure creation
  const handleCreateProcedure = (e?: React.FormEvent) => {
    if (e) e.preventDefault();
    const rawClean = newProcName.trim().replace(/[^a-zA-Z0-9_]/g, '_');
    const cleanName = rawClean || 'CustomRoutine';

    const targetModule =
      selectedModuleOption === '__new__'
        ? customModuleName.trim().replace(/[^a-zA-Z0-9_]/g, '_') || 'Module_Custom'
        : selectedModuleOption;

    const formattedSignature = `${newProcScope === 'Private' ? 'Private ' : ''}${newProcType} ${cleanName}()`;
    const code = previewGeneratedCode;

    const newMacro: LegacyMacro = {
      id: `custom-macro-${Date.now()}-${Math.random().toString(36).substring(2, 6)}`,
      name: formattedSignature,
      moduleName: targetModule,
      moduleType:
        targetModule.toLowerCase().includes('sheet') || targetModule.toLowerCase().includes('workbook')
          ? 'document'
          : 'standard',
      procedureType: newProcType,
      description: newProcDescription.trim() || 'Custom procedure created in XLSM Cloud Studio.',
      triggerEvent: newProcType === 'Function' ? 'User-Defined Formula (UDF)' : 'Button Call / Macro Runner',
      originalVba: code,
      code: code,
      lineCount: code.split('\n').length,
      isEnhanced: true,
      enhancementTags: ['User Created']
    };

    const updated = [...macros, newMacro];
    setMacros(updated);
    onUpdateMacros(updated);
    setSelectedMacroId(newMacro.id);
    setActiveViewTab('editor');
    setIsAddModalOpen(false);
    showNotification(`Created procedure "${formattedSignature}" in ${targetModule}!`);
  };

  // Run in-browser test simulation
  const handleRunSimulation = () => {
    if (!activeMacro) return;
    setIsSimulating(true);
    setTimeout(() => {
      const report = simulateVbaExecution(activeMacro.code, dataset);
      setSimulationReport(report);
      setIsSimulating(false);
      if (!report.success) {
        setShowTrace(true);
        showNotification(`Simulation detected ${report.errors.length} syntax/runtime fault(s) in "${activeMacro.name}".`);
      } else if (report.warnings.length > 0) {
        showNotification(`Simulation verified with ${report.warnings.length} warning(s).`);
      } else {
        showNotification(`Simulation verified: 0 faults detected in "${activeMacro.name}".`);
      }
    }, 200);
  };

  // VBE Compile VBAProject action (Debug -> Compile VBAProject)
  const handleCompileProject = () => {
    if (!activeMacro) return;
    const report = simulateVbaExecution(activeMacro.code, dataset);
    setSimulationReport(report);
    if (!report.success && report.errors.length > 0) {
      setCompileStatus('error');
      setCompileMessage(report.errors[0].message);
      setHighlightedLine(report.errors[0].line);
      showNotification(`Compile Error: ${report.errors[0].message}`);
      if (editorTextareaRef.current) {
        const lineH = 24;
        editorTextareaRef.current.scrollTop = Math.max(0, (report.errors[0].line - 3) * lineH);
      }
    } else {
      setCompileStatus('success');
      setCompileMessage(`✓ VBAProject compiled successfully. 0 syntax or typo errors detected across ${report.totalLines} lines.`);
      setHighlightedLine(null);
      showNotification(`VBAProject compiled: 0 syntax or typo faults.`);
    }
  };

  // VBE Immediate Window expression evaluator (? <expr>)
  const handleExecuteImmediate = (cmdToRun?: string) => {
    const cmd = (cmdToRun !== undefined ? cmdToRun : immediateInput).trim();
    if (!cmd) return;
    const output = evaluateImmediateExpression(cmd, dataset);
    const isError = output.toLowerCase().includes('compile error') || output.toLowerCase().includes('syntax error');
    setImmediateHistory((prev) => [
      ...prev,
      {
        id: `imm-${Date.now()}-${Math.random().toString(36).substring(2, 5)}`,
        command: cmd,
        output,
        timestamp: new Date().toLocaleTimeString(),
        isError
      }
    ]);
    if (cmdToRun === undefined) {
      setImmediateInput('');
    }
  };

  // Sync scrolling between line numbers gutter and code editor
  const handleEditorScroll = (e: React.UIEvent<HTMLTextAreaElement>) => {
    if (lineGutterRef.current) {
      lineGutterRef.current.scrollTop = e.currentTarget.scrollTop;
    }
  };

  // Jump and highlight specific line
  const handleGoToLine = (line: number) => {
    setHighlightedLine(line);
    if (editorTextareaRef.current) {
      const lineH = 24;
      editorTextareaRef.current.scrollTop = Math.max(0, (line - 4) * lineH);
      editorTextareaRef.current.focus();
    }
  };

  // Insert code snippet from Object Browser into active macro
  const handleInsertSnippet = (snippet: string) => {
    if (!activeMacro) return;
    const newCode = activeMacro.code + '\n    ' + snippet;
    handleCodeChange(newCode);
    showNotification(`Inserted '${snippet}' into ${activeMacro.name}.`);
    setShowObjectBrowser(false);
  };

  // Download upgraded .xlsm workbook
  const handleDownloadUpgradedXlsm = async () => {
    setDownloading(true);
    try {
      if (dataset.rawZip) {
        const upgradedBlob = await repackUpgradedXlsm(dataset.rawZip, macros);
        downloadWorkbookFile(upgradedBlob, `Upgraded_${dataset.filename || 'Workbook.xlsm'}`);
        showNotification(`Downloaded upgraded ${dataset.filename} with enhanced macros!`);
      } else {
        // Fallback: build a package containing modified macros and spreadsheet data
        const textData = macros
          .map((m) => `' Module: ${m.moduleName}\n' Procedure: ${m.name}\n${m.code}\n`)
          .join('\n\n');
        const blob = new Blob([textData], { type: 'application/vnd.ms-excel.sheet.macroEnabled.12' });
        downloadWorkbookFile(blob, `Upgraded_${dataset.filename || 'Workbook.xlsm'}`);
        showNotification(`Downloaded upgraded workbook package!`);
      }
    } catch (err: any) {
      console.error('Failed to repack workbook:', err);
      showNotification('Failed to repack workbook file. Exporting .bas source instead.');
      handleExportAllVba();
    } finally {
      setDownloading(false);
    }
  };

  // Download VBA source module (.bas)
  const handleExportAllVba = () => {
    const fullCode = macros
      .map(
        (m) =>
          `' ========================================================\n' Module: ${m.moduleName}\n' Procedure: ${m.name}\n' Trigger: ${m.triggerEvent}\n' Status: ${m.isEnhanced ? 'ENHANCED' : 'ORIGINAL'}\n' ========================================================\n${m.code}\n`
      )
      .join('\n\n');

    const blob = new Blob([fullCode], { type: 'text/plain' });
    downloadWorkbookFile(blob, `${dataset.name.replace(/\s+/g, '_')}_Upgraded_Macros.bas`);
    showNotification('Exported all modified and enhanced macros to .bas file!');
  };

  const isModified = activeMacro && activeMacro.code !== activeMacro.originalVba;

  return (
    <div className="space-y-4">
      {/* Top Banner: Ingestion & Upgrade Actions */}
      <div className="rounded-xl border border-slate-800 bg-slate-900/90 p-5 shadow-sm">
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4">
          <div>
            <div className="flex items-center gap-2.5">
              <div className="p-2 rounded-lg bg-indigo-950/80 border border-indigo-500/30 text-indigo-400">
                <Code2 className="w-5 h-5" />
              </div>
              <div>
                <div className="flex flex-wrap items-center gap-2">
                  <h3 className="text-base font-semibold text-white">
                    VBA Macro & Module Engineering Studio
                  </h3>
                  <span className="text-[10px] font-mono px-2 py-0.5 rounded bg-indigo-950 text-indigo-300 border border-indigo-500/30">
                    MS Office VBE Engine
                  </span>
                </div>
                <div className="flex flex-wrap items-center gap-2 mt-0.5 text-xs text-slate-400">
                  <span>Ingest, inspect, compile, debug, and download upgraded macros.</span>
                  <span className="text-slate-600">·</span>
                  <span>Developed by <span className="text-indigo-400 font-semibold">Ammar Thaqif</span></span>
                </div>
              </div>
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-2.5 font-mono text-xs">
            <button
              onClick={openAddProcedureModal}
              className="flex items-center gap-1.5 px-3 py-1.5 bg-indigo-950/80 hover:bg-indigo-900 border border-indigo-500/40 rounded-lg text-indigo-300 hover:text-white transition-colors"
            >
              <Plus className="w-3.5 h-3.5" />
              <span>New Procedure</span>
            </button>

            <button
              onClick={handleExportAllVba}
              className="flex items-center gap-1.5 px-3 py-1.5 bg-slate-800 hover:bg-slate-700 border border-slate-700 rounded-lg text-slate-200 transition-colors"
            >
              <FileText className="w-3.5 h-3.5 text-indigo-400" />
              <span>Export Code (.bas)</span>
            </button>

            <button
              onClick={handleDownloadUpgradedXlsm}
              disabled={downloading}
              className="flex items-center gap-1.5 px-3.5 py-1.5 bg-emerald-600 hover:bg-emerald-500 text-white font-semibold rounded-lg shadow-sm transition-colors"
            >
              <Download className="w-3.5 h-3.5" />
              <span>{downloading ? 'Packing .xlsm...' : 'Download Upgraded .xlsm'}</span>
            </button>
          </div>
        </div>

        {/* Real-time status notification toast */}
        {statusNotification && (
          <div className="mt-3 p-2.5 rounded-lg bg-emerald-950/60 border border-emerald-500/30 text-emerald-300 text-xs flex items-center gap-2 animate-in fade-in">
            <Check className="w-4 h-4 shrink-0 text-emerald-400" />
            <span>{statusNotification}</span>
          </div>
        )}
      </div>

      {/* Main Dual-Pane Studio Layout */}
      <div className="grid grid-cols-1 lg:grid-cols-12 gap-5 items-start">
        {/* Left Pane: Hierarchical Module & Procedure Tree (4 cols) */}
        <div className="lg:col-span-4 rounded-xl border border-slate-800 bg-slate-900/90 p-4 shadow-sm space-y-4">
          <div className="flex items-center justify-between pb-2 border-b border-slate-800">
            <div className="flex items-center gap-2">
              <FolderTree className="w-4 h-4 text-indigo-400" />
              <h4 className="text-xs font-semibold uppercase tracking-wider text-slate-300">
                Workbook Project Tree
              </h4>
            </div>
            <button
              onClick={openAddProcedureModal}
              title="Add a new VBA procedure or module"
              className="flex items-center gap-1 px-2 py-0.5 text-[11px] font-mono text-indigo-300 hover:text-white bg-indigo-950/80 hover:bg-indigo-900 border border-indigo-500/30 rounded transition-colors"
            >
              <Plus className="w-3 h-3" />
              <span>New Procedure</span>
            </button>
          </div>

          <div className="space-y-3 max-h-[640px] overflow-y-auto pr-1">
            {macros.length === 0 && (
              <div className="p-6 text-center border border-dashed border-slate-800 rounded-lg">
                <p className="text-xs text-slate-400 mb-2">No macros or modules found in workbook.</p>
                <button
                  type="button"
                  onClick={openAddProcedureModal}
                  className="text-xs text-indigo-400 hover:text-indigo-300 font-medium inline-flex items-center gap-1"
                >
                  <Plus className="w-3.5 h-3.5" />
                  <span>Add New Procedure</span>
                </button>
              </div>
            )}

            {/* 1. Standard Modules */}
            {Object.keys(groupedModules.standard).length > 0 && (
              <div className="space-y-1">
                <button
                  onClick={() =>
                    setExpandedModules((prev) => ({ ...prev, standard: !prev.standard }))
                  }
                  className="w-full flex items-center justify-between py-1 px-1.5 text-xs font-semibold text-slate-400 hover:text-slate-200"
                >
                  <div className="flex items-center gap-1.5">
                    {expandedModules.standard ? (
                      <ChevronDown className="w-3.5 h-3.5" />
                    ) : (
                      <ChevronRight className="w-3.5 h-3.5" />
                    )}
                    <span>Standard Modules</span>
                  </div>
                  <span className="text-[10px] font-mono text-slate-500">
                    {Object.values(groupedModules.standard).flat().length}
                  </span>
                </button>

                {expandedModules.standard && (
                  <div className="pl-3 space-y-2 border-l border-slate-800 ml-2">
                    {Object.entries(groupedModules.standard).map(([moduleName, procs]) => (
                      <div key={moduleName} className="space-y-1">
                        <div className="text-[11px] font-mono font-medium text-slate-400 px-1 py-0.5 flex items-center justify-between group/mod">
                          <div className="flex items-center gap-1.5 truncate">
                            <Layers className="w-3 h-3 text-slate-500 shrink-0" />
                            <span className="truncate">{moduleName}</span>
                          </div>
                          <button
                            type="button"
                            onClick={(e) => promptDeleteModule(moduleName, procs, e)}
                            title={`Delete module "${moduleName}" (${procs.length} routine${procs.length === 1 ? '' : 's'})`}
                            className="p-1 text-slate-500 hover:text-rose-400 hover:bg-rose-950/60 rounded transition-colors shrink-0"
                          >
                            <Trash2 className="w-3 h-3" />
                          </button>
                        </div>
                        {procs.map((m) => (
                          <div
                            key={m.id}
                            onClick={() => setSelectedMacroId(m.id)}
                            className={`group p-2 rounded-lg text-xs cursor-pointer transition-all ${
                              selectedMacroId === m.id
                                ? 'bg-indigo-950/70 text-white border border-indigo-500/50 shadow-sm'
                                : 'text-slate-300 hover:bg-slate-800/60'
                            }`}
                          >
                            <div className="flex items-center justify-between gap-1.5">
                              <span className="font-mono font-semibold truncate">{m.name}</span>
                              <div className="flex items-center gap-1 shrink-0">
                                <span className="text-[9px] font-mono px-1 py-0.2 rounded bg-slate-800 text-slate-400 uppercase">
                                  {m.procedureType}
                                </span>
                                <button
                                  type="button"
                                  onClick={(e) => promptDeleteProcedure(m, e)}
                                  title={`Delete procedure "${m.name}"`}
                                  className="p-1 text-slate-500 hover:text-rose-400 hover:bg-rose-950/60 rounded transition-colors"
                                >
                                  <Trash2 className="w-3 h-3" />
                                </button>
                              </div>
                            </div>
                            <div className="mt-1 flex items-center gap-2 text-[10px] font-mono text-slate-500">
                              <span>{m.lineCount || m.code.split('\n').length} lines</span>
                              {m.isEnhanced && (
                                <span className="text-emerald-400 font-semibold">Enhanced</span>
                              )}
                              {m.code !== m.originalVba && !m.isEnhanced && (
                                <span className="text-amber-400">Edited</span>
                              )}
                            </div>
                          </div>
                        ))}
                      </div>
                    ))}
                  </div>
                )}
              </div>
            )}

            {/* 2. Document Modules (ThisWorkbook, Sheets) */}
            {Object.keys(groupedModules.document).length > 0 && (
              <div className="space-y-1">
                <button
                  onClick={() =>
                    setExpandedModules((prev) => ({ ...prev, document: !prev.document }))
                  }
                  className="w-full flex items-center justify-between py-1 px-1.5 text-xs font-semibold text-slate-400 hover:text-slate-200"
                >
                  <div className="flex items-center gap-1.5">
                    {expandedModules.document ? (
                      <ChevronDown className="w-3.5 h-3.5" />
                    ) : (
                      <ChevronRight className="w-3.5 h-3.5" />
                    )}
                    <span>Microsoft Excel Objects</span>
                  </div>
                  <span className="text-[10px] font-mono text-slate-500">
                    {Object.values(groupedModules.document).flat().length}
                  </span>
                </button>

                {expandedModules.document && (
                  <div className="pl-3 space-y-2 border-l border-slate-800 ml-2">
                    {Object.entries(groupedModules.document).map(([moduleName, procs]) => (
                      <div key={moduleName} className="space-y-1">
                        <div className="text-[11px] font-mono font-medium text-slate-400 px-1 py-0.5 flex items-center justify-between group/mod">
                          <div className="flex items-center gap-1.5 truncate">
                            <Layers className="w-3 h-3 text-slate-500 shrink-0" />
                            <span className="truncate">{moduleName}</span>
                          </div>
                          <button
                            type="button"
                            onClick={(e) => promptDeleteModule(moduleName, procs, e)}
                            title={`Delete module "${moduleName}" (${procs.length} routine${procs.length === 1 ? '' : 's'})`}
                            className="p-1 text-slate-500 hover:text-rose-400 hover:bg-rose-950/60 rounded transition-colors shrink-0"
                          >
                            <Trash2 className="w-3 h-3" />
                          </button>
                        </div>
                        {procs.map((m) => (
                          <div
                            key={m.id}
                            onClick={() => setSelectedMacroId(m.id)}
                            className={`group p-2 rounded-lg text-xs cursor-pointer transition-all ${
                              selectedMacroId === m.id
                                ? 'bg-indigo-950/70 text-white border border-indigo-500/50 shadow-sm'
                                : 'text-slate-300 hover:bg-slate-800/60'
                            }`}
                          >
                            <div className="flex items-center justify-between gap-1.5">
                              <span className="font-mono font-semibold truncate">{m.name}</span>
                              <div className="flex items-center gap-1 shrink-0">
                                <span className="text-[9px] font-mono px-1 py-0.2 rounded bg-slate-800 text-slate-400 uppercase">
                                  {m.procedureType}
                                </span>
                                <button
                                  type="button"
                                  onClick={(e) => promptDeleteProcedure(m, e)}
                                  title={`Delete procedure "${m.name}"`}
                                  className="p-1 text-slate-500 hover:text-rose-400 hover:bg-rose-950/60 rounded transition-colors"
                                >
                                  <Trash2 className="w-3 h-3" />
                                </button>
                              </div>
                            </div>
                            <div className="mt-1 flex items-center gap-2 text-[10px] font-mono text-slate-500">
                              <span>{m.lineCount || m.code.split('\n').length} lines</span>
                              {m.isAutoRun && (
                                <span className="text-amber-400 font-semibold">Auto-Run</span>
                              )}
                              {m.isEnhanced && (
                                <span className="text-emerald-400 font-semibold">Enhanced</span>
                              )}
                            </div>
                          </div>
                        ))}
                      </div>
                    ))}
                  </div>
                )}
              </div>
            )}

            {/* 3. Class & UserForm Modules if present */}
            {groupedModules.class && Object.keys(groupedModules.class).length > 0 && (
              <div className="space-y-1">
                <button
                  onClick={() =>
                    setExpandedModules((prev) => ({ ...prev, class: !prev.class }))
                  }
                  className="w-full flex items-center justify-between py-1 px-1.5 text-xs font-semibold text-slate-400 hover:text-slate-200"
                >
                  <div className="flex items-center gap-1.5">
                    {expandedModules.class ? (
                      <ChevronDown className="w-3.5 h-3.5" />
                    ) : (
                      <ChevronRight className="w-3.5 h-3.5" />
                    )}
                    <span>Class Modules</span>
                  </div>
                  <span className="text-[10px] font-mono text-slate-500">
                    {Object.values(groupedModules.class).flat().length}
                  </span>
                </button>

                {expandedModules.class && (
                  <div className="pl-3 space-y-2 border-l border-slate-800 ml-2">
                    {Object.entries(groupedModules.class).map(([moduleName, procs]) => (
                      <div key={moduleName} className="space-y-1">
                        <div className="text-[11px] font-mono font-medium text-slate-400 px-1 py-0.5 flex items-center justify-between group/mod">
                          <div className="flex items-center gap-1.5 truncate">
                            <Layers className="w-3 h-3 text-slate-500 shrink-0" />
                            <span className="truncate">{moduleName}</span>
                          </div>
                          <button
                            type="button"
                            onClick={(e) => promptDeleteModule(moduleName, procs, e)}
                            title={`Delete module "${moduleName}" (${procs.length} routine${procs.length === 1 ? '' : 's'})`}
                            className="p-1 text-slate-500 hover:text-rose-400 hover:bg-rose-950/60 rounded transition-colors shrink-0"
                          >
                            <Trash2 className="w-3 h-3" />
                          </button>
                        </div>
                        {procs.map((m) => (
                          <div
                            key={m.id}
                            onClick={() => setSelectedMacroId(m.id)}
                            className={`group p-2 rounded-lg text-xs cursor-pointer transition-all ${
                              selectedMacroId === m.id
                                ? 'bg-indigo-950/70 text-white border border-indigo-500/50 shadow-sm'
                                : 'text-slate-300 hover:bg-slate-800/60'
                            }`}
                          >
                            <div className="flex items-center justify-between gap-1.5">
                              <span className="font-mono font-semibold truncate">{m.name}</span>
                              <div className="flex items-center gap-1 shrink-0">
                                <span className="text-[9px] font-mono px-1 py-0.2 rounded bg-slate-800 text-slate-400 uppercase">
                                  {m.procedureType}
                                </span>
                                <button
                                  type="button"
                                  onClick={(e) => promptDeleteProcedure(m, e)}
                                  title={`Delete procedure "${m.name}"`}
                                  className="p-1 text-slate-500 hover:text-rose-400 hover:bg-rose-950/60 rounded transition-colors"
                                >
                                  <Trash2 className="w-3 h-3" />
                                </button>
                              </div>
                            </div>
                            <div className="mt-1 flex items-center gap-2 text-[10px] font-mono text-slate-500">
                              <span>{m.lineCount || m.code.split('\n').length} lines</span>
                            </div>
                          </div>
                        ))}
                      </div>
                    ))}
                  </div>
                )}
              </div>
            )}
          </div>
        </div>

        {/* Right Pane: Code Editor & Enhancement Suite (8 cols) */}
        <div className="lg:col-span-8 rounded-xl border border-slate-800 bg-slate-900/90 p-5 shadow-sm space-y-4">
          {!activeMacro ? (
            <div className="flex flex-col items-center justify-center py-20 px-4 text-center">
              <div className="p-3.5 rounded-xl bg-slate-800/70 border border-slate-700/60 text-slate-400 mb-3">
                <FileCode className="w-8 h-8 text-indigo-400" />
              </div>
              <h4 className="text-base font-semibold text-white">No Procedure Selected</h4>
              <p className="text-xs text-slate-400 max-w-sm mt-1 mb-5 leading-relaxed">
                {macros.length === 0
                  ? 'All procedures in this workbook have been deleted. You can create a new subroutine or function to start automating.'
                  : 'Select a procedure from the Workbook Project Tree on the left to inspect, edit, or optimize its code.'}
              </p>
              <button
                type="button"
                onClick={openAddProcedureModal}
                className="flex items-center gap-1.5 px-4 py-2 bg-indigo-600 hover:bg-indigo-500 text-white text-xs font-semibold rounded-lg shadow-sm transition-colors"
              >
                <Plus className="w-4 h-4" />
                <span>Create New Procedure</span>
              </button>
            </div>
          ) : (
            <>
              {/* Header & Meta */}
              <div className="flex flex-wrap items-center justify-between gap-3 pb-3 border-b border-slate-800">
                <div>
                  <div className="flex flex-wrap items-center gap-2">
                    <FileCode className="w-4 h-4 text-indigo-400" />
                    <h3 className="text-sm font-bold text-white font-mono">{activeMacro.name}</h3>
                    <div className="flex items-center gap-1 bg-slate-800 border border-slate-700 rounded px-2 py-0.5">
                      <span className="text-[10px] font-mono text-slate-300">
                        Module: {activeMacro.moduleName}
                      </span>
                      <button
                        type="button"
                        onClick={() => {
                          const moduleProcs = macros.filter((m) => m.moduleName === activeMacro.moduleName);
                          promptDeleteModule(activeMacro.moduleName, moduleProcs);
                        }}
                        title={`Delete module "${activeMacro.moduleName}" and its procedures`}
                        className="p-0.5 text-slate-500 hover:text-rose-400 rounded transition-colors"
                      >
                        <Trash2 className="w-2.5 h-2.5" />
                      </button>
                    </div>
                    {isModified && (
                      <span className="text-[10px] font-mono px-1.5 py-0.5 rounded bg-amber-950/80 text-amber-300 border border-amber-500/30">
                        Modified
                      </span>
                    )}
                  </div>
                  <p className="text-xs text-slate-400 mt-1">{activeMacro.description}</p>
                </div>

                {/* View Mode Toggle */}
                <div className="flex items-center gap-1 bg-slate-950 p-0.5 rounded-lg border border-slate-800 text-xs">
                  <button
                    onClick={() => setActiveViewTab('editor')}
                    className={`px-3 py-1 rounded font-medium transition-colors ${
                      activeViewTab === 'editor'
                        ? 'bg-slate-800 text-white'
                        : 'text-slate-400 hover:text-slate-200'
                    }`}
                  >
                    Code Editor
                  </button>
                  <button
                    onClick={() => setActiveViewTab('diff')}
                    className={`px-3 py-1 rounded font-medium transition-colors flex items-center gap-1 ${
                      activeViewTab === 'diff'
                        ? 'bg-slate-800 text-white'
                        : 'text-slate-400 hover:text-slate-200'
                    }`}
                  >
                    <ArrowRightLeft className="w-3 h-3" />
                    <span>Compare Original</span>
                  </button>
                </div>
              </div>

              {/* Enhancement Toolbar */}
              <div className="flex flex-wrap items-center justify-between gap-2 p-2.5 bg-slate-950/60 rounded-lg border border-slate-800 text-xs">
                <span className="text-[11px] font-mono text-slate-400 uppercase font-semibold">
                  1-Click Enhancements:
                </span>

                <div className="flex flex-wrap items-center gap-1.5">
                  <button
                    onClick={handleApplySpeedBoost}
                    title="Inject ScreenUpdating = False and xlCalculationManual to prevent screen freeze"
                    className="flex items-center gap-1 px-2.5 py-1 bg-slate-800 hover:bg-slate-700 text-slate-200 rounded border border-slate-700 transition-colors"
                  >
                    <Zap className="w-3.5 h-3.5 text-amber-400" />
                    <span>Speed Optimization</span>
                  </button>

                  <button
                    onClick={handleApplyErrorHandler}
                    title="Inject On Error GoTo ErrHandler and recovery exit logic"
                    className="flex items-center gap-1 px-2.5 py-1 bg-slate-800 hover:bg-slate-700 text-slate-200 rounded border border-slate-700 transition-colors"
                  >
                    <ShieldAlert className="w-3.5 h-3.5 text-indigo-400" />
                    <span>Error Handling</span>
                  </button>

                  <button
                    onClick={handleApplyProfiler}
                    title="Inject Timer profiler to log execution latency in seconds"
                    className="flex items-center gap-1 px-2.5 py-1 bg-slate-800 hover:bg-slate-700 text-slate-200 rounded border border-slate-700 transition-colors"
                  >
                    <Clock className="w-3.5 h-3.5 text-emerald-400" />
                    <span>Profiler Timer</span>
                  </button>

                  <button
                    onClick={handleAutoFormat}
                    title="Re-indent nested blocks, loops, and procedures"
                    className="flex items-center gap-1 px-2.5 py-1 bg-slate-800 hover:bg-slate-700 text-slate-200 rounded border border-slate-700 transition-colors"
                  >
                    <Sparkles className="w-3.5 h-3.5 text-cyan-400" />
                    <span>Format Code</span>
                  </button>

                  {isModified && (
                    <button
                      onClick={handleRevert}
                      title="Revert back to original raw VBA"
                      className="flex items-center gap-1 px-2.5 py-1 bg-amber-950/60 hover:bg-amber-900/60 text-amber-300 rounded border border-amber-500/30 transition-colors"
                    >
                      <RotateCcw className="w-3.5 h-3.5" />
                      <span>Revert</span>
                    </button>
                  )}

                  <button
                    type="button"
                    onClick={() => promptDeleteProcedure(activeMacro)}
                    title={`Delete procedure "${activeMacro.name}"`}
                    className="flex items-center gap-1 px-2.5 py-1 bg-rose-950/50 hover:bg-rose-900/70 text-rose-300 rounded border border-rose-500/30 transition-colors"
                  >
                    <Trash2 className="w-3.5 h-3.5 text-rose-400" />
                    <span>Delete Procedure</span>
                  </button>
                </div>
              </div>

              {/* Active Applied Tags */}
              {activeMacro.enhancementTags && activeMacro.enhancementTags.length > 0 && (
                <div className="flex flex-wrap items-center gap-1.5 text-xs font-mono">
                  <span className="text-slate-500 text-[11px]">Applied Upgrades:</span>
                  {activeMacro.enhancementTags.map((tag, i) => (
                    <span
                      key={i}
                      className="px-2 py-0.5 rounded bg-emerald-950 text-emerald-300 border border-emerald-500/30 text-[10px]"
                    >
                      ✓ {tag}
                    </span>
                  ))}
                </div>
              )}

              {/* VBE Navigation & Action Ribbon */}
              <div className="flex flex-wrap items-center justify-between gap-2.5 p-2 bg-slate-950 rounded-lg border border-slate-800 text-xs">
                {/* Left: VBE Object & Procedure Navigation Dropdowns */}
                <div className="flex flex-wrap items-center gap-2">
                  <div className="flex items-center gap-1.5 bg-slate-900 px-2 py-1 rounded border border-slate-700/80">
                    <span className="text-[10px] font-mono text-slate-500 uppercase">Obj:</span>
                    <select
                      className="bg-transparent text-slate-200 font-mono text-xs focus:outline-none cursor-pointer"
                      defaultValue="(General)"
                    >
                      <option value="(General)" className="bg-slate-900 text-white">(General)</option>
                      <option value="Worksheet" className="bg-slate-900 text-white">Worksheet</option>
                      <option value="Workbook" className="bg-slate-900 text-white">Workbook</option>
                      <option value="Application" className="bg-slate-900 text-white">Application</option>
                    </select>
                  </div>

                  <div className="flex items-center gap-1.5 bg-slate-900 px-2 py-1 rounded border border-slate-700/80">
                    <span className="text-[10px] font-mono text-slate-500 uppercase">Proc:</span>
                    <select
                      value={activeMacro.id}
                      onChange={(e) => setSelectedMacroId(e.target.value)}
                      className="bg-transparent text-indigo-300 font-mono text-xs font-semibold focus:outline-none cursor-pointer max-w-[200px] truncate"
                    >
                      {currentModuleProcedures.map((p) => (
                        <option key={p.id} value={p.id} className="bg-slate-900 text-white">
                          {p.name}
                        </option>
                      ))}
                    </select>
                  </div>
                </div>

                {/* Right: VBE Action Buttons */}
                <div className="flex flex-wrap items-center gap-1.5 font-mono text-xs">
                  <button
                    type="button"
                    onClick={handleCompileProject}
                    title="Debug -> Compile VBAProject: Checks code for syntax and typo errors"
                    className={`flex items-center gap-1 px-2.5 py-1 rounded border transition-colors ${
                      compileStatus === 'error'
                        ? 'bg-rose-950/80 text-rose-300 border-rose-500/50 hover:bg-rose-900'
                        : compileStatus === 'success'
                        ? 'bg-emerald-950/80 text-emerald-300 border-emerald-500/50 hover:bg-emerald-900'
                        : 'bg-slate-900 hover:bg-slate-800 text-slate-200 border-slate-700'
                    }`}
                  >
                    <Cpu className="w-3.5 h-3.5 text-cyan-400" />
                    <span>Compile VBAProject</span>
                  </button>

                  <button
                    type="button"
                    onClick={handleRunSimulation}
                    disabled={isSimulating}
                    title="Run Sub / UserForm (F5)"
                    className="flex items-center gap-1 px-2.5 py-1 bg-indigo-600 hover:bg-indigo-500 disabled:opacity-60 text-white font-semibold rounded shadow-sm transition-colors"
                  >
                    {isSimulating ? (
                      <RefreshCw className="w-3.5 h-3.5 animate-spin" />
                    ) : (
                      <Play className="w-3.5 h-3.5 fill-current" />
                    )}
                    <span>Run (F5)</span>
                  </button>

                  <button
                    type="button"
                    onClick={() => setShowImmediateWindow(!showImmediateWindow)}
                    title="View Immediate Window (Ctrl+G)"
                    className={`flex items-center gap-1 px-2 py-1 rounded border transition-colors ${
                      showImmediateWindow
                        ? 'bg-indigo-950 text-indigo-300 border-indigo-500/50'
                        : 'bg-slate-900 hover:bg-slate-800 text-slate-300 border-slate-700'
                    }`}
                  >
                    <Terminal className="w-3.5 h-3.5 text-emerald-400" />
                    <span>Immediate (Ctrl+G)</span>
                  </button>

                  <button
                    type="button"
                    onClick={() => setShowObjectBrowser(true)}
                    title="View Object Browser (F2)"
                    className="flex items-center gap-1 px-2 py-1 bg-slate-900 hover:bg-slate-800 text-slate-300 rounded border border-slate-700 transition-colors"
                  >
                    <BookOpen className="w-3.5 h-3.5 text-amber-400" />
                    <span>Object Browser (F2)</span>
                  </button>
                </div>
              </div>

              {/* Compile Status Feedback Bar */}
              {compileStatus !== 'idle' && compileMessage && (
                <div
                  className={`p-2.5 rounded-lg border text-xs flex items-center justify-between gap-3 animate-in fade-in duration-150 ${
                    compileStatus === 'error'
                      ? 'bg-rose-950/60 border-rose-500/40 text-rose-300'
                      : 'bg-emerald-950/60 border-emerald-500/40 text-emerald-300'
                  }`}
                >
                  <div className="flex items-center gap-2">
                    {compileStatus === 'error' ? (
                      <AlertCircle className="w-4 h-4 text-rose-400 shrink-0" />
                    ) : (
                      <CheckCircle2 className="w-4 h-4 text-emerald-400 shrink-0" />
                    )}
                    <span className="font-mono">{compileMessage}</span>
                  </div>
                  <div className="flex items-center gap-2">
                    {compileStatus === 'error' && highlightedLine && (
                      <button
                        type="button"
                        onClick={() => handleGoToLine(highlightedLine)}
                        className="px-2 py-0.5 rounded bg-rose-900 hover:bg-rose-800 text-rose-200 text-[10px] font-mono font-bold transition-colors"
                      >
                        Go to Line {highlightedLine}
                      </button>
                    )}
                    <button
                      type="button"
                      onClick={() => setCompileStatus('idle')}
                      className="p-1 hover:text-white rounded"
                    >
                      <X className="w-3.5 h-3.5" />
                    </button>
                  </div>
                </div>
              )}

              {/* Editor Body */}
              {activeViewTab === 'editor' ? (
                <div className="rounded-lg border border-slate-800 bg-slate-950 overflow-hidden shadow-inner">
                  <div className="flex items-center justify-between px-3 py-1.5 bg-slate-900/90 border-b border-slate-800 text-[11px] font-mono text-slate-400">
                    <div className="flex items-center gap-2">
                      <FileCode className="w-3.5 h-3.5 text-indigo-400" />
                      <span>{activeMacro.moduleName} · {activeMacro.name}</span>
                    </div>
                    <div className="flex items-center gap-3">
                      {errorLinesSet.size > 0 && (
                        <span className="text-rose-400 font-semibold flex items-center gap-1">
                          <span className="w-2 h-2 rounded-full bg-rose-500 animate-pulse"></span>
                          {errorLinesSet.size} line error{errorLinesSet.size > 1 ? 's' : ''}
                        </span>
                      )}
                      <span>{codeLines.length} lines · UTF-8 Plaintext</span>
                    </div>
                  </div>

                  <div className="relative flex">
                    {/* Line Numbers & Error Gutter */}
                    <div
                      ref={lineGutterRef}
                      className="select-none py-3 px-1.5 w-12 bg-slate-900/70 border-r border-slate-800 text-slate-500 text-right space-y-0 overflow-hidden shrink-0 font-mono text-xs"
                      style={{ lineHeight: '1.625rem' }}
                    >
                      {codeLines.map((_, i) => {
                        const lineNum = i + 1;
                        const isError = errorLinesSet.has(lineNum);
                        const isHighlighted = highlightedLine === lineNum;
                        return (
                          <div
                            key={lineNum}
                            onClick={() => handleGoToLine(lineNum)}
                            className={`flex items-center justify-end gap-1 px-1 rounded cursor-pointer transition-colors ${
                              isError
                                ? 'bg-rose-950/90 text-rose-300 font-bold border-r-2 border-rose-500'
                                : isHighlighted
                                ? 'bg-indigo-950/90 text-indigo-300 font-bold'
                                : 'hover:text-slate-300'
                            }`}
                            title={isError ? `Compile error on line ${lineNum}` : `Line ${lineNum}`}
                          >
                            {isError && <span className="text-rose-400 text-[9px]">●</span>}
                            <span>{lineNum}</span>
                          </div>
                        );
                      })}
                    </div>

                    {/* Textarea Code Input */}
                    <textarea
                      ref={editorTextareaRef}
                      value={activeMacro.code || ''}
                      onChange={(e) => handleCodeChange(e.target.value)}
                      onScroll={handleEditorScroll}
                      spellCheck={false}
                      rows={16}
                      className="w-full bg-slate-950 py-3 px-3 font-mono text-xs text-slate-100 placeholder-slate-600 focus:outline-none resize-y leading-[1.625rem] selection:bg-indigo-600/50 whitespace-pre overflow-x-auto"
                    />
                  </div>
                </div>
              ) : (
                /* Diff / Compare View */
                <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                  <div className="space-y-1">
                    <div className="text-[11px] font-mono text-amber-400 font-semibold px-1">
                      Original Ingested VBA
                    </div>
                    <div className="p-3 bg-slate-950 rounded-lg border border-slate-800 font-mono text-xs text-slate-400 overflow-x-auto max-h-[380px] leading-relaxed">
                      <pre>{activeMacro.originalVba}</pre>
                    </div>
                  </div>
                  <div className="space-y-1">
                    <div className="text-[11px] font-mono text-emerald-400 font-semibold px-1">
                      Upgraded / Modified VBA
                    </div>
                    <div className="p-3 bg-slate-950 rounded-lg border border-emerald-500/20 font-mono text-xs text-emerald-200 overflow-x-auto max-h-[380px] leading-relaxed">
                      <pre>{activeMacro.code}</pre>
                    </div>
                  </div>
                </div>
              )}

              {/* VBE Immediate Window (Ctrl+G) */}
              {showImmediateWindow && (
                <div className="rounded-lg border border-slate-800 bg-slate-950 overflow-hidden animate-in fade-in duration-200">
                  <div className="flex items-center justify-between px-3 py-1.5 bg-slate-900 border-b border-slate-800 text-[11px] font-mono text-slate-300">
                    <div className="flex items-center gap-2">
                      <Terminal className="w-3.5 h-3.5 text-emerald-400" />
                      <span className="font-semibold">Immediate Window (Ctrl+G)</span>
                      <span className="text-[10px] text-slate-500">· Type '? expr' to print</span>
                    </div>
                    <div className="flex items-center gap-2">
                      <button
                        type="button"
                        onClick={() => setImmediateHistory([])}
                        className="text-[10px] text-slate-400 hover:text-slate-200 px-1.5 py-0.5 rounded bg-slate-800 hover:bg-slate-700 transition-colors"
                      >
                        Clear History
                      </button>
                      <button
                        type="button"
                        onClick={() => setShowImmediateWindow(false)}
                        className="text-slate-400 hover:text-white"
                      >
                        <X className="w-3.5 h-3.5" />
                      </button>
                    </div>
                  </div>

                  {/* Immediate Quick Test Pills */}
                  <div className="flex flex-wrap items-center gap-1.5 p-2 bg-slate-900/50 border-b border-slate-800/80 text-[11px] font-mono">
                    <span className="text-slate-500 text-[10px]">Quick Test:</span>
                    <button
                      type="button"
                      onClick={() => handleExecuteImmediate('? ActiveSheet.Name')}
                      className="px-2 py-0.5 rounded bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white transition-colors"
                    >
                      ? ActiveSheet.Name
                    </button>
                    <button
                      type="button"
                      onClick={() => handleExecuteImmediate('? Sheets.Count')}
                      className="px-2 py-0.5 rounded bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white transition-colors"
                    >
                      ? Sheets.Count
                    </button>
                    <button
                      type="button"
                      onClick={() => handleExecuteImmediate('? Application.ScreenUpdating')}
                      className="px-2 py-0.5 rounded bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white transition-colors"
                    >
                      ? ScreenUpdating
                    </button>
                    <button
                      type="button"
                      onClick={() => handleExecuteImmediate('? 150 * 1.15')}
                      className="px-2 py-0.5 rounded bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white transition-colors"
                    >
                      ? 150 * 1.15
                    </button>
                    <button
                      type="button"
                      onClick={() => handleExecuteImmediate('? Tru')}
                      className="px-2 py-0.5 rounded bg-rose-950/70 hover:bg-rose-900 border border-rose-500/30 text-rose-300 transition-colors"
                      title="Test typo detection in immediate window"
                    >
                      ? Tru (Test Typo)
                    </button>
                    <button
                      type="button"
                      onClick={() => handleExecuteImmediate('ActiveSheet.Calculate')}
                      className="px-2 py-0.5 rounded bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white transition-colors"
                    >
                      ActiveSheet.Calculate
                    </button>
                  </div>

                  {/* Immediate History Output */}
                  <div className="p-3 max-h-48 overflow-y-auto space-y-2 font-mono text-xs">
                    {immediateHistory.map((item) => (
                      <div key={item.id} className="space-y-0.5">
                        <div className="flex items-center justify-between text-slate-400 text-[11px]">
                          <span className="text-indigo-400 font-semibold">{item.command}</span>
                          <span className="text-[10px] text-slate-600">{item.timestamp}</span>
                        </div>
                        <div
                          className={`p-1.5 rounded text-xs select-text ${
                            item.isError
                              ? 'bg-rose-950/60 text-rose-300 border border-rose-500/30'
                              : 'bg-slate-900/80 text-emerald-300'
                          }`}
                        >
                          {item.output}
                        </div>
                      </div>
                    ))}
                    {immediateHistory.length === 0 && (
                      <div className="text-slate-600 text-xs text-center py-2">
                        Immediate Window empty. Type expression below to evaluate.
                      </div>
                    )}
                  </div>

                  {/* Prompt Input Form */}
                  <form
                    onSubmit={(e) => {
                      e.preventDefault();
                      handleExecuteImmediate();
                    }}
                    className="flex items-center gap-2 p-2 bg-slate-900 border-t border-slate-800"
                  >
                    <span className="text-indigo-400 font-mono font-bold pl-1">&gt;</span>
                    <input
                      type="text"
                      value={immediateInput}
                      onChange={(e) => setImmediateInput(e.target.value)}
                      placeholder="e.g. ? Range(&quot;A1&quot;).Value or ? Sheets.Count"
                      className="flex-1 bg-slate-950 border border-slate-700 rounded px-2.5 py-1 text-xs font-mono text-white placeholder-slate-500 focus:outline-none focus:border-indigo-500"
                    />
                    <button
                      type="submit"
                      disabled={!immediateInput.trim()}
                      className="px-3 py-1 bg-indigo-600 hover:bg-indigo-500 disabled:opacity-50 text-white text-xs font-semibold rounded font-mono transition-colors"
                    >
                      Eval
                    </button>
                  </form>
                </div>
              )}

              {/* Simulation & Diagnostics Suite */}
              <div className="pt-3 border-t border-slate-800 space-y-3">
                <div className="flex flex-wrap items-center justify-between gap-3">
                  <div className="flex items-center gap-2">
                    <button
                      type="button"
                      onClick={handleRunSimulation}
                      disabled={isSimulating}
                      className="flex items-center gap-1.5 px-3.5 py-1.5 bg-indigo-600 hover:bg-indigo-500 disabled:opacity-60 text-white text-xs font-semibold rounded-lg shadow-sm transition-colors"
                    >
                      {isSimulating ? (
                        <RefreshCw className="w-3.5 h-3.5 animate-spin" />
                      ) : (
                        <Play className="w-3.5 h-3.5 fill-current" />
                      )}
                      <span>{isSimulating ? 'Analyzing & Simulating...' : 'Simulate Macro Execution'}</span>
                    </button>

                    {simulationReport && (
                      <button
                        type="button"
                        onClick={() => setShowTrace(!showTrace)}
                        className={`flex items-center gap-1.5 px-2.5 py-1.5 text-xs rounded-lg border transition-colors ${
                          showTrace
                            ? 'bg-slate-800 text-white border-slate-700'
                            : 'bg-slate-950 text-slate-400 hover:text-slate-200 border-slate-800'
                        }`}
                      >
                        <Terminal className="w-3.5 h-3.5 text-indigo-400" />
                        <span>{showTrace ? 'Hide Trace' : 'Execution Trace'}</span>
                      </button>
                    )}
                  </div>

                  {simulationReport && (
                    <div className="flex items-center gap-2">
                      {simulationReport.success ? (
                        <div className="flex items-center gap-1.5 px-2.5 py-1 rounded-lg bg-emerald-950/70 border border-emerald-500/40 text-emerald-300 text-xs font-mono">
                          <CheckCircle2 className="w-3.5 h-3.5 text-emerald-400 shrink-0" />
                          <span>Passed ({simulationReport.executionTimeMs}ms · 0 Faults)</span>
                        </div>
                      ) : (
                        <div className="flex items-center gap-1.5 px-2.5 py-1 rounded-lg bg-rose-950/70 border border-rose-500/50 text-rose-300 text-xs font-mono font-semibold">
                          <AlertCircle className="w-3.5 h-3.5 text-rose-400 shrink-0" />
                          <span>
                            Failed ({simulationReport.errors.length} Error{simulationReport.errors.length > 1 ? 's' : ''})
                          </span>
                        </div>
                      )}
                      <button
                        type="button"
                        onClick={() => setSimulationReport(null)}
                        className="p-1 text-slate-500 hover:text-slate-300 rounded transition-colors"
                        title="Dismiss simulation report"
                      >
                        <X className="w-3.5 h-3.5" />
                      </button>
                    </div>
                  )}
                </div>

                {/* Detailed Simulation Report Box */}
                {simulationReport && (
                  <div className="space-y-3 animate-in fade-in duration-200">
                    {/* 1. Errors Display */}
                    {simulationReport.errors.length > 0 && (
                      <div className="rounded-xl border border-rose-500/40 bg-rose-950/20 p-4 space-y-3">
                        <div className="flex items-center justify-between pb-2 border-b border-rose-500/20">
                          <div className="flex items-center gap-2 text-rose-300 text-xs font-semibold">
                            <AlertCircle className="w-4 h-4 text-rose-400" />
                            <span>VBA Syntax & Execution Faults ({simulationReport.errors.length})</span>
                          </div>
                          <span className="text-[10px] font-mono text-rose-400/80 bg-rose-950/80 border border-rose-500/30 px-2 py-0.5 rounded">
                            Simulation Blocked
                          </span>
                        </div>

                        <div className="space-y-2.5 max-h-64 overflow-y-auto pr-1">
                          {simulationReport.errors.map((err, idx) => (
                            <div
                              key={idx}
                              className="bg-slate-950/90 rounded-lg border border-rose-500/30 p-3 space-y-1.5 text-xs font-mono"
                            >
                              <div className="flex items-center justify-between gap-2">
                                <div className="flex items-center gap-2">
                                  <button
                                    type="button"
                                    onClick={() => handleGoToLine(err.line)}
                                    className="px-2 py-0.5 rounded bg-rose-950 hover:bg-rose-900 text-rose-300 border border-rose-500/40 text-[10px] font-bold cursor-pointer transition-colors flex items-center gap-1"
                                    title={`Scroll to Line ${err.line}`}
                                  >
                                    <span>Line {err.line}</span>
                                    <span className="text-[9px] text-rose-400">↗</span>
                                  </button>
                                  <span className="text-rose-200 font-semibold">{err.code}</span>
                                </div>
                                <span className="text-[10px] text-slate-500">Syntax Error</span>
                              </div>

                              <p className="text-xs text-slate-300 font-sans font-normal leading-relaxed">
                                {err.message}
                              </p>

                              {err.lineSnippet && (
                                <div className="p-2 rounded bg-slate-900 border border-slate-800 text-rose-300 text-[11px] truncate">
                                  <code>&gt; {err.lineSnippet}</code>
                                </div>
                              )}

                              {err.suggestion && (
                                <div className="text-[11px] text-cyan-400/90 font-sans flex items-start gap-1.5 pt-0.5">
                                  <span className="font-semibold text-cyan-300 shrink-0">💡 Suggested Fix:</span>
                                  <span>{err.suggestion}</span>
                                </div>
                              )}
                            </div>
                          ))}
                        </div>
                      </div>
                    )}

                    {/* 2. Warnings Display */}
                    {simulationReport.warnings.length > 0 && (
                      <div className="rounded-xl border border-amber-500/40 bg-amber-950/20 p-3.5 space-y-2.5">
                        <div className="flex items-center justify-between pb-1.5 border-b border-amber-500/20">
                          <div className="flex items-center gap-2 text-amber-300 text-xs font-semibold">
                            <AlertTriangle className="w-4 h-4 text-amber-400" />
                            <span>Warnings & Reference Checks ({simulationReport.warnings.length})</span>
                          </div>
                        </div>

                        <div className="space-y-2 max-h-48 overflow-y-auto pr-1">
                          {simulationReport.warnings.map((warn, idx) => (
                            <div
                              key={idx}
                              className="bg-slate-950/80 rounded-lg border border-amber-500/30 p-2.5 space-y-1 text-xs"
                            >
                              <div className="flex items-center gap-2 font-mono">
                                <button
                                  type="button"
                                  onClick={() => handleGoToLine(warn.line)}
                                  className="px-2 py-0.5 rounded bg-amber-950 hover:bg-amber-900 text-amber-300 border border-amber-500/40 text-[10px] font-bold cursor-pointer transition-colors flex items-center gap-1"
                                  title={`Scroll to Line ${warn.line}`}
                                >
                                  <span>Line {warn.line}</span>
                                  <span className="text-[9px] text-amber-400">↗</span>
                                </button>
                                <span className="text-amber-200 font-medium">{warn.code}</span>
                              </div>
                              <p className="text-slate-300">{warn.message}</p>
                              {warn.suggestion && (
                                <p className="text-[11px] text-cyan-300/90 font-sans">
                                  💡 {warn.suggestion}
                                </p>
                              )}
                            </div>
                          ))}
                        </div>
                      </div>
                    )}

                    {/* 3. Success Environment & Context Metrics */}
                    {simulationReport.success && (
                      <div className="rounded-xl border border-emerald-500/40 bg-emerald-950/20 p-4 space-y-3">
                        <div className="flex items-center justify-between pb-2 border-b border-emerald-500/20">
                          <div className="flex items-center gap-2 text-emerald-300 text-xs font-semibold">
                            <CheckCircle2 className="w-4 h-4 text-emerald-400" />
                            <span>Simulation Execution Verified</span>
                          </div>
                          <span className="text-[11px] font-mono text-emerald-400">
                            {simulationReport.executionTimeMs} ms execution latency
                          </span>
                        </div>

                        <div className="grid grid-cols-2 sm:grid-cols-4 gap-2.5 text-xs font-mono">
                          <div className="p-2 rounded bg-slate-950 border border-slate-800">
                            <span className="text-slate-500 text-[10px] block">Analyzed Code</span>
                            <span className="text-slate-200 font-semibold">{simulationReport.totalLines} lines</span>
                          </div>
                          <div className="p-2 rounded bg-slate-950 border border-slate-800">
                            <span className="text-slate-500 text-[10px] block">Statements</span>
                            <span className="text-slate-200 font-semibold">{simulationReport.statementsAnalyzed} statements</span>
                          </div>
                          <div className="p-2 rounded bg-slate-950 border border-slate-800">
                            <span className="text-slate-500 text-[10px] block">Declared Vars</span>
                            <span className="text-slate-200 font-semibold">{simulationReport.variablesDetected.length} vars</span>
                          </div>
                          <div className="p-2 rounded bg-slate-950 border border-slate-800">
                            <span className="text-slate-500 text-[10px] block">Active Sheet</span>
                            <span className="text-emerald-300 font-semibold truncate block">
                              {dataset.activeSheetName || dataset.sheets[0]?.name || 'Sheet1'}
                            </span>
                          </div>
                        </div>

                        {/* Referenced Sheets */}
                        {simulationReport.referencedSheets.length > 0 && (
                          <div className="pt-1 flex flex-wrap items-center gap-2 text-xs font-mono">
                            <span className="text-slate-400 text-[11px]">Worksheet Targets:</span>
                            {simulationReport.referencedSheets.map((sh, idx) => (
                              <span
                                key={idx}
                                className={`px-2 py-0.5 rounded border text-[10px] flex items-center gap-1 ${
                                  sh.existsInWorkbook
                                    ? 'bg-emerald-950/60 text-emerald-300 border-emerald-500/30'
                                    : 'bg-rose-950/60 text-rose-300 border-rose-500/30'
                                }`}
                              >
                                <span>{sh.name}</span>
                                <span>({sh.existsInWorkbook ? 'Found' : 'Missing'})</span>
                              </span>
                            ))}
                          </div>
                        )}
                      </div>
                    )}

                    {/* 4. Execution Trace Console (collapsible) */}
                    {showTrace && (
                      <div className="rounded-xl border border-slate-800 bg-slate-950 p-3.5 space-y-2">
                        <div className="flex items-center justify-between text-xs font-mono text-slate-400 pb-1.5 border-b border-slate-800">
                          <div className="flex items-center gap-1.5">
                            <Terminal className="w-3.5 h-3.5 text-indigo-400" />
                            <span>Simulation Trace & Runtime Mock</span>
                          </div>
                          <span className="text-[10px]">{simulationReport.executionTrace.length} log events</span>
                        </div>
                        <div className="p-2.5 bg-slate-900 rounded font-mono text-[11px] text-slate-300 max-h-48 overflow-y-auto space-y-1">
                          {simulationReport.executionTrace.map((log, idx) => (
                            <div key={idx} className="leading-relaxed">
                              <span className="text-slate-600 select-none mr-2">
                                {String(idx + 1).padStart(2, '0')}:
                              </span>
                              <span
                                className={
                                  log.includes('Failed') || log.includes('error') || log.includes('aborted')
                                    ? 'text-rose-400'
                                    : log.includes('Warning')
                                    ? 'text-amber-400'
                                    : log.includes('Success') || log.includes('Verified') || log.includes('passed')
                                    ? 'text-emerald-400'
                                    : 'text-slate-300'
                                }
                              >
                                {log}
                              </span>
                            </div>
                          ))}
                        </div>
                      </div>
                    )}
                  </div>
                )}
              </div>
            </>
          )}
        </div>
      </div>

      {/* Add New Procedure Modal Dialog */}
      {isAddModalOpen && (
        <div
          className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-950/80 backdrop-blur-sm animate-in fade-in"
          onClick={(e) => {
            if (e.target === e.currentTarget) setIsAddModalOpen(false);
          }}
        >
          <div className="bg-slate-900 border border-slate-700/80 rounded-2xl shadow-2xl max-w-2xl w-full max-h-[92vh] overflow-y-auto flex flex-col">
            {/* Modal Header */}
            <div className="p-5 border-b border-slate-800 flex items-center justify-between">
              <div className="flex items-center gap-3">
                <div className="p-2 rounded-lg bg-indigo-950/80 border border-indigo-500/40 text-indigo-400">
                  <Code2 className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="text-base font-semibold text-white">Add New VBA Procedure</h3>
                  <p className="text-xs text-slate-400 mt-0.5">
                    Generate and insert a new subroutine or function into your workbook module.
                  </p>
                </div>
              </div>
              <button
                onClick={() => setIsAddModalOpen(false)}
                className="p-1.5 rounded-lg text-slate-400 hover:text-white hover:bg-slate-800 transition-colors"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* Modal Form Body */}
            <form onSubmit={handleCreateProcedure} className="p-6 space-y-4">
              {/* Row 1: Procedure Name & Scope */}
              <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
                <div className="sm:col-span-2 space-y-1.5">
                  <label className="text-xs font-medium text-slate-300">
                    Procedure Name <span className="text-indigo-400">*</span>
                  </label>
                  <input
                    type="text"
                    value={newProcName}
                    onChange={(e) => setNewProcName(e.target.value)}
                    placeholder="e.g. ExportMonthlyAuditReport"
                    className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-xs font-mono text-white placeholder-slate-500 focus:outline-none focus:border-indigo-500"
                    autoFocus
                    required
                  />
                  <p className="text-[11px] text-slate-500">
                    Alphanumeric and underscores only (spaces converted automatically).
                  </p>
                </div>

                <div className="space-y-1.5">
                  <label className="text-xs font-medium text-slate-300">Scope</label>
                  <select
                    value={newProcScope}
                    onChange={(e) => setNewProcScope(e.target.value as 'Public' | 'Private')}
                    className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-xs font-mono text-slate-200 focus:outline-none focus:border-indigo-500"
                  >
                    <option value="Public">Public (Accessible globally)</option>
                    <option value="Private">Private (Module-only)</option>
                  </select>
                </div>
              </div>

              {/* Row 2: Procedure Type & Boilerplate Template */}
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                <div className="space-y-1.5">
                  <label className="text-xs font-medium text-slate-300">Procedure Type</label>
                  <div className="grid grid-cols-2 gap-2">
                    <button
                      type="button"
                      onClick={() => setNewProcType('Sub')}
                      className={`py-2 px-3 rounded-lg border text-xs font-mono font-medium transition-colors ${
                        newProcType === 'Sub'
                          ? 'border-indigo-500/80 bg-indigo-950/70 text-indigo-300'
                          : 'border-slate-800 bg-slate-950 text-slate-400 hover:text-slate-200'
                      }`}
                    >
                      Sub (Subroutine)
                    </button>
                    <button
                      type="button"
                      onClick={() => setNewProcType('Function')}
                      className={`py-2 px-3 rounded-lg border text-xs font-mono font-medium transition-colors ${
                        newProcType === 'Function'
                          ? 'border-indigo-500/80 bg-indigo-950/70 text-indigo-300'
                          : 'border-slate-800 bg-slate-950 text-slate-400 hover:text-slate-200'
                      }`}
                    >
                      Function (Returns Value)
                    </button>
                  </div>
                </div>

                <div className="space-y-1.5">
                  <label className="text-xs font-medium text-slate-300">Starter Template</label>
                  <select
                    value={selectedTemplate}
                    onChange={(e) => setSelectedTemplate(e.target.value as any)}
                    disabled={newProcType === 'Function'}
                    className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-xs text-slate-200 focus:outline-none focus:border-indigo-500 disabled:opacity-50"
                  >
                    <option value="basic">Standard Skeleton (Clean Worksheet Call)</option>
                    <option value="loop">Row Iteration Loop (UsedRange For-Next)</option>
                    <option value="fast_calc">Speed-Boosted Batch (Manual Calc & Events)</option>
                  </select>
                </div>
              </div>

              {/* Row 3: Target Module */}
              <div className="space-y-1.5">
                <label className="text-xs font-medium text-slate-300">Target Module</label>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                  <select
                    value={selectedModuleOption}
                    onChange={(e) => setSelectedModuleOption(e.target.value)}
                    className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-xs font-mono text-slate-200 focus:outline-none focus:border-indigo-500"
                  >
                    {existingModules.map((mod) => (
                      <option key={mod} value={mod}>
                        {mod}
                      </option>
                    ))}
                    <option value="__new__">+ Create New Module...</option>
                  </select>

                  {selectedModuleOption === '__new__' && (
                    <input
                      type="text"
                      value={customModuleName}
                      onChange={(e) => setCustomModuleName(e.target.value)}
                      placeholder="e.g. Module3_Reporting"
                      className="w-full bg-slate-950 border border-indigo-500/50 rounded-lg px-3 py-2 text-xs font-mono text-white placeholder-slate-500 focus:outline-none"
                      required
                    />
                  )}
                </div>
              </div>

              {/* Description */}
              <div className="space-y-1.5">
                <label className="text-xs font-medium text-slate-300">Description</label>
                <input
                  type="text"
                  value={newProcDescription}
                  onChange={(e) => setNewProcDescription(e.target.value)}
                  placeholder="Summary of what this routine calculates or automates..."
                  className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-xs text-slate-200 placeholder-slate-500 focus:outline-none focus:border-indigo-500"
                />
              </div>

              {/* Live Code Preview */}
              <div className="space-y-1.5 pt-1">
                <div className="flex items-center justify-between text-[11px] font-mono text-slate-400">
                  <span>Generated VBA Code Preview</span>
                  <span>{previewGeneratedCode.split('\n').length} lines</span>
                </div>
                <div className="p-3 bg-slate-950 rounded-lg border border-slate-800 font-mono text-xs text-slate-300 max-h-48 overflow-y-auto leading-relaxed select-text">
                  <pre>{previewGeneratedCode}</pre>
                </div>
              </div>

              {/* Form Action Buttons */}
              <div className="pt-3 border-t border-slate-800 flex items-center justify-end gap-3">
                <button
                  type="button"
                  onClick={() => setIsAddModalOpen(false)}
                  className="px-4 py-2 text-xs font-medium text-slate-400 hover:text-white bg-slate-800 hover:bg-slate-700 rounded-lg transition-colors"
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  className="flex items-center gap-1.5 px-4 py-2 text-xs font-semibold text-white bg-indigo-600 hover:bg-indigo-500 rounded-lg shadow-sm transition-colors"
                >
                  <Plus className="w-3.5 h-3.5" />
                  <span>Create Procedure</span>
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* Delete Confirmation Modal Dialog */}
      {deleteTarget && (
        <div
          className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-950/80 backdrop-blur-sm animate-in fade-in"
          onClick={(e) => {
            if (e.target === e.currentTarget) setDeleteTarget(null);
          }}
        >
          <div className="bg-slate-900 border border-slate-700/80 rounded-2xl shadow-2xl max-w-md w-full overflow-hidden flex flex-col">
            {/* Modal Header */}
            <div className="p-5 border-b border-slate-800 flex items-center justify-between">
              <div className="flex items-center gap-3">
                <div className="p-2.5 rounded-xl bg-rose-950/80 border border-rose-500/40 text-rose-400">
                  <Trash2 className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="text-base font-semibold text-white">
                    {deleteTarget.type === 'procedure' ? 'Delete Procedure' : 'Delete Module'}
                  </h3>
                  <p className="text-xs text-slate-400 mt-0.5">
                    Confirm permanent deletion
                  </p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setDeleteTarget(null)}
                className="p-1.5 rounded-lg text-slate-400 hover:text-white hover:bg-slate-800 transition-colors"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* Modal Content */}
            <div className="p-5 space-y-3">
              {deleteTarget.type === 'procedure' ? (
                <div>
                  <p className="text-xs text-slate-300 leading-relaxed">
                    Are you sure you want to delete procedure{' '}
                    <span className="font-mono font-semibold text-white bg-slate-800 px-1.5 py-0.5 rounded border border-slate-700">
                      {deleteTarget.name}
                    </span>
                    {deleteTarget.moduleName && (
                      <>
                        {' '}from module{' '}
                        <span className="font-mono text-indigo-300 font-medium">
                          {deleteTarget.moduleName}
                        </span>
                      </>
                    )}
                    ?
                  </p>
                  <p className="text-[11px] text-rose-300/90 mt-3 bg-rose-950/30 border border-rose-500/20 rounded-lg p-2.5">
                    This will permanently remove the subroutine or function code from your workbook macros.
                  </p>
                </div>
              ) : (
                <div>
                  <p className="text-xs text-slate-300 leading-relaxed">
                    Are you sure you want to delete the module{' '}
                    <span className="font-mono font-semibold text-white bg-slate-800 px-1.5 py-0.5 rounded border border-slate-700">
                      {deleteTarget.name}
                    </span>
                    ?
                  </p>
                  <div className="text-[11px] text-rose-300 mt-3 bg-rose-950/40 border border-rose-500/30 rounded-lg p-3 space-y-1">
                    <div className="font-semibold flex items-center gap-1.5 text-rose-200">
                      <AlertTriangle className="w-3.5 h-3.5 text-rose-400 shrink-0" />
                      <span>Permanent Module Deletion</span>
                    </div>
                    <p className="text-rose-300/90 leading-relaxed">
                      This will delete the entire module file and all{' '}
                      <strong className="text-white font-mono">{deleteTarget.count || 0}</strong>{' '}
                      associated routine(s) inside it.
                    </p>
                  </div>
                </div>
              )}
            </div>

            {/* Modal Action Buttons */}
            <div className="p-4 bg-slate-950 border-t border-slate-800 flex items-center justify-end gap-3">
              <button
                type="button"
                onClick={() => setDeleteTarget(null)}
                className="px-4 py-2 text-xs font-medium text-slate-400 hover:text-white bg-slate-800 hover:bg-slate-700 rounded-lg transition-colors"
              >
                Cancel
              </button>
              <button
                type="button"
                onClick={handleConfirmDelete}
                className="flex items-center gap-1.5 px-4 py-2 text-xs font-semibold text-white bg-rose-600 hover:bg-rose-500 rounded-lg shadow-sm transition-colors"
              >
                <Trash2 className="w-3.5 h-3.5" />
                <span>
                  {deleteTarget.type === 'procedure' ? 'Delete Procedure' : 'Delete Module'}
                </span>
              </button>
            </div>
          </div>
        </div>
      )}

      {/* VBE Object Browser Modal (F2) */}
      {showObjectBrowser && (
        <div
          className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-950/80 backdrop-blur-sm animate-in fade-in"
          onClick={(e) => {
            if (e.target === e.currentTarget) setShowObjectBrowser(false);
          }}
        >
          <div className="bg-slate-900 border border-slate-700/80 rounded-2xl shadow-2xl max-w-4xl w-full max-h-[90vh] flex flex-col overflow-hidden">
            {/* Modal Header */}
            <div className="p-4 bg-slate-950 border-b border-slate-800 flex items-center justify-between">
              <div className="flex items-center gap-3">
                <div className="p-2 rounded-lg bg-amber-950/80 border border-amber-500/40 text-amber-400">
                  <BookOpen className="w-5 h-5" />
                </div>
                <div>
                  <div className="flex items-center gap-2">
                    <h3 className="text-base font-semibold text-white">Object Browser (F2)</h3>
                    <span className="text-[10px] font-mono px-2 py-0.5 rounded bg-slate-800 text-slate-300 border border-slate-700">
                      MS Office VBE Reference
                    </span>
                  </div>
                  <p className="text-xs text-slate-400 mt-0.5">
                    Browse Excel Object Model &amp; VBA runtime classes, methods, properties, and constants.
                  </p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setShowObjectBrowser(false)}
                className="p-1.5 rounded-lg text-slate-400 hover:text-white hover:bg-slate-800 transition-colors"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* Filter Toolbar */}
            <div className="p-3 bg-slate-900 border-b border-slate-800 flex flex-wrap items-center justify-between gap-3 text-xs">
              <div className="flex items-center gap-2">
                <span className="text-[11px] font-mono text-slate-400">Library:</span>
                <select
                  value={selectedObjLibrary}
                  onChange={(e) => setSelectedObjLibrary(e.target.value as any)}
                  className="bg-slate-950 border border-slate-700 rounded-lg px-2.5 py-1 text-slate-200 font-mono text-xs focus:outline-none focus:border-indigo-500"
                >
                  <option value="ALL">&lt;All Libraries&gt;</option>
                  <option value="Excel">Excel (Microsoft Excel 16.0 Object Library)</option>
                  <option value="VBA">VBA (Visual Basic for Applications)</option>
                </select>
              </div>

              <div className="relative flex-1 max-w-xs">
                <Search className="w-3.5 h-3.5 text-slate-500 absolute left-2.5 top-2.5" />
                <input
                  type="text"
                  value={objSearchQuery}
                  onChange={(e) => setObjSearchQuery(e.target.value)}
                  placeholder="Search classes or members..."
                  className="w-full bg-slate-950 border border-slate-700 rounded-lg pl-8 pr-3 py-1 text-xs font-mono text-white placeholder-slate-500 focus:outline-none focus:border-indigo-500"
                />
                {objSearchQuery && (
                  <button
                    type="button"
                    onClick={() => setObjSearchQuery('')}
                    className="absolute right-2 top-2 text-slate-400 hover:text-white"
                  >
                    <X className="w-3 h-3" />
                  </button>
                )}
              </div>
            </div>

            {/* Split View: Classes vs Members */}
            <div className="grid grid-cols-1 md:grid-cols-12 flex-1 min-h-[300px] max-h-[380px] overflow-hidden">
              {/* Classes Column (5 cols) */}
              <div className="md:col-span-5 border-r border-slate-800 bg-slate-950/50 p-2 overflow-y-auto space-y-1">
                <div className="px-2 py-1 text-[10px] font-mono uppercase text-slate-500 font-semibold tracking-wider">
                  Classes ({filteredObjectBrowserClasses.length})
                </div>
                {filteredObjectBrowserClasses.map((cls) => {
                  const isSelected = currentObjClass?.name === cls.name;
                  return (
                    <div
                      key={cls.name}
                      onClick={() => {
                        setSelectedObjClass(cls.name);
                        setSelectedObjMember(cls.members[0]?.name || '');
                      }}
                      className={`px-2.5 py-1.5 rounded-lg text-xs font-mono cursor-pointer flex items-center justify-between transition-colors ${
                        isSelected
                          ? 'bg-indigo-950 text-indigo-200 border border-indigo-500/40 font-semibold'
                          : 'text-slate-300 hover:bg-slate-800/60'
                      }`}
                    >
                      <div className="flex items-center gap-2 truncate">
                        <Layers className={`w-3.5 h-3.5 shrink-0 ${isSelected ? 'text-indigo-400' : 'text-slate-500'}`} />
                        <span className="truncate">{cls.name}</span>
                      </div>
                      <span className="text-[10px] text-slate-500 shrink-0 font-normal">
                        {cls.members.length}
                      </span>
                    </div>
                  );
                })}
              </div>

              {/* Members Column (7 cols) */}
              <div className="md:col-span-7 bg-slate-950/20 p-2 overflow-y-auto space-y-1">
                <div className="px-2 py-1 text-[10px] font-mono uppercase text-slate-500 font-semibold tracking-wider">
                  Members of '{currentObjClass?.name || 'Class'}' ({currentObjClass?.members.length || 0})
                </div>
                {currentObjClass?.members.map((m) => {
                  const isSelected = currentObjMember?.name === m.name;
                  return (
                    <div
                      key={m.name}
                      onClick={() => setSelectedObjMember(m.name)}
                      className={`px-2.5 py-1.5 rounded-lg text-xs font-mono cursor-pointer flex items-center justify-between transition-colors ${
                        isSelected
                          ? 'bg-indigo-950/90 text-indigo-200 border border-indigo-500/40 font-semibold'
                          : 'text-slate-300 hover:bg-slate-800/60'
                      }`}
                    >
                      <div className="flex items-center gap-2 truncate">
                        <span
                          className={`text-[9px] px-1 py-0.2 rounded font-bold uppercase shrink-0 ${
                            m.type === 'Property'
                              ? 'bg-sky-950 text-sky-300 border border-sky-500/30'
                              : m.type === 'Method'
                              ? 'bg-emerald-950 text-emerald-300 border border-emerald-500/30'
                              : m.type === 'Function'
                              ? 'bg-indigo-950 text-indigo-300 border border-indigo-500/30'
                              : 'bg-amber-950 text-amber-300 border border-amber-500/30'
                          }`}
                        >
                          {m.type}
                        </span>
                        <span className="truncate">{m.name}</span>
                      </div>
                      <span className="text-[10px] text-slate-500 truncate max-w-[120px] font-sans">
                        {m.description.slice(0, 30)}...
                      </span>
                    </div>
                  );
                })}
              </div>
            </div>

            {/* Bottom Detail Pane */}
            {currentObjMember && (
              <div className="p-4 bg-slate-950 border-t border-slate-800 space-y-2">
                <div className="flex flex-wrap items-center justify-between gap-2">
                  <div className="flex items-center gap-2">
                    <span className="text-xs font-bold font-mono text-white">
                      {currentObjClass?.name}.{currentObjMember.name}
                    </span>
                    <span className="text-[10px] font-mono px-2 py-0.5 rounded bg-slate-800 text-indigo-300">
                      {currentObjMember.signature}
                    </span>
                  </div>

                  <div className="flex items-center gap-2">
                    <button
                      type="button"
                      onClick={() => {
                        navigator.clipboard.writeText(currentObjMember.snippet);
                        showNotification(`Copied '${currentObjMember.snippet}' to clipboard.`);
                      }}
                      className="flex items-center gap-1 px-2.5 py-1 bg-slate-800 hover:bg-slate-700 text-slate-200 text-xs rounded transition-colors"
                    >
                      <Copy className="w-3 h-3" />
                      <span>Copy</span>
                    </button>
                    {activeMacro && (
                      <button
                        type="button"
                        onClick={() => handleInsertSnippet(currentObjMember.snippet)}
                        className="flex items-center gap-1 px-3 py-1 bg-indigo-600 hover:bg-indigo-500 text-white text-xs font-semibold rounded shadow-sm transition-colors"
                      >
                        <Plus className="w-3 h-3" />
                        <span>Insert into Procedure</span>
                      </button>
                    )}
                  </div>
                </div>

                <p className="text-xs text-slate-300 leading-relaxed font-sans">
                  {currentObjMember.description}
                </p>

                <div className="p-2 bg-slate-900 rounded border border-slate-800 text-xs font-mono text-emerald-300">
                  <code>{currentObjMember.snippet}</code>
                </div>
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  );
};
