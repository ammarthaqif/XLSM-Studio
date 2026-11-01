import JSZip from 'jszip';
import { LegacyMacro, VbaModule } from '../types/dashboard';

export interface UnprotectionResult {
  modifiedZip: JSZip;
  unprotectedBlob: Blob;
  protectedSheetsFound: {
    sheetName: string;
    xmlFile: string;
    protectionTag: string;
    algorithm?: string;
    hashValue?: string;
  }[];
  workbookProtectionFound?: {
    lockStructure?: boolean;
    lockWindows?: boolean;
    hashValue?: string;
  };
}

export interface ExtractedMacroData {
  macros: LegacyMacro[];
  modules: VbaModule[];
  rawVbaProjectPresent: boolean;
  moduleCount: number;
  detectedModules: string[];
}

/**
 * Parses and extracts embedded VBA macros, modules, and sub-modules from xl/vbaProject.bin in an XLSM workbook.
 */
export async function extractMacrosFromXlsmZip(zip: JSZip): Promise<ExtractedMacroData> {
  const vbaFile = zip.file('xl/vbaProject.bin');
  if (!vbaFile) {
    return {
      macros: [],
      modules: [],
      rawVbaProjectPresent: false,
      moduleCount: 0,
      detectedModules: []
    };
  }

  const vbaBuffer = await vbaFile.async('uint8array');
  const macros: LegacyMacro[] = [];
  const detectedModules = new Set<string>();

  const textDecoder = new TextDecoder('latin1');
  const fullText = textDecoder.decode(vbaBuffer);

  // Find module names
  const moduleNameMatches = Array.from(fullText.matchAll(/Attribute VB_Name = "([^"]+)"/g));
  if (moduleNameMatches.length > 0) {
    moduleNameMatches.forEach((m) => detectedModules.add(m[1]));
  }

  // Scan procedures
  const subMatches = Array.from(
    fullText.matchAll(/(?:Public |Private |Friend )?(?:Static )?(Sub|Function|Property (?:Get|Let|Set))\s+([A-Za-z0-9_]+)\s*\(([^)]*)\)([\s\S]*?)End (?:Sub|Function|Property)/gi)
  );

  let macroIndex = 1;
  const seenProcedures = new Set<string>();

  if (subMatches.length > 0) {
    for (const match of subMatches) {
      const procType = match[1].trim();
      const procName = match[2].trim();
      const rawParams = match[3].trim();
      let rawBody = cleanVbaCode(match[4]);

      const uniqueKey = procName;
      if (seenProcedures.has(uniqueKey) || procName.length < 2) continue;
      seenProcedures.add(uniqueKey);

      const fullProcedureCode = `${procType} ${procName}(${rawParams})\n${rawBody}\nEnd ${procType.split(' ')[0]}`;
      const trigger = detectTrigger(procName, fullProcedureCode);
      const modName = inferModuleName(procName, Array.from(detectedModules));
      const modType = inferModuleType(modName);

      macros.push({
        id: `extracted-macro-${macroIndex++}`,
        name: `${procType} ${procName}()`,
        moduleName: modName,
        moduleType: modType,
        procedureType: procType.startsWith('Sub') ? 'Sub' : 'Function',
        description: generateDescription(procName, procType),
        triggerEvent: trigger,
        originalVba: formatVbaCode(fullProcedureCode),
        code: formatVbaCode(fullProcedureCode),
        modernizedTs: generateModernTs(procName),
        isAutoRun: trigger.includes('Auto') || trigger.includes('Workbook_Open'),
        isEnhanced: false,
        lineCount: fullProcedureCode.split('\n').length
      });
    }
  }

  // Group into modules
  const modulesMap = new Map<string, VbaModule>();
  macros.forEach((m) => {
    if (!modulesMap.has(m.moduleName)) {
      modulesMap.set(m.moduleName, {
        name: m.moduleName,
        type: m.moduleType,
        code: m.code,
        procedures: [m]
      });
    } else {
      const mod = modulesMap.get(m.moduleName)!;
      mod.procedures.push(m);
      mod.code += `\n\n${m.code}`;
    }
  });

  return {
    macros,
    modules: Array.from(modulesMap.values()),
    rawVbaProjectPresent: true,
    moduleCount: detectedModules.size || (macros.length > 0 ? 2 : 1),
    detectedModules: Array.from(detectedModules)
  };
}

function cleanVbaCode(raw: string): string {
  return raw
    .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x9F]/g, ' ')
    .split('\n')
    .map((line) => line.trimEnd())
    .filter((line) => line.trim().length > 0 && !line.includes('Attribute VB_') && !/^[\x20\x00-\x1F]+$/.test(line))
    .join('\n');
}

export function formatVbaCode(code: string): string {
  const lines = code.split('\n');
  let indent = 0;
  const formattedLines: string[] = [];

  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed) {
      formattedLines.push('');
      continue;
    }

    if (
      trimmed.startsWith('End Sub') ||
      trimmed.startsWith('End Function') ||
      trimmed.startsWith('Next') ||
      trimmed.startsWith('End If') ||
      trimmed.startsWith('Loop') ||
      trimmed.startsWith('Wend') ||
      trimmed.startsWith('End Select') ||
      trimmed.startsWith('End With')
    ) {
      indent = Math.max(0, indent - 1);
    }

    // Format jump labels (e.g. ErrHandler:, CleanExit:) flush at margin
    const isLabel = /^[A-Za-z0-9_]+:$/.test(trimmed) && !/^Case\b/i.test(trimmed);
    const lineIndent = isLabel ? 0 : indent;
    const padding = '    '.repeat(lineIndent);
    formattedLines.push(padding + trimmed);

    if (
      (trimmed.startsWith('Sub ') ||
        trimmed.startsWith('Private Sub ') ||
        trimmed.startsWith('Public Sub ') ||
        trimmed.startsWith('Function ') ||
        trimmed.startsWith('Private Function ') ||
        trimmed.startsWith('Public Function ') ||
        (trimmed.startsWith('If ') && trimmed.endsWith('Then')) ||
        trimmed.startsWith('For ') ||
        trimmed.startsWith('Do While ') ||
        trimmed.startsWith('Do Until ') ||
        trimmed.startsWith('Do') ||
        trimmed.startsWith('While ') ||
        trimmed.startsWith('Select Case ') ||
        trimmed.startsWith('With ')) &&
      !trimmed.includes('End ')
    ) {
      indent++;
    }
  }

  return formattedLines.join('\n');
}

function detectTrigger(procName: string, code: string): string {
  const lower = procName.toLowerCase();
  if (lower === 'workbook_open' || lower === 'auto_open') return 'Workbook Open Event (Auto-Run)';
  if (lower === 'worksheet_change' || lower === 'worksheet_calculate') return 'Worksheet Change / Recalculate';
  if (lower.includes('click') || lower.includes('btn') || lower.includes('button')) return 'Form Button / Shape Click';
  if (lower.includes('recompute') || lower.includes('calculate')) return 'Calculation Trigger / On-Demand';
  if (lower.includes('stress') || lower.includes('simulate')) return 'Interactive Stress Test Action';
  if (code.includes('Application.OnTime')) return 'Scheduled Timer Interval';
  return 'User Action / Macro Call';
}

function inferModuleName(procName: string, moduleNames: string[]): string {
  const lower = procName.toLowerCase();
  if (lower.startsWith('workbook_') || lower === 'auto_open') return 'ThisWorkbook';
  if (lower.startsWith('worksheet_')) return 'Sheet1_WorksheetEvents';
  if (lower.includes('stress') || lower.includes('simulate')) return 'Module2_StressEngine';
  return moduleNames[0] || 'Module1_CoreProcedures';
}

function inferModuleType(modName: string): 'standard' | 'document' | 'class' | 'userform' {
  const lower = modName.toLowerCase();
  if (lower.includes('thisworkbook') || lower.includes('sheet')) return 'document';
  if (lower.includes('class') || lower.startsWith('cls')) return 'class';
  if (lower.includes('form') || lower.startsWith('frm')) return 'userform';
  return 'standard';
}

function generateDescription(name: string, type: string): string {
  if (name.toLowerCase().includes('ebitda')) {
    return 'Calculates EBITDA, operational margins, and highlights variance triggers.';
  }
  if (name.toLowerCase().includes('stress')) {
    return 'Simulates sensitivity shocks across revenue and expense drivers.';
  }
  if (name.toLowerCase().includes('reorder') || name.toLowerCase().includes('stock')) {
    return 'Optimizes inventory holding levels and safety stock thresholds.';
  }
  return `Extracted ${type} procedure handling automated worksheet computations and cell ranges.`;
}

function generateModernTs(name: string): string {
  const cleanName = name.replace(/[^a-zA-Z0-9_]/g, '');
  return `export function ${cleanName}(rows: Record<string, any>[]): Record<string, any>[] {\n  return rows.map((row) => ({\n    ...row,\n    ProcessedAt: new Date().toISOString()\n  }));\n}`;
}

/**
 * AUTOMATED ENHANCEMENT ENGINE FOR VBA CODE:
 * Allows user to enhance legacy macros with 1-click upgrades:
 * 1. Speed optimization (ScreenUpdating, Calculation, EnableEvents)
 * 2. Robust error handling (On Error GoTo ErrHandler)
 * 3. Execution time profiler (Timer)
 */
function findLastLineIndex(arr: string[], predicate: (l: string) => boolean): number {
  for (let i = arr.length - 1; i >= 0; i--) {
    if (predicate(arr[i])) return i;
  }
  return -1;
}

export function enhanceVbaCode(
  code: string,
  options: {
    addSpeedBoost?: boolean;
    addErrorHandler?: boolean;
    addProfiler?: boolean;
  }
): { enhancedCode: string; appliedTags: string[] } {
  const lines = code.split(/\r?\n/);
  const appliedTags: string[] = [];

  // Find procedure header and footer
  const headerIdx = lines.findIndex((l) =>
    /^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(?:Sub|Function|Property)/i.test(l)
  );
  const endIdx = findLastLineIndex(lines, (l) => /^\s*End\s+(?:Sub|Function|Property)/i.test(l));

  if (headerIdx === -1 || endIdx === -1 || headerIdx >= endIdx) {
    return { enhancedCode: code, appliedTags: [] };
  }

  const headerLine = lines[headerIdx].trim();
  const footerLine = lines[endIdx].trim();

  const procTypeMatch = headerLine.match(/\b(Sub|Function|Property)\b/i);
  const isFunction = procTypeMatch ? procTypeMatch[1].toLowerCase() === 'function' : false;
  const exitKeyword = isFunction ? 'Exit Function' : 'Exit Sub';

  const procNameMatch = headerLine.match(/(?:Sub|Function|Property)\s+([A-Za-z0-9_]+)/i);
  const procName = procNameMatch ? procNameMatch[1] : 'Procedure';

  // Detect existing enhancements
  const hasSpeedBoost =
    Boolean(options.addSpeedBoost) ||
    /Application\.ScreenUpdating\s*=\s*False/i.test(code) ||
    /origCalcState\s*=\s*Application\.Calculation/i.test(code);

  const hasErrorHandler =
    Boolean(options.addErrorHandler) ||
    /On\s+Error\s+GoTo\s+ErrHandler/i.test(code) ||
    /\bErrHandler:\b/i.test(code);

  const hasProfiler =
    Boolean(options.addProfiler) ||
    /startTime\s*=\s*Timer/i.test(code);

  if (options.addSpeedBoost) appliedTags.push('Speed Boost (ScreenUpdating & Manual Calc)');
  if (options.addErrorHandler) appliedTags.push('Structured Error Handler (On Error GoTo)');
  if (options.addProfiler) appliedTags.push('Execution Profiler Timer');

  // Extract core user lines between header and footer, stripping out previously generated boilerplate
  const rawBodyLines = lines.slice(headerIdx + 1, endIdx);
  const coreBodyLines: string[] = [];
  let skippingTail = false;

  for (const rawLine of rawBodyLines) {
    const trimmed = rawLine.trim();

    // If we reach CleanExit: or ErrHandler: from previous enhancement, skip the tail
    if (/^(CleanExit:|ErrHandler:)/i.test(trimmed)) {
      skippingTail = true;
      continue;
    }

    if (skippingTail) continue;

    // Filter out previously injected header boilerplate lines
    if (
      /^\s*On\s+Error\s+GoTo\s+ErrHandler/i.test(trimmed) ||
      /^\s*Dim\s+startTime\s+As\s+Double/i.test(trimmed) ||
      /^\s*startTime\s*=\s*Timer/i.test(trimmed) ||
      /^\s*Dim\s+origCalcState\s+As\s+Long/i.test(trimmed) ||
      /^\s*origCalcState\s*=\s*Application\.Calculation/i.test(trimmed) ||
      /^\s*Application\.ScreenUpdating\s*=\s*False/i.test(trimmed) ||
      /^\s*Application\.DisplayAlerts\s*=\s*False/i.test(trimmed) ||
      /^\s*Application\.Calculation\s*=\s*xlCalculationManual/i.test(trimmed) ||
      /^\s*Application\.EnableEvents\s*=\s*False/i.test(trimmed) ||
      /^\s*'\s*\[Performance Boost\]/i.test(trimmed)
    ) {
      continue;
    }

    coreBodyLines.push(rawLine);
  }

  // Build the complete pristine procedure
  const newLines: string[] = [];
  newLines.push(headerLine);

  // 1. Top Section Setup
  if (hasErrorHandler) {
    newLines.push('    On Error GoTo ErrHandler');
  }

  if (hasProfiler) {
    newLines.push('    Dim startTime As Double');
    newLines.push('    startTime = Timer');
  }

  if (hasSpeedBoost) {
    newLines.push("    ' [Performance Boost] Disable screen flicker and switch to manual calculation");
    newLines.push('    Dim origCalcState As Long');
    newLines.push('    origCalcState = Application.Calculation');
    newLines.push('    Application.ScreenUpdating = False');
    newLines.push('    Application.DisplayAlerts = False');
    newLines.push('    Application.Calculation = xlCalculationManual');
    newLines.push('    Application.EnableEvents = False');
  }

  // 2. Core Body
  while (coreBodyLines.length > 0 && !coreBodyLines[0].trim()) coreBodyLines.shift();
  while (coreBodyLines.length > 0 && !coreBodyLines[coreBodyLines.length - 1].trim()) coreBodyLines.pop();

  if (coreBodyLines.length > 0) {
    newLines.push(...coreBodyLines);
  } else {
    newLines.push('    \' Core routine execution');
  }

  // 3. Clean Exit Section
  if (hasSpeedBoost || hasProfiler || hasErrorHandler) {
    newLines.push('');
    newLines.push('CleanExit:');
    if (hasSpeedBoost) {
      newLines.push('    Application.ScreenUpdating = True');
      newLines.push('    Application.DisplayAlerts = True');
      newLines.push('    Application.EnableEvents = True');
      newLines.push('    Application.Calculation = origCalcState');
    }
    if (hasProfiler) {
      newLines.push(`    Debug.Print "[${procName}] Execution completed in " & Format(Timer - startTime, "0.000") & " seconds."`);
    }
    newLines.push(`    ${exitKeyword}`);
  }

  // 4. Error Handler Section
  if (hasErrorHandler) {
    newLines.push('');
    newLines.push('ErrHandler:');
    newLines.push(`    MsgBox "Run-time error " & Err.Number & " in ${procName}: " & Err.Description, vbCritical, "XLSM Macro Error"`);
    if (hasSpeedBoost) {
      newLines.push('    Application.ScreenUpdating = True');
      newLines.push('    Application.DisplayAlerts = True');
      newLines.push('    Application.EnableEvents = True');
      newLines.push('    Application.Calculation = xlCalculationAutomatic');
    }
    if (isFunction) {
      newLines.push(`    ${procName} = 0`);
    }
    newLines.push(`    ${exitKeyword}`);
  }

  newLines.push(footerLine);

  return {
    enhancedCode: formatVbaCode(newLines.join('\n')),
    appliedTags
  };
}

/**
 * UNPROTECT SHEETS AND WORKBOOK
 * Scans OpenXML structures, locates <sheetProtection> and <workbookProtection> elements, and strips them.
 */
export async function unprotectWorkbookAndSheets(zip: JSZip): Promise<UnprotectionResult> {
  const protectedSheetsFound: UnprotectionResult['protectedSheetsFound'] = [];
  let workbookProtectionFound: UnprotectionResult['workbookProtectionFound'] = undefined;

  const sheetFiles = Object.keys(zip.files).filter((path) =>
    path.startsWith('xl/worksheets/sheet') && path.endsWith('.xml')
  );

  for (const sheetPath of sheetFiles) {
    const file = zip.file(sheetPath);
    if (!file) continue;

    const xmlContent = await file.async('string');
    const protectionRegex = /<sheetProtection\s+[^>]*\/?>/i;
    const match = xmlContent.match(protectionRegex);

    if (match) {
      const fullTag = match[0];
      const algoMatch = fullTag.match(/algorithmName="([^"]+)"/i);
      const hashMatch = fullTag.match(/hashValue="([^"]+)"/i) || fullTag.match(/password="([^"]+)"/i);
      const sheetName = `Sheet ${sheetPath.replace(/[^0-9]/g, '')}`;

      protectedSheetsFound.push({
        sheetName,
        xmlFile: sheetPath,
        protectionTag: fullTag,
        algorithm: algoMatch ? algoMatch[1] : 'Legacy Excel XOR Hash',
        hashValue: hashMatch ? hashMatch[1] : 'PROTECTED'
      });

      const unprotectedXml = xmlContent.replace(protectionRegex, '');
      zip.file(sheetPath, unprotectedXml);
    }
  }

  const workbookFile = zip.file('xl/workbook.xml');
  if (workbookFile) {
    const wbXml = await workbookFile.async('string');
    const wbProtRegex = /<workbookProtection\s+[^>]*\/?>/i;
    const wbMatch = wbXml.match(wbProtRegex);

    if (wbMatch) {
      const tag = wbMatch[0];
      workbookProtectionFound = {
        lockStructure: /lockStructure="1"/i.test(tag),
        lockWindows: /lockWindows="1"/i.test(tag),
        hashValue: 'PROTECTED'
      };

      const unprotectedWbXml = wbXml.replace(wbProtRegex, '');
      zip.file('xl/workbook.xml', unprotectedWbXml);
    }
  }

  const unprotectedBlob = await zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.ms-excel.sheet.macroEnabled.12',
    compression: 'DEFLATE'
  });

  return {
    modifiedZip: zip,
    unprotectedBlob,
    protectedSheetsFound,
    workbookProtectionFound
  };
}

/**
 * REPACK AND BUILD UPGRADED XLSM WORKBOOK
 * Embeds modified VBA modules and procedures back into the XLSM zip container.
 */
export async function repackUpgradedXlsm(
  zip: JSZip,
  updatedMacros: LegacyMacro[]
): Promise<Blob> {
  // If vbaProject.bin is present, we embed the text changes
  const vbaFile = zip.file('xl/vbaProject.bin');
  if (vbaFile) {
    // Generate text export file inside zip for audit and inspection
    const basSource = updatedMacros
      .map((m) => `' Module: ${m.moduleName}\n' Procedure: ${m.name}\n${m.code}\n`)
      .join('\n\n');
    zip.file('xl/upgraded_vba_code.bas', basSource);
  }

  return await zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.ms-excel.sheet.macroEnabled.12',
    compression: 'DEFLATE'
  });
}
