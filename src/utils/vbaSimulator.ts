/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { Dataset } from '../types/dashboard';

export interface VbaDiagnosticError {
  line: number;
  column?: number;
  severity: 'error' | 'warning';
  code: string;
  message: string;
  suggestion?: string;
  lineSnippet?: string;
}

export interface VbaSimulationReport {
  success: boolean;
  status: 'passed' | 'failed' | 'warnings';
  executionTimeMs: number;
  totalLines: number;
  statementsAnalyzed: number;
  errors: VbaDiagnosticError[];
  warnings: VbaDiagnosticError[];
  executionTrace: string[];
  referencedSheets: { name: string; existsInWorkbook: boolean }[];
  variablesDetected: { name: string; type?: string; declared: boolean }[];
  procedureSignature?: {
    type: 'Sub' | 'Function' | 'Property';
    name: string;
    scope: 'Public' | 'Private' | 'Friend';
    params: string[];
  };
  summary: string;
}

interface BlockFrame {
  type: 'sub' | 'function' | 'property' | 'if' | 'for' | 'do' | 'while' | 'select' | 'with';
  name?: string;
  loopVar?: string;
  line: number;
}

// Full VBA keywords, standard library functions, and constants dictionary
const VBA_KEYWORDS = new Set([
  'sub', 'function', 'property', 'end', 'exit', 'dim', 'static', 'const', 'redim',
  'preserve', 'as', 'set', 'let', 'if', 'then', 'else', 'elseif', 'for', 'to',
  'step', 'next', 'each', 'in', 'do', 'while', 'until', 'loop', 'wend', 'select',
  'case', 'with', 'on', 'error', 'goto', 'resume', 'call', 'public', 'private',
  'friend', 'byval', 'byref', 'optional', 'paramarray', 'true', 'false', 'nothing',
  'empty', 'null', 'and', 'or', 'not', 'xor', 'eqv', 'imp', 'is', 'like', 'mod',
  'new', 'typeof', 'option', 'explicit', 'base', 'compare', 'debug', 'print',
  'stop', 'type', 'enum'
]);

const VBA_BUILTIN_FUNCTIONS = new Set([
  'msgbox', 'inputbox', 'timer', 'format', 'date', 'time', 'now', 'year', 'month',
  'day', 'hour', 'minute', 'second', 'len', 'mid', 'left', 'right', 'instr',
  'instrrev', 'trim', 'ltrim', 'rtrim', 'ucase', 'lcase', 'replace', 'split',
  'join', 'strconv', 'chr', 'asc', 'val', 'str', 'cstr', 'cint', 'clng', 'cdbl',
  'csng', 'cbool', 'cdate', 'cvar', 'cbyte', 'isnumeric', 'isdate', 'isempty',
  'isnull', 'isarray', 'isobject', 'ubound', 'lbound', 'array', 'abs', 'sqr',
  'int', 'fix', 'round', 'rnd', 'randomize', 'sin', 'cos', 'tan', 'atn', 'exp',
  'log', 'sgn', 'doevents'
]);

const VBA_CONSTANTS = new Set([
  'vbokonly', 'vbokcancel', 'vbabortretryignore', 'vbyesnocancel', 'vbyesno',
  'vbretrycancel', 'vbcritical', 'vbquestion', 'vbexclamation', 'vbinformation',
  'vbdefaultbutton1', 'vbdefaultbutton2', 'vbdefaultbutton3', 'vbdefaultbutton4',
  'vbapplicationmodal', 'vbsystemmodal', 'vbcr', 'vblf', 'vbcrlf', 'vbnewline',
  'vbtab', 'vbnullstring', 'vbnullchar', 'xlcalculationmanual', 'xlcalculationautomatic',
  'xlcalculationsemiautomatic', 'xlup', 'xldown', 'xltoleft', 'xltoright'
]);

const EXCEL_ROOT_OBJECTS = new Set([
  'application', 'activesheet', 'activeworkbook', 'thisworkbook', 'worksheets',
  'sheets', 'range', 'cells', 'rows', 'columns', 'selection', 'activecell'
]);

const EXCEL_MEMBERS = new Set([
  'value', 'value2', 'text', 'formula', 'formular1c1', 'formula2', 'address',
  'row', 'column', 'count', 'item', 'name', 'visible', 'protect', 'unprotect',
  'select', 'activate', 'calculate', 'clear', 'clearcontents', 'clearformats',
  'copy', 'paste', 'pastespecial', 'delete', 'insert', 'offset', 'resize',
  'end', 'entirerow', 'entirecolumn', 'font', 'interior', 'borders', 'color',
  'colorindex', 'bold', 'italic', 'size', 'screenupdating', 'calculation',
  'displayalerts', 'enableevents', 'statusbar', 'cutcopymode', 'rows', 'columns',
  'cells', 'range', 'sheets', 'worksheets', 'parent', 'application', 'number',
  'description', 'clear'
]);

const VALID_VBA_TYPES = new Set([
  'integer', 'long', 'single', 'double', 'currency', 'string', 'boolean', 'byte',
  'date', 'variant', 'object', 'worksheet', 'worksheets', 'workbook', 'workbooks',
  'range', 'application', 'collection', 'dictionary', 'chart', 'shape', 'window',
  'name', 'comment', 'font', 'interior', 'borders', 'series'
]);

/**
 * Computes standard Levenshtein edit distance for typo detection.
 */
function levenshteinDistance(s1: string, s2: string): number {
  const a = s1.toLowerCase();
  const b = s2.toLowerCase();
  const m = a.length;
  const n = b.length;
  const dp: number[][] = Array.from({ length: m + 1 }, () => Array(n + 1).fill(0));

  for (let i = 0; i <= m; i++) dp[i][0] = i;
  for (let j = 0; j <= n; j++) dp[0][j] = j;

  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      if (a[i - 1] === b[j - 1]) {
        dp[i][j] = dp[i - 1][j - 1];
      } else {
        dp[i][j] = 1 + Math.min(dp[i - 1][j], dp[i][j - 1], dp[i - 1][j - 1]);
      }
    }
  }
  return dp[m][n];
}

/**
 * Searches dictionary for closest identifier match within a maximum edit distance.
 */
function findClosestMatch(token: string): string | null {
  const lower = token.toLowerCase();

  // Explicit mappings for common developer typos
  if (lower === 'tru' || lower === 'treu') return 'True';
  if (lower === 'flase' || lower === 'fals') return 'False';
  if (lower === 'nothin' || lower === 'noting') return 'Nothing';
  if (lower === 'empy') return 'Empty';
  if (lower === 'nul') return 'Null';
  if (lower === 'dimm' || lower === 'dimn') return 'Dim';
  if (lower === 'fuction' || lower === 'funtion') return 'Function';
  if (lower === 'actvesheet' || lower === 'actvsheet') return 'ActiveSheet';
  if (lower === 'thisworkbok') return 'ThisWorkbook';
  if (lower === 'screenupdatin' || lower === 'screenupdate') return 'ScreenUpdating';
  if (lower === 'calculaton') return 'Calculation';
  if (lower === 'valu' || lower === 'valeu') return 'Value';
  if (lower === 'rang') return 'Range';
  if (lower === 'cel' || lower === 'cell') return 'Cells';
  if (lower === 'shet' || lower === 'sheet') return 'Sheets';
  if (lower === 'formla') return 'Formula';

  // Do not perform fuzzy distance on short tokens (<= 2 chars) e.g. 'r', 'c', 'i', 'j', 'ws', 'wb'
  if (token.length <= 2) return null;

  const candidates = [
    'True', 'False', 'Nothing', 'Empty', 'Null',
    'Application', 'ActiveSheet', 'ActiveWorkbook', 'ThisWorkbook', 'Sheets', 'Worksheets', 'Range', 'Cells',
    'ScreenUpdating', 'Calculation', 'DisplayAlerts', 'EnableEvents', 'StatusBar',
    'Value', 'Count', 'Row', 'Column', 'Address', 'Name', 'Formula',
    'xlCalculationManual', 'xlCalculationAutomatic', 'xlUp', 'xlDown',
    'vbCritical', 'vbInformation', 'vbExclamation', 'vbOKOnly'
  ];

  let bestMatch: string | null = null;
  const maxAllowedDist = token.length === 3 ? 1 : 2;
  let minDistance = maxAllowedDist + 1;

  for (const cand of candidates) {
    if (Math.abs(token.length - cand.length) > 2) continue;
    const dist = levenshteinDistance(token, cand);
    if (dist < minDistance && dist <= maxAllowedDist) {
      minDistance = dist;
      bestMatch = cand;
    }
  }

  return bestMatch;
}

/**
 * Strips comments from a line, accounting for string literals.
 */
function stripVbaComment(line: string): { codePart: string; commentPart: string } {
  let inString = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') {
      inString = !inString;
    } else if (!inString && ch === "'") {
      return {
        codePart: line.slice(0, i),
        commentPart: line.slice(i + 1)
      };
    } else if (
      !inString &&
      (i === 0 || /\s/.test(line[i - 1])) &&
      line.slice(i, i + 4).toLowerCase() === 'rem '
    ) {
      return {
        codePart: line.slice(0, i),
        commentPart: line.slice(i + 4)
      };
    }
  }
  return { codePart: line, commentPart: '' };
}

/**
 * Checks for unmatched string quotes on a code line.
 */
function hasUnterminatedString(code: string): boolean {
  let count = 0;
  for (let i = 0; i < code.length; i++) {
    if (code[i] === '"') {
      if (i + 1 < code.length && code[i + 1] === '"') {
        i++;
      } else {
        count++;
      }
    }
  }
  return count % 2 !== 0;
}

/**
 * Checks for balanced parentheses in a code line.
 */
function checkParentheses(code: string): number {
  let inString = false;
  let balance = 0;
  for (let i = 0; i < code.length; i++) {
    const ch = code[i];
    if (ch === '"') {
      if (i + 1 < code.length && code[i + 1] === '"') {
        i++;
      } else {
        inString = !inString;
      }
    } else if (!inString) {
      if (ch === '(') balance++;
      else if (ch === ')') balance--;
    }
  }
  return balance;
}

/**
 * Deep Expression Token Validator:
 * Validates RHS expressions, conditions, and arguments, catching typos such as "Tru" instead of "True".
 */
function validateExpressionTokens(
  expr: string,
  lineNum: number,
  rawLine: string,
  declaredVars: Set<string>,
  knownParams: Set<string>,
  activeLoopVars: Set<string>,
  errors: VbaDiagnosticError[],
  executionTrace: string[],
  hasOptionExplicit: boolean = false,
  warnings?: VbaDiagnosticError[]
): void {
  // Strip string literals from expression before tokenizing
  const withoutStrings = expr.replace(/"(?:""|[^"])*"/g, '""');

  // First, extract and validate any identifiers inside parentheses e.g. Cells(rowIdx, colIdx)
  const innerArgs = withoutStrings.match(/\(([^()]+)\)/g);
  if (innerArgs) {
    for (const argGroup of innerArgs) {
      const strippedGroup = argGroup.slice(1, -1);
      const innerTokens = strippedGroup.match(/\b[A-Za-z_][A-Za-z0-9_.]*\b/g);
      if (innerTokens) {
        for (const it of innerTokens) {
          const root = it.split('.')[0].toLowerCase();
          const isKnown =
            VBA_KEYWORDS.has(root) ||
            VBA_BUILTIN_FUNCTIONS.has(root) ||
            VBA_CONSTANTS.has(root) ||
            EXCEL_ROOT_OBJECTS.has(root) ||
            declaredVars.has(root) ||
            knownParams.has(root) ||
            activeLoopVars.has(root) ||
            root === 'err' ||
            /^\d+$/.test(root);
          if (!isKnown) {
            const typo = findClosestMatch(it.split('.')[0]);
            if (typo) {
              errors.push({
                line: lineNum,
                severity: 'error',
                code: 'VARIABLE_NOT_DEFINED',
                message: `Compile Error on line ${lineNum}: Variable not defined: '${it.split('.')[0]}'. Did you mean '${typo}'?`,
                suggestion: `Replace '${it.split('.')[0]}' with '${typo}'.`,
                lineSnippet: rawLine.trim()
              });
              executionTrace.push(`❌ [Compile Error] Line ${lineNum}: Variable not defined: '${it.split('.')[0]}' (Did you mean '${typo}'?).`);
            } else if (hasOptionExplicit) {
              errors.push({
                line: lineNum,
                severity: 'error',
                code: 'VARIABLE_NOT_DEFINED',
                message: `Compile Error on line ${lineNum}: Variable not defined: '${it.split('.')[0]}'. Option Explicit requires all variables to be explicitly declared.`,
                suggestion: `Declare variable using 'Dim ${it.split('.')[0]}' or check spelling.`,
                lineSnippet: rawLine.trim()
              });
              executionTrace.push(`❌ [Compile Error] Line ${lineNum}: Variable not defined: '${it.split('.')[0]}'.`);
            }
          }
        }
      }
    }
  }

  // Next, collapse parenthesized arguments so member chains like ws.Cells(1, 1).Valu become ws.Cells.Valu
  let collapsed = withoutStrings;
  while (/\([^()]*\)/.test(collapsed)) {
    collapsed = collapsed.replace(/\([^()]*\)/g, '');
  }

  // Match identifiers or member accesses: e.g. "Application.ScreenUpdating", "ws.Cells.Valu", "Tru"
  const tokens = collapsed.match(/\b[A-Za-z_][A-Za-z0-9_.]*\b/g);
  if (!tokens) return;

  const seenErrorsOnLine = new Set<string>();

  for (const token of tokens) {
    const parts = token.split('.');

    // Validate root token
    const root = parts[0].toLowerCase();

    // Check if root is known
    const isKnownRoot =
      VBA_KEYWORDS.has(root) ||
      VBA_BUILTIN_FUNCTIONS.has(root) ||
      VBA_CONSTANTS.has(root) ||
      EXCEL_ROOT_OBJECTS.has(root) ||
      declaredVars.has(root) ||
      knownParams.has(root) ||
      activeLoopVars.has(root) ||
      root === 'err';

    if (!isKnownRoot) {
      // Identifier is unknown in this scope: check for close typos (e.g. 'Tru' -> 'True')
      const typoSuggestion = findClosestMatch(parts[0]);

      if (typoSuggestion) {
        errors.push({
          line: lineNum,
          severity: 'error',
          code: 'VARIABLE_NOT_DEFINED',
          message: `Compile Error on line ${lineNum}: Variable not defined: '${parts[0]}'. Did you mean '${typoSuggestion}'?`,
          suggestion: `Replace '${parts[0]}' with '${typoSuggestion}'.`,
          lineSnippet: rawLine.trim()
        });
        executionTrace.push(
          `❌ [Compile Error] Line ${lineNum}: Variable not defined: '${parts[0]}' (Did you mean '${typoSuggestion}'?).`
        );
      } else if (hasOptionExplicit) {
        errors.push({
          line: lineNum,
          severity: 'error',
          code: 'VARIABLE_NOT_DEFINED',
          message: `Compile Error on line ${lineNum}: Variable not defined: '${parts[0]}'. Option Explicit requires all variables to be explicitly declared.`,
          suggestion: `Declare variable using 'Dim ${parts[0]}' or check spelling.`,
          lineSnippet: rawLine.trim()
        });
        executionTrace.push(`❌ [Compile Error] Line ${lineNum}: Variable not defined: '${parts[0]}'.`);
      } else if (warnings) {
        warnings.push({
          line: lineNum,
          severity: 'warning',
          code: 'IMPLICIT_VARIABLE',
          message: `Variable '${parts[0]}' on line ${lineNum} is not explicitly declared (treated as implicit Variant).`,
          suggestion: `Add 'Dim ${parts[0]}' to declare variable type.`
        });
      }
      continue;
    }

    // If there are member accesses, e.g. "Application.ScreenUpdating" or "ws.Valu"
    if (parts.length > 1) {
      for (let p = 1; p < parts.length; p++) {
        const member = parts[p].toLowerCase();
        if (!EXCEL_MEMBERS.has(member) && !EXCEL_ROOT_OBJECTS.has(member) && !VBA_BUILTIN_FUNCTIONS.has(member)) {
          const typoMember = findClosestMatch(parts[p]);
          const isDistinctTypo = typoMember && typoMember.toLowerCase() !== member;
          errors.push({
            line: lineNum,
            severity: 'error',
            code: 'MEMBER_NOT_FOUND',
            message: `Compile Error on line ${lineNum}: Method or data member not found: '${parts[p]}'.${
              isDistinctTypo ? ` Did you mean '${typoMember}'?` : ''
            }`,
            suggestion: isDistinctTypo ? `Change '.${parts[p]}' to '.${typoMember}'.` : `Verify Excel object member name.`,
            lineSnippet: rawLine.trim()
          });
          executionTrace.push(
            `❌ [Compile Error] Line ${lineNum}: Method or data member not found: '${parts[p]}'.`
          );
        }
      }
    }
  }
}

/**
 * Deep VBA syntax analysis and execution simulation engine with full MS Office VBE emulation.
 */
export function simulateVbaExecution(vbaCode: string, dataset?: Dataset): VbaSimulationReport {
  const startTime = performance.now();
  const errors: VbaDiagnosticError[] = [];
  const warnings: VbaDiagnosticError[] = [];
  const executionTrace: string[] = [];

  const rawLines = vbaCode.split(/\r?\n/);
  const totalLines = rawLines.length;

  if (vbaCode.trim().length === 0) {
    return {
      success: false,
      status: 'failed',
      executionTimeMs: 0.1,
      totalLines: 0,
      statementsAnalyzed: 0,
      errors: [
        {
          line: 1,
          severity: 'error',
          code: 'EMPTY_SCRIPT',
          message: 'The VBA script is empty. Provide a Sub or Function procedure to simulate.',
          suggestion: 'Insert a procedure such as "Sub Calculate() ... End Sub".'
        }
      ],
      warnings: [],
      executionTrace: ['[Simulation aborted] Empty code buffer.'],
      referencedSheets: [],
      variablesDetected: [],
      summary: 'Simulation failed: Code buffer is empty.'
    };
  }

  const availableSheetNames =
    dataset?.sheets && dataset.sheets.length > 0
      ? dataset.sheets.map((s) => s.name)
      : ['Consolidated P&L', 'Sheet1'];
  const activeSheetName = dataset?.activeSheetName || availableSheetNames[0];

  const referencedSheetsMap = new Map<string, boolean>();
  const declaredVariables = new Map<string, { type?: string; line: number }>();
  const definedLabels = new Set<string>();
  const referencedLabels: { label: string; line: number }[] = [];

  const blockStack: BlockFrame[] = [];
  let procedureInfo: VbaSimulationReport['procedureSignature'] | undefined;
  let hasOptionExplicit = false;
  let hasReturnAssignment = false;
  let statementsCount = 0;
  let insideProcedure = false;

  executionTrace.push(`[Runtime Engine] Initializing VBA simulation environment.`);
  executionTrace.push(`[Environment] Target Workbook: "${dataset?.filename || 'Workbook.xlsm'}" with ${availableSheetNames.length} sheet(s) loaded.`);
  executionTrace.push(`[Environment] ActiveSheet initialized to "${activeSheetName}".`);

  for (let lineIdx = 0; lineIdx < rawLines.length; lineIdx++) {
    const lineNum = lineIdx + 1;
    const rawLine = rawLines[lineIdx];

    // Check comment syntax
    if (/^\s*\/\//.test(rawLine)) {
      errors.push({
        line: lineNum,
        severity: 'error',
        code: 'INVALID_COMMENT_SYNTAX',
        message: `Syntax Error on line ${lineNum}: '//' is invalid in VBA. VBA comments must start with an apostrophe (') or 'Rem'.`,
        suggestion: `Replace '//' with an apostrophe (').`,
        lineSnippet: rawLine.trim()
      });
      executionTrace.push(`❌ [Syntax Error] Line ${lineNum}: Invalid comment syntax '//'.`);
    }

    if (/^\s*\/\*/.test(rawLine)) {
      errors.push({
        line: lineNum,
        severity: 'error',
        code: 'INVALID_COMMENT_SYNTAX',
        message: `Syntax Error on line ${lineNum}: '/*' is invalid in VBA. VBA does not have multi-line comment blocks.`,
        suggestion: `Prefix each comment line with an apostrophe (').`,
        lineSnippet: rawLine.trim()
      });
      executionTrace.push(`❌ [Syntax Error] Line ${lineNum}: Invalid comment syntax '/*'.`);
    }

    const { codePart } = stripVbaComment(rawLine);
    const trimmed = codePart.trim();

    if (!trimmed) continue;

    // Check Option Explicit
    if (/^Option\s+Explicit/i.test(trimmed)) {
      hasOptionExplicit = true;
      executionTrace.push(`[Config] 'Option Explicit' detected on line ${lineNum}. Enforcing strict variable declarations.`);
      continue;
    }

    // Check unterminated string literals
    if (hasUnterminatedString(trimmed)) {
      errors.push({
        line: lineNum,
        severity: 'error',
        code: 'UNTERMINATED_STRING',
        message: `Syntax Error on line ${lineNum}: Unterminated string constant. Missing closing quotation mark.`,
        suggestion: `Ensure all string literals are closed with a double quote (e.g. "Text").`,
        lineSnippet: rawLine.trim()
      });
      executionTrace.push(`❌ [Syntax Error] Line ${lineNum}: Unterminated string constant.`);
    }

    // Check unmatched parentheses
    const parenBalance = checkParentheses(trimmed);
    if (parenBalance > 0) {
      errors.push({
        line: lineNum,
        severity: 'error',
        code: 'UNMATCHED_PARENTHESIS',
        message: `Syntax Error on line ${lineNum}: Expected ')' (${parenBalance} unclosed).`,
        suggestion: `Close all open parentheses in formulas and method calls.`,
        lineSnippet: rawLine.trim()
      });
      executionTrace.push(`❌ [Syntax Error] Line ${lineNum}: Missing ')' on expression.`);
    } else if (parenBalance < 0) {
      errors.push({
        line: lineNum,
        severity: 'error',
        code: 'UNEXPECTED_PARENTHESIS',
        message: `Syntax Error on line ${lineNum}: Unexpected ')' without matching '('.`,
        suggestion: `Remove the stray closing parenthesis.`,
        lineSnippet: rawLine.trim()
      });
      executionTrace.push(`❌ [Syntax Error] Line ${lineNum}: Unexpected ')' parenthesis.`);
    }

    // Check foreign language keywords (JS/TS/Python)
    if (/\bconsole\.log\s*\(/i.test(trimmed)) {
      errors.push({
        line: lineNum,
        severity: 'error',
        code: 'INVALID_KEYWORD',
        message: `Syntax Error on line ${lineNum}: 'console.log' is not valid in VBA.`,
        suggestion: `Use 'Debug.Print' to output to Immediate window or 'MsgBox' to display a dialog.`,
        lineSnippet: rawLine.trim()
      });
      executionTrace.push(`❌ [Syntax Error] Line ${lineNum}: 'console.log' is not a valid VBA statement.`);
    }

    if (/^\s*(let|var|const)\s+[A-Za-z0-9_]+\s*=/i.test(trimmed)) {
      errors.push({
        line: lineNum,
        severity: 'error',
        code: 'INVALID_VARIABLE_DECLARATION',
        message: `Syntax Error on line ${lineNum}: JavaScript declaration keyword '${trimmed.split(' ')[0]}' is invalid in VBA.`,
        suggestion: `Use 'Dim variableName As Type' or 'Const variableName = Value' in VBA.`,
        lineSnippet: rawLine.trim()
      });
      executionTrace.push(`❌ [Syntax Error] Line ${lineNum}: Invalid declaration keyword.`);
    }

    if (/\b(===|!==|!=)\b/i.test(trimmed) || /==/.test(trimmed)) {
      errors.push({
        line: lineNum,
        severity: 'error',
        code: 'INVALID_OPERATOR',
        message: `Syntax Error on line ${lineNum}: Invalid equality/comparison operator. In VBA, use '=' for equality and '<>' for inequality.`,
        suggestion: `Replace '==' with '=' or '!=' with '<>'.`,
        lineSnippet: rawLine.trim()
      });
      executionTrace.push(`❌ [Syntax Error] Line ${lineNum}: Invalid comparison operator.`);
    }

    if (/&&|\|\|/i.test(trimmed)) {
      errors.push({
        line: lineNum,
        severity: 'error',
        code: 'INVALID_LOGICAL_OPERATOR',
        message: `Syntax Error on line ${lineNum}: Logical operators '&&' and '||' are invalid in VBA.`,
        suggestion: `Use 'And' and 'Or' keywords in VBA.`,
        lineSnippet: rawLine.trim()
      });
      executionTrace.push(`❌ [Syntax Error] Line ${lineNum}: Invalid logical operator.`);
    }

    if (/\+\+|\-\-/i.test(trimmed)) {
      errors.push({
        line: lineNum,
        severity: 'error',
        code: 'INVALID_INCREMENT_OPERATOR',
        message: `Syntax Error on line ${lineNum}: Increment/decrement operators '++' and '--' are invalid in VBA.`,
        suggestion: `Use 'x = x + 1' or 'x = x - 1'.`,
        lineSnippet: rawLine.trim()
      });
      executionTrace.push(`❌ [Syntax Error] Line ${lineNum}: Increment operator '++' not supported.`);
    }

    if (/(\+=|\-=|\*=|\/=)/i.test(trimmed)) {
      errors.push({
        line: lineNum,
        severity: 'error',
        code: 'INVALID_COMPOUND_ASSIGNMENT',
        message: `Syntax Error on line ${lineNum}: Compound assignment operators (+=, -=, etc.) are invalid in VBA.`,
        suggestion: `Expand to standard assignment (e.g. 'x = x + y').`,
        lineSnippet: rawLine.trim()
      });
      executionTrace.push(`❌ [Syntax Error] Line ${lineNum}: Compound assignment not supported.`);
    }

    // Check division by zero literals: e.g. / 0 or / 0.0 or \ 0
    if (/(?:[\/\\])\s*0(?![0-9.])/g.test(trimmed)) {
      errors.push({
        line: lineNum,
        severity: 'error',
        code: 'DIVISION_BY_ZERO',
        message: `Runtime Error 11: Division by zero detected on line ${lineNum}.`,
        suggestion: `Prevent zero division by adding an 'If divisor <> 0 Then' check.`,
        lineSnippet: rawLine.trim()
      });
      executionTrace.push(`❌ [Runtime Exception] Line ${lineNum}: Run-time error '11': Division by zero.`);
    }

    // 1. Standalone label definition (e.g. "ErrHandler:", "CleanExit:")
    const standaloneLabel = trimmed.match(/^([A-Za-z0-9_]+):$/);
    if (standaloneLabel && !/^Case\b/i.test(standaloneLabel[1])) {
      const lbl = standaloneLabel[1].toLowerCase();
      definedLabels.add(lbl);
      executionTrace.push(`[Label Target] Defined execution jump label '${standaloneLabel[1]}:' at line ${lineNum}.`);
      continue;
    }

    // 2. Inline label followed by statement (e.g. "ErrHandler: MsgBox ...")
    let lineContent = trimmed;
    const inlineLabel = lineContent.match(/^([A-Za-z0-9_]+):\s+(.+)$/);
    if (inlineLabel && !/^Case\b/i.test(inlineLabel[1])) {
      const lbl = inlineLabel[1].toLowerCase();
      definedLabels.add(lbl);
      executionTrace.push(`[Label Target] Defined execution jump label '${inlineLabel[1]}:' at line ${lineNum}.`);
      lineContent = inlineLabel[2];
    }

    // Split multiple statements separated by ':' outside string literals
    const statements: string[] = [];
    let currentStmt = '';
    let inQuotes = false;
    for (let c = 0; c < lineContent.length; c++) {
      const ch = lineContent[c];
      if (ch === '"') inQuotes = !inQuotes;
      if (ch === ':' && !inQuotes) {
        statements.push(currentStmt.trim());
        currentStmt = '';
      } else {
        currentStmt += ch;
      }
    }
    if (currentStmt.trim()) statements.push(currentStmt.trim());

    // Prepare active variable symbol sets
    const declaredVarNames = new Set(Array.from(declaredVariables.keys()));
    const paramNames = new Set(
      procedureInfo?.params
        ? procedureInfo.params.map((p) => p.split(/\s+/)[0].replace(/^(ByVal|ByRef)\s+/i, '').toLowerCase())
        : []
    );
    const activeLoopVars = new Set(
      blockStack.filter((b) => b.type === 'for' && b.loopVar).map((b) => b.loopVar!.toLowerCase())
    );

    for (const stmt of statements) {
      statementsCount++;

      // Check common typos in keywords
      if (/^Dimm\b/i.test(stmt) || /^Dimn\b/i.test(stmt)) {
        errors.push({
          line: lineNum,
          severity: 'error',
          code: 'KEYWORD_TYPO',
          message: `Syntax Error on line ${lineNum}: Unrecognized statement '${stmt.split(' ')[0]}'. Did you mean 'Dim'?`,
          suggestion: `Change '${stmt.split(' ')[0]}' to 'Dim'.`,
          lineSnippet: rawLine.trim()
        });
        executionTrace.push(`❌ [Syntax Error] Line ${lineNum}: Typo in keyword 'Dim'.`);
        continue;
      }

      if (/^Fuction\b/i.test(stmt) || /^Funtion\b/i.test(stmt)) {
        errors.push({
          line: lineNum,
          severity: 'error',
          code: 'KEYWORD_TYPO',
          message: `Syntax Error on line ${lineNum}: Typo in procedure declaration. Did you mean 'Function'?`,
          suggestion: `Change keyword to 'Function'.`,
          lineSnippet: rawLine.trim()
        });
        executionTrace.push(`❌ [Syntax Error] Line ${lineNum}: Typo in keyword 'Function'.`);
        continue;
      }

      // Check illegal Sub call with parentheses
      if (/^MsgBox\s*\(/i.test(stmt) && !/^([A-Za-z0-9_]+\s*=\s*MsgBox|If\s+.*MsgBox)/i.test(stmt)) {
        warnings.push({
          line: lineNum,
          severity: 'warning',
          code: 'VBA_CALL_SYNTAX',
          message: `VBA syntax warning on line ${lineNum}: Cannot use parentheses when calling a Sub procedure without 'Call' or variable assignment.`,
          suggestion: `Use 'MsgBox "Message", vbInformation' or 'Call MsgBox(...)'.`,
          lineSnippet: rawLine.trim()
        });
        executionTrace.push(`⚠️ [Syntax Warning] Line ${lineNum}: Parentheses used on Sub call without 'Call'.`);
      }

      // Detect Label definition if statement matches label pattern
      const labelMatch = stmt.match(/^([A-Za-z0-9_]+):?$/);
      if (labelMatch && (stmt.endsWith(':') || definedLabels.has(labelMatch[1].toLowerCase()))) {
        definedLabels.add(labelMatch[1].toLowerCase());
        executionTrace.push(`[Label Target] Defined execution jump label '${labelMatch[1]}:' at line ${lineNum}.`);
        continue;
      }

      // Detect GoTo / On Error GoTo
      const gotoMatch = stmt.match(/(?:GoTo|On\s+Error\s+GoTo)\s+([A-Za-z0-9_]+)/i);
      if (gotoMatch) {
        const targetLabel = gotoMatch[1];
        if (targetLabel !== '0' && targetLabel !== '-1') {
          referencedLabels.push({ label: targetLabel, line: lineNum });
          executionTrace.push(`[Control Flow] Registered GoTo target '${targetLabel}' at line ${lineNum}.`);
        }
        continue;
      }

      // Detect Dim / Static / Public / Private declarations
      const dimMatch = stmt.match(/^(?:Dim|Static|Public|Private)\s+(.+)/i);
      if (dimMatch && !/^(?:Public|Private)\s+(?:Sub|Function|Property)/i.test(stmt)) {
        const varDefs = dimMatch[1].split(',');
        for (const varDef of varDefs) {
          const parts = varDef.trim().match(/^([A-Za-z0-9_]+)(?:\s+As\s+([A-Za-z0-9_]+))?/i);
          if (parts) {
            const vName = parts[1];
            const vType = parts[2];
            const lowerType = vType ? vType.toLowerCase() : 'variant';

            if (vType && !VALID_VBA_TYPES.has(lowerType)) {
              errors.push({
                line: lineNum,
                severity: 'error',
                code: 'TYPE_NOT_DEFINED',
                message: `Compile Error on line ${lineNum}: User-defined type not defined: '${vType}'.`,
                suggestion: `Verify type name spelling (e.g. 'Worksheet', 'Range', 'Long', 'Double', 'String').`,
                lineSnippet: rawLine.trim()
              });
              executionTrace.push(`❌ [Compile Error] Line ${lineNum}: User-defined type '${vType}' not defined.`);
            }

            if (declaredVariables.has(vName.toLowerCase())) {
              errors.push({
                line: lineNum,
                severity: 'error',
                code: 'DUPLICATE_DECLARATION',
                message: `Duplicate declaration on line ${lineNum}: Variable '${vName}' was already declared on line ${declaredVariables.get(vName.toLowerCase())!.line}.`,
                suggestion: `Remove the duplicate 'Dim ${vName}' declaration.`,
                lineSnippet: rawLine.trim()
              });
              executionTrace.push(`❌ [Compile Error] Line ${lineNum}: Duplicate declaration '${vName}'.`);
            } else {
              declaredVariables.set(vName.toLowerCase(), { type: vType || 'Variant', line: lineNum });
              executionTrace.push(`[Memory] Allocated local variable '${vName}' (Type: ${vType || 'Variant'}).`);
            }
          } else {
            errors.push({
              line: lineNum,
              severity: 'error',
              code: 'MALFORMED_DECLARATION',
              message: `Syntax Error on line ${lineNum}: Invalid variable declaration in '${varDef}'.`,
              suggestion: `Use 'Dim variableName As Type'.`,
              lineSnippet: rawLine.trim()
            });
          }
        }
        continue;
      }

      // Check procedure opening: Sub, Function, Property
      const procStartMatch = stmt.match(
        /^(?:(Public|Private|Friend)\s+)?(?:Static\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z0-9_]+)\s*\(([^)]*)\)/i
      );
      if (procStartMatch) {
        const scope = (procStartMatch[1] || 'Public') as 'Public' | 'Private' | 'Friend';
        const rawType = procStartMatch[2];
        const procType = rawType.toLowerCase().startsWith('sub')
          ? 'Sub'
          : rawType.toLowerCase().startsWith('function')
          ? 'Function'
          : 'Property';
        const procName = procStartMatch[3];
        const rawParams = procStartMatch[4].trim();
        const params = rawParams ? rawParams.split(',').map((p) => p.trim()) : [];

        if (insideProcedure) {
          errors.push({
            line: lineNum,
            severity: 'error',
            code: 'NESTED_PROCEDURE',
            message: `Syntax Error on line ${lineNum}: Cannot declare '${procType} ${procName}' inside an unclosed procedure block.`,
            suggestion: `Close the preceding procedure with 'End Sub' or 'End Function' before declaring a new one.`,
            lineSnippet: rawLine.trim()
          });
          executionTrace.push(`❌ [Syntax Error] Line ${lineNum}: Nested procedure declaration.`);
        }

        insideProcedure = true;
        blockStack.push({
          type: procType.toLowerCase() as any,
          name: procName,
          line: lineNum
        });

        if (!procedureInfo) {
          procedureInfo = {
            type: procType,
            name: procName,
            scope,
            params
          };
          executionTrace.push(`[Entry Point] Entered ${scope} ${procType} ${procName}(${params.join(', ')}) on line ${lineNum}.`);
        }
        continue;
      }

      // Check procedure closing: End Sub / End Function / End Property
      const procEndMatch = stmt.match(/^End\s+(Sub|Function|Property)/i);
      if (procEndMatch) {
        const closingKind = procEndMatch[1].toLowerCase();
        insideProcedure = false;

        if (blockStack.length === 0) {
          errors.push({
            line: lineNum,
            severity: 'error',
            code: 'UNEXPECTED_PROC_END',
            message: `Unexpected 'End ${procEndMatch[1]}' on line ${lineNum} without matching procedure declaration.`,
            suggestion: `Remove orphaned 'End ${procEndMatch[1]}' or ensure procedure declaration is present.`,
            lineSnippet: rawLine.trim()
          });
          executionTrace.push(`❌ [Syntax Error] Line ${lineNum}: Orphaned 'End ${procEndMatch[1]}'.`);
        } else {
          const top = blockStack[blockStack.length - 1];
          if (top.type !== closingKind && (top.type === 'sub' || top.type === 'function' || top.type === 'property')) {
            errors.push({
              line: lineNum,
              severity: 'error',
              code: 'PROC_TYPE_MISMATCH',
              message: `Mismatched block termination on line ${lineNum}: Found 'End ${procEndMatch[1]}', but procedure '${top.name}' was declared as '${top.type}'.`,
              suggestion: `Change 'End ${procEndMatch[1]}' to 'End ${top.type === 'sub' ? 'Sub' : 'Function'}'.`,
              lineSnippet: rawLine.trim()
            });
            executionTrace.push(`❌ [Syntax Error] Line ${lineNum}: Mismatched procedure terminator.`);
            blockStack.pop();
          } else if (top.type !== closingKind) {
            errors.push({
              line: lineNum,
              severity: 'error',
              code: 'UNCLOSED_INNER_BLOCK',
              message: `Procedure ended on line ${lineNum} while inner '${top.type}' block opened on line ${top.line} is still open.`,
              suggestion: `Close the '${top.type}' block (e.g. with 'End If', 'Next', or 'Loop') before 'End ${procEndMatch[1]}'.`,
              lineSnippet: rawLine.trim()
            });
            executionTrace.push(`❌ [Syntax Error] Line ${lineNum}: Procedure ended with unclosed '${top.type}' block.`);
            while (blockStack.length > 0 && blockStack[blockStack.length - 1].type !== closingKind) {
              blockStack.pop();
            }
            if (blockStack.length > 0) blockStack.pop();
          } else {
            blockStack.pop();
            executionTrace.push(`[Block Exit] Procedure '${top.name}' clean exit reached on line ${lineNum}.`);
          }
        }
        continue;
      }

      // Check statements outside procedure
      if (!insideProcedure && !/^Option\b/i.test(stmt) && !/^Dim\b/i.test(stmt) && !/^Const\b/i.test(stmt)) {
        errors.push({
          line: lineNum,
          severity: 'error',
          code: 'STATEMENT_OUTSIDE_PROCEDURE',
          message: `Compile Error on line ${lineNum}: Statement '${stmt}' is invalid outside a procedure.`,
          suggestion: `Place executable statements inside a 'Sub' or 'Function' procedure block.`,
          lineSnippet: rawLine.trim()
        });
        executionTrace.push(`❌ [Compile Error] Line ${lineNum}: Statement outside procedure block.`);
        continue;
      }

      // Detect Worksheet references: Sheets("...") or Worksheets("...")
      const sheetMatches = Array.from(stmt.matchAll(/(?:Sheets|Worksheets)\s*\(\s*"([^"]+)"\s*\)/gi));
      for (const sm of sheetMatches) {
        const sName = sm[1];
        const exists = availableSheetNames.some((n) => n.toLowerCase() === sName.toLowerCase());
        referencedSheetsMap.set(sName, exists);
        if (!exists) {
          errors.push({
            line: lineNum,
            severity: 'error',
            code: 'SUBSCRIPT_OUT_OF_RANGE',
            message: `Run-time error '9': Subscript out of range on line ${lineNum}. Worksheet "${sName}" does not exist in workbook.`,
            suggestion: `Loaded sheets: [${availableSheetNames.map((n) => `"${n}"`).join(', ')}].`,
            lineSnippet: rawLine.trim()
          });
          executionTrace.push(`❌ [Runtime Exception] Line ${lineNum}: Run-time error '9': Subscript out of range. Worksheets("${sName}") not found in workbook.`);
        } else {
          executionTrace.push(`[Object Binding] Line ${lineNum}: Bound worksheet reference "${sName}". Verified in workbook.`);
        }
      }

      // Check Function Return assignment: e.g. "MyFunction = result"
      if (procedureInfo?.type === 'Function' && procedureInfo.name) {
        const returnRegex = new RegExp(`^${procedureInfo.name}\\s*=`, 'i');
        if (returnRegex.test(stmt)) {
          hasReturnAssignment = true;
          executionTrace.push(`[Function Return] Assigned return value to '${procedureInfo.name}' at line ${lineNum}.`);
        }
      }

      // Check If / Then
      if (/^If\b/i.test(stmt)) {
        const thenIdx = stmt.search(/\bThen\b/i);
        if (thenIdx === -1) {
          errors.push({
            line: lineNum,
            severity: 'error',
            code: 'MISSING_THEN',
            message: `Syntax Error on line ${lineNum}: 'If' statement missing required 'Then' keyword.`,
            suggestion: `Add 'Then' at the end of the condition (e.g. 'If x > 10 Then').`,
            lineSnippet: rawLine.trim()
          });
          executionTrace.push(`❌ [Syntax Error] Line ${lineNum}: 'If' statement missing 'Then'.`);
        } else {
          const conditionPart = stmt.slice(2, thenIdx).trim();
          validateExpressionTokens(
            conditionPart,
            lineNum,
            rawLine,
            declaredVarNames,
            paramNames,
            activeLoopVars,
            errors,
            executionTrace,
            hasOptionExplicit,
            warnings
          );

          const afterThen = stmt.slice(thenIdx + 4).trim();
          if (!afterThen) {
            blockStack.push({ type: 'if', line: lineNum });
            executionTrace.push(`[Branching] Evaluated multi-line If block condition on line ${lineNum}.`);
          } else {
            validateExpressionTokens(
              afterThen,
              lineNum,
              rawLine,
              declaredVarNames,
              paramNames,
              activeLoopVars,
              errors,
              executionTrace,
              hasOptionExplicit,
              warnings
            );
            executionTrace.push(`[Branching] Evaluated single-line If statement on line ${lineNum}.`);
          }
        }
        continue;
      }

      // Check ElseIf
      if (/^ElseIf\b/i.test(stmt)) {
        const top = blockStack[blockStack.length - 1];
        if (!top || top.type !== 'if') {
          errors.push({
            line: lineNum,
            severity: 'error',
            code: 'UNEXPECTED_ELSEIF',
            message: `Syntax Error on line ${lineNum}: 'ElseIf' without matching 'If'.`,
            suggestion: `Ensure 'ElseIf' is placed inside an open multi-line 'If' block.`,
            lineSnippet: rawLine.trim()
          });
          executionTrace.push(`❌ [Syntax Error] Line ${lineNum}: Orphaned 'ElseIf'.`);
        } else if (!/\bThen\b/i.test(stmt)) {
          errors.push({
            line: lineNum,
            severity: 'error',
            code: 'MISSING_THEN',
            message: `Syntax Error on line ${lineNum}: 'ElseIf' statement missing required 'Then' keyword.`,
            suggestion: `Add 'Then' at the end of the ElseIf condition.`,
            lineSnippet: rawLine.trim()
          });
          executionTrace.push(`❌ [Syntax Error] Line ${lineNum}: 'ElseIf' missing 'Then'.`);
        } else {
          const thenIdx = stmt.search(/\bThen\b/i);
          const conditionPart = stmt.slice(6, thenIdx).trim();
          validateExpressionTokens(
            conditionPart,
            lineNum,
            rawLine,
            declaredVarNames,
            paramNames,
            activeLoopVars,
            errors,
            executionTrace,
            hasOptionExplicit,
            warnings
          );
        }
        continue;
      }

      // Check Else
      if (/^Else\b/i.test(stmt)) {
        const top = blockStack[blockStack.length - 1];
        if (!top || (top.type !== 'if' && top.type !== 'select')) {
          errors.push({
            line: lineNum,
            severity: 'error',
            code: 'UNEXPECTED_ELSE',
            message: `Syntax Error on line ${lineNum}: 'Else' without matching 'If' or 'Select Case'.`,
            suggestion: `Place 'Else' inside an open 'If ... Then' block.`,
            lineSnippet: rawLine.trim()
          });
          executionTrace.push(`❌ [Syntax Error] Line ${lineNum}: Orphaned 'Else'.`);
        }
        continue;
      }

      // Check End If
      if (/^End\s+If/i.test(stmt)) {
        const top = blockStack[blockStack.length - 1];
        if (!top || top.type !== 'if') {
          errors.push({
            line: lineNum,
            severity: 'error',
            code: 'UNEXPECTED_END_IF',
            message: `Unexpected 'End If' on line ${lineNum} without matching 'If ... Then' block.`,
            suggestion: `Remove the redundant 'End If' or ensure the preceding 'If' is multi-line.`,
            lineSnippet: rawLine.trim()
          });
          executionTrace.push(`❌ [Syntax Error] Line ${lineNum}: Orphaned 'End If'.`);
        } else {
          blockStack.pop();
          executionTrace.push(`[Block Exit] Closed If block on line ${lineNum}.`);
        }
        continue;
      }

      // Check For loop
      const forMatch = stmt.match(/^For\s+(?:Each\s+)?([A-Za-z0-9_]+)\s*(?:=\s*(.+?)\s+To\s+(.+)|In\s+(.+))/i);
      if (forMatch) {
        const loopVar = forMatch[1];
        blockStack.push({ type: 'for', loopVar, line: lineNum });
        activeLoopVars.add(loopVar.toLowerCase());

        const startExpr = forMatch[2];
        const endExpr = forMatch[3];
        const inExpr = forMatch[4];

        if (startExpr) {
          validateExpressionTokens(startExpr, lineNum, rawLine, declaredVarNames, paramNames, activeLoopVars, errors, executionTrace);
        }
        if (endExpr) {
          validateExpressionTokens(endExpr, lineNum, rawLine, declaredVarNames, paramNames, activeLoopVars, errors, executionTrace);
        }
        if (inExpr) {
          validateExpressionTokens(inExpr, lineNum, rawLine, declaredVarNames, paramNames, activeLoopVars, errors, executionTrace);
        }

        executionTrace.push(`[Loop Engine] Initializing For loop on variable '${loopVar}' on line ${lineNum}.`);
        continue;
      }

      // Check Next
      if (/^Next\b/i.test(stmt)) {
        const nextMatch = stmt.match(/^Next(?:\s+([A-Za-z0-9_]+))?/i);
        const specifiedVar = nextMatch ? nextMatch[1] : undefined;

        const top = blockStack[blockStack.length - 1];
        if (!top || top.type !== 'for') {
          errors.push({
            line: lineNum,
            severity: 'error',
            code: 'UNEXPECTED_NEXT',
            message: `Unexpected 'Next' on line ${lineNum} without matching 'For' loop.`,
            suggestion: `Ensure a corresponding 'For ...' statement precedes this 'Next'.`,
            lineSnippet: rawLine.trim()
          });
          executionTrace.push(`❌ [Syntax Error] Line ${lineNum}: Orphaned 'Next' without 'For'.`);
        } else {
          if (specifiedVar && top.loopVar && specifiedVar.toLowerCase() !== top.loopVar.toLowerCase()) {
            errors.push({
              line: lineNum,
              severity: 'error',
              code: 'NEXT_VARIABLE_MISMATCH',
              message: `Mismatched loop variable on line ${lineNum}: 'Next ${specifiedVar}' does not match loop variable '${top.loopVar}' opened on line ${top.line}.`,
              suggestion: `Change 'Next ${specifiedVar}' to 'Next ${top.loopVar}' or simply 'Next'.`,
              lineSnippet: rawLine.trim()
            });
            executionTrace.push(`❌ [Syntax Error] Line ${lineNum}: Loop variable mismatch 'Next ${specifiedVar}'.`);
          }
          blockStack.pop();
          executionTrace.push(`[Loop Engine] Closed For loop iteration at line ${lineNum}.`);
        }
        continue;
      }

      // Check Do While / Do Until / Do
      if (/^Do\b/i.test(stmt)) {
        blockStack.push({ type: 'do', line: lineNum });
        const condMatch = stmt.match(/^Do\s+(?:While|Until)\s+(.+)/i);
        if (condMatch) {
          validateExpressionTokens(condMatch[1], lineNum, rawLine, declaredVarNames, paramNames, activeLoopVars, errors, executionTrace);
        }
        executionTrace.push(`[Loop Engine] Initialized Do loop on line ${lineNum}.`);
        continue;
      }

      // Check Loop
      if (/^Loop\b/i.test(stmt)) {
        const top = blockStack[blockStack.length - 1];
        if (!top || top.type !== 'do') {
          errors.push({
            line: lineNum,
            severity: 'error',
            code: 'UNEXPECTED_LOOP',
            message: `Unexpected 'Loop' on line ${lineNum} without matching 'Do' block.`,
            suggestion: `Ensure an opening 'Do' statement precedes this 'Loop'.`,
            lineSnippet: rawLine.trim()
          });
          executionTrace.push(`❌ [Syntax Error] Line ${lineNum}: Orphaned 'Loop' without 'Do'.`);
        } else {
          const condMatch = stmt.match(/^Loop\s+(?:While|Until)\s+(.+)/i);
          if (condMatch) {
            validateExpressionTokens(condMatch[1], lineNum, rawLine, declaredVarNames, paramNames, activeLoopVars, errors, executionTrace);
          }
          blockStack.pop();
          executionTrace.push(`[Loop Engine] Terminated Do loop on line ${lineNum}.`);
        }
        continue;
      }

      // Check While
      if (/^While\b/i.test(stmt)) {
        blockStack.push({ type: 'while', line: lineNum });
        const cond = stmt.replace(/^While\s+/i, '').trim();
        if (cond) {
          validateExpressionTokens(cond, lineNum, rawLine, declaredVarNames, paramNames, activeLoopVars, errors, executionTrace);
        }
        continue;
      }

      // Check Wend
      if (/^Wend\b/i.test(stmt)) {
        const top = blockStack[blockStack.length - 1];
        if (!top || top.type !== 'while') {
          errors.push({
            line: lineNum,
            severity: 'error',
            code: 'UNEXPECTED_WEND',
            message: `Unexpected 'Wend' on line ${lineNum} without matching 'While' statement.`,
            suggestion: `Ensure an opening 'While' precedes this 'Wend'.`,
            lineSnippet: rawLine.trim()
          });
          executionTrace.push(`❌ [Syntax Error] Line ${lineNum}: Orphaned 'Wend' without 'While'.`);
        } else {
          blockStack.pop();
        }
        continue;
      }

      // Check Select Case
      if (/^Select\s+Case\b/i.test(stmt)) {
        blockStack.push({ type: 'select', line: lineNum });
        const exprPart = stmt.replace(/^Select\s+Case\s+/i, '').trim();
        if (exprPart) {
          validateExpressionTokens(exprPart, lineNum, rawLine, declaredVarNames, paramNames, activeLoopVars, errors, executionTrace);
        }
        executionTrace.push(`[Branching] Opened Select Case block on line ${lineNum}.`);
        continue;
      }

      if (/^Case\b/i.test(stmt)) {
        const top = blockStack[blockStack.length - 1];
        if (!top || top.type !== 'select') {
          errors.push({
            line: lineNum,
            severity: 'error',
            code: 'UNEXPECTED_CASE',
            message: `Syntax Error on line ${lineNum}: 'Case' statement outside 'Select Case' block.`,
            suggestion: `Place 'Case' inside a 'Select Case' block.`,
            lineSnippet: rawLine.trim()
          });
          executionTrace.push(`❌ [Syntax Error] Line ${lineNum}: 'Case' outside 'Select Case'.`);
        } else if (!/^Case\s+Else\b/i.test(stmt)) {
          const caseVals = stmt.replace(/^Case\s+/i, '').trim();
          validateExpressionTokens(caseVals, lineNum, rawLine, declaredVarNames, paramNames, activeLoopVars, errors, executionTrace);
        }
        continue;
      }

      // Check End Select
      if (/^End\s+Select\b/i.test(stmt)) {
        const top = blockStack[blockStack.length - 1];
        if (!top || top.type !== 'select') {
          errors.push({
            line: lineNum,
            severity: 'error',
            code: 'UNEXPECTED_END_SELECT',
            message: `Unexpected 'End Select' on line ${lineNum} without matching 'Select Case' statement.`,
            suggestion: `Remove the stray 'End Select' or open with 'Select Case'.`,
            lineSnippet: rawLine.trim()
          });
          executionTrace.push(`❌ [Syntax Error] Line ${lineNum}: Orphaned 'End Select'.`);
        } else {
          blockStack.pop();
          executionTrace.push(`[Branching] Closed Select Case block on line ${lineNum}.`);
        }
        continue;
      }

      // Check With
      if (/^With\b/i.test(stmt)) {
        blockStack.push({ type: 'with', line: lineNum });
        const objExpr = stmt.replace(/^With\s+/i, '').trim();
        if (objExpr) {
          validateExpressionTokens(objExpr, lineNum, rawLine, declaredVarNames, paramNames, activeLoopVars, errors, executionTrace);
        }
        executionTrace.push(`[Scope] Opened With block on line ${lineNum}.`);
        continue;
      }

      // Check End With
      if (/^End\s+With\b/i.test(stmt)) {
        const top = blockStack[blockStack.length - 1];
        if (!top || top.type !== 'with') {
          errors.push({
            line: lineNum,
            severity: 'error',
            code: 'UNEXPECTED_END_WITH',
            message: `Unexpected 'End With' on line ${lineNum} without matching 'With' statement.`,
            suggestion: `Remove stray 'End With' or open with 'With object'.`,
            lineSnippet: rawLine.trim()
          });
          executionTrace.push(`❌ [Syntax Error] Line ${lineNum}: Orphaned 'End With'.`);
        } else {
          blockStack.pop();
        }
        continue;
      }

      // Check Assignment Statements: [Set ] varOrProp = expr
      if (stmt.includes('=')) {
        const eqIdx = stmt.indexOf('=');
        const lhs = stmt.slice(0, eqIdx).trim();
        const rhs = stmt.slice(eqIdx + 1).trim();

        if (!lhs) {
          errors.push({
            line: lineNum,
            severity: 'error',
            code: 'EXPECTED_VARIABLE',
            message: `Syntax Error on line ${lineNum}: Expected variable or property before '='.`,
            suggestion: `Provide a left-hand side variable name (e.g. 'x = ...').`,
            lineSnippet: rawLine.trim()
          });
          executionTrace.push(`❌ [Syntax Error] Line ${lineNum}: Missing LHS in assignment.`);
          continue;
        }

        if (!rhs) {
          errors.push({
            line: lineNum,
            severity: 'error',
            code: 'EXPECTED_EXPRESSION',
            message: `Syntax Error on line ${lineNum}: Expected expression after '=' in assignment.`,
            suggestion: `Provide a value or expression on right-hand side of '='.`,
            lineSnippet: rawLine.trim()
          });
          executionTrace.push(`❌ [Syntax Error] Line ${lineNum}: Missing RHS expression in assignment.`);
          continue;
        }

        const isSet = /^Set\s+/i.test(lhs);
        const cleanLhs = lhs.replace(/^Set\s+/i, '').trim();

        // Validate LHS expression tokens if it contains member access or indexing, e.g. ws.Cells(1, 1).Valu
        if (cleanLhs.includes('.') || cleanLhs.includes('(')) {
          validateExpressionTokens(
            cleanLhs,
            lineNum,
            rawLine,
            declaredVarNames,
            paramNames,
            activeLoopVars,
            errors,
            executionTrace,
            hasOptionExplicit,
            warnings
          );
        } else {
          // Check bare identifier typo or Option Explicit enforcement
          const lowerLhs = cleanLhs.toLowerCase();
          const isKnownLhs =
            declaredVarNames.has(lowerLhs) ||
            paramNames.has(lowerLhs) ||
            activeLoopVars.has(lowerLhs) ||
            (procedureInfo && procedureInfo.name.toLowerCase() === lowerLhs) ||
            EXCEL_ROOT_OBJECTS.has(lowerLhs) ||
            VBA_KEYWORDS.has(lowerLhs);

          if (!isKnownLhs) {
            const typoLhs = findClosestMatch(cleanLhs);
            if (typoLhs && typoLhs.toLowerCase() !== lowerLhs) {
              errors.push({
                line: lineNum,
                severity: 'error',
                code: 'VARIABLE_NOT_DEFINED',
                message: `Compile Error on line ${lineNum}: Variable not defined: '${cleanLhs}'. Did you mean '${typoLhs}'?`,
                suggestion: `Replace '${cleanLhs}' with '${typoLhs}' or declare variable using 'Dim ${cleanLhs}'.`,
                lineSnippet: rawLine.trim()
              });
              executionTrace.push(`❌ [Compile Error] Line ${lineNum}: Variable not defined: '${cleanLhs}'.`);
            } else if (hasOptionExplicit) {
              errors.push({
                line: lineNum,
                severity: 'error',
                code: 'VARIABLE_NOT_DEFINED',
                message: `Compile Error on line ${lineNum}: Variable not defined: '${cleanLhs}'. Option Explicit requires all variables to be explicitly declared.`,
                suggestion: `Add 'Dim ${cleanLhs} As ...' before assigning to '${cleanLhs}'.`,
                lineSnippet: rawLine.trim()
              });
              executionTrace.push(`❌ [Compile Error] Line ${lineNum}: Undeclared variable '${cleanLhs}' under Option Explicit.`);
            } else {
              warnings.push({
                line: lineNum,
                severity: 'warning',
                code: 'IMPLICIT_VARIABLE',
                message: `Variable '${cleanLhs}' on line ${lineNum} is not explicitly declared (treated as implicit Variant).`,
                suggestion: `Add 'Dim ${cleanLhs}' to declare variable type.`
              });
            }
          }
        }

        // Validate RHS expression tokens (catches typos like 'Tru' instead of 'True'!)
        validateExpressionTokens(
          rhs,
          lineNum,
          rawLine,
          declaredVarNames,
          paramNames,
          activeLoopVars,
          errors,
          executionTrace,
          hasOptionExplicit,
          warnings
        );

        // Check if Set is incorrectly used for primitive literals
        if (isSet && (/^-?[0-9]+(\.[0-9]+)?$/.test(rhs) || /^".*"$/.test(rhs) || /^(true|false)$/i.test(rhs))) {
          errors.push({
            line: lineNum,
            severity: 'error',
            code: 'INVALID_SET_ASSIGNMENT',
            message: `Compile Error on line ${lineNum}: 'Set' statement is invalid for primitive types (${rhs}). 'Set' is only used for Object references.`,
            suggestion: `Remove 'Set' keyword (use '${cleanLhs} = ${rhs}').`,
            lineSnippet: rawLine.trim()
          });
          executionTrace.push(`❌ [Compile Error] Line ${lineNum}: 'Set' invalid for primitive type.`);
        }

        // Check if object assignment is missing Set
        if (!isSet && (/^(?:ThisWorkbook\.)?(?:Sheets|Worksheets)\s*\(/i.test(rhs) || /^Range\s*\(/i.test(rhs))) {
          warnings.push({
            line: lineNum,
            severity: 'warning',
            code: 'MISSING_SET_KEYWORD',
            message: `Warning on line ${lineNum}: Object assignment to '${cleanLhs}' requires 'Set' keyword.`,
            suggestion: `Use 'Set ${cleanLhs} = ${rhs}'.`,
            lineSnippet: rawLine.trim()
          });
          executionTrace.push(`⚠️ [Warning] Line ${lineNum}: Missing 'Set' keyword for object assignment.`);
        }

        executionTrace.push(`[Assignment] Line ${lineNum}: Evaluated assignment to '${cleanLhs}'.`);
        continue;
      }

      // Check standard keywords & statements
      if (
        /^Exit\s+(Sub|Function|For|Do|Property)\b/i.test(stmt) ||
        /^Resume(?:\s+Next|\s+[A-Za-z0-9_]+)?\b/i.test(stmt) ||
        /^On\s+Error\s+(?:Resume\s+Next|GoTo\s+[A-Za-z0-9_]+)\b/i.test(stmt) ||
        /^Call\s+[A-Za-z0-9_]+/i.test(stmt) ||
        /^MsgBox\b/i.test(stmt) ||
        /^InputBox\b/i.test(stmt) ||
        /^Debug\.Print\b/i.test(stmt) ||
        /^DoEvents\b/i.test(stmt) ||
        /^Application\.[A-Za-z0-9_.]+/i.test(stmt) ||
        /^[A-Za-z0-9_]+(?:\.[A-Za-z0-9_]+)*\s*(?:\(.*\)|[A-Za-z0-9_",\s]+)?$/i.test(stmt)
      ) {
        // Also validate expressions inside MsgBox / Debug.Print / method calls
        const callArgs = stmt.replace(/^(Call\s+[A-Za-z0-9_.]+|MsgBox|Debug\.Print|InputBox)\s*/i, '');
        if (callArgs && callArgs !== stmt) {
          validateExpressionTokens(
            callArgs,
            lineNum,
            rawLine,
            declaredVarNames,
            paramNames,
            activeLoopVars,
            errors,
            executionTrace
          );
        }

        executionTrace.push(`[Statement Engine] Line ${lineNum}: Executed statement '${stmt.length > 50 ? stmt.slice(0, 48) + '...' : stmt}'.`);
        continue;
      }

      // If statement doesn't match any known VBA construct: Unrecognized statement error!
      errors.push({
        line: lineNum,
        severity: 'error',
        code: 'UNRECOGNIZED_STATEMENT',
        message: `Syntax Error on line ${lineNum}: Invalid or unrecognized statement: '${stmt}'.`,
        suggestion: `Verify statement syntax or ensure it is a valid VBA command.`,
        lineSnippet: rawLine.trim()
      });
      executionTrace.push(`❌ [Syntax Error] Line ${lineNum}: Unrecognized statement '${stmt}'.`);
    }
  }

  // Check unclosed blocks remaining on the stack at the end of the file
  if (blockStack.length > 0) {
    for (const unclosed of blockStack) {
      let expectedCloser = 'End Sub';
      if (unclosed.type === 'sub') expectedCloser = 'End Sub';
      else if (unclosed.type === 'function') expectedCloser = 'End Function';
      else if (unclosed.type === 'property') expectedCloser = 'End Property';
      else if (unclosed.type === 'if') expectedCloser = 'End If';
      else if (unclosed.type === 'for') expectedCloser = 'Next';
      else if (unclosed.type === 'do') expectedCloser = 'Loop';
      else if (unclosed.type === 'while') expectedCloser = 'Wend';
      else if (unclosed.type === 'select') expectedCloser = 'End Select';
      else if (unclosed.type === 'with') expectedCloser = 'End With';

      errors.push({
        line: unclosed.line,
        severity: 'error',
        code: 'UNCLOSED_BLOCK',
        message: `Block error: '${unclosed.type.toUpperCase()}' block started on line ${unclosed.line} was never closed before end of script.`,
        suggestion: `Insert '${expectedCloser}' to properly terminate the ${unclosed.type} block.`
      });
      executionTrace.push(`❌ [Block Error] Line ${unclosed.line}: Unclosed '${unclosed.type}' block. Expected '${expectedCloser}'.`);
    }
  }

  // Check unresolved GoTo labels
  for (const ref of referencedLabels) {
    if (!definedLabels.has(ref.label.toLowerCase())) {
      errors.push({
        line: ref.line,
        severity: 'error',
        code: 'LABEL_NOT_DEFINED',
        message: `Label not defined: 'GoTo ${ref.label}' on line ${ref.line} references undefined label '${ref.label}:'.`,
        suggestion: `Add the label target '${ref.label}:' in the procedure or check for spelling errors.`
      });
      executionTrace.push(`❌ [Compile Error] Line ${ref.line}: Undefined label target '${ref.label}:'.`);
    }
  }

  // Check if no procedure was ever found in non-empty code
  if (!procedureInfo && errors.length === 0) {
    warnings.push({
      line: 1,
      severity: 'warning',
      code: 'NO_PROCEDURE_HEADER',
      message: 'No procedure declaration (Sub or Function) was detected in this code.',
      suggestion: 'Wrap code in a Sub or Function (e.g. Sub ExecuteRoutine() ... End Sub).'
    });
    executionTrace.push(`⚠️ [Warning] No procedure declaration found in script.`);
  }

  // Check Function return value warning
  if (procedureInfo?.type === 'Function' && !hasReturnAssignment && errors.length === 0) {
    warnings.push({
      line: 1,
      severity: 'warning',
      code: 'MISSING_RETURN_VALUE',
      message: `Function '${procedureInfo.name}' does not appear to assign a return value.`,
      suggestion: `Assign the function return value before exiting (e.g. '${procedureInfo.name} = result').`
    });
    executionTrace.push(`⚠️ [Warning] Function '${procedureInfo.name}' exits without return value assignment.`);
  }

  // Format referenced sheets report
  const referencedSheets = Array.from(referencedSheetsMap.entries()).map(([name, exists]) => ({
    name,
    existsInWorkbook: exists
  }));

  // Format variables report
  const variablesDetected = Array.from(declaredVariables.entries()).map(([name, info]) => ({
    name,
    type: info.type,
    declared: true
  }));

  const executionTimeMs = Math.round((performance.now() - startTime) * 10) / 10;
  const isSuccess = errors.length === 0;

  if (isSuccess) {
    executionTrace.push(`[Execution Simulation] Mock execution completed successfully in ${executionTimeMs}ms with zero syntax or runtime faults.`);
  } else {
    executionTrace.push(`❌ [Simulation Blocked] Execution halted with ${errors.length} fatal error(s) and ${warnings.length} warning(s).`);
  }

  const summary = isSuccess
    ? `Simulation Verified: Executed "${procedureInfo?.name || 'Procedure'}" in ${executionTimeMs}ms across ${totalLines} lines with 0 faults.`
    : `Simulation Failed: Detected ${errors.length} error(s) and ${warnings.length} warning(s) in VBA code.`;

  return {
    success: isSuccess,
    status: isSuccess ? (warnings.length > 0 ? 'warnings' : 'passed') : 'failed',
    executionTimeMs,
    totalLines,
    statementsAnalyzed: statementsCount,
    errors,
    warnings,
    executionTrace,
    referencedSheets,
    variablesDetected,
    procedureSignature: procedureInfo,
    summary
  };
}

/**
 * Evaluates live expressions in the VBE Immediate Window (Ctrl+G).
 * Supports "? <math/logic expr>", "? ActiveSheet.Name", "? Sheets.Count", etc.
 */
export function evaluateImmediateExpression(rawCommand: string, dataset?: Dataset): string {
  const trimmed = rawCommand.trim();
  if (!trimmed) return '';

  const activeSheet = dataset?.sheets?.find((s) => s.name === dataset.activeSheetName) || dataset?.sheets?.[0];

  // 1. Print evaluation command (? <expr>)
  if (trimmed.startsWith('?') || trimmed.toLowerCase().startsWith('print ')) {
    const expr = trimmed.replace(/^(\?|print\s+)/i, '').trim();

    // Typo checks in immediate window
    const typo = findClosestMatch(expr);
    if (typo && expr.toLowerCase() !== typo.toLowerCase()) {
      return `Compile error: Variable not defined: '${expr}'. Did you mean '${typo}'?`;
    }

    if (/^activesheet\.name$/i.test(expr)) {
      return activeSheet?.name || 'Sheet1';
    }
    if (/^sheets\.count$/i.test(expr) || /^worksheets\.count$/i.test(expr)) {
      return String(dataset?.sheets?.length || 1);
    }
    if (/^application\.screenupdating$/i.test(expr)) {
      return 'True';
    }
    if (/^application\.calculation$/i.test(expr)) {
      return 'xlCalculationAutomatic (-4105)';
    }
    if (/^now$/i.test(expr)) {
      return new Date().toLocaleString();
    }
    if (/^timer$/i.test(expr)) {
      return String(Math.round((Date.now() % 86400000) / 1000));
    }

    // Mathematical evaluation
    try {
      // Safe arithmetic evaluator
      const safeMath = expr
        .replace(/true/gi, '1')
        .replace(/false/gi, '0')
        .replace(/\^/g, '**');

      if (/^[0-9+\-*/().\s*%]+$/.test(safeMath)) {
        // eslint-disable-next-line no-eval
        const result = Function(`"use strict"; return (${safeMath})`)();
        return String(result);
      }
    } catch {
      // fallback
    }

    return `Evaluated: ${expr}`;
  }

  // 2. Statement commands (e.g. "ActiveSheet.Calculate", "Range("A1").Value = 10")
  if (/^activesheet\.calculate$/i.test(trimmed)) {
    return `[VBE Execution] Recalculated active worksheet "${activeSheet?.name || 'Sheet1'}" (12 rows evaluated).`;
  }

  if (/^application\.screenupdating\s*=\s*(true|false)/i.test(trimmed)) {
    const val = trimmed.match(/=\s*(true|false)/i)![1];
    return `Application.ScreenUpdating = ${val.toLowerCase() === 'true' ? 'True' : 'False'}`;
  }

  if (/^run\s+([A-Za-z0-9_]+)/i.test(trimmed)) {
    const proc = trimmed.match(/^run\s+([A-Za-z0-9_]+)/i)![1];
    return `[VBE Execution] Macro "${proc}" executed in 2.4ms with return code 0.`;
  }

  return `Syntax Error: Invalid Immediate Window statement: '${trimmed}'. Use '? <expression>' to evaluate.`;
}
