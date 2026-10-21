/**
 * [非生产、仅规则镜像 / NON-PRODUCTION RULE MIRROR]
 * 本文件仅作为离线测试与客户端规则契约参考，生产环境实际执行拦截由 src/VbaRiskScanner.cs 保证。
 * 任何在此文件修改的规则，必须同步更新 src/VbaRiskScanner.cs，并通过 test_suite_unit.cjs 契约断言。
 * 
 * 等级划分：
 * - L0: 只读查询，不修改工作表与工作簿内容。无需快照，直接执行。
 * - L1: 普通原地修改（写单元格、改公式、设置样式等）。自动创建快照，可一键撤销。
 * - L2: 不可逆/越界高危操作（删除工作表、操作外部文件/工作簿、系统API等）。
 *       严格阻断静默执行，返回 needsConfirm 与风险理由，待用户确认后才可执行。
 */

export type VbaRiskLevel = 'L0' | 'L1' | 'L2';

export interface VbaRiskAnalysis {
  riskLevel: VbaRiskLevel;
  riskCode?: string;
  reasons: string[];
  targetDetail?: string;
}

function hasMultiStatementColon(line: string): boolean {
  let inQuote = false;
  for (let i = 0; i < line.length; i++) {
    const c = line[i];
    if (c === '"') {
      inQuote = !inQuote;
    } else if (!inQuote && c === ':') {
      if (i + 1 < line.length && line[i + 1] === '=') {
        i++;
        continue;
      }
      return true;
    }
  }
  return false;
}

export function classifyVbaRisk(vbaCode: string, targetWorkbookFullName?: string): VbaRiskAnalysis {
  if (!vbaCode || !vbaCode.trim()) {
    return {
      riskLevel: 'L0',
      reasons: ['代码为空']
    };
  }

  // 预处理：移除注释行，防止注释里的关键词误触发
  const codeLines = vbaCode.split('\n');
  const codeWithoutComments = codeLines
    .map(line => {
      const trimmed = line.trim();
      if (trimmed.startsWith("'") || trimmed.toLowerCase().startsWith('rem ')) {
        return '';
      }
      // 移除行尾注释 (简单去除单引号后内容，注意避免字面量内的单引号)
      const quoteIndex = line.indexOf("'");
      if (quoteIndex >= 0) {
        // 如果前面双引号数量是偶数，说明单引号在字符串外
        const prefix = line.substring(0, quoteIndex);
        const doubleQuotesCount = (prefix.match(/"/g) || []).length;
        if (doubleQuotesCount % 2 === 0) {
          return prefix;
        }
      }
      return line;
    })
    .join('\n');

  const l2Reasons: string[] = [];
  let riskCode: string | undefined;

  // 1. 删除工作表 (不可逆)
  if (/(?:Worksheets|Sheets)\s*\([^)]*\)\s*\.Delete/i.test(codeWithoutComments) ||
      /(?:ActiveSheet|ThisWorkbook\.ActiveSheet)\.Delete/i.test(codeWithoutComments)) {
    // 尝试提取被删除的工作表名
    const match = codeWithoutComments.match(/(?:Worksheets|Sheets)\s*\(\s*["']([^"']+)["']\s*\)\s*\.Delete/i);
    const sheetName = match ? match[1] : '工作表';
    l2Reasons.push(`包含删除工作表指令 (将删除工作表《${sheetName}》)`);
    riskCode = 'delete_sheet';
  }

  // 2. 外部工作簿操作 (Open / SaveAs / Close)
  if (/\bWorkbooks\s*\.\s*(?:Open|SaveAs|Close)\b/i.test(codeWithoutComments)) {
    l2Reasons.push('包含外部工作簿打开/另存/关闭操作 (Workbooks.Open/SaveAs/Close)');
    if (!riskCode) riskCode = 'external_workbook';
  }

  // 3. 高危外部 COM 对象 (FileSystemObject / 网络请求等)
  if (/CreateObject\s*\(\s*["'](?:Scripting\.FileSystemObject|MSXML2\.[^"']+|WinHttp\.[^"']+|InternetExplorer\.Application|WScript\.Shell)["']/i.test(codeWithoutComments) ||
      /New\s+FileSystemObject\b/i.test(codeWithoutComments)) {
    l2Reasons.push('包含文件系统或外部网络组件调用 (FileSystemObject/WinHttp/Shell)');
    if (!riskCode) riskCode = 'system_com';
  }

  // 4. 篡改代码工程 (VBProject)
  if (/\b(?:VBProject|ThisWorkbook\.VBProject)\b/i.test(codeWithoutComments)) {
    l2Reasons.push('包含对 VBA 代码工程自身的访问与修改 (VBProject)');
    if (!riskCode) riskCode = 'modify_vbproject';
  }

  // 5. 原生文件系统操作语句 (Kill / FileCopy / Name ... As ...)
  if (/\b(?:Kill|FileCopy|RmDir)\b/i.test(codeWithoutComments) ||
      /\bName\b\s+.+\s+\bAs\b/i.test(codeWithoutComments) ||
      /\bOpen\b\s+["'][^"']+["']\s+For\s+(?:Output|Append|Binary)/i.test(codeWithoutComments)) {
    l2Reasons.push('包含文件删除或外部磁盘写入语句 (Kill/FileCopy/Open For Output)');
    if (!riskCode) {
      riskCode = /\bKill\b/i.test(codeWithoutComments) ? 'delete_file' : 'file_io';
    }
  }

  // 命中任一 L2 规则即判定为 L2
  if (l2Reasons.length > 0) {
    let targetDetail = '不可逆操作';
    if (l2Reasons[0].includes('删除工作表')) {
      const m = l2Reasons[0].match(/《([^》]+)》/);
      targetDetail = m ? `将删除工作表《${m[1]}》` : '将删除工作表';
    } else if (l2Reasons[0].includes('文件')) {
      targetDetail = '涉及外部文件操作';
    }
    return {
      riskLevel: 'L2',
      riskCode: riskCode || 'dangerous_action',
      reasons: l2Reasons,
      targetDetail
    };
  }

  // 判定 L1: 普通写操作
  const hasWriteOperation = 
    /\.(?:Value|Value2|Formula|FormulaR1C1|Formula2)\s*=/i.test(codeWithoutComments) ||
    /\.(?:Interior|Font|Borders)\b/i.test(codeWithoutComments) ||
    /\.(?:Clear|ClearContents|ClearFormats)\b/i.test(codeWithoutComments) ||
    /(?:Rows|Columns)\s*\([^)]*\)\s*\.(?:Insert|Delete)/i.test(codeWithoutComments) ||
    /\.(?:NumberFormat|NumberFormatLocal)\s*=/i.test(codeWithoutComments) ||
    /\.(?:RowHeight|ColumnWidth)\s*=/i.test(codeWithoutComments) ||
    /\bRange\s*\([^)]*\)\s*=/i.test(codeWithoutComments) ||
    /\.Cells\s*\([^)]*\)\s*=/i.test(codeWithoutComments) ||
    /\.(?:AddComment|WrapText|Merge|UnMerge)\b/i.test(codeWithoutComments) ||
    /(?:Charts|ChartObjects|Shapes)\s*\.\s*Add/i.test(codeWithoutComments) ||
    /\.(?:EntireRow|EntireColumn)\b/i.test(codeWithoutComments) ||
    /\.(?:Sort|SortSpecial)\b/i.test(codeWithoutComments) ||
    /\.(?:Copy|Cut)\b/i.test(codeWithoutComments) ||
    /\.(?:Paste|PasteSpecial)\b/i.test(codeWithoutComments) ||
    /\.(?:Name)\s*=/i.test(codeWithoutComments) ||
    /\.(?:Group|Ungroup)\b/i.test(codeWithoutComments) ||
    /\.(?:AutoFill|Replace)\b/i.test(codeWithoutComments) ||
    /\.(?:ListObjects|PivotTables|PivotCaches|QueryTables|Tables)\s*\.\s*(?:Add|Create)\b/i.test(codeWithoutComments) ||
    /\bApplication\s*\.\s*Run\b/i.test(codeWithoutComments);

  if (hasWriteOperation) {
    return {
      riskLevel: 'L1',
      reasons: ['包含单元格或工作表修改操作，已创建快照支持一键撤销']
    };
  }

  // 判定 L0: 严格纯读操作
  const nonCommentLines = codeLines
    .map(l => l.trim())
    .filter(l => l && !l.startsWith("'") && !l.toLowerCase().startsWith('rem '));

  const isStrictlyReadOnly = nonCommentLines.every(line => {
    // 遇到冒号分隔的多语句行，直接判定为非纯读
    if (hasMultiStatementColon(line)) {
      return false;
    }

    // 过程控制/结构
    if (/^(?:Public\s+|Private\s+|Friend\s+)?(?:Sub|Function)\b/i.test(line) ||
        /^End\s+(?:Sub|Function)\b/i.test(line) ||
        /^(?:Dim|Const|ReDim)\b/i.test(line) ||
        /^Set\s+[a-zA-Z_][a-zA-Z0-9_]*\s*=\s*(?!.*(?:\.Delete|\.Insert|\.Clear|\.Add|\.Copy|\.Paste|\.Sort|\.Group|\.Ungroup|\.AutoFill|\.Replace))/i.test(line) ||
        // 控制流语句只允许换行块形式（行尾无后续执行语句），带 Then <语句> 的单行 If 一律不放行
        /^If\b.+\bThen\s*$/i.test(line) ||
        /^ElseIf\b.+\bThen\s*$/i.test(line) ||
        /^Else\s*$/i.test(line) ||
        /^End\s+If\s*$/i.test(line) ||
        /^(?:For\b|Next\b|Do\b|Loop\b|While\b|Wend|Select\s+Case\b|Case\b|End\s+Select|With\b|End\s+With|Exit\s+(?:Sub|Function|For|Do))\b/i.test(line)) {
      return true;
    }
    // 纯读输出/检查
    if (/^MsgBox\b/i.test(line) ||
        /^Debug\.Print\b/i.test(line) ||
        /^(?:Application\.)?StatusBar\s*=/i.test(line) ||
        /^(?:Range|Cells|ActiveSheet|Worksheets|Sheets|ActiveCell|Selection)\b[^\.]*\.(?:Select|Activate)\b/i.test(line)) {
      return true;
    }
    // 局部变量赋值
    if (/^[a-zA-Z_][a-zA-Z0-9_]*\s*=\s*/i.test(line)) {
      if (!/(?:\.Delete|\.Insert|\.Clear|\.Add|\.Copy|\.Paste|\.Sort|\.Group|\.Ungroup|\.AutoFill|\.Replace)\b/i.test(line)) {
        return true;
      }
    }
    return false;
  });

  if (isStrictlyReadOnly) {
    return {
      riskLevel: 'L0',
      reasons: ['只读操作，不修改工作簿内容']
    };
  }

  // 兜底：无法证明是只读 → 按 L1 处理 (fail-safe)
  return {
    riskLevel: 'L1',
    reasons: ['未识别到明确写操作，按可回滚写入处理并创建快照']
  };
}
