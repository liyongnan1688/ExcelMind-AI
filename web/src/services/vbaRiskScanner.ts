/**
 * vbaRiskScanner.ts
 * 
 * VBA 代码风险等级静态扫描器 (P5 规范)
 * 在代码执行前对 VBA 进行静态规则分级判定，实现风险导向的执行与确认闭环。
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
  reasons: string[];
  targetDetail?: string;
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

  // 1. 删除工作表 (不可逆)
  if (/(?:Worksheets|Sheets)\s*\([^)]*\)\s*\.Delete/i.test(codeWithoutComments) ||
      /(?:ActiveSheet|ThisWorkbook\.ActiveSheet)\.Delete/i.test(codeWithoutComments)) {
    // 尝试提取被删除的工作表名
    const match = codeWithoutComments.match(/(?:Worksheets|Sheets)\s*\(\s*["']([^"']+)["']\s*\)\s*\.Delete/i);
    const sheetName = match ? match[1] : '工作表';
    l2Reasons.push(`包含删除工作表指令 (将删除工作表《${sheetName}》)`);
  }

  // 2. 外部工作簿操作 (Open / SaveAs / Close)
  if (/\bWorkbooks\s*\.\s*(?:Open|SaveAs|Close)\b/i.test(codeWithoutComments)) {
    l2Reasons.push('包含外部工作簿打开/另存/关闭操作 (Workbooks.Open/SaveAs/Close)');
  }

  // 3. 高危外部 COM 对象 (FileSystemObject / 网络请求等)
  if (/CreateObject\s*\(\s*["'](?:Scripting\.FileSystemObject|MSXML2\.[^"']+|WinHttp\.[^"']+|InternetExplorer\.Application|WScript\.Shell)["']/i.test(codeWithoutComments) ||
      /New\s+FileSystemObject\b/i.test(codeWithoutComments)) {
    l2Reasons.push('包含文件系统或外部网络组件调用 (FileSystemObject/WinHttp/Shell)');
  }

  // 4. 篡改代码工程 (VBProject)
  if (/\b(?:VBProject|ThisWorkbook\.VBProject)\b/i.test(codeWithoutComments)) {
    l2Reasons.push('包含对 VBA 代码工程自身的访问与修改 (VBProject)');
  }

  // 5. 原生文件系统操作语句 (Kill / FileCopy / Name ... As ...)
  if (/\b(?:Kill|FileCopy|RmDir)\b/i.test(codeWithoutComments) ||
      /\bName\b\s+.+\s+\bAs\b/i.test(codeWithoutComments) ||
      /\bOpen\b\s+["'][^"']+["']\s+For\s+(?:Output|Append|Binary)/i.test(codeWithoutComments)) {
    l2Reasons.push('包含文件删除或外部磁盘写入语句 (Kill/FileCopy/Open For Output)');
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
    /(?:Charts|ChartObjects|Shapes)\s*\.\s*Add/i.test(codeWithoutComments);

  if (hasWriteOperation) {
    return {
      riskLevel: 'L1',
      reasons: ['包含单元格或工作表修改操作，已创建快照支持一键撤销']
    };
  }

  // 判定 L0: 确定为纯读操作
  const isPureReadOnly = 
    /\b(?:MsgBox|Debug\.Print|Application\.StatusBar)\b/i.test(codeWithoutComments) ||
    /\bFor\s+Each\b/i.test(codeWithoutComments) ||
    /\bSelect\s+Case\b/i.test(codeWithoutComments) ||
    /\bIf\b.+\bThen\b/i.test(codeWithoutComments);

  if (isPureReadOnly || codeWithoutComments.trim().length > 0) {
    // 如果没有写操作，检查是否只是空函数或只读运算
    if (!hasWriteOperation) {
      return {
        riskLevel: 'L0',
        reasons: ['只读操作，不修改工作簿内容']
      };
    }
  }

  // 兜底：无法确定时按 L2 安全拦截
  return {
    riskLevel: 'L2',
    reasons: ['无法静态识别该代码的写入行为，按安全要求需确认后执行'],
    targetDetail: '潜在未知修改'
  };
}
