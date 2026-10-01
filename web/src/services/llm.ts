import { loadLlmConfig, type LlmConfig } from './config';
import type { WorkbookReadback, VbaExecutionData } from './bridge';

export type UserIntent = 'CHAT' | 'AUTOMATION' | 'AMBIGUOUS';

export interface ParsedStreamOutput {
  explanation: string;
  vbaCode: string;
  hasCode: boolean;
  isTruncated: boolean;
}

export interface ScopeRiskResult {
  hasRisk: boolean;
  riskType?: 'ALL_CELLS_FORMAT' | 'ALL_CELLS_CLEAR' | 'ALL_COLUMNS_FORMAT';
  matchedSnippet?: string;
  advice?: string;
}

export interface ExtractedVbaResult {
  code: string;
  isTruncated: boolean;
  status: 'valid' | 'no_code' | 'truncated' | 'invalid_structure' | 'scope_risk';
  error?: string;
  scopeRisk?: ScopeRiskResult;
}

/**
 * 意图识别器：严禁仅凭是否含有 '表格' 或 'Sub' 等简单词判定
export type UserIntent = 'CHAT' | 'AUTOMATION' | 'AMBIGUOUS';


/**
 * 获取当前本地系统时钟信息
 */
export function getCurrentDateTimeInfo(): string {
  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth() + 1;
  const date = now.getDate();
  const days = ['星期日', '星期一', '星期二', '星期三', '星期四', '星期五', '星期六'];
  const dayOfWeek = days[now.getDay()];
  const hours = String(now.getHours()).padStart(2, '0');
  const minutes = String(now.getMinutes()).padStart(2, '0');
  const seconds = String(now.getSeconds()).padStart(2, '0');
  return `${year}年${month}月${date}日 ${dayOfWeek} ${hours}:${minutes}:${seconds}`;
}

/**
 * 构建纯自然语言对话通道系统提示词
 */
export function buildChatSystemPrompt(): string {
  const timeInfo = getCurrentDateTimeInfo();
  return `你是一名精通 Microsoft Excel 和 VBA 的专业顾问。
当前用户设备的本地系统真实时间: 【${timeInfo}】。
当前处于【问答咨询与解释通道】。请用专业、亲切、通俗易懂的中文直接解答用户的问题、解释代码含义或进行日常交流。
若用户咨询当前日期、今天几号、明天星期几、节假日或与时间相关的推算，请严格基于上述系统真实时间准确回答。
【核心边界规则】：
1. 本通道为问答与咨询通道，不会自动在 Excel 中执行宏代码。若用户要求编写、展示或提供 VBA 代码供参考，可以用 Markdown 代码块（如 \`\`\`vba）清晰展示，无需引导执行。
2. 若用户询问代码含义，请用文字清晰分步剖析。
3. 若用户需要直接在当前工作簿中自动化修改或生成数据，请提示用户直接下达操作指令（例如“直接在表格中操作”、“帮我拆分D列”等），系统会自动切换至自动化执行通道直接在工作表中完成操作。`;
}

/**
 * 提示词规则与限制结构化编目（供界面呈现、审计日志与合规校验使用）
 */
export interface PromptRuleItem {
  name: string;
  category: '运行入口协议' | '输出格式协议' | '执行交互边界' | '模型能力限制';
  description: string;
  isHostCapabilityLimitation: boolean;
  limitationReason?: string;
}

export function getPromptConstraintsCatalog(): PromptRuleItem[] {
  return [
    {
      name: '运行宿主环境',
      category: '运行入口协议',
      description: '代码在 Windows 桌面 Excel 进程中编译并执行。',
      isHostCapabilityLimitation: false,
    },
    {
      name: '主入口过程规范',
      category: '运行入口协议',
      description: '建议声明为 Sub Main(targetWb As Workbook) 或 Sub Main()，显式操作目标工作簿。',
      isHostCapabilityLimitation: false,
    },
    {
      name: '过程与算法自由度',
      category: '运行入口协议',
      description: '允许自由声明辅助 Sub、Function、常量、自定义类型或选择最优算法，无范式限制。',
      isHostCapabilityLimitation: false,
    },
    {
      name: '纯源码输出',
      category: '输出格式协议',
      description: '直接输出完整、可执行的纯 VBA 源码，严禁输出对话前言或后记。',
      isHostCapabilityLimitation: false,
    },
    {
      name: '免 Markdown 代码围栏',
      category: '输出格式协议',
      description: '自动化通道不要包裹在 ```vba 围栏中，直接输出代码本身以利解析。',
      isHostCapabilityLimitation: false,
    },
    {
      name: '严禁伪代码与占位符',
      category: '输出格式协议',
      description: '严禁输出 \'TODO、未实现的占位符代码。',
      isHostCapabilityLimitation: false,
    },
    {
      name: '语法控制流严格闭合',
      category: '执行交互边界',
      description: '变量声明需规范，控制块必须严格闭合配对（For/Next、If/End If 等）。',
      isHostCapabilityLimitation: false,
    },
    {
      name: '无法完成时单行说明',
      category: '执行交互边界',
      description: '现有上下文确实无法完成时，直接单行说明原因，严禁生成伪造代码。',
      isHostCapabilityLimitation: false,
    },
    {
      name: '严禁调用 MsgBox / InputBox 阻塞式交互弹窗',
      category: '模型能力限制',
      description: '严禁调用 MsgBox、InputBox 等会导致流程挂起的阻塞式交互对话框。',
      isHostCapabilityLimitation: true,
      limitationReason: '【宿主能力限制说明】当前宿主自动化执行内核运行在非交互式后台调用管道中，阻塞式 UI 弹窗会导致 Excel 主线程永久挂死，属于宿主执行器施加的安全能力边界，并非中性协议，亦非模型自然选择。',
    },
  ];
}

/**
 * 构建自动化执行通道系统提示词
 * 严格按照【运行入口协议】、【输出格式协议】、【执行交互边界】与【模型能力限制】四类组织
 */
export function buildAutomationSystemPrompt(
  targetWorkbookName: string,
  sheets: string[],
  activeSheet: string = '',
  usedRange: string = ''
): string {
  const timeInfo = getCurrentDateTimeInfo();
  return `你是一名精通 Microsoft Windows 桌面 Excel 和 VBA 自动化的专业 AI 助手。
当前用户设备的本地系统真实时间: 【${timeInfo}】。
当前任务目标工作簿: "${targetWorkbookName || '当前活动工作簿'}"。
目标工作簿包含的工作表: [${sheets.join(', ')}]。
当前活动工作表: "${activeSheet || '默认'}"，使用区域: "${usedRange || '空'}"。

【运行入口协议】：
1. 运行环境：你的代码将在 Windows 桌面 Excel 进程中编译并执行。
2. 主入口声明：主执行过程建议声明为 Sub Main(targetWb As Workbook)，通过 targetWb 显式操作目标工作簿中的工作表（如 targetWb.Worksheets("${activeSheet || '默认'}") 或 targetWb.ActiveSheet）；亦可使用标准无参入口 Sub Main()，此时宏执行器将在激活目标工作簿后调用该过程，通过 ActiveSheet 或 ActiveWorkbook 操作。
3. 算法与辅助过程自由度：除主过程入口外，你可以根据任务需求自由定义任何辅助 Sub、Function、常量、自定义类型或选择最佳算法，无任何固定范式限制。

【输出格式协议】：
1. 直接输出完整、可执行的纯 VBA 源码，严禁输出任何解释说明文字、对话前言或后记；
2. 严禁输出 Markdown 代码围栏（例如不要包裹在 \`\`\`vba 或 \`\`\` 中，直接输出 VBA 代码本身）；
3. 严禁输出伪代码、未完成的占位符（如 'TODO、'此处补充代码）。

【执行交互边界】：
1. 控制流与语法严格闭合：如需声明变量，请完整规范声明；代码逻辑必须完整闭合（For 与 Next、If 与 End If 等严格配对）；
2. 无法完成时单行拒绝：如果根据现有上下文确实无法完成任务，请不要生成伪造代码，直接输出单行说明原因。

【模型能力限制（宿主环境安全约束）】：
1. 严禁调用 MsgBox、InputBox 等会导致自动化流程挂起的阻塞式交互对话框（说明：由于插件执行内核处于后台自动化链路，此类弹窗会阻塞 Excel UI 线程并导致自动化执行永久挂死，属于宿主执行器施加的能力限制，并非中性协议或模型自然选择）。`;
}

/**
 * 实时解析大模型流式输出：
 * 支持两种协议：纯 VBA 源码优先，同时向下兼容单层 Markdown ```vba ... ``` 围栏
 */
export function parseStreamOutput(raw: string): ParsedStreamOutput {
  if (!raw) {
    return { explanation: '', vbaCode: '', hasCode: false, isTruncated: false };
  }

  const fenceRegex = /```(?:vba|vb)?\s*/i;
  const match = raw.match(fenceRegex);

  if (match && match.index !== undefined) {
    // 包含代码围栏
    const explanation = raw.slice(0, match.index).trim();
    const codeStartIndex = match.index + match[0].length;
    const remaining = raw.slice(codeStartIndex);
    const endFenceIndex = remaining.indexOf('```');

    if (endFenceIndex !== -1) {
      return {
        explanation,
        vbaCode: remaining.slice(0, endFenceIndex),
        hasCode: true,
        isTruncated: false,
      };
    } else {
      return {
        explanation,
        vbaCode: remaining,
        hasCode: true,
        isTruncated: true,
      };
    }
  }

  // 纯文本输出模式：检查是否本身就是纯 VBA 代码
  const isPureVba = /^\s*(?:Option\s+Explicit|Attribute\s+|'(?:[^\r\n]*)|(?:\b(?:Public\s+|Private\s+)?(?:Sub|Function)\b))/im.test(raw);
  if (isPureVba) {
    return {
      explanation: '',
      vbaCode: raw,
      hasCode: true,
      isTruncated: false,
    };
  }

  return {
    explanation: raw.trim(),
    vbaCode: '',
    hasCode: false,
    isTruncated: false,
  };
}

/**
 * 严格提取并校验 VBA 代码：
 * 1. 优先按纯代码协议提取，向下兼容单层 Markdown 代码围栏（仅剥除最外层标记）；
 * 2. 严禁静默删除、替换、补全或重写模型生成的 VBA（保留所有引号、注释、声明与换行）；
 * 3. 严格识别截断、未闭合过程与已知范围失控风险。
 */
export function extractVbaCode(content: string, finishReason?: string): ExtractedVbaResult {
  if (!content || typeof content !== 'string') {
    return { code: '', isTruncated: false, status: 'no_code', error: '模型响应内容为空' };
  }

  // 1. API 级别截断判定
  if (finishReason === 'length') {
    return {
      code: content,
      isTruncated: true,
      status: 'truncated',
      error: '模型响应达到最大 Token 长度上限被硬截断，输出代码不完整，已安全停止执行。',
    };
  }

  const trimmed = content.trim();

  let extractedCode = '';
  const fenceRegex = /```(?:vba|vb)?\s*([\s\S]*?)(?:```|$)/i;
  const fenceMatch = trimmed.match(fenceRegex);

  if (fenceMatch && fenceMatch.index !== undefined) {
    // 检查是否具备闭合的 ```
    const codePart = trimmed.slice(fenceMatch.index + 3);
    if (!codePart.includes('```')) {
      return {
        code: fenceMatch[1],
        isTruncated: true,
        status: 'truncated',
        error: '模型响应被截断 (代码围栏未闭合 ```)，已安全停止执行。',
      };
    }
    // 仅剥除外层一层围栏，内部所有字符 100% 原样保留，绝不篡改
    extractedCode = fenceMatch[1];
  } else {
    // 纯文本形态：检查是否以典型的 VBA 过程或语句开始
    const isVbaText = /^\s*(?:Option\s+Explicit|Attribute\s+|'(?:[^\r\n]*)|(?:\b(?:Public\s+|Private\s+)?(?:Sub|Function)\b))/im.test(trimmed);
    if (isVbaText) {
      extractedCode = trimmed;
    } else {
      return {
        code: '',
        isTruncated: false,
        status: 'no_code',
        error: '模型回复未包含可执行的 VBA 过程源码（普通文字回答）。',
      };
    }
  }

  // 2. 检查基本过程完整性
  const hasSubOrFunction = /(?:^|\n)\s*(?:Public\s+|Private\s+)?(?:Sub|Function)\s+[a-zA-Z0-9_\u4e00-\u9fa5]+/i.test(extractedCode);
  if (!hasSubOrFunction) {
    return {
      code: extractedCode,
      isTruncated: false,
      status: 'invalid_structure',
      error: '提取出的代码中未包含有效的 Sub 或 Function 过程定义。',
    };
  }

  // 3. 检查未闭合的 Sub / Function
  const hasSubStart = /(?:^|\n)\s*(?:Public\s+|Private\s+)?Sub\s+/i.test(extractedCode);
  const hasSubEnd = /(?:^|\n)\s*End\s+Sub\b/i.test(extractedCode);
  const hasFnStart = /(?:^|\n)\s*(?:Public\s+|Private\s+)?Function\s+/i.test(extractedCode);
  const hasFnEnd = /(?:^|\n)\s*End\s+Function\b/i.test(extractedCode);

  if ((hasSubStart && !hasSubEnd) || (hasFnStart && !hasFnEnd)) {
    return {
      code: extractedCode,
      isTruncated: true,
      status: 'truncated',
      error: '代码过程未闭合 (缺少配对的 End Sub 或 End Function)，可能模型生成被中途截断，已安全拦截未执行。',
    };
  }

  // 4. 执行前高危范围检查
  const scopeRisk = checkVbaScopeRisk(extractedCode);
  if (scopeRisk.hasRisk) {
    return {
      code: extractedCode,
      isTruncated: false,
      status: 'scope_risk',
      scopeRisk,
      error: `影响范围失控警告: ${scopeRisk.matchedSnippet}。${scopeRisk.advice}`,
    };
  }

  // 100% 原始代码返回，不得为了编译通过而静默改写任何字符
  return {
    code: extractedCode,
    isTruncated: false,
    status: 'valid',
    scopeRisk,
  };
}

/**
 * 执行前影响范围检查：重点检测对整张工作表 171 亿单元格做重度格式化、批量删除的高危操作
 * 杜绝 ws.Cells.Borders.LineStyle = xlContinuous 等合法语法但失控导致 Excel 挂死的代码
 */
export function checkVbaScopeRisk(code: string): ScopeRiskResult {
  if (!code) return { hasRisk: false };

  // 1. 全表单元格边框、背景色或条件格式
  const allCellsFormatRegex =
    /(?:(?:ws|ActiveSheet|Worksheets\([^)]+\)|targetWb\.ActiveSheet)\s*\.\s*Cells|(?<!\.)\bCells)\s*\.\s*(?:Borders|Interior|FormatConditions)\b/i;
  const match1 = code.match(allCellsFormatRegex);
  if (match1) {
    return {
      hasRisk: true,
      riskType: 'ALL_CELLS_FORMAT',
      matchedSnippet: match1[0],
      advice: `代码包含对整张工作表全部单元格的格式化操作（${match1[0]}）。Excel单张表包含171亿个单元格，对全表Cells直接设置边框或背景色会耗尽系统资源导致Excel深度卡死。请将边框与背景色限定在实际业务数据区域（如 ws.Range(...) 或 Range(ws.Cells(r1, c1), ws.Cells(r2, c2))）。`,
    };
  }

  // 2. 全表单元格清空格式或删除
  const allCellsClearRegex =
    /(?:(?:ws|ActiveSheet|Worksheets\([^)]+\)|targetWb\.ActiveSheet)\s*\.\s*Cells|(?<!\.)\bCells)\s*\.\s*(?:ClearFormats|Delete)\b/i;
  const match2 = code.match(allCellsClearRegex);
  if (match2) {
    return {
      hasRisk: true,
      riskType: 'ALL_CELLS_CLEAR',
      matchedSnippet: match2[0],
      advice: `代码包含对全表单元格的批量删除或清格式操作（${match2[0]}）。请改为仅对数据表使用区域（ws.UsedRange）或指定数据Range操作。`,
    };
  }

  // 3. 整列/整行批量格式化（整列含104万行，批量设置边框极易引发性能灾难）
  const allColumnsFormatRegex =
    /(?:(?:ws|ActiveSheet|Worksheets\([^)]+\))\s*\.\s*)?(?:Columns(?:\([^)]+\))?|Rows(?:\([^)]+\))?)\s*\.\s*(?:Borders|Interior|FormatConditions)\b/i;
  const match3 = code.match(allColumnsFormatRegex);
  if (match3) {
    return {
      hasRisk: true,
      riskType: 'ALL_COLUMNS_FORMAT',
      matchedSnippet: match3[0],
      advice: `代码尝试对整列/整行全部单元格设置边框或背景色（${match3[0]}）。整列包含104万个单元格，请仅在有效数据行范围内设置格式。`,
    };
  }

  return { hasRisk: false };
}

/**
 * 执行后状态核验：将写后读回 (Readback) 与用户实际诉求比较
 * 严格区分“宏已运行”与“任务完成”，主观美观标为“效果待确认”
 */
export function verifyExecutionResult(
  prompt: string,
  readback: WorkbookReadback | undefined
): { status: 'verified' | 'unconfirmed' | 'failed'; note: string } {
  if (!readback) {
    return {
      status: 'unconfirmed',
      note: '未能读回目标工作簿实际变更状态，具体效果请在工作表中核对。',
    };
  }

  if (!readback.targetVerified) {
    return {
      status: 'failed',
      note: `目标工作簿身份核验失败（预期目标: ${readback.targetWorkbookName}）。`,
    };
  }

  if (readback.otherWorkbooksAffected) {
    return {
      status: 'failed',
      note: readback.affectedWorkbooksWarning || '安全告警：检测到非目标工作簿受到附带修改！',
    };
  }

  const notes: string[] = [];

  // 1. 坐标起点核验
  const startMatch = prompt.match(/(?:从|在)\s*([A-Za-z]+[0-9]+)/i);
  if (startMatch) {
    const expectedStart = startMatch[1].toUpperCase();
    if (readback.startCell && readback.startCell.toUpperCase() !== expectedStart) {
      notes.push(`指令指定从 ${expectedStart} 开始，实际检测起始于 ${readback.startCell}`);
    }
  }

  // 2. 公式核验
  const wantsFormula = /(?:公式|求和|sum|计算|平均|average|vlookup|xlookup)/i.test(prompt);
  if (wantsFormula && !readback.hasFormulas) {
    notes.push('指令包含公式诉求，实际检测区域未发现标准 Excel 公式（可能已直接写入数值）');
  }

  // 3. 具体单元格写入核验 (如：在C3，输入文字“错误” 或 在A1写入测试)
  const cellWriteMatch = prompt.match(/(?:在|向)\s*([A-Za-z]+[0-9]+)[，,\s]*(?:写入|输入|填写|填入|置入)(?:内容|文字|数值)?\s*[“"']?([^”"'\s，。！!]+)[”"']?/i);
  if (cellWriteMatch) {
    const targetCell = cellWriteMatch[1].toUpperCase();
    const expectedVal = cellWriteMatch[2];
    if (readback.sampleValues && readback.sampleValues.length > 0) {
      const found = readback.sampleValues.some((v) => v.includes(expectedVal));
      if (found) {
        return {
          status: 'verified',
          note: `已成功在目标区域 (${readback.usedRangeAddress || targetCell}) 写入【${expectedVal}】，数据核验一致。`,
        };
      }
    }
  }

  // 4. 开放式业务任务（分列、图表、美化、排版等）：客观描述当前区域，诚实标为待人工核验，绝不单凭 rowCount > 0 冒充验证通过
  const wantsTransformOrStyle = /(?:美化|商务|好看|排版|样式|颜色|边框|图表|柱状图|折线图|饼图|分列|拆分|排序|筛选|整理)/i.test(prompt);
  if (wantsTransformOrStyle) {
    const features: string[] = [];
    if (readback.hasBorders) features.push('已包含边框');
    if (readback.hasInteriorColor) features.push('已包含单元格填充');
    const featureDesc = features.length > 0 ? `（${features.join('、')}）` : '';
    return {
      status: 'unconfirmed',
      note: `宏已执行完成，工作表使用区域为 ${readback.usedRangeAddress || '已更新'}${featureDesc}。视觉与业务呈现效果待人工确认。`,
    };
  }

  if (notes.length > 0) {
    return {
      status: 'unconfirmed',
      note: `宏已执行完成，但与指令存在客观核验差异：${notes.join('；')}，请在工作表中核对。`,
    };
  }

  return {
    status: 'unconfirmed',
    note: `宏已执行完成（当前使用区域: ${readback.usedRangeAddress || ''}，共 ${readback.rowCount || 0} 行 ${readback.columnCount || 0} 列），具体效果请在工作表中核对确认。`,
  };
}

export interface ChatHistoryItem {
  role: 'user' | 'assistant';
  content: string;
}

export interface ApiAuditInfo {
  maxTokensStatus: string;
  thinkingBudgetStatus: string;
  historyCountSent: number;
  totalHistoryAvailable: number;
  historyStrategy: string;
  isHistoryCompressedOrStripped: boolean;
  actualPayloadSummary: {
    model: string;
    temperatureSent: boolean;
    temperatureVal?: number;
    maxTokensSent: boolean;
    maxTokensVal?: number;
    thinkingMode?: string;
    thinkingBudgetSent: boolean;
    thinkingBudgetVal?: number;
    messagesCount: number;
  };
}

export interface LlmStreamResponse {
  fullText: string;
  finishReason?: string;
  audit: ApiAuditInfo;
}

export async function callLlmStream(
  prompt: string,
  targetWorkbookName: string,
  sheets: string[],
  activeSheet: string,
  usedRange: string,
  intent: UserIntent,
  onChunk: (text: string) => void,
  history: ChatHistoryItem[] = []
): Promise<LlmStreamResponse> {
  const config: LlmConfig = loadLlmConfig();

  if (!config.apiKey && config.provider !== 'ollama') {
    throw new Error('请先在右上角【API 设置】中填写 API Key！');
  }

  let baseUrl = config.baseUrl.trim().replace(/\/+$/, '');
  if (!baseUrl.endsWith('/chat/completions')) {
    baseUrl += '/chat/completions';
  }

  // 根据通道选择对应的系统提示词
  const systemPrompt =
    intent === 'CHAT'
      ? buildChatSystemPrompt()
      : buildAutomationSystemPrompt(targetWorkbookName, sheets, activeSheet, usedRange);

  // 历史上下文策略：
  // 1. 保留最近对话完整原文（含 VBA 代码块与文字说明，严禁静默正则剔除代码块）；
  // 2. 确保模型能基于上一轮生成的宏进行修改、调整、补充或排错；
  // 3. 保留最近 6 条有效对话历史，不暗改、不压缩、不摘要
  const MAX_HISTORY_TURNS = 6;
  const actualHistory = history
    .slice(-MAX_HISTORY_TURNS)
    .filter((h) => h && h.content && h.content.trim().length > 0)
    .map((h) => ({
      role: h.role,
      content: h.content, // 100% 原始文本保留，含任何 ```vba 宏代码
    }));

  const payload: Record<string, any> = {
    model: config.model,
    messages: [
      { role: 'system', content: systemPrompt },
      ...actualHistory,
      { role: 'user', content: prompt },
    ],
    stream: true,
  };

  // 输出预算策略：用户明确配置 maxTokens 时才发送；未配置时不发送，交由模型服务端采用其默认限制
  let maxTokensStatus = '未发送（采用服务端模型默认限制）';
  let maxTokensVal: number | undefined = undefined;
  if (typeof config.maxTokens === 'number' && config.maxTokens > 0) {
    payload.max_tokens = config.maxTokens;
    maxTokensVal = config.maxTokens;
    maxTokensStatus = `用户配置值: ${config.maxTokens}`;
  }

  // 思考模式预算策略：
  // 1. disabled 时发送 type: disabled；
  // 2. budget 时仅当用户明确配置 thinkingBudget > 0 才发送 budget_tokens，绝不暗设 2048；
  // 3. auto 时不发送 thinking 字段，完全交给模型服务端决定
  let thinkingBudgetStatus = '未发送';
  let thinkingBudgetVal: number | undefined = undefined;
  if (config.thinkingMode === 'disabled') {
    payload.thinking = { type: 'disabled' };
    thinkingBudgetStatus = '已关闭思考链 (type: disabled)';
  } else if (config.thinkingMode === 'budget') {
    const thinkingObj: Record<string, any> = { type: 'enabled' };
    if (typeof config.thinkingBudget === 'number' && config.thinkingBudget > 0) {
      thinkingObj.budget_tokens = config.thinkingBudget;
      thinkingBudgetVal = config.thinkingBudget;
      thinkingBudgetStatus = `用户配置值: ${config.thinkingBudget}`;
    } else {
      thinkingBudgetStatus = '未发送 budget_tokens（仅发送 type: enabled，采用服务端默认限制）';
    }
    payload.thinking = thinkingObj;
  }

  // 用户有配置温度才传入，未配置则不传，使用模型默认行为
  let temperatureVal: number | undefined = undefined;
  if (typeof config.temperature === 'number' && !isNaN(config.temperature)) {
    payload.temperature = config.temperature;
    temperatureVal = config.temperature;
  }

  const apiAudit: ApiAuditInfo = {
    maxTokensStatus,
    thinkingBudgetStatus,
    historyCountSent: actualHistory.length,
    totalHistoryAvailable: history.length,
    historyStrategy: `保留最近 ${MAX_HISTORY_TURNS} 条完整对话原文（含 VBA 代码块，无暗改、无截断剥离）`,
    isHistoryCompressedOrStripped: false,
    actualPayloadSummary: {
      model: config.model,
      temperatureSent: temperatureVal !== undefined,
      temperatureVal,
      maxTokensSent: maxTokensVal !== undefined,
      maxTokensVal,
      thinkingMode: config.thinkingMode || 'auto',
      thinkingBudgetSent: thinkingBudgetVal !== undefined,
      thinkingBudgetVal,
      messagesCount: payload.messages.length,
    },
  };

  const headers: Record<string, string> = {
    'Content-Type': 'application/json',
  };

  if (config.apiKey) {
    headers['Authorization'] = `Bearer ${config.apiKey}`;
  }

  // 超时守护：深度思考模型规划时间较长，心跳超时设为 60 秒，避免误杀
  const abortController = new AbortController();
  let chunkTimer: any = null;

  const resetHeartbeat = () => {
    if (chunkTimer) clearTimeout(chunkTimer);
    chunkTimer = setTimeout(() => {
      abortController.abort(new Error('LLM_HEARTBEAT_TIMEOUT'));
    }, 60000);
  };

  resetHeartbeat();

  let response: Response;
  try {
    response = await fetch(baseUrl, {
      method: 'POST',
      headers,
      body: JSON.stringify(payload),
      signal: abortController.signal,
    });
  } catch (fetchErr: any) {
    if (chunkTimer) clearTimeout(chunkTimer);
    if (fetchErr.name === 'AbortError' || fetchErr.message?.includes('TIMEOUT')) {
      throw new Error('大模型连接超时，请检查网络或模型服务状态后重试。');
    }
    throw fetchErr;
  }

  if (!response.ok) {
    if (chunkTimer) clearTimeout(chunkTimer);
    const errText = await response.text().catch(() => '');
    throw new Error(`API 请求失败 (${response.status}): ${errText || response.statusText}`);
  }

  if (!response.body) {
    if (chunkTimer) clearTimeout(chunkTimer);
    throw new Error('未获取到流式响应数据体');
  }

  const reader = response.body.getReader();
  const decoder = new TextDecoder('utf-8');
  let fullText = '';
  let finishReason: string | undefined = undefined;
  let buffer = '';

  try {
    while (true) {
      resetHeartbeat();
      const { done, value } = await reader.read();
      if (done) break;

      buffer += decoder.decode(value, { stream: true });
      const lines = buffer.split('\n');
      buffer = lines.pop() || '';

      for (const line of lines) {
        const trimmed = line.trim();
        if (!trimmed || trimmed.startsWith(':')) continue;
        if (trimmed === 'data: [DONE]') continue;

        if (trimmed.startsWith('data: ')) {
          const jsonStr = trimmed.slice(6);
          try {
            const data = JSON.parse(jsonStr);
            const delta = data.choices?.[0]?.delta?.content || '';
            const fr = data.choices?.[0]?.finish_reason;
            if (fr) finishReason = fr;
            if (delta) {
              fullText += delta;
              onChunk(fullText);
            }
          } catch {}
        }
      }
    }

    // 处理流结束后缓冲区中可能残留的最后一行 SSE 数据
    if (buffer.trim()) {
      const lastLine = buffer.trim();
      if (lastLine.startsWith('data: ') && lastLine !== 'data: [DONE]') {
        try {
          const data = JSON.parse(lastLine.slice(6));
          const delta = data.choices?.[0]?.delta?.content || '';
          const fr = data.choices?.[0]?.finish_reason;
          if (fr) finishReason = fr;
          if (delta) {
            fullText += delta;
            onChunk(fullText);
          }
        } catch {}
      }
    }
  } catch (streamErr: any) {
    if (streamErr.name === 'AbortError' || streamErr.message?.includes('TIMEOUT')) {
      throw new Error('流式生成传输中断（60秒未收到新响应），请点击重试。');
    }
    throw streamErr;
  } finally {
    if (chunkTimer) clearTimeout(chunkTimer);
  }

  return { fullText, finishReason, audit: apiAudit };
}
