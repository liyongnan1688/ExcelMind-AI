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
  error?: string;
  scopeRisk?: ScopeRiskResult;
}

/**
 * 意图识别器：严禁仅凭是否含有 '表格' 或 'Sub' 等简单词判定
 * 精确区分：
 * 1. CHAT：问答、解释、闲聊、代码解读（即使粘贴了 VBA 代码询问含义）
 * 2. AUTOMATION：明确要求修改/操作当前工作簿
 * 3. AMBIGUOUS：意图不明确，需简短澄清
 */
export function detectIntent(text: string): UserIntent {
  if (!text || typeof text !== 'string') return 'AMBIGUOUS';
  const trimmed = text.trim();
  if (!trimmed) return 'AMBIGUOUS';

  // 1. 优先匹配纯聊天、问候、身份咨询与元问题
  if (
    /(?:自我介绍|介绍(?:一下|下)?(?:自己)?|你是谁|你是[？\?]|你能做(?:什么|啥)|有什么功能|功能介绍|使用说明)/i.test(
      trimmed
    ) ||
    /^(?:你好|您好|hi|hello|hey|help|帮助)[\s!！?？~]*$/i.test(trimmed)
  ) {
    return 'CHAT';
  }

  // 2. 匹配对刚才操作/步骤/代码的解释与复盘要求
  if (
    /(?:解释|说明|介绍|复盘|讲讲|说说|请教|了解|分析)(?:一下|下)?(?:刚才|刚刚|上一[步次]|前面)?.*(?:做了什么|执行了什么|干了什么|代码|操作|步骤|原因|原理|逻辑)/i.test(
      trimmed
    ) ||
    /(?:刚才|刚刚|上一[步次]|前面).*(?:做了什么|执行了什么|干了什么|代码|是什么意思|是干嘛的|原理)/i.test(trimmed)
  ) {
    return 'CHAT';
  }

  // 3. 匹配用户粘贴代码或提及代码询问含义（必须只解释、不执行）
  const asksCodeMeaningRegex =
    /(?:这段代码|这个宏|这几行代码|这段VBA|以下代码|这段宏|Sub\s+[\s\S]+End\s+Sub)[\s\S]*(?:什么意思|含义|解释|怎么理解|干嘛|干什么|作用|读懂|请教|为什么|如何理解)/i;
  if (asksCodeMeaningRegex.test(trimmed)) {
    return 'CHAT';
  }

  // 4. 用户若以“这段代码什么意思”、“帮我看看这段代码”等提问，且文本中包含过程结构
  if (
    /(?:什么意思|怎么理解|是干什么的|有何作用)[\?？]*$/.test(trimmed) &&
    (trimmed.includes('Sub') || trimmed.includes('Range') || trimmed.includes('Dim'))
  ) {
    return 'CHAT';
  }

  // 5. 纯知识性疑问句（如“什么是数据透视表”、“如何使用VLOOKUP”、“怎么计算均值”），未要求直接在工作簿操作
  const generalKnowledgeRegex =
    /^(?:什么是|如何理解|为什么|怎么用|怎么使用|函数用法|公式怎么写|区别是什么|有什么区别)/i;
  if (generalKnowledgeRegex.test(trimmed)) {
    return 'CHAT';
  }

  // 6. 明确的工作簿修改/自动化指令判定
  const hasActionVerb =
    /(?:新建|创建|生成|制作|写入|填充|输入|添加|插入|删除|清除|清空|修改|替换|设置|调整|美化|排版|对齐|边框|底色|颜色|格式化|计算|求和|统计|汇总|排序|筛选|做个|画个|建立|构建)/i.test(
      trimmed
    );
  const hasTargetNoun =
    /(?:表|表格|数据|列|行|单元格|矩阵|图表|柱状图|折线图|饼图|公式|看板|乘法表|清单|明细|工作表|sheet)/i.test(
      trimmed
    );
  const hasCoordinateSpec = /(?:在|从)\s*[A-Za-z]+[0-9]+(?:\s*:\s*[A-Za-z]+[0-9]+)?/i.test(trimmed);
  const hasDirectImperative = /(?:把|将)\s*.+\s*(?:改|设|调|删|变|换|排序|求和|汇总|居中|加粗)/i.test(trimmed);

  if ((hasActionVerb && hasTargetNoun) || hasCoordinateSpec || hasDirectImperative) {
    return 'AUTOMATION';
  }

  // 7. 模糊无明确谓语动词的简短词组（如仅输入“乘法表”、“表格”、“VBA”）
  if (trimmed.length <= 4 || (!hasActionVerb && !trimmed.includes('？') && !trimmed.includes('?'))) {
    return 'AMBIGUOUS';
  }

  // 兜底为问答咨询通道，防止未知文本被误当自动化执行
  return 'CHAT';
}

/**
 * 构建纯自然语言对话通道系统提示词
 */
export function buildChatSystemPrompt(): string {
  return `你是一名精通 Microsoft Excel 和 VBA 的专业顾问。
当前处于【问答咨询与解释通道】。请用专业、亲切、通俗易懂的中文直接解答用户的问题、解释代码含义或进行日常交流。
【核心边界规则】：
1. 本通道只进行纯文本自然语言解答，绝对不要输出任何可被执行的自动化代码块，严禁输出任何 \`\`\`vba 代码块。
2. 若用户询问代码含义，请用文字清晰分步剖析，不要诱导执行。`;
}

/**
 * 构建自动化执行通道系统提示词
 * 彻底移除 D1:L9、35行上限、严禁逐格操作等带有偏置和惩罚性的误导提示词
 */
export function buildAutomationSystemPrompt(
  targetWorkbookName: string,
  sheets: string[],
  activeSheet: string = '',
  usedRange: string = ''
): string {
  return `你是一名精通 Microsoft Windows 桌面 Excel 和 VBA 自动化的专业 AI 助手。
当前任务目标工作簿: "${targetWorkbookName || '当前活动工作簿'}"。
目标工作簿包含的工作表: [${sheets.join(', ')}]。
当前活动工作表: "${activeSheet || '默认'}"，使用区域: "${usedRange || '空'}"。

【核心执行协议与规范】：
1. 根据用户的自然语言需求，自主决定最合适的高效实现方案（可自由使用循环、数组、公式、格式、图表、筛选、数据透视表及辅助过程等，不受限固定模板与行数）。
2. 主过程可以声明接收目标工作簿参数（如 Sub Main(targetWb As Workbook)），也可以编写无参主过程（如 Sub Main()）；允许定义多个辅助过程与函数。
3. 代码必须是完整可编译运行的标准 VBA，语法严格遵循 VB6/VBA 规范（仔细检查括号与属性调用的位置如 ws.Columns(1).ColumnWidth，提前退出请使用 Exit Sub/Function，禁止书写非法的自定义 End 标签 如 End CleanExit 等），包裹在单个 \`\`\`vba ... \`\`\` 代码块中，以 End Sub 正常闭合。
4. 【安全约束】：严禁调用 MsgBox、Application.Quit 或弹出阻塞式交互确认框。
5. 【结构与输出】：直接输出完整可执行的标准 VBA 代码，包裹在 \`\`\`vba ... \`\`\` 代码块中，在代码块前后仅提供简明扼要的说明，避免冗长说明以确保代码完整不被截断。`;
}

/**
 * 实时解析大模型流式输出
 */
export function parseStreamOutput(raw: string): ParsedStreamOutput {
  if (!raw) {
    return { explanation: '', vbaCode: '', hasCode: false, isTruncated: false };
  }

  const fenceRegex = /```(?:vba|vb)?\s*/i;
  const match = raw.match(fenceRegex);

  if (!match || match.index === undefined) {
    return { explanation: raw.trim(), vbaCode: '', hasCode: false, isTruncated: false };
  }

  const explanation = raw.slice(0, match.index).trim();
  const codeStartIndex = match.index + match[0].length;
  const remaining = raw.slice(codeStartIndex);

  const endFenceIndex = remaining.indexOf('```');
  let vbaCode = '';
  let isTruncated = false;

  if (endFenceIndex !== -1) {
    vbaCode = remaining.slice(0, endFenceIndex).trim();
  } else {
    vbaCode = remaining.trim();
    isTruncated = true; // 尚未闭合
  }

  return { explanation, vbaCode, hasCode: true, isTruncated };
}

/**
 * 严格提取并校验 VBA 代码
 * 绝不允许从聊天文字中凭借 includes('Sub') 乱猜宏
 */
export function extractVbaCode(content: string): ExtractedVbaResult {
  if (!content) {
    return { code: '', isTruncated: false, error: '响应内容为空' };
  }

  const openFenceMatch = content.match(/```(?:vba|vb)?\s*/i);
  if (!openFenceMatch || openFenceMatch.index === undefined) {
    return { code: '', isTruncated: false, error: '模型响应未包含规范的 VBA 代码块' };
  }

  const codeStart = openFenceMatch.index + openFenceMatch[0].length;
  const rest = content.slice(codeStart);
  const closeFenceIndex = rest.indexOf('```');

  if (closeFenceIndex === -1) {
    return {
      code: rest.trim(),
      isTruncated: true,
      error: '模型响应被截断 (代码块未闭合 ```)，已安全停止执行。',
    };
  }

  const rawCode = rest.slice(0, closeFenceIndex).trim();

  // 严格结构校验：必须具备过程声明，且 Sub/Function 与 End Sub/End Function 数量必须严格配对
  const subMatches = rawCode.match(/(?:^|\n)\s*(?:Public\s+|Private\s+)?Sub\s+[a-zA-Z0-9_\u4e00-\u9fa5]+\s*\(/gi) || [];
  const endSubMatches = rawCode.match(/(?:^|\n)\s*End\s+Sub\b/gi) || [];
  const fnMatches = rawCode.match(/(?:^|\n)\s*(?:Public\s+|Private\s+)?Function\s+[a-zA-Z0-9_\u4e00-\u9fa5]+\s*\(/gi) || [];
  const endFnMatches = rawCode.match(/(?:^|\n)\s*End\s+Function\b/gi) || [];

  if (subMatches.length === 0) {
    return {
      code: rawCode,
      isTruncated: false,
      error: '代码块中未包含有效的 Sub 过程声明。',
    };
  }

  if (subMatches.length > endSubMatches.length || fnMatches.length > endFnMatches.length) {
    return {
      code: rawCode,
      isTruncated: true,
      error: `代码结构不完整 (检测到 ${subMatches.length} 个 Sub、${endSubMatches.length} 个 End Sub；${fnMatches.length} 个 Function、${endFnMatches.length} 个 End Function)，可能由于模型生成被截断引起，已安全拦截未执行。`,
    };
  }

  // 检查非法的 End 语句 (如 End CleanExit, End Try)
  const invalidEndMatch = rawCode.match(/^\s*End\s+(?!Sub\b|Function\b|Property\b|If\b|With\b|Select\b|Type\b|Enum\b)([A-Za-z0-9_]+)/im);
  if (invalidEndMatch) {
    return {
      code: rawCode,
      isTruncated: false,
      error: `代码包含非标准 VBA 语法语句 '${invalidEndMatch[0].trim()}'（跳出请使用 Exit Sub/Function），已安全拦截未注入。`,
    };
  }

  // 清理行前可能的 markdown 符号
  const cleanedCode = rawCode.replace(/^[ \t]*[*\-•][ \t]+/gm, '');

  const scopeRisk = checkVbaScopeRisk(cleanedCode);

  return {
    code: cleanedCode,
    isTruncated: false,
    scopeRisk,
    error: scopeRisk.hasRisk ? `影响范围失控警告: ${scopeRisk.matchedSnippet}` : undefined,
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
      note: '未能读回目标工作簿实际变更状态，效果待确认。',
    };
  }

  if (!readback.targetVerified) {
    return {
      status: 'failed',
      note: `目标工作簿身份核验失败（预期: ${readback.targetWorkbookName}）。`,
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
    notes.push('指令包含公式计算诉求，实际区域内未检测到标准 Excel 公式（可能直接写入了数值）');
  }

  // 3. 主观美观排版与图表提示（如实标为待人工确认，不冒充万能语义验收器）
  const wantsBeautyOrChart = /(?:美化|商务|好看|排版|样式|颜色|边框|图表|柱状图|折线图|饼图)/i.test(prompt);
  if (wantsBeautyOrChart) {
    const styleFeatures: string[] = [];
    if (readback.hasBorders) styleFeatures.push('检测到边框');
    if (readback.hasInteriorColor) styleFeatures.push('检测到背景填充');
    notes.push(`视觉样式(${styleFeatures.join('、') || '已渲染'})，视觉与版式呈现需人工确认`);
  }

  if (notes.some((n) => n.includes('实际检测起始于') || n.includes('未检测到标准 Excel 公式'))) {
    return {
      status: 'unconfirmed',
      note: `宏已运行，但与指令存在客观差异：${notes.join('；')}`,
    };
  }

  if (wantsBeautyOrChart || notes.length > 0) {
    return {
      status: 'unconfirmed',
      note: `宏已运行（更新区域: ${readback.usedRangeAddress || '已更新'}），${notes.join('；')}。`,
    };
  }

  // 对无法建立确定性断言的常规开放式任务，诚实显示“宏已运行，效果待确认”
  return {
    status: 'unconfirmed',
    note: `宏已运行（更新区域: ${readback.usedRangeAddress || ''}，共 ${readback.rowCount || 0} 行 ${readback.columnCount || 0} 列），具体效果请在工作表中人工核验确认。`,
  };
}

export interface ChatHistoryItem {
  role: 'user' | 'assistant';
  content: string;
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
): Promise<string> {
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

  // 整理历史上下文（排除残留的巨幅代码块）
  const cleanHistory = history
    .slice(-4)
    .map((h) => ({
      role: h.role,
      content: h.content.replace(/```(?:vba|vb)?[\s\S]*?```/gi, '').trim(),
    }))
    .filter((h) => !!h.content);

  const maxTokens = config.maxTokens && config.maxTokens > 0 ? config.maxTokens : 16384;
  const payload: Record<string, any> = {
    model: config.model,
    messages: [
      { role: 'system', content: systemPrompt },
      ...cleanHistory,
      { role: 'user', content: prompt },
    ],
    max_tokens: maxTokens,
    stream: true,
  };

  // 支持思考模式配置：disabled (关闭深度思考，全速输出代码) / budget (设定思考预算) / auto (默认)
  if (config.thinkingMode === 'disabled') {
    payload.thinking = { type: 'disabled' };
  } else if (config.thinkingMode === 'budget') {
    payload.thinking = {
      type: 'enabled',
      budget_tokens: config.thinkingBudget && config.thinkingBudget > 0 ? config.thinkingBudget : 2048,
    };
  }

  // 用户有配置温度才传入，未配置则不传，使用模型默认行为
  if (typeof config.temperature === 'number' && !isNaN(config.temperature)) {
    payload.temperature = config.temperature;
  }

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
            if (delta) {
              fullText += delta;
              onChunk(fullText);
            }
          } catch {}
        }
      }
    }
  } catch (streamErr: any) {
    if (streamErr.name === 'AbortError' || streamErr.message?.includes('TIMEOUT')) {
      throw new Error('流式生成传输中断（15秒未收到新响应），请点击重试。');
    }
    throw streamErr;
  } finally {
    if (chunkTimer) clearTimeout(chunkTimer);
  }

  return fullText;
}
