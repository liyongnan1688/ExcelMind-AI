import { loadLlmConfig, type LlmConfig } from './config';
import type { WorkbookReadback, VbaExecutionData } from './bridge';

export type UserIntent = 'CHAT' | 'AUTOMATION' | 'AMBIGUOUS';

export interface ParsedStreamOutput {
  explanation: string;
  vbaCode: string;
  hasCode: boolean;
  isTruncated: boolean;
}

export interface ExtractedVbaResult {
  code: string;
  isTruncated: boolean;
  error?: string;
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

【核心执行要求】：
1. 准确理解用户的具体业务意图：严格按照用户指定的目标、坐标、内容类型与格式执行任务，绝不可擅自将具体业务诉求降级或替换为无关结构（例如：用户要求算式乘法表口诀时，必须生成带算式文本的表格；用户要求从指定单元格开始时，必须从该坐标起笔；用户明确要求纯数字乘积矩阵时，才填入数字矩阵）。
2. 视觉排版与美化：当用户要求表格美化或制作完整报表时，请应用清晰、典雅的商务排版风格（如清晰表头、合适列宽、对齐方式、细边框与柔和底色），使表格美观易读。
3. VBA 规范：
   - 必须且仅编写一个主过程: \`Sub LeeTaskEntry()\`，以 \`End Sub\` 完整闭合。
   - 严禁调用 MsgBox，严禁使用 Application.Quit，严禁弹出任何交互确认框。
   - 代码必须针对目标工作簿及其中的工作表进行操作。
4. 回复格式协议：
   先用 1~2 句话中文概括将要执行的操作，随后直接给出标准的 \`\`\`vba ... \`\`\` 代码块。严禁将代码分散在聊天文字中，严禁输出未闭合的代码。`;
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

  // 严格结构校验：必须具备过程声明与完整的 End Sub
  const hasSubDecl = /(?:Public\s+|Private\s+)?Sub\s+[a-zA-Z0-9_\u4e00-\u9fa5]+\s*\(/i.test(rawCode);
  const hasEndSub = /End\s+Sub/i.test(rawCode);

  if (!hasSubDecl) {
    return {
      code: rawCode,
      isTruncated: false,
      error: '代码块中未包含有效的 Sub 过程声明。',
    };
  }

  if (!hasEndSub) {
    return {
      code: rawCode,
      isTruncated: true,
      error: '代码结构不完整 (缺少闭合 End Sub)，可能由于响应截断引起，已拦截未执行。',
    };
  }

  // 清理行前可能的 markdown 符号
  const cleanedCode = rawCode.replace(/^[ \t]*[*\-•][ \t]+/gm, '');

  return {
    code: cleanedCode,
    isTruncated: false,
  };
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

  const p = prompt.toLowerCase();
  const notes: string[] = [];

  // 1. 坐标起点核验
  const startMatch = prompt.match(/(?:从|在)\s*([A-Za-z]+[0-9]+)/i);
  if (startMatch) {
    const expectedStart = startMatch[1].toUpperCase();
    if (readback.startCell && readback.startCell.toUpperCase() !== expectedStart) {
      notes.push(`要求从 ${expectedStart} 开始，实际起始于 ${readback.startCell}`);
    }
  }

  // 2. 算式文本 vs 纯数字乘积核验
  const wantsEquation = /(?:算式|口诀|乘法口诀|×|\*|=)/i.test(prompt);
  if (wantsEquation && readback.sampleValues && readback.sampleValues.length > 0) {
    const hasEquationText = readback.sampleValues.some(
      (v) => v.includes('×') || v.includes('*') || v.includes('=') || v.includes('得')
    );
    if (!hasEquationText) {
      notes.push('检测到填入内容为纯数字矩阵，未生成算式文本');
    }
  }

  // 3. 主观美观排版提示
  const wantsBeauty = /(?:美化|商务|好看|排版|样式|颜色|边框)/i.test(prompt);
  if (wantsBeauty) {
    const styleFeatures: string[] = [];
    if (readback.hasBorders) styleFeatures.push('已添加边框');
    if (readback.hasInteriorColor) styleFeatures.push('已应用背景填充');
    notes.push(`视觉样式(${styleFeatures.join('、') || '基础样式'})已应用，效果待人工确认`);
  }

  if (notes.some((n) => n.includes('未生成算式文本') || n.includes('实际起始于'))) {
    return {
      status: 'unconfirmed',
      note: `宏已运行，但与指令存在差异：${notes.join('；')}`,
    };
  }

  if (wantsBeauty) {
    return {
      status: 'unconfirmed',
      note: `宏已运行（区域: ${readback.usedRangeAddress || '已更新'}），视觉排版效果待人工确认。`,
    };
  }

  return {
    status: 'verified',
    note: `宏已运行，数据区域 ${readback.usedRangeAddress || ''} 验证通过。`,
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

  const payload: Record<string, any> = {
    model: config.model,
    messages: [
      { role: 'system', content: systemPrompt },
      ...cleanHistory,
      { role: 'user', content: prompt },
    ],
    max_tokens: 4096,
    stream: true,
  };

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

  // 超时守护
  const abortController = new AbortController();
  let chunkTimer: any = null;

  const resetHeartbeat = () => {
    if (chunkTimer) clearTimeout(chunkTimer);
    chunkTimer = setTimeout(() => {
      abortController.abort(new Error('LLM_HEARTBEAT_TIMEOUT'));
    }, 15000);
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
