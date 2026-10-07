import type { VbaExecutionData, SelectionSnapshot } from './bridge';
import type { ChatHistoryItem } from './llm';

export interface ChatMessage {
  id: string;
  role: 'user' | 'assistant';
  content: string; // 界面展示文本（供 UI 气泡渲染：可为简明摘要、纯文本说明或中断提示）
  rawContent?: string; // 大模型原始完整回复文本（保留完整 Sub/Function 源码、前言后语，严禁混入宿主包装器或执行摘要，用于多轮上下文）
  mode?: 'AUTOMATION' | 'CHAT'; // 模式标识
  prompt?: string;
  streamVbaCode?: string; // 流式生成的代码
  execution?: VbaExecutionData | null;
  isExecuting?: boolean;
  selectionSnapshot?: SelectionSnapshot; // 附加的只读选区快照
  macroReference?: MacroReferenceData; // 显式引用的前序宏快照
}

/**
 * 从当前所有消息中提取用于多轮对话的大模型历史上下文：
 * 1. 过滤掉正在处理的当前请求消息与系统欢迎语；
 * 2. 助手角色优先提取 rawContent（模型原始完整回复），避免界面执行摘要污染模型历史；
 * 3. 严格不混入宿主包装器（如 LeeHostRunner）或执行状态标记；
 * 4. 纯文字拒绝、执行失败或无代码回复同样如实保留真实原始文本。
 */
export function buildChatHistory(
  messages: ChatMessage[],
  currentIds?: { userMsgId?: string; assistantMsgId?: string }
): ChatHistoryItem[] {
  return messages
    .filter((m) => {
      if (!m || !m.content) return false;
      if (m.id === 'msg_welcome') return false;
      if (currentIds) {
        if (currentIds.userMsgId && m.id === currentIds.userMsgId) return false;
        if (currentIds.assistantMsgId && m.id === currentIds.assistantMsgId) return false;
      }
      return true;
    })
    .map((m) => {
      let contentToSend = m.content;
      if (m.role === 'assistant') {
        // 关键核心：优先使用大模型原始完整回复 rawContent；仅当 rawContent 未记录时才退化为 content
        if (typeof m.rawContent === 'string' && m.rawContent.trim().length > 0) {
          contentToSend = m.rawContent;
        }
      }
      return {
        role: m.role,
        content: contentToSend,
      };
    });
}

export interface MacroReferenceCheckResult {
  isMacroReferenced: boolean;
  isHistoricalMacroExcluded: boolean;
  warning?: string;
}

/**
 * 检查用户输入是否包含对前序宏代码的引用，以及该宏是否已被历史滑动窗口排除：
 * 若用户意图引用此前生成的宏，但纳入实际发送的历史消息中不含可执行代码，且被排除的历史中曾经存在代码，
 * 则明确给出警示，告知模型未收到该宏，不假装模型已知。
 */
export function checkMacroReferenceContext(
  text: string,
  actualHistory: ChatHistoryItem[],
  excludedHistory: ChatHistoryItem[]
): MacroReferenceCheckResult {
  if (!text || typeof text !== 'string') {
    return { isMacroReferenced: false, isHistoricalMacroExcluded: false };
  }

  // 匹配前序宏引用短语
  const referencePattern = /(刚才|上面|之前|上次|上一轮|现有的|原先的|刚刚|前面).*(宏|代码|vba|sub|脚本|程序)/i;
  const isMacroReferenced = referencePattern.test(text);
  if (!isMacroReferenced) {
    return { isMacroReferenced: false, isHistoricalMacroExcluded: false };
  }

  const containsVbaCode = (item: ChatHistoryItem) => {
    if (item.role !== 'assistant' || !item.content) return false;
    return /(?:Sub\s+\w+|Function\s+\w+|```(?:vba|vb)?)/i.test(item.content);
  };

  const includedHasMacro = actualHistory.some(containsVbaCode);
  const excludedHasMacro = excludedHistory.some(containsVbaCode);

  if (!includedHasMacro && excludedHasMacro) {
    return {
      isMacroReferenced: true,
      isHistoricalMacroExcluded: true,
      warning: '提示：检测到可能引用了前序宏，但包含该宏的消息已超出系统发送范围（仅保留最近 6 条消息）。模型本次请求未收到该宏源码，建议直接将原代码复制到输入框中。',
    };
  }

  // 注：未命中关键词不代表“上下文一定完整”，仅代表未触发超窗宏显式引用警示
  return { isMacroReferenced: true, isHistoricalMacroExcluded: false };
}

export interface MacroReferenceData {
  id: string; // 唯一引用标识，如 'ref_1727850000000_abcd'
  sourceMessageId?: string; // 来源消息 ID
  promptSummary?: string; // 宏生成时的原指令摘要
  vbaCode: string; // 模型原始 VBA 正文（绝不含宿主包装器）
  status: 'complete' | 'incomplete'; // 代码结构完整性
  procedureName?: string; // 主过程名（如 Sub Main 或 Sub CalculateSales）
  charCount: number; // 正文字符数（统计口径: UTF-16 code units, str.length）
  lineCount: number; // 代码行数
  isHistoricalExcluded?: boolean; // 是否已超出最近 6 条消息发送窗口
}

export interface FormatMacroReferenceResult {
  formattedText: string;
  totalChars: number;
  isAlreadyInHistory: boolean;
  duplicateNotice?: string;
}

/**
 * 格式化用户显式引用的前序宏，生成只读附加上下文数据段：
 * 1. 严格保留原始源码，不自动改写、不压缩、不截断；
 * 2. 检查是否已经在最近历史消息中（前 6 条内），若在则明确提示；
 * 3. 统计口径说明：明确标注为 JavaScript string.length 字符数（UTF-16 代码单元），绝不伪称 Token 数。
 */
export function formatMacroReferenceForPrompt(
  ref: MacroReferenceData,
  actualHistory: ChatHistoryItem[] = []
): FormatMacroReferenceResult {
  if (!ref || !ref.vbaCode || typeof ref.vbaCode !== 'string') {
    return { formattedText: '', totalChars: 0, isAlreadyInHistory: false };
  }

  const cleanCode = ref.vbaCode.trim();
  const procName = ref.procedureName || extractProcedureName(cleanCode) || '未命名宏';

  // 检查是否已包含在发送范围内的历史消息中
  const isAlreadyInHistory = actualHistory.some(
    (item) => item.role === 'assistant' && item.content && item.content.includes(cleanCode)
  );

  const duplicateNotice = isAlreadyInHistory
    ? '（注：该宏已在最近 6 条发送历史中；本次作为用户显式指定的重点参考再次附带）'
    : '';

  const lines: string[] = [];
  lines.push('<referenced_vba_context>');
  lines.push('【用户显式引用的前序宏代码（参考源码，来源保真，未受宿主包装器篡改）】');
  lines.push(`- 引用标识: ${ref.id}`);
  lines.push(`- 过程名称: ${procName}`);
  lines.push(`- 完整性状态: ${ref.status === 'incomplete' ? '代码结构不完整（仅供参考，不可直接当作可运行宏）' : '代码结构闭合'}`);
  lines.push(`- 源码字符量: ${cleanCode.length} 字符 (统计口径: JavaScript UTF-16 code units，非 Token 估算)`);
  if (ref.promptSummary) {
    lines.push(`- 原始需求背景: "${ref.promptSummary.replace(/[\r\n]+/g, ' ').slice(0, 100)}"`);
  }
  if (duplicateNotice) {
    lines.push(`- 历史状态说明: ${duplicateNotice}`);
  }
  lines.push('\n```vba');
  lines.push(cleanCode);
  lines.push('```');
  lines.push('</referenced_vba_context>');

  const formattedText = lines.join('\n');
  return {
    formattedText,
    totalChars: formattedText.length,
    isAlreadyInHistory,
    duplicateNotice: isAlreadyInHistory ? duplicateNotice : undefined,
  };
}

export function extractProcedureName(code: string): string {
  if (!code || typeof code !== 'string') return '';
  const match = code.match(/(?:Public\s+|Private\s+)?(?:Sub|Function)\s+([a-zA-Z0-9_]+)/i);
  return match ? match[1] : '';
}

