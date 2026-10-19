import type { ChatHistoryItem } from './llm';

/**
 * 历史会话压缩配置选项
 */
export interface CompactionOptions {
  /** 保留未压缩的最新尾部消息条数 (默认: 6) */
  maxTailMessages?: number;
  /** 触发自动压缩的字符阈值预估 (默认: 30,000 字符，约合 8,000~10,000 tokens) */
  triggerCharThreshold?: number;
  /** 额外关注点提示 (例如特定工作表或关键变量名) */
  focusNote?: string;
}

/**
 * 压缩结果结构
 */
export interface CompactionResult {
  /** 压缩后供发给模型的历史消息列表 */
  history: ChatHistoryItem[];
  /** 是否实际触发了压缩 */
  isCompacted: boolean;
  /** 提取生成的结构化摘要正文 (未触发压缩时为 undefined) */
  summaryText?: string;
  /** 被压缩折叠的历史消息数量 */
  compactedCount: number;
}

/**
 * 保守预估文本 Token 数量
 * 中英混排与代码混合场景下，1 token 约对应 3.5 个字符
 */
export function estimateTokens(text: string): number {
  if (!text || typeof text !== 'string') return 0;
  return Math.ceil(text.length / 3.5);
}

/**
 * 预估整组历史消息所占用的 Token 总数
 */
export function estimateHistoryTokens(history: ChatHistoryItem[]): number {
  if (!Array.isArray(history)) return 0;
  return history.reduce((sum, item) => sum + estimateTokens(item?.content || ''), 0);
}

/**
 * 检查错误信息是否为上下文溢出异常 (用于三级压缩中的熔断自动重试)
 */
export function isContextOverflowError(errorMsg: string): boolean {
  if (!errorMsg || typeof errorMsg !== 'string') return false;
  const lower = errorMsg.toLowerCase();
  return (
    lower.includes('context_length_exceeded') ||
    lower.includes('context window') ||
    lower.includes('prompt is too long') ||
    lower.includes('maximum context length') ||
    lower.includes('token limit') ||
    lower.includes('too many tokens') ||
    lower.includes('exceeds token limit')
  );
}

/**
 * 执行证据与存证结构 (Immutable Execution Ledger)
 */
export interface ExecutionEvidence {
  snapshotIds: string[];
  failedIncidents: string[];
  verifiedChanges: string[];
}

/**
 * 从消息中提取不可丢失的关键执行事实证据 (快照ID、失败步骤、已核验变更)
 */
export function extractExecutionEvidence(content: string): ExecutionEvidence {
  const snapshotIds: string[] = [];
  const failedIncidents: string[] = [];
  const verifiedChanges: string[] = [];

  if (!content) return { snapshotIds, failedIncidents, verifiedChanges };

  // 1. 提取物理快照 ID (如 snap_20261019_120000, snapshot-xxx)
  const snapRegex = /(?:snap(?:shot)?_[0-9a-zA-Z_]+|snapshot-[0-9a-zA-Z_\-]+)/gi;
  let snapMatch;
  while ((snapMatch = snapRegex.exec(content)) !== null) {
    const id = snapMatch[0];
    if (!snapshotIds.includes(id)) {
      snapshotIds.push(id);
    }
  }

  // 2. 提取执行失败、预检拦截或运行期报错记录
  const lines = content.split('\n');
  for (const line of lines) {
    const trimmed = line.trim();
    if (
      trimmed.length > 4 &&
      trimmed.length < 120 &&
      /(未执行|编译预检未通过|执行异常|运行期异常|执行遇到错误|失败阶段|拦截|Err\s*\d+|1004)/i.test(trimmed) &&
      !trimmed.startsWith('【')
    ) {
      const clean = trimmed.replace(/^[-*•\s\d.⚠️]+/, '').trim();
      if (clean && !failedIncidents.includes(clean) && failedIncidents.length < 6) {
        failedIncidents.push(clean);
      }
    }
  }

  // 3. 提取已核验通过的变更区域与已运行宏
  const verifiedRegex = /(?:区域\s*([a-zA-Z]{1,3}[0-9]{1,7}:[a-zA-Z]{1,3}[0-9]{1,7})\s*验证通过|宏【([^】]+)】已运行完成)/gi;
  let verMatch;
  while ((verMatch = verifiedRegex.exec(content)) !== null) {
    const change = verMatch[0];
    if (!verifiedChanges.includes(change) && verifiedChanges.length < 6) {
      verifiedChanges.push(change);
    }
  }

  return { snapshotIds, failedIncidents, verifiedChanges };
}

/**
 * 从单段消息或代码中结构化提取关键特征 (工作表、宏过程名、关键意图)
 */
export function extractMessageInsights(content: string): {
  subNames: string[];
  sheetNames: string[];
  rules: string[];
} {
  const subNames: string[] = [];
  const sheetNames: string[] = [];
  const rules: string[] = [];

  if (!content) return { subNames, sheetNames, rules };

  // 1. 提取 Sub / Function 过程名
  const procRegex = /(?:Sub|Function)\s+([a-zA-Z0-9_\u4e00-\u9fa5]+)\s*\(/gi;
  let procMatch;
  while ((procMatch = procRegex.exec(content)) !== null) {
    if (procMatch[1] && !subNames.includes(procMatch[1])) {
      subNames.push(procMatch[1]);
    }
  }

  // 2. 提取 Worksheets("...") 出现的表名
  const sheetRegex = /Worksheets\s*\(\s*["']([^"']+)["']\s*\)/gi;
  let sheetMatch;
  while ((sheetMatch = sheetRegex.exec(content)) !== null) {
    if (sheetMatch[1] && !sheetNames.includes(sheetMatch[1])) {
      sheetNames.push(sheetMatch[1]);
    }
  }

  // 3. 提取包含约束/规则的关键行 (如 "保留两位小数"、"千分位"、"标黄"、"去重" 等)
  const lines = content.split('\n');
  for (const line of lines) {
    const trimmed = line.trim();
    if (
      trimmed.length > 5 &&
      trimmed.length < 80 &&
      /(必须|严禁|要求|格式|保留|千分位|标黄|红字|排序|过滤|对齐|Sheet|表头)/i.test(trimmed)
    ) {
      if (!rules.includes(trimmed) && rules.length < 5) {
        rules.push(trimmed.replace(/^[-*•\s\d.]+/, ''));
      }
    }
  }

  return { subNames, sheetNames, rules };
}

/**
 * 构建结构化历史会话摘要检查点 (Compacted Summary)
 * 严格保留执行事实证据（快照ID、失败记录、生效范围）置于不可压缩专区
 */
export function buildStructuredSummary(
  messagesToCompact: ChatHistoryItem[],
  existingSummary?: string,
  focusNote?: string
): string {
  const allSubs: string[] = [];
  const allSheets: string[] = [];
  const allRules: string[] = [];
  const keyMilestones: string[] = [];

  const allSnapshotIds: string[] = [];
  const allFailedIncidents: string[] = [];
  const allVerifiedChanges: string[] = [];

  // 如果存在早先周期的摘要，先行吸收其证据链，保证跨周期多轮压缩绝对不丢快照与失败记录
  if (existingSummary) {
    const prevEvidence = extractExecutionEvidence(existingSummary);
    prevEvidence.snapshotIds.forEach((id) => !allSnapshotIds.includes(id) && allSnapshotIds.push(id));
    prevEvidence.failedIncidents.forEach((f) => !allFailedIncidents.includes(f) && allFailedIncidents.push(f));
    prevEvidence.verifiedChanges.forEach((v) => !allVerifiedChanges.includes(v) && allVerifiedChanges.push(v));
  }

  for (let i = 0; i < messagesToCompact.length; i++) {
    const msg = messagesToCompact[i];
    if (!msg || !msg.content) continue;

    // 提取常规元数据
    const insights = extractMessageInsights(msg.content);
    insights.subNames.forEach((s) => !allSubs.includes(s) && allSubs.push(s));
    insights.sheetNames.forEach((s) => !allSheets.includes(s) && allSheets.push(s));
    insights.rules.forEach((r) => !allRules.includes(r) && allRules.push(r));

    // 提取关键证据事实
    const evidence = extractExecutionEvidence(msg.content);
    evidence.snapshotIds.forEach((id) => !allSnapshotIds.includes(id) && allSnapshotIds.push(id));
    evidence.failedIncidents.forEach((f) => !allFailedIncidents.includes(f) && allFailedIncidents.push(f));
    evidence.verifiedChanges.forEach((v) => !allVerifiedChanges.includes(v) && allVerifiedChanges.push(v));

    if (msg.role === 'user') {
      const cleanPrompt = msg.content
        .split('\n')[0]
        .trim()
        .replace(/^[-*•\s\d.]+/, '');
      if (cleanPrompt.length > 2 && cleanPrompt.length < 50 && !cleanPrompt.startsWith('【')) {
        if (!keyMilestones.includes(cleanPrompt) && keyMilestones.length < 6) {
          keyMilestones.push(cleanPrompt);
        }
      }
    }
  }

  let summary = `【历史会话摘要检查点 (Compacted Session Summary)】\n`;
  summary += `> 本记录由三级智能压缩管道自动折叠生成，包含了前序 ${messagesToCompact.length} 条会话的关键事实、不可压缩证据与约定：\n\n`;

  // 证据账本专区 (Immutable Execution Ledger) - 永远置顶不可丢失
  summary += `【关键执行事实与证据锚点 (Immutable Execution Ledger)】：\n`;
  if (allSnapshotIds.length > 0) {
    summary += `- 关联物理快照 ID (用于撤销与回滚): [${allSnapshotIds.join(', ')}]\n`;
  } else {
    summary += `- 关联物理快照 ID: 暂无快照记录\n`;
  }

  if (allFailedIncidents.length > 0) {
    summary += `- 曾发生过的异常与失败步骤:\n`;
    allFailedIncidents.forEach((fail) => {
      summary += `  * ⚠️ ${fail}\n`;
    });
  }

  if (allVerifiedChanges.length > 0) {
    summary += `- 已核验完成的工作表变更:\n`;
    allVerifiedChanges.forEach((v) => {
      summary += `  * ✓ ${v}\n`;
    });
  }
  summary += `\n`;

  if (keyMilestones.length > 0) {
    summary += `1. **核心任务与执行步骤**：\n`;
    keyMilestones.forEach((m) => {
      summary += `   - ${m}\n`;
    });
  }

  if (allSheets.length > 0) {
    summary += `2. **已涉及操作的工作表**：\n   - [${allSheets.join(', ')}]\n`;
  }

  if (allSubs.length > 0) {
    summary += `3. **已定义/执行的宏过程签名**：\n   - ${allSubs.map((s) => `\`Sub ${s}\``).join(', ')}\n`;
  }

  if (allRules.length > 0) {
    summary += `4. **用户约定的关键格式与业务约束**：\n`;
    allRules.forEach((r) => {
      summary += `   - ${r}\n`;
    });
  }

  if (focusNote && focusNote.trim()) {
    summary += `5. **重点关注项**：\n   - ${focusNote.trim()}\n`;
  }

  if (existingSummary && existingSummary.trim()) {
    summary += `\n---\n【更早周期的历史锚点】：\n${existingSummary.trim()}\n`;
  }

  return summary.trim();
}

/**
 * 历史会话智能压缩主函数
 * 
 * 策略规范：
 * 1. 阈值内无损保留：当总字符/Token 未超限且条数在合理范围时，100% 保持原汁原味历史；
 * 2. 尾部优先保障：永远完整保留最近 maxTailMessages 条真实会话；
 * 3. 头部智能提取：前序超出窗口的消息进行特征提取与结构化沉淀，生成清晰锚点注入历史首部；
 * 4. 证据链不可压缩：快照 ID、错误信息与生效区域通过 Immutable Ledger 强制保留；
 * 5. 角色配对安全：注入的锚点组织为成对的 user 摘要 + assistant 确认，杜绝非法角色序列。
 */
export function compactHistory(
  history: ChatHistoryItem[],
  options: CompactionOptions = {}
): CompactionResult {
  if (!Array.isArray(history) || history.length === 0) {
    return { history: [], isCompacted: false, compactedCount: 0 };
  }

  const maxTailMessages = options.maxTailMessages ?? 6;
  const triggerCharThreshold = options.triggerCharThreshold ?? 30000;

  // 统计字符总量
  const totalChars = history.reduce((sum, item) => sum + (item?.content?.length || 0), 0);

  // 未超过消息条数上限且未超过字符容量阈值：无需压缩，直接放行
  if (history.length <= maxTailMessages && totalChars <= triggerCharThreshold) {
    return {
      history: [...history],
      isCompacted: false,
      compactedCount: 0,
    };
  }

  // 确定切分边界：保留最后 maxTailMessages 条作为活跃尾部
  const splitIndex = Math.max(0, history.length - maxTailMessages);
  const messagesToCompact = history.slice(0, splitIndex);
  const tailMessages = history.slice(splitIndex);

  if (messagesToCompact.length === 0) {
    return {
      history: [...tailMessages],
      isCompacted: false,
      compactedCount: 0,
    };
  }

  // 检查 messagesToCompact 中是否已经存在先前的摘要
  let existingSummary: string | undefined = undefined;
  const filteredMessagesToCompact: ChatHistoryItem[] = [];

  for (const item of messagesToCompact) {
    if (item.content && item.content.includes('【历史会话摘要检查点')) {
      existingSummary = item.content;
    } else {
      filteredMessagesToCompact.push(item);
    }
  }

  // 生成新的结构化检查点摘要 (包含不可压缩证据账本)
  const summaryText = buildStructuredSummary(
    filteredMessagesToCompact.length > 0 ? filteredMessagesToCompact : messagesToCompact,
    existingSummary,
    options.focusNote
  );

  // 组装压缩后的历史：[成对的摘要锚点] + [完整保留的活跃尾部]
  const compactedHistory: ChatHistoryItem[] = [
    {
      role: 'user',
      content: summaryText,
    },
    {
      role: 'assistant',
      content: '已充分理解并锁定上述历史会话检查点、所有已约定的业务规则与执行证据账本。请继续给出下一步具体指令。',
    },
    ...tailMessages,
  ];

  return {
    history: compactedHistory,
    isCompacted: true,
    summaryText,
    compactedCount: messagesToCompact.length,
  };
}
