import { loadLlmConfig, type LlmConfig } from './config';

export interface ParsedStreamOutput {
  explanation: string;
  vbaCode: string;
  hasCode: boolean;
}

/**
 * 实时解析大模型流式输出：将纯文本说明与 VBA 代码块彻底剥离
 * 保证聊天气泡内只显示一两句简报，代码完全收纳到折叠卡片中
 */
export function parseStreamOutput(raw: string): ParsedStreamOutput {
  if (!raw) {
    return { explanation: '', vbaCode: '', hasCode: false };
  }

  // 匹配 ```vba 或 ```vb 或 ``` 开头
  const fenceRegex = /```(?:vba|vb)?\s*/i;
  const match = raw.match(fenceRegex);

  if (!match || match.index === undefined) {
    // 检查是否没有 markdown 围栏但直接写了 Sub RunTask
    const subMatch = raw.match(/(?:Public\s+|Private\s+)?Sub\s+/i);
    if (subMatch && subMatch.index !== undefined) {
      const explanation = raw.slice(0, subMatch.index).trim();
      const vbaCode = raw.slice(subMatch.index).trim();
      return { explanation, vbaCode, hasCode: true };
    }
    return { explanation: raw.trim(), vbaCode: '', hasCode: false };
  }

  // 代码块之前的内容作为自然语言说明
  const explanation = raw.slice(0, match.index).trim();
  const codeStartIndex = match.index + match[0].length;
  const remaining = raw.slice(codeStartIndex);

  // 检查是否有结束闭合 ```
  const endFenceIndex = remaining.indexOf('```');
  let vbaCode = '';
  if (endFenceIndex !== -1) {
    vbaCode = remaining.slice(0, endFenceIndex).trim();
  } else {
    // 仍在流式吐出中，取当前所有未闭合代码
    vbaCode = remaining.trim();
  }

  return { explanation, vbaCode, hasCode: true };
}

export function hasValidVbaTask(code: string): boolean {
  if (!code || typeof code !== 'string') return false;
  const trimmed = code.trim();
  // 必须同时严格包含 Sub 过程定义与闭合的 End Sub
  const hasSub = /(?:Public\s+|Private\s+)?Sub\s+[a-zA-Z0-9_]+\s*\(/i.test(trimmed);
  const hasEndSub = /End\s+Sub/i.test(trimmed);
  return hasSub && hasEndSub;
}

export function extractVbaCode(content: string): string {
  if (!content) return '';
  let code = '';
  // 优先匹配标准 markdown 围栏
  const match = content.match(/```(?:vba|vb)?\s*([\s\S]*?)\s*```/i);
  if (match) {
    code = match[1].trim();
  } else {
    // 次选匹配闭合的 Sub ... End Sub
    const subMatch = content.match(/(?:Public\s+|Private\s+)?Sub\s+[\s\S]*?End\s+Sub/i);
    if (subMatch) {
      code = subMatch[0].trim();
    }
  }

  // 移除非法行前标记
  if (code) {
    code = code.replace(/^[ \t]*[*\-•][ \t]+/gm, '');
  }

  return code;
}

export function buildSystemPrompt(workbookName: string, sheets: string[]): string {
  return `你是一名精通 Microsoft Windows 桌面 Excel 和 VBA 自动化的专业 AI 助手。
当前连接的工作簿: "${workbookName || '活动工作簿'}"。
当前工作簿包含的工作表: [${sheets.join(', ')}]。

【核心交互机制 - 必须严格区分以下两种场景】：

场景 1. 【自然语言对话 / 咨询 / 闲聊 / 答疑】
- 当用户询问你的身份（如“你是谁”、“你是？”）、能力介绍、Excel 使用技巧、函数公式解释或日常打招呼时；
- 必须使用专业、亲切、通俗易懂的中文直接解答；
- 【极其重要】：此类对话中绝对不要输出任何 \`\`\`vba 代码块，也绝不要在结尾附加空代码或反引号。保持对话自然顺畅。

场景 2. 【Excel 操作 / 自动化指令 / 表格生成与美化】
- 当用户要求新建表格、写入数据、计算汇总、格式调整、数据整理等具体任务时；
- 【准确理解中文业务习惯】：
  例如用户要求“九九乘法表”，必须生成标准的乘法口诀表达式（格式形如: "1×1=1"、i & "×" & j & "=" & (i*j)），阶梯形或 9×9 完整矩阵，搭配优雅的商务排版（经典深蓝/翠绿表头、白色加粗文字、居中对齐、浅色背景与清晰边框），绝非仅填入纯数字！
- 【VBA 规范】：
  1. 固定使用纯英文过程入口: \`Sub RunTask()\`，必须包含完整的闭合 \`End Sub\`。
  2. 必须使用 \`ActiveWorkbook\` 操作当前活动文档，严禁使用 \`ThisWorkbook\`（避免产生加载项工程指向偏差）。
  3. 严禁调用 MsgBox，严禁编写 Application.ScreenUpdating 等环境控制语句（C# 宿主已全局接管）。
  4. 格式美化请采用区域整块设置（例如: ws.Range("D1:L9").Interior.Color = ...），严禁逐格循环涂色，总代码控制在 40 行内。
- 【回复格式】：
  先用 1 句话中文概括将要执行的操作（例如：“正在新建九九乘法表并应用商务美化排版...”），随后直接给出标准的 \`\`\`vba ... \`\`\` 代码块。`;
}

export interface ChatHistoryItem {
  role: 'user' | 'assistant';
  content: string;
}

export async function callLlmStream(
  prompt: string,
  workbookName: string,
  sheets: string[],
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

  const systemPrompt = buildSystemPrompt(workbookName, sheets);

  // 整理历史上下文（最多取最近 4 条，避免 token 过载）
  const cleanHistory = history.slice(-4).map((h) => ({
    role: h.role,
    // 过滤掉历史消息中可能残留的长篇代码围栏，只保留对话说明
    content: h.content.replace(/```(?:vba|vb)?[\s\S]*?```/gi, '').trim(),
  })).filter((h) => !!h.content);

  const payload = {
    model: config.model,
    messages: [
      { role: 'system', content: systemPrompt },
      ...cleanHistory,
      { role: 'user', content: prompt },
    ],
    temperature: 0.2,
    max_tokens: 4096,
    stream: true,
  };

  const headers: Record<string, string> = {
    'Content-Type': 'application/json',
  };

  if (config.apiKey) {
    headers['Authorization'] = `Bearer ${config.apiKey}`;
  }

  // 超时控制器：首包 30 秒超时，每块 15 秒心跳超时守护
  const abortController = new AbortController();
  let chunkTimer: any = null;

  const resetHeartbeat = () => {
    if (chunkTimer) clearTimeout(chunkTimer);
    chunkTimer = setTimeout(() => {
      abortController.abort(new Error('LLM_HEARTBEAT_TIMEOUT'));
    }, 15000); // 15秒无数据包则判定超时中断
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
          } catch {
            // 忽略单帧 JSON 解析容错
          }
        }
      }
    }
  } catch (streamErr: any) {
    if (streamErr.name === 'AbortError' || streamErr.message?.includes('TIMEOUT')) {
      // 若已有部分有效代码，保留以避免用户全盘重来
      if (fullText.includes('Sub RunTask')) {
        return fullText;
      }
      throw new Error('流式生成传输中断（15秒未收到新响应），请点击重试。');
    }
    throw streamErr;
  } finally {
    if (chunkTimer) clearTimeout(chunkTimer);
  }

  return fullText;
}
