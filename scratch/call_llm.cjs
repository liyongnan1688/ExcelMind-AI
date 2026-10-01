const fs = require('fs');
const https = require('https');
const crypto = require('crypto');

const systemPromptFile = process.argv[2];
let userPrompt = process.argv[3];
const outputFile = process.argv[4];
const metaOutputFile = process.argv[5]; // 可选：输出元数据包含 usage、finish_reason、retryCount 等

if (!systemPromptFile || !userPrompt) {
  console.error("Usage: node call_llm.cjs <systemPromptFile> <userPrompt|userPromptFile> [outputFile] [metaOutputFile]");
  process.exit(1);
}

let messages = null;
let systemPrompt = "";
if (fs.existsSync(systemPromptFile)) {
  try {
    const raw = fs.readFileSync(systemPromptFile, 'utf8');
    const parsed = JSON.parse(raw);
    if (Array.isArray(parsed)) {
      messages = parsed;
    } else {
      systemPrompt = raw;
    }
  } catch (e) {
    systemPrompt = fs.readFileSync(systemPromptFile, 'utf8');
  }
} else {
  systemPrompt = systemPromptFile;
}

if (fs.existsSync(userPrompt)) {
  try {
    userPrompt = fs.readFileSync(userPrompt, 'utf8');
  } catch (e) {}
}

// 读取 LocalStorage 配置
const logPath = process.env.LOCALAPPDATA + '\\LeeExcel\\WebView2Profile\\EBWebView\\Default\\Local Storage\\leveldb\\000003.log';
let config = { provider: 'deepseek', model: 'deepseek-flash', baseUrl: 'https://api.deepseek.com/v1', apiKey: '' };

if (fs.existsSync(logPath)) {
  const logContent = fs.readFileSync(logPath, 'utf8');
  const matches = logContent.match(/\{[^{}]*provider[^{}]*\}/g);
  if (matches && matches.length > 0) {
    try {
      config = JSON.parse(matches[matches.length - 1]);
    } catch (e) {}
  }
}

const apiKey = config.apiKey;
const model = config.model || 'deepseek-flash';
const baseUrl = (config.baseUrl || 'https://api.deepseek.com/v1').replace(/\/+$/, '');

// 根据用户配置或默认值确定生成预算与思考参数
// 充分保证生成预算，避免复杂任务被截断
const maxTokens = config.maxTokens ? parseInt(config.maxTokens) : 16384;
const thinkingMode = process.env.THINKING_MODE || config.thinkingMode || 'auto'; // 'auto' | 'disabled' | 'budget'
const thinkingBudget = process.env.THINKING_BUDGET ? parseInt(process.env.THINKING_BUDGET) : (config.thinkingBudget ? parseInt(config.thinkingBudget) : 2048);

function makePayload(attempt) {
  const p = {
    model: model,
    messages: messages ? messages : [
      { role: 'system', content: systemPrompt },
      { role: 'user', content: userPrompt }
    ],
    max_tokens: maxTokens,
    stream: false
  };

  if (thinkingMode === 'disabled') {
    p.thinking = { type: 'disabled' };
  } else if (thinkingMode === 'budget') {
    p.thinking = { type: 'enabled', budget_tokens: thinkingBudget };
  }
  // 若为 'auto'，则不主动传入 thinking 字段，完全尊重模型端默认能力

  if (typeof config.temperature === 'number' && !isNaN(config.temperature)) {
    p.temperature = config.temperature;
  }

  return p;
}

function requestLlm(payload) {
  return new Promise((resolve, reject) => {
    const url = new URL(baseUrl + '/chat/completions');
    const req = https.request(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + apiKey
      },
      timeout: 180000 // 充足的 3 分钟超时
    }, (res) => {
      let body = '';
      res.on('data', chunk => body += chunk);
      res.on('end', () => {
        try {
          if (res.statusCode !== 200) {
            return reject(new Error(`HTTP error ${res.statusCode}: ${body}`));
          }
          const json = JSON.parse(body);
          if (json.error) {
            return reject(new Error(`API error: ${json.error.message || JSON.stringify(json.error)}`));
          }
          resolve(json);
        } catch (err) {
          reject(new Error(`Failed to parse response: ${err.message}`));
        }
      });
    });

    req.on('error', (e) => reject(new Error(`HTTP request failed: ${e.message}`)));
    req.write(JSON.stringify(payload));
    req.end();
  });
}

(async () => {
  const MAX_RETRIES = 2; // 最多自动重试 2 次
  let attempts = 0;
  let finalJson = null;
  let totalCost = { prompt_tokens: 0, completion_tokens: 0, reasoning_tokens: 0, total_tokens: 0 };
  let retryHistory = [];

  while (attempts <= MAX_RETRIES) {
    attempts++;
    const payload = makePayload(attempts);
    const startTime = Date.now();
    try {
      const resp = await requestLlm(payload);
      const elapsed = Date.now() - startTime;
      const choice = resp.choices?.[0];
      const finishReason = choice?.finish_reason || 'unknown';
      const content = choice?.message?.content || '';
      const reasoningContent = choice?.message?.reasoning_content || '';
      const usage = resp.usage || {};

      totalCost.prompt_tokens += (usage.prompt_tokens || 0);
      totalCost.completion_tokens += (usage.completion_tokens || 0);
      totalCost.reasoning_tokens += (usage.completion_tokens_details?.reasoning_tokens || 0);
      totalCost.total_tokens += (usage.total_tokens || 0);

      retryHistory.push({
        attempt: attempts,
        finishReason: finishReason,
        contentLength: content.length,
        reasoningLength: reasoningContent.length,
        usage: usage,
        elapsedMs: elapsed
      });

      // 截断判定：若 finish_reason 为 length，或 content 为空且被截断，绝不拼装执行半截代码！
      if (finishReason === 'length' || (!content && reasoningContent)) {
        console.warn(`[Attempt ${attempts}] Response truncated (finish_reason=${finishReason}, content_len=${content.length}). Will retry from scratch if attempts remain.`);
        if (attempts <= MAX_RETRIES) {
          continue;
        } else {
          finalJson = resp;
          break;
        }
      }

      // 获取到有效内容
      finalJson = resp;
      break;
    } catch (err) {
      console.warn(`[Attempt ${attempts}] Request failed: ${err.message}`);
      retryHistory.push({
        attempt: attempts,
        error: err.message,
        elapsedMs: Date.now() - startTime
      });
      if (attempts > MAX_RETRIES) {
        console.error(`All ${MAX_RETRIES + 1} attempts failed.`);
        process.exit(1);
      }
    }
  }

  if (!finalJson) {
    console.error("No valid response obtained.");
    process.exit(1);
  }

  let choice = finalJson.choices?.[0];
  let content = choice?.message?.content || '';
  let finishReason = choice?.finish_reason || 'unknown';
  let usage = finalJson.usage || {};

  // ==========================================
  // 执行前影响范围检查与失控纠偏（不静默篡改源码）
  // ==========================================
  function checkVbaScopeRisk(code) {
    if (!code) return { hasRisk: false };
    // 1. 全表单元格边框、背景色或条件格式
    const allCellsFormatRegex = /(?:(?:ws|ActiveSheet|Worksheets\([^)]+\)|targetWb\.ActiveSheet)\s*\.\s*Cells|(?<!\.)\bCells)\s*\.\s*(?:Borders|Interior|FormatConditions)\b/i;
    const match1 = code.match(allCellsFormatRegex);
    if (match1) {
      return {
        hasRisk: true,
        riskType: 'ALL_CELLS_FORMAT',
        matchedSnippet: match1[0],
        advice: `代码包含对整张工作表全部单元格的格式化操作（${match1[0]}）。Excel包含逾171亿个单元格，对Cells直接设置边框或背景色会耗尽系统资源导致Excel卡死。请将边框与背景色限定在实际业务数据区域（如 ws.Range(...) 或 Range(ws.Cells(r1, c1), ws.Cells(r2, c2))）。`
      };
    }

    // 2. 全表单元格清空格式或删除
    const allCellsClearRegex = /(?:(?:ws|ActiveSheet|Worksheets\([^)]+\)|targetWb\.ActiveSheet)\s*\.\s*Cells|(?<!\.)\bCells)\s*\.\s*(?:ClearFormats|Delete)\b/i;
    const match2 = code.match(allCellsClearRegex);
    if (match2) {
      return {
        hasRisk: true,
        riskType: 'ALL_CELLS_CLEAR',
        matchedSnippet: match2[0],
        advice: `代码包含对全表单元格的批量删除或清格式操作（${match2[0]}）。请改为仅对数据表使用区域（ws.UsedRange）或指定数据Range操作。`
      };
    }

    // 3. 整列/整行批量格式化（整列含104万行，批量设置边框极易引发性能灾难）
    const allColumnsFormatRegex = /(?:(?:ws|ActiveSheet|Worksheets\([^)]+\))\s*\.\s*)?(?:Columns(?:\([^)]+\))?|Rows(?:\([^)]+\))?)\s*\.\s*(?:Borders|Interior|FormatConditions)\b/i;
    const match3 = code.match(allColumnsFormatRegex);
    if (match3) {
      return {
        hasRisk: true,
        riskType: 'ALL_COLUMNS_FORMAT',
        matchedSnippet: match3[0],
        advice: `代码尝试对整列/整行全部单元格设置边框或背景色（${match3[0]}）。整列包含104万个单元格，请仅在有效数据行范围内设置格式。`
      };
    }

    return { hasRisk: false };
  }

  function extractCode(text) {
    if (!text || typeof text !== 'string') return { code: '', isTruncated: false, error: '内容为空' };
    const trimmed = text.trim();
    const fenceRegex = /```(?:vba|vb)?\s*([\s\S]*?)(?:```|$)/i;
    const match = trimmed.match(fenceRegex);

    if (match && match.index !== undefined) {
      const codePart = trimmed.slice(match.index + 3);
      if (!codePart.includes('```')) {
        return {
          code: match[1],
          isTruncated: true,
          error: '代码围栏未闭合 (``` 截断)'
        };
      }
      return { code: match[1], isTruncated: false };
    }

    // 纯 VBA 源码形态（无围栏）
    const isVbaText = /^\s*(?:Option\s+Explicit|Attribute\s+|'(?:[^\r\n]*)|(?:\b(?:Public\s+|Private\s+)?(?:Sub|Function)\b))/im.test(trimmed);
    if (isVbaText) {
      const hasSubStart = /(?:^|\n)\s*(?:Public\s+|Private\s+)?Sub\s+/i.test(trimmed);
      const hasSubEnd = /(?:^|\n)\s*End\s+Sub\b/i.test(trimmed);
      const hasFnStart = /(?:^|\n)\s*(?:Public\s+|Private\s+)?Function\s+/i.test(trimmed);
      const hasFnEnd = /(?:^|\n)\s*End\s+Function\b/i.test(trimmed);
      if ((hasSubStart && !hasSubEnd) || (hasFnStart && !hasFnEnd)) {
        return { code: trimmed, isTruncated: true, error: '纯 VBA 源码未闭合 (缺少 End Sub / End Function)' };
      }
      return { code: trimmed, isTruncated: false };
    }

    return { code: '', isTruncated: false, error: '未检测到合法的纯 VBA 源码或代码围栏' };
  }

  const extractResult = extractCode(content);

  // 严格截断阻断：若 finishReason 为 length 或 extractResult.isTruncated，拒绝写出代码并退出
  if (finishReason === 'length' || extractResult.isTruncated) {
    const errMsg = finishReason === 'length' 
      ? `模型响应达到 Token 上限被硬截断 (finish_reason=length)` 
      : `代码截断异常 (${extractResult.error})`;
    console.error(`[Fatal Truncation] ${errMsg}。已安全中止，绝不向执行器提供残缺代码！`);
    if (metaOutputFile) {
      fs.writeFileSync(metaOutputFile, JSON.stringify({
        status: 'FAILED_TRUNCATED',
        error: errMsg,
        finishReason: finishReason,
        isTruncated: true,
        usage: usage
      }, null, 2), 'utf8');
    }
    process.exit(1);
  }

  let scopeAudit = {
    hasRisk: false,
    riskSnippet: null,
    riskAdvice: null
  };

  if (extractResult.code) {
    const initialScopeRisk = checkVbaScopeRisk(extractResult.code);
    if (initialScopeRisk.hasRisk) {
      console.warn(`[Scope Warning] 检测到全局影响范围失控风险: ${initialScopeRisk.matchedSnippet}`);
      scopeAudit.hasRisk = true;
      scopeAudit.riskSnippet = initialScopeRisk.matchedSnippet;
      scopeAudit.riskAdvice = initialScopeRisk.advice;
    }
  }

  function hashText(t) {
    if (!t) return "";
    return crypto.createHash('sha256').update(t, 'utf8').digest('hex');
  }

  const activeUsage = usage;
  const activeReasoningTokens = activeUsage.completion_tokens_details?.reasoning_tokens || 0;
  const activeCompletionTokens = activeUsage.completion_tokens || 0;
  const activeContentTokens = activeCompletionTokens >= activeReasoningTokens
    ? (activeCompletionTokens - activeReasoningTokens)
    : activeCompletionTokens;

  const metadata = {
    provider: config.provider,
    model: config.model,
    maxTokens: maxTokens,
    thinkingMode: thinkingMode,
    finishReason: finishReason,
    totalAttempts: attempts,
    retryCount: attempts - 1,
    retryHistory: retryHistory,
    contentLength: content.length,
    rawUsage: activeUsage,
    tokenAccounting: {
      promptTokens: activeUsage.prompt_tokens || 0,
      completionTokens: activeCompletionTokens,
      reasoningTokens: activeReasoningTokens,
      contentTokens: activeContentTokens,
      totalTokens: activeUsage.total_tokens || 0,
      mathNote: `completion(${activeCompletionTokens}) = reasoning(${activeReasoningTokens}) + content(${activeContentTokens})`
    },
    accumulatedCost: totalCost,
    scopeAudit: scopeAudit,
    evidenceSeparation: {
      rawModelResponseLength: content.length,
      extractedModelCode: extractResult.code || "",
      extractedModelCodeLength: (extractResult.code || "").length,
      extractedModelCodeHash: hashText(extractResult.code || ""),
      isTruncated: extractResult.isTruncated,
      extractionError: extractResult.error || null,
      note: "本探针仅记录模型回复与提取代码。实际执行代码(executedVbaCode)、包装器(wrapperCode)与执行结果必须由宿主执行器(VbaRunner)回报提供，绝不提前假定执行。"
    }
  };

  if (outputFile) {
    fs.writeFileSync(outputFile, content, 'utf8');
  } else {
    process.stdout.write(content);
  }

  if (metaOutputFile) {
    fs.writeFileSync(metaOutputFile, JSON.stringify(metadata, null, 2), 'utf8');
  }

  console.log(`Success! Attempts: ${attempts}, Finish: ${finishReason}, ContentLen: ${content.length}, CodeLen: ${(extractResult.code || '').length}, TotalTokens: ${activeUsage.total_tokens}`);
})();
