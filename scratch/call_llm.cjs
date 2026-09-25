const fs = require('fs');
const https = require('https');
const crypto = require('crypto');

const systemPromptFile = process.argv[2];
const userPrompt = process.argv[3];
const outputFile = process.argv[4];
const metaOutputFile = process.argv[5]; // 可选：输出元数据包含 usage、finish_reason、retryCount 等

if (!systemPromptFile || !userPrompt) {
  console.error("Usage: node call_llm.cjs <systemPromptFile> <userPrompt> [outputFile] [metaOutputFile]");
  process.exit(1);
}

const systemPrompt = fs.readFileSync(systemPromptFile, 'utf8');

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
    messages: [
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
    if (!text) return '';
    const openMatch = text.match(/```(?:vba|vb)?\s*/i);
    if (!openMatch) return '';
    const start = openMatch.index + openMatch[0].length;
    const closeIdx = text.indexOf('```', start);
    return closeIdx < 0 ? text.substring(start).trim() : text.substring(start, closeIdx).trim();
  }

  let originalCodeVersion = extractCode(content);
  let regeneratedCodeVersion = null;
  let scopeAudit = {
    hasRisk: false,
    riskSnippet: null,
    riskAdvice: null,
    correctionTriggered: false
  };

  const initialScopeRisk = checkVbaScopeRisk(originalCodeVersion);
  if (initialScopeRisk.hasRisk) {
    console.warn(`[Scope Warning] 检测到全局影响范围失控风险: ${initialScopeRisk.matchedSnippet}`);
    scopeAudit.hasRisk = true;
    scopeAudit.riskSnippet = initialScopeRisk.matchedSnippet;
    scopeAudit.riskAdvice = initialScopeRisk.advice;
    scopeAudit.correctionTriggered = true;

    // 带有明确范围反馈向模型发起针对性重生，绝不静默暗改源码
    const fixPrompt = `${userPrompt}\n\n【重要范围规范修正】：\n上一版代码包含【${initialScopeRisk.matchedSnippet}】。Excel包含171亿个单元格，直接对全表Cells设置边框或底色会导致Excel进程长时间无响应挂起。\n请保留您的所有业务逻辑、公式计算、指标卡片与表格设计，但务必将边框、背景色等格式化语句限定在实际业务数据区域（例如使用 ws.Range(...) 或 Range(ws.Cells(...), ws.Cells(...))），请重新输出一份完整可运行的标准 VBA 代码。`;

    console.log("-> 正在请模型结合范围诊断重新生成完整 VBA (保持算法自由，纠偏失控范围)...");
    const regenPayload = {
      model: model,
      messages: [
        { role: 'system', content: systemPrompt },
        { role: 'assistant', content: content },
        { role: 'user', content: fixPrompt }
      ],
      max_tokens: maxTokens,
      stream: false
    };
    if (thinkingMode === 'disabled') regenPayload.thinking = { type: 'disabled' };
    else if (thinkingMode === 'budget') regenPayload.thinking = { type: 'enabled', budget_tokens: thinkingBudget };

    try {
      const regenStartTime = Date.now();
      const regenResp = await requestLlm(regenPayload);
      const regenElapsed = Date.now() - regenStartTime;
      const regenChoice = regenResp.choices?.[0];
      const regenContent = regenChoice?.message?.content || '';
      const regenUsage = regenResp.usage || {};

      totalCost.prompt_tokens += (regenUsage.prompt_tokens || 0);
      totalCost.completion_tokens += (regenUsage.completion_tokens || 0);
      totalCost.reasoning_tokens += (regenUsage.completion_tokens_details?.reasoning_tokens || 0);
      totalCost.total_tokens += (regenUsage.total_tokens || 0);

      retryHistory.push({
        attempt: 'scope_correction',
        finishReason: regenChoice?.finish_reason,
        contentLength: regenContent.length,
        usage: regenUsage,
        elapsedMs: regenElapsed,
        triggeredBySnippet: initialScopeRisk.matchedSnippet
      });

      if (regenContent) {
        regeneratedCodeVersion = extractCode(regenContent);
        content = regenContent; // 最终候选版
        finishReason = regenChoice?.finish_reason || 'stop';

        // 校验重新生成的版本是否仍包含失控范围
        const secondScopeRisk = checkVbaScopeRisk(regeneratedCodeVersion);
        if (secondScopeRisk.hasRisk) {
          console.warn(`[Scope Warning] 重新生成的版本仍包含失控操作: ${secondScopeRisk.matchedSnippet}。将移除代码块，不予放行执行！`);
          scopeAudit.hasRiskAfterRegen = true;
          scopeAudit.secondRiskSnippet = secondScopeRisk.matchedSnippet;
          // 关键：移除 content 中的代码块，使下游 extractVbaCode 无法提取可执行代码
          // 保留模型的文字说明部分，仅删除 ```vba...``` 围栏
          content = content.replace(/```(?:vba|vb)?\s*[\s\S]*?```/gi,
            '\n\n【安全拦截】模型两次生成的 VBA 均包含对整张工作表的全局格式化操作，已阻止自动执行。请手动简化需求或明确数据区域后重试。\n');
          finishReason = 'scope_blocked';
        } else {
          scopeAudit.hasRiskAfterRegen = false;
          console.log("-> 重新生成成功！第二版代码已纠偏，影响范围已限定在数据区域。");
        }
      }
    } catch (regenErr) {
      console.warn("重新生成请求失败: " + regenErr.message);
    }
  }

  function hashText(t) {
    if (!t) return "";
    return crypto.createHash('sha256').update(t, 'utf8').digest('hex');
  }

  const executedCode = extractCode(content);
  const activeUsage = (regeneratedCodeVersion && retryHistory.length > 0 && retryHistory[retryHistory.length - 1].usage)
    ? retryHistory[retryHistory.length - 1].usage
    : usage;

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
      // 数学关系说明：completion_tokens = reasoning_tokens + content_tokens
      // DeepSeek API: completion_tokens 包含 reasoning_tokens (思考草稿) + content_tokens (正文输出)
      // reasoning_tokens 来自 usage.completion_tokens_details.reasoning_tokens
      // content_tokens = completion_tokens - reasoning_tokens
      mathNote: `completion(${activeCompletionTokens}) = reasoning(${activeReasoningTokens}) + content(${activeContentTokens})`
    },
    accumulatedCost: totalCost,
    scopeAudit: scopeAudit,
    versionControl: {
      firstVersionCode: originalCodeVersion,
      firstVersionHash: hashText(originalCodeVersion),
      regeneratedVersionCode: regeneratedCodeVersion,
      regeneratedVersionHash: hashText(regeneratedCodeVersion),
      executedVersionCode: executedCode,
      executedVersionHash: hashText(executedCode)
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

  console.log(`Success! Attempts: ${attempts}, Finish: ${finishReason}, ContentLen: ${content.length}, ReasoningTokens: ${activeReasoningTokens}, ContentTokens: ${activeContentTokens}, TotalTokens: ${activeUsage.total_tokens}`);
})();
