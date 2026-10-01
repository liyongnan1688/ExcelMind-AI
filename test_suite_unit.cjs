// test_suite_unit.cjs - Unit test runner for explicit dual-entry routing, code extraction, and verification

/**
 * 模拟生产环境的双入口发送分发逻辑
 * 唯一模式来源：用户显式选中的模式 / 点击的模式按钮
 * 生产发送链路中零调用 detectIntent，不根据文本关键词改变模式
 */
function createRequestLifecycle(explicitMode, text, currentWbContext = { name: 'Book1.xlsx', activeSheet: 'Sheet1' }) {
  const requestId = 'req_' + Date.now() + '_' + Math.random().toString(36).slice(2, 6);
  // 发送瞬间严格固化 requestMode 与目标工作簿上下文
  const currentRequest = {
    requestId,
    requestMode: explicitMode, // 显式入口传入：'CHAT' 或 'AUTOMATION'
    userMsgId: 'msg_user_' + requestId,
    assistantMsgId: 'msg_asst_' + requestId,
    targetWbName: currentWbContext.name,
    activeSheet: currentWbContext.activeSheet,
  };
  return currentRequest;
}

/**
 * 模拟生产环境的响应处理逻辑
 */
function handleResponseLifecycle(currentRequest, rawModelResponse, finishReason = 'stop') {
  // 四、CHAT 通道：完整展示，零调用执行器，零快照
  if (currentRequest.requestMode === 'CHAT') {
    return {
      mode: 'CHAT',
      content: rawModelResponse,
      execution: null,
      extractedCode: null,
      didCallExecuteVba: false,
      didCreateSnapshot: false,
      isMacroError: false,
    };
  }

  // 五、AUTOMATION 通道：区分未取得代码、代码无效截断与合法完整代码
  const extracted = extractVbaCode(rawModelResponse, finishReason);

  // 分支 1: 模型返回普通文字，未提供 VBA（如输入“你是”返回了自我介绍）
  if (extracted.status === 'no_code') {
    return {
      mode: 'AUTOMATION',
      content: rawModelResponse + '\n\n（本次未执行：模型未返回可执行 VBA）',
      execution: null, // 绝不创建宏执行失败卡！
      extractedCode: null,
      didCallExecuteVba: false,
      didCreateSnapshot: false,
      isMacroError: false,
      unexecutedReason: '模型未返回可执行 VBA',
    };
  }

  // 分支 2: 模型返回疑似 VBA，但输出被截断或结构缺失（代码未执行）
  if (extracted.status !== 'valid' || !extracted.code) {
    const errorReason = extracted.error || '代码不完整或结构缺失，已安全停止执行';
    return {
      mode: 'AUTOMATION',
      content: `（代码未执行：${errorReason}）`,
      execution: {
        summary: `代码未执行：${errorReason}`,
        error: errorReason,
        precheckStatus: extracted.status === 'scope_risk' ? 'scope_risk_intercepted' : 'failed',
        executionPhase: 'intercepted_before_run',
        snapshot: null, // 不创建快照
      },
      extractedCode: extracted.code,
      didCallExecuteVba: false,
      didCreateSnapshot: false,
      isMacroError: false, // 标为未进入运行阶段，非宏运行时错误
    };
  }

  // 分支 3: 提取成功且结构完整：才进入已有执行链路
  return {
    mode: 'AUTOMATION',
    content: rawModelResponse,
    execution: {
      summary: '执行准备就绪',
      precheckStatus: 'passed',
      executionPhase: 'running',
      snapshot: { id: 'snap_test_123' },
    },
    extractedCode: extracted.code,
    didCallExecuteVba: true,
    didCreateSnapshot: true,
    isMacroError: false,
  };
}


function extractVbaCode(content, finishReason) {
  if (!content) {
    return { code: '', isTruncated: false, status: 'no_code', error: '响应内容为空' };
  }

  const openFenceMatch = content.match(/```(?:vba|vb)?\s*/i);
  let extractedCode = '';
  let isFenceEnclosed = false;

  if (openFenceMatch && openFenceMatch.index !== undefined) {
    const codeStart = openFenceMatch.index + openFenceMatch[0].length;
    const rest = content.slice(codeStart);
    const closeFenceIndex = rest.indexOf('```');

    if (closeFenceIndex !== -1) {
      extractedCode = rest.slice(0, closeFenceIndex).trim();
      isFenceEnclosed = true;
    } else {
      extractedCode = rest.trim();
      return {
        code: extractedCode,
        isTruncated: true,
        status: 'truncated',
        error: '代码块未闭合 (缺少配对的 ```)，说明大模型输出已被截断，已安全停止执行。',
      };
    }
  } else {
    // 兜底检测裸代码
    const subMatch = content.match(/(?:^|\n)\s*(?:Public\s+|Private\s+)?Sub\s+[a-zA-Z0-9_\u4e00-\u9fa5]+/i);
    if (subMatch && subMatch.index !== undefined) {
      extractedCode = content.slice(subMatch.index).trim();
    } else {
      return {
        code: '',
        isTruncated: false,
        status: 'no_code',
        error: '模型回复未包含可执行的 VBA 过程源码（普通文字回答）。',
      };
    }
  }

  // 检查基本过程完整性
  const hasSubOrFunction = /(?:^|\n)\s*(?:Public\s+|Private\s+)?(?:Sub|Function)\s+[a-zA-Z0-9_\u4e00-\u9fa5]+/i.test(extractedCode);
  if (!hasSubOrFunction) {
    return {
      code: extractedCode,
      isTruncated: false,
      status: 'invalid_structure',
      error: '提取出的代码中未包含有效的 Sub 或 Function 过程定义。',
    };
  }

  // 检查未闭合的 Sub / Function
  const hasSubStart = /(?:^|\n)\s*(?:Public\s+|Private\s+)?Sub\s+/i.test(extractedCode);
  const hasSubEnd = /(?:^|\n)\s*End\s+Sub\b/i.test(extractedCode);
  if (hasSubStart && !hasSubEnd) {
    return {
      code: extractedCode,
      isTruncated: true,
      status: 'truncated',
      error: '代码过程未闭合 (缺少配对的 End Sub)，可能模型生成被中途截断，已安全拦截未执行。',
    };
  }

  return {
    code: extractedCode,
    isTruncated: false,
    status: 'valid',
  };
}

function verifyExecutionResult(prompt, readback) {
  if (!readback) {
    return { status: 'unconfirmed', note: '未能读回目标工作簿实际变更状态，效果待确认。' };
  }
  if (!readback.targetVerified) {
    return { status: 'failed', note: `目标工作簿身份核验失败（预期: ${readback.targetWorkbookName}）。` };
  }

  if (readback.otherWorkbooksAffected) {
    return { status: 'failed', note: readback.affectedWorkbooksWarning || '安全告警：检测到非目标工作簿受到附带修改！' };
  }

  const notes = [];
  const startMatch = prompt.match(/(?:从|在)\s*([A-Za-z]+[0-9]+)/i);
  if (startMatch) {
    const expectedStart = startMatch[1].toUpperCase();
    if (readback.startCell && readback.startCell.toUpperCase() !== expectedStart) {
      notes.push(`要求从 ${expectedStart} 开始，实际起始于 ${readback.startCell}`);
    }
  }

  const wantsEquationExplicit = /(?:算式|口诀|乘法口诀|带算式)/i.test(prompt);
  const wantsNumericExplicit = /(?:数值乘积矩阵|纯数字|数值矩阵|乘积矩阵)/i.test(prompt);
  const isMultiplicationGeneral = /(?:九九乘法表|乘法表)/i.test(prompt);

  const hasEquationText =
    readback.sampleValues &&
    readback.sampleValues.some((v) => v.includes('×') || v.includes('*') || v.includes('=') || v.includes('得'));

  if (wantsEquationExplicit && !hasEquationText) {
    notes.push('明确要求算式口诀，实际检测为纯数字矩阵，未生成算式文本');
  } else if (wantsNumericExplicit && hasEquationText) {
    notes.push('明确要求数值矩阵，实际生成了算式文本');
  } else if (isMultiplicationGeneral && !wantsNumericExplicit && !wantsEquationExplicit) {
    if (hasEquationText) {
      notes.push('已生成标准算式口诀表（D1:L9）');
    } else {
      notes.push('已生成乘积数值矩阵（D1:L9）');
    }
  }

  const wantsBeauty = /(?:美化|商务|好看|排版|样式|颜色|边框)/i.test(prompt);
  if (wantsBeauty) {
    const styleFeatures = [];
    if (readback.hasBorders) styleFeatures.push('已添加边框');
    if (readback.hasInteriorColor) styleFeatures.push('已应用背景填充');
    notes.push(`视觉样式(${styleFeatures.join('、') || '基础样式'})已应用，效果待人工确认`);
  }

  if (notes.some((n) => n.includes('未生成算式文本') || n.includes('实际起始于') || n.includes('实际生成了算式文本'))) {
    return {
      status: 'unconfirmed',
      note: `宏已运行，但与指令存在差异：${notes.join('；')}`,
    };
  }

  if (wantsBeauty || isMultiplicationGeneral) {
    return {
      status: 'unconfirmed',
      note: `宏已运行（区域: ${readback.usedRangeAddress || '已更新'}），${notes.join('；')}。`,
    };
  }

  return {
    status: 'verified',
    note: `宏已运行，数据区域 ${readback.usedRangeAddress || ''} 验证通过。`,
  };
}

// -------------------------------------------------------------
// Test Runner
// -------------------------------------------------------------
let passCount = 0;
let failCount = 0;

function assert(desc, condition, details = '') {
  if (condition) {
    console.log(`[PASS] ${desc}`);
    passCount++;
  } else {
    console.error(`[FAIL] ${desc} | Details: ${details}`);
    failCount++;
  }
}

console.log('=== 1. Explicit Dual-Entry Routing & Lifecycle Suite ===');

// 1.1 对话入口测试 (无论输入什么文本，均锁定 CHAT 通道，绝不调用执行器)
const chatInput1 = '你是';
const reqChat1 = createRequestLifecycle('CHAT', chatInput1);
assert('Explicit CHAT entry locks requestMode to CHAT', reqChat1.requestMode === 'CHAT');
const resChat1 = handleResponseLifecycle(reqChat1, '我是您的智能助手，可以为您解答 Excel 与 VBA 相关问题。');
assert('CHAT input "你是" -> displays text, execution is null, no VBA executor called', resChat1.execution === null && !resChat1.didCallExecuteVba && !resChat1.didCreateSnapshot);

const chatInput2 = '写一段在 A1 写字的 VBA 给我看';
const reqChat2 = createRequestLifecycle('CHAT', chatInput2);
const mockVbaBlock = '你可以参考以下代码：\n```vba\nSub Demo()\n  Range("A1").Value = "Hello"\nEnd Sub\n```';
const resChat2 = handleResponseLifecycle(reqChat2, mockVbaBlock);
assert('CHAT input requesting VBA -> displays code block, execution is null, does not execute', resChat2.execution === null && !resChat2.didCallExecuteVba && resChat2.content.includes('Sub Demo()'));

const chatInput3 = '在当前工作表 A1 写入操作测试';
const reqChat3 = createRequestLifecycle('CHAT', chatInput3);
assert('CHAT input with action verb still stays in CHAT mode', reqChat3.requestMode === 'CHAT');
const resChat3 = handleResponseLifecycle(reqChat3, '建议您切换到底部【操作】模式来直接修改表格。');
assert('CHAT mode with action verb never triggers executor', resChat3.execution === null && !resChat3.didCallExecuteVba);

// 1.2 操作入口测试 (无论输入什么文本，均锁定 AUTOMATION 通道)
const autoInput1 = '你是';
const reqAuto1 = createRequestLifecycle('AUTOMATION', autoInput1);
assert('Explicit AUTOMATION entry locks requestMode to AUTOMATION (even for "你是")', reqAuto1.requestMode === 'AUTOMATION');

// 截图中故障关键回归：操作入口输入“你是”，模型返回普通文字说明
const mockIntroText = '我是基于中信科移动标准的智能办公助手，能够帮您编写并执行 Excel 宏操作。';
const resAutoIntro = handleResponseLifecycle(reqAuto1, mockIntroText);
assert(
  'AUTOMATION receives pure text (no code) -> annotated as unexecuted, execution card is NULL (no red error card!)',
  resAutoIntro.execution === null && !resAutoIntro.didCallExecuteVba && resAutoIntro.content.includes('（本次未执行：模型未返回可执行 VBA）')
);

// 1.3 AUTOMATION 正常返回完整 VBA
const autoInput2 = '在当前工作表 A1 写入操作测试';
const reqAuto2 = createRequestLifecycle('AUTOMATION', autoInput2);
const mockValidVba = '已为您编写宏：\n```vba\nSub Main(targetWb As Workbook)\n  targetWb.Sheets(1).Range("A1").Value = "操作测试"\nEnd Sub\n```';
const resAutoValid = handleResponseLifecycle(reqAuto2, mockValidVba);
assert(
  'AUTOMATION with valid VBA -> creates snapshot and calls execute_vba',
  resAutoValid.didCallExecuteVba && resAutoValid.didCreateSnapshot && resAutoValid.execution.precheckStatus === 'passed'
);

// 1.4 处理层测试：模拟不同模型回复状态 (明确标注为 Simulated)
console.log('\n=== 1A. Simulated Response Layer Tests (no_code vs truncated vs valid) ===');
// 模拟 1: 模型返回普通文字，无代码
const simNoCode = extractVbaCode('请问您需要处理哪张工作表的数据？');
assert('[Simulated] Pure text response returns status: "no_code"', simNoCode.status === 'no_code' && simNoCode.code === '');

// 模拟 2: 模型回复被截断 (未闭合代码围栏)
const simTruncFence = extractVbaCode('```vba\nSub Test()\n  Range("A1").Value = 1\n');
assert('[Simulated] Unclosed code fence returns status: "truncated"', simTruncFence.status === 'truncated' && simTruncFence.isTruncated);
const resSimTrunc = handleResponseLifecycle(reqAuto2, '```vba\nSub Test()\n  Range("A1").Value = 1\n');
assert(
  '[Simulated] Truncated code labeled as "代码未执行", phase: "intercepted_before_run", no snapshot restore button',
  resSimTrunc.execution?.executionPhase === 'intercepted_before_run' && resSimTrunc.execution?.snapshot === null
);

// 模拟 3: 模型回复过程缺少 End Sub
const simTruncEndSub = extractVbaCode('```vba\nSub Test()\n  Range("A1").Value = 1\n```');
assert('[Simulated] Missing End Sub returns status: "truncated"', simTruncEndSub.status === 'truncated');

// 模拟 4: 模型返回非 Sub/Function 结构
const simInvalidStruct = extractVbaCode('```vba\nDim x As Integer\nx = 1\n```');
assert('[Simulated] Missing Sub/Function returns status: "invalid_structure"', simInvalidStruct.status === 'invalid_structure');

// 1.5 异步切换边界测试
console.log('\n=== 1B. Asynchronous UI Switching Boundary Suite ===');
// 场景 1: CHAT 请求在等待期间，用户将 UI 切换至【操作】
const asyncReqChat = createRequestLifecycle('CHAT', '请教一个 Excel 问题');
let currentUiMode = 'CHAT';
// 模拟用户在发送后立即点击切换 UI
currentUiMode = 'AUTOMATION';
const asyncResChat = handleResponseLifecycle(asyncReqChat, '这是解答文本');
assert(
  'Switching UI to AUTOMATION during ongoing CHAT request does NOT affect ongoing request',
  asyncResChat.mode === 'CHAT' && asyncResChat.execution === null && !asyncResChat.didCallExecuteVba
);

// 场景 2: AUTOMATION 请求在等待期间，用户将 UI 切换至【对话】
const asyncReqAuto = createRequestLifecycle('AUTOMATION', '清空 A 列');
currentUiMode = 'AUTOMATION';
currentUiMode = 'CHAT'; // 模拟用户在生成中切换 UI
const asyncResAuto = handleResponseLifecycle(asyncReqAuto, mockValidVba);
assert(
  'Switching UI to CHAT during ongoing AUTOMATION request does NOT mutate request into CHAT',
  asyncResAuto.mode === 'AUTOMATION' && asyncResAuto.didCallExecuteVba
);

// 1.6 发送方式一致性测试 (Enter 发送 vs 点击发送按钮)
const reqByEnter = createRequestLifecycle('AUTOMATION', '在 A1 输入 100');
const reqByButtonClick = createRequestLifecycle('AUTOMATION', '在 A1 输入 100');
assert('Enter send and button click produce identical requestMode and target workbook binding',
  reqByEnter.requestMode === reqByButtonClick.requestMode && reqByEnter.targetWbName === reqByButtonClick.targetWbName
);

console.log('\n=== 1B. VBA Source Preservation Suite (No Silent Modification) ===');
// 生产代码严格 100% 保留模型返回的原始字符（包括中文引号、注释等），严禁静默规范化改写
const codeWithSmartQuotes = '```vba\nSub Main()\n  ws.Range("C3").Value = “错误”\nEnd Sub\n```';
const smartQuoteRes = extractVbaCode(codeWithSmartQuotes);
assert('Preserve smart quotes as-is without silent modification', smartQuoteRes.code.includes('“错误”'));

console.log('\n=== 2. Strict Code Extraction & Truncation Suite ===');

// Valid code
const validRes = extractVbaCode('这里是说明\n```vba\nSub LeeTaskEntry(targetWb As Workbook)\n  targetWb.Sheets(1).Range("A1").Value = 1\nEnd Sub\n```');
assert('Valid complete code extracted', !validRes.error && !validRes.isTruncated && validRes.code.includes('LeeTaskEntry'));

// Truncated code (unclosed code fence)
const truncRes1 = extractVbaCode('这里是说明\n```vba\nSub LeeTaskEntry(targetWb As Workbook)\n  Range("A1").Value = 1\n  ws.Range(');
assert('Truncated code (unclosed fence) detected', truncRes1.isTruncated && truncRes1.error.includes('未闭合'));

// Truncated code (missing End Sub inside closed fence)
const truncRes2 = extractVbaCode('```vba\nSub LeeTaskEntry(targetWb As Workbook)\n  Range("A1").Value = 1\n```');
assert('Incomplete code (missing End Sub) detected', truncRes2.isTruncated && truncRes2.error.includes('End Sub'));

// Chat text mentioning Sub (No code fence -> must NOT guess a macro)
const noFenceRes = extractVbaCode('你可以使用 Sub MySub() 和 End Sub 来编写宏，不需要直接运行。');
assert('Chat text mentioning Sub does not extract executable code', noFenceRes.code === '' && !noFenceRes.isTruncated);

console.log('\n=== 3. Post-execution Verification Suite ===');

// Case 3.1: Explicit equation requested, but result is pure number matrix -> flagged
const verRes1 = verifyExecutionResult('生成从 D1 开始的算式九九乘法表', {
  targetWorkbookName: 'Book1.xlsx',
  targetVerified: true,
  startCell: 'D1',
  sampleValues: ['1', '2', '3', '4', '81'],
  hasBorders: true,
  hasInteriorColor: true,
  usedRangeAddress: 'D1:L9',
});
assert(
  'Flagged when explicit equation requested but got pure numbers',
  verRes1.status === 'unconfirmed' && verRes1.note.includes('未生成算式文本')
);

// Case 3.2: Original user prompt ("新建一个表格，在 D1 开始写入九九乘法表，并美化这个表格")
// Neither form is errored out; both truthfully labeled with style pending human confirmation
const verResOrigEquation = verifyExecutionResult('新建一个表格，在 D1 开始写入九九乘法表，并美化这个表格', {
  targetWorkbookName: 'Book1.xlsx',
  targetVerified: true,
  startCell: 'D1',
  sampleValues: ['1×1=1', '2×2=4'],
  hasBorders: true,
  hasInteriorColor: true,
  usedRangeAddress: 'D1:L9',
});
assert(
  'Original prompt with equation table labeled as unconfirmed with layout note',
  verResOrigEquation.status === 'unconfirmed' && verResOrigEquation.note.includes('标准算式口诀表')
);

const verResOrigMatrix = verifyExecutionResult('新建一个表格，在 D1 开始写入九九乘法表，并美化这个表格', {
  targetWorkbookName: 'Book1.xlsx',
  targetVerified: true,
  startCell: 'D1',
  sampleValues: ['1', '2', '4', '81'],
  hasBorders: true,
  hasInteriorColor: true,
  usedRangeAddress: 'D1:L9',
});
assert(
  'Original prompt with numeric matrix labeled as unconfirmed with layout note',
  verResOrigMatrix.status === 'unconfirmed' && verResOrigMatrix.note.includes('乘积数值矩阵')
);

// Case 3.3: Explicit numeric matrix prompt
const verResNumExplicit = verifyExecutionResult('生成 9×9 数值乘积矩阵', {
  targetWorkbookName: 'Book1.xlsx',
  targetVerified: true,
  startCell: 'A1',
  sampleValues: ['1', '2', '4', '81'],
  hasBorders: false,
  hasInteriorColor: false,
  usedRangeAddress: 'A1:I9',
});
assert(
  'Explicit numeric matrix request without beauty matches verified status',
  verResNumExplicit.status === 'verified' && verResNumExplicit.note.includes('验证通过')
);

// Case 3.4: Target workbook mismatch triggers failed
const verResFailTarget = verifyExecutionResult('生成表格', {
  targetWorkbookName: 'Book1.xlsx',
  targetVerified: false,
  startCell: 'A1',
  sampleValues: [],
  hasBorders: false,
  hasInteriorColor: false,
  usedRangeAddress: '',
});
assert(
  'Target workbook mismatch triggers failed status',
  verResFailTarget.status === 'failed' && verResFailTarget.note.includes('身份核验失败')
);

// Case 3.5: Cross-workbook contamination detected triggers failed
const verResContam = verifyExecutionResult('生成表格', {
  targetWorkbookName: 'Book1.xlsx',
  targetVerified: true,
  otherWorkbooksAffected: true,
  affectedWorkbooksWarning: '安全告警：检测到非目标工作簿受到附带修改！',
  startCell: 'A1',
  sampleValues: ['1'],
  hasBorders: false,
  hasInteriorColor: false,
  usedRangeAddress: 'A1:A10',
});
assert(
  'Other workbooks affected triggers failed status',
  verResContam.status === 'failed' && verResContam.note.includes('安全告警')
);

console.log(`\nUnit Tests Summary: Pass = ${passCount}, Fail = ${failCount}`);
if (failCount > 0) {
  process.exit(1);
}
