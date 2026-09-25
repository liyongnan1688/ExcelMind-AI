// test_suite_unit.cjs - Unit test runner for intent classification, parser, and verification

function detectIntent(text) {
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

  const asksCodeMeaningRegex =
    /(?:这段代码|这个宏|这几行代码|这段VBA|以下代码|这段宏|Sub\s+[\s\S]+End\s+Sub)[\s\S]*(?:什么意思|含义|解释|怎么理解|干嘛|干什么|作用|读懂|请教|为什么|如何理解)/i;
  if (asksCodeMeaningRegex.test(trimmed)) {
    return 'CHAT';
  }

  if (
    /(?:什么意思|怎么理解|是干什么的|有何作用)[\?？]*$/.test(trimmed) &&
    (trimmed.includes('Sub') || trimmed.includes('Range') || trimmed.includes('Dim'))
  ) {
    return 'CHAT';
  }

  const generalKnowledgeRegex =
    /^(?:什么是|如何理解|为什么|怎么用|怎么使用|函数用法|公式怎么写|区别是什么|有什么区别)/i;
  if (generalKnowledgeRegex.test(trimmed)) {
    return 'CHAT';
  }

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

  if (trimmed.length <= 4 || (!hasActionVerb && !trimmed.includes('？') && !trimmed.includes('?'))) {
    return 'AMBIGUOUS';
  }

  return 'CHAT';
}

function extractVbaCode(content) {
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

  const cleanedCode = rawCode.replace(/^[ \t]*[*\-•][ \t]+/gm, '');
  return {
    code: cleanedCode,
    isTruncated: false,
  };
}

function verifyExecutionResult(prompt, readback) {
  if (!readback) {
    return { status: 'unconfirmed', note: '未能读回目标工作簿实际变更状态，效果待确认。' };
  }
  if (!readback.targetVerified) {
    return { status: 'failed', note: `目标工作簿身份核验失败（预期: ${readback.targetWorkbookName}）。` };
  }

  const notes = [];
  const startMatch = prompt.match(/(?:从|在)\s*([A-Za-z]+[0-9]+)/i);
  if (startMatch) {
    const expectedStart = startMatch[1].toUpperCase();
    if (readback.startCell && readback.startCell.toUpperCase() !== expectedStart) {
      notes.push(`要求从 ${expectedStart} 开始，实际起始于 ${readback.startCell}`);
    }
  }

  const wantsEquation = /(?:算式|口诀|乘法口诀|×|\*|=)/i.test(prompt);
  if (wantsEquation && readback.sampleValues && readback.sampleValues.length > 0) {
    const hasEquationText = readback.sampleValues.some(
      (v) => v.includes('×') || v.includes('*') || v.includes('=') || v.includes('得')
    );
    if (!hasEquationText) {
      notes.push('检测到填入内容为纯数字矩阵，未生成算式文本');
    }
  }

  const wantsBeauty = /(?:美化|商务|好看|排版|样式|颜色|边框)/i.test(prompt);
  if (wantsBeauty) {
    const styleFeatures = [];
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

console.log('=== 1. Intent Detection Suite (Multi-phrasing) ===');
const chatTests = [
  '你是？',
  '你是谁，能帮我做什么？',
  '自我介绍一下',
  '解释刚才做了什么',
  '说说刚才那步宏执行了什么操作',
  '请问刚才的代码是什么原理',
  'Sub HighlightRows()\n ActiveSheet.Range("A1").Interior.Color = vbYellow\nEnd Sub\n这段代码什么意思？',
  '帮我看看这段宏是干什么的：\nSub Test()\nMsgBox "hi"\nEnd Sub',
  '什么是数据透视表？',
  '在Excel中如何使用SUMIF函数？',
];

for (const q of chatTests) {
  const intent = detectIntent(q);
  assert(`Chat input: "${q.slice(0, 30).replace(/\n/g, ' ')}" -> CHAT`, intent === 'CHAT', `Got ${intent}`);
}

const autoTests = [
  '新建一个表格，在 D1 开始写入九九乘法表，并美化这个表格',
  '请从 D1 单元格开始生成阶梯式算式九九乘法表',
  '生成 9×9 数值乘积矩阵',
  '制作一个九九乘法数字矩阵表',
  '为已有业务数据调整列宽、对齐、表头和数字格式',
  '给当前表格加上边框并设置隔行变色',
  '在末尾行计算销售总额公式',
  '根据这几列数据生成柱状图',
];

for (const a of autoTests) {
  const intent = detectIntent(a);
  assert(`Auto input: "${a.slice(0, 30)}" -> AUTOMATION`, intent === 'AUTOMATION', `Got ${intent}`);
}

const ambigTests = ['乘法表', 'VBA', '表格', '宏'];
for (const m of ambigTests) {
  const intent = detectIntent(m);
  assert(`Ambiguous input: "${m}" -> AMBIGUOUS`, intent === 'AMBIGUOUS', `Got ${intent}`);
}

console.log('\n=== 2. Strict Code Extraction & Truncation Suite ===');

// Valid code
const validRes = extractVbaCode('这里是说明\n```vba\nSub LeeTaskEntry()\n  Range("A1").Value = 1\nEnd Sub\n```');
assert('Valid complete code extracted', !validRes.error && !validRes.isTruncated && validRes.code.includes('LeeTaskEntry'));

// Truncated code (unclosed code fence)
const truncRes1 = extractVbaCode('这里是说明\n```vba\nSub LeeTaskEntry()\n  Range("A1").Value = 1\n  ws.Range(');
assert('Truncated code (unclosed fence) detected', truncRes1.isTruncated && truncRes1.error.includes('未闭合'));

// Truncated code (missing End Sub inside closed fence)
const truncRes2 = extractVbaCode('```vba\nSub LeeTaskEntry()\n  Range("A1").Value = 1\n```');
assert('Incomplete code (missing End Sub) detected', truncRes2.isTruncated && truncRes2.error.includes('End Sub'));

// Chat text mentioning Sub (No code fence -> must NOT guess a macro)
const noFenceRes = extractVbaCode('你可以使用 Sub MySub() 和 End Sub 来编写宏，不需要直接运行。');
assert('Chat text mentioning Sub does not extract executable code', noFenceRes.code === '' && !noFenceRes.isTruncated);

console.log('\n=== 3. Post-execution Verification Suite ===');

// Case: User asked for equation, but result is pure number matrix
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
  'Flagged when user requested equations but got pure numbers',
  verRes1.status === 'unconfirmed' && verRes1.note.includes('纯数字矩阵')
);

// Case: User asked for start at D1, but result starts at A1
const verRes2 = verifyExecutionResult('在 D1 开始写入数据', {
  targetWorkbookName: 'Book1.xlsx',
  targetVerified: true,
  startCell: 'A1',
  sampleValues: ['Test'],
  hasBorders: false,
  hasInteriorColor: false,
  usedRangeAddress: 'A1:A10',
});
assert(
  'Flagged when starting cell does not match user requirement',
  verRes2.status === 'unconfirmed' && verRes2.note.includes('要求从 D1 开始，实际起始于 A1')
);

// Case: Subjective beauty request -> status is 'unconfirmed' (never false green complete)
const verRes3 = verifyExecutionResult('生成表格并美化', {
  targetWorkbookName: 'Book1.xlsx',
  targetVerified: true,
  startCell: 'A1',
  sampleValues: ['Data'],
  hasBorders: true,
  hasInteriorColor: true,
  usedRangeAddress: 'A1:C5',
});
assert(
  'Subjective beauty is marked as unconfirmed (effect pending human confirmation)',
  verRes3.status === 'unconfirmed' && verRes3.note.includes('效果待人工确认')
);

// Case: Target workbook identity mismatch
const verRes4 = verifyExecutionResult('生成表格', {
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
  verRes4.status === 'failed' && verRes4.note.includes('身份核验失败')
);

console.log(`\nUnit Tests Summary: Pass = ${passCount}, Fail = ${failCount}`);
if (failCount > 0) {
  process.exit(1);
}
