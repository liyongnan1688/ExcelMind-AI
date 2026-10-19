/**
 * 双轨自主调度决策器 (Dual-Track Decision Engine)
 * 
 * 核心设计：
 * - 轨道一 (Track A: Native COM Fast-Path Dispatch / 原生轻量快速分派通道)：
 *   轻量感知、只读探测、全簿透视、公式血缘审计与外部数据接入。毫秒级即时响应，零宏注入、零快照、零弹窗。
 * - 轨道二 (Track B: Heavy VBA & Physical Snapshot)：
 *   重度计算、复杂批量清洗、图表渲染与跨表形变。先做物理全量快照，再进行 100% 源码保真编译注入与写后读回。
 * - 纯咨询通道 (CHAT)：
 *   检测到疑问修饰词（如“怎么”、“如何”、“为什么”、“how”、“why”）或技术答疑意图时，智能降级为纯文本咨询，绝对不误触发宏执行与物理快照。
 */

import { type McpToolCallRequest } from './mcpClient';
import { TOOL_REGISTRY } from './toolRegistry';
import type { SelectionSnapshot } from './bridge';

export type ExecutionTrackType =
  | 'TRACK_A_OVERVIEW'       // 全簿结构透视
  | 'TRACK_A_FORMULA'        // 公式血缘与穿透审计
  | 'TRACK_A_MCP'            // 外部数据接入预览
  | 'TRACK_B_HEAVY_VBA'      // 重度表格修改 / 批量数据清洗 / 图表生成
  | 'CHAT';                  // 纯技术咨询与闲聊 (不触发自动化写入)

export interface DecisionResult {
  track: ExecutionTrackType;
  confidence: number;
  reason: string;
  suggestedMcpCall?: McpToolCallRequest;
}

/**
 * 用户输入意图分析与轨道决策
 */
export function classifyExecutionTrack(
  prompt: string,
  selectionSnapshot?: SelectionSnapshot,
  hasReferencedMacro: boolean = false
): DecisionResult {
  const p = (prompt || '').trim().toLowerCase();

  // 1. 如果显式引用了前序宏，说明意图是延续或修改宏，锁定 Track B
  if (hasReferencedMacro) {
    return {
      track: 'TRACK_B_HEAVY_VBA',
      confidence: 0.95,
      reason: '用户显式引用了前序宏，将执行宏代码延续与运行',
    };
  }

  // 2. 检查是否为公式血缘与依赖审计意图 (中英双语)
  const formulaAuditKeywords = /(依赖|血缘|谁引用了|引用了谁|从属|前序|公式解释|白话解释|为什么报错|#n\/a|#div\/0|#value|公式体检|查错|precedent|dependent|trace|formula\s*audit|audit\s*formula)/i;
  if (formulaAuditKeywords.test(p)) {
    const cellMatch = p.match(/([a-zA-Z]{1,3}[0-9]{1,7})/);
    const cellAddr = cellMatch ? cellMatch[1].toUpperCase() : (selectionSnapshot?.context?.address?.split(':')[0] || 'A1');
    const sheet = selectionSnapshot?.context?.sheetName || '';

    return {
      track: 'TRACK_A_FORMULA',
      confidence: 0.9,
      reason: `检测到公式依赖分析需求，命中轻量原生 COM 审计工具 (${sheet ? sheet + '!' : ''}${cellAddr})`,
      suggestedMcpCall: {
        toolName: TOOL_REGISTRY.excel_trace_formula_dependencies.name,
        arguments: {
          sheetName: sheet,
          cellAddress: cellAddr,
        },
      },
    };
  }

  // 3. 检查是否为工作簿全景架构透视意图 (中英双语)
  const overviewKeywords = /(工作簿架构|全簿|所有工作表|有多少个表|表头结构|全貌|结构概览|各表清单|公式数量|查看表名|workbook\s*overview|sheet\s*list|all\s*sheets|structure\s*overview)/i;
  if (overviewKeywords.test(p)) {
    return {
      track: 'TRACK_A_OVERVIEW',
      confidence: 0.88,
      reason: '检测到全局工作簿感知需求，命中轻量原生 COM 结构透视工具',
      suggestedMcpCall: {
        toolName: TOOL_REGISTRY.excel_get_workbook_overview.name,
        arguments: {},
      },
    };
  }

  // 4. 检查是否为外部数据接入 / 开放数据源读取
  const mcpKeywords = /(http|api|json|csv|外部数据|抓取|接入接口|external\s*data|fetch\s*data)/i;
  if (mcpKeywords.test(p) && (p.includes('http://') || p.includes('https://') || p.includes('.csv') || p.includes('.json'))) {
    const urlMatch = p.match(/(https?:\/\/[^\s"'<>]+)/i);
    const pathMatch = p.match(/([a-zA-Z]:\\[^\s"'<>]+\.(?:csv|json))/i);
    const target = urlMatch ? urlMatch[1] : (pathMatch ? pathMatch[1] : '');

    if (target) {
      return {
        track: 'TRACK_A_MCP',
        confidence: 0.85,
        reason: '检测到外部数据源接入需求，命中通用数据安全预览网关',
        suggestedMcpCall: {
          toolName: TOOL_REGISTRY.excel_preview_external_data.name,
          arguments: {
            sourceType: target.startsWith('http') ? 'http' : (target.endsWith('.csv') ? 'csv' : 'json'),
            pathOrUrl: target,
          },
        },
      };
    }
  }

  // 5. 疑问修饰词检查：如果包含“怎么/如何/为什么/why/how”等疑问词，优先判定为技术咨询与指导，防止误触发代码执行
  const questionKeywords = /(怎么|如何|为什么|为啥|为何|请教|说明|能否|解释一下|是什么|how\s+to|how\s+do|why\s+is|what\s+is|can\s+you\s+explain)/i;
  if (questionKeywords.test(p)) {
    return {
      track: 'CHAT',
      confidence: 0.9,
      reason: '检测到用户带有疑问或解释咨询词（如“怎么/如何/why”），智能判定为业务答疑，避免误触发宏执行',
    };
  }

  // 6. 检查是否包含实质表格修改动词 (写入、清洗、修改、求和、标红、删除、合并、分裂、生成) (中英双语)
  const heavyKeywords = /(新建|创建|删除|修改|填充|清洗|标红|高亮|排序|过滤|合并|拆分|分裂|加总|生成图表|柱状图|折线图|饼图|宏|vba|写进|覆盖|create|insert|delete|clean|highlight|sort|filter|merge|split|generate\s*chart)/i;
  if (heavyKeywords.test(p)) {
    return {
      track: 'TRACK_B_HEAVY_VBA',
      confidence: 0.9,
      reason: '检测到表格写操作或批量形变需求，进入物理快照锁定与 VBA 注入执行轨道',
    };
  }

  // 7. 默认纯技术咨询
  return {
    track: 'CHAT',
    confidence: 0.7,
    reason: '未检测到明确的表格操作动词，按纯文本咨询通道响应',
  };
}
