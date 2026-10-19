/**
 * 统一工具注册表与契约中心 (Unified Tool Registry)
 * 
 * 单一真实源 (Single Source of Truth)：
 * 集中管理所有本地 COM 原生只读工具与外部数据接口的名称、描述、Schema、参数校验与调度映射。
 * 彻底消除 decisionEngine、mcpClient、bridge 与 NativeBridge 之间的分散硬编码。
 */

import {
  bridge,
  type WorkbookOverviewResult,
  type FormulaAuditResult,
  type ExternalDataPreviewResult,
} from './bridge';

export interface ToolValidationResult {
  valid: boolean;
  error?: string;
}

export interface ToolExecutionResult {
  ok: boolean;
  toolName: string;
  blueprintSummary?: string;
  content: Array<{
    type: 'text' | 'json';
    text?: string;
    json?: any;
  }>;
  error?: string;
}

export interface ToolDescriptor {
  name: string;
  displayName: string;
  description: string;
  bridgeAction: string;
  isReadOnly: boolean;
  category: 'perception' | 'audit' | 'external_data';
  inputSchema: {
    type: 'object';
    properties: Record<string, {
      type: string;
      description?: string;
      enum?: string[];
      default?: any;
    }>;
    required?: string[];
  };
  validate?: (args: Record<string, any>) => ToolValidationResult;
  execute: (args: Record<string, any>) => Promise<ToolExecutionResult>;
}

/**
 * 校验 Excel 单元格地址格式 (如 A1, $B$10, AA100)
 */
export function isValidCellAddress(addr: any): boolean {
  if (typeof addr !== 'string' || !addr.trim()) return false;
  const clean = addr.replace(/\$/g, '').trim();
  return /^[a-zA-Z]{1,3}[0-9]{1,7}$/.test(clean);
}

export const TOOL_REGISTRY: Record<string, ToolDescriptor> = {
  excel_get_workbook_overview: {
    name: 'excel_get_workbook_overview',
    displayName: '全簿全局架构透视',
    description: '只读透视当前活动工作簿的所有工作表清单、可见性、UsedRange 区域、公式单元格密度与表头预览。无需注入 VBA，毫秒级响应。',
    bridgeAction: 'get_workbook_overview',
    isReadOnly: true,
    category: 'perception',
    inputSchema: {
      type: 'object',
      properties: {
        targetWorkbookName: {
          type: 'string',
          description: '可选指定的目标工作簿文件名；若省略则自动获取当前活动工作簿',
        },
        targetWorkbookFullName: {
          type: 'string',
          description: '可选指定的目标工作簿磁盘全路径',
        },
      },
      required: [],
    },
    validate: (args) => {
      if (args.targetWorkbookName && typeof args.targetWorkbookName !== 'string') {
        return { valid: false, error: 'targetWorkbookName 必须为字符串' };
      }
      return { valid: true };
    },
    execute: async (args) => {
      try {
        const res: WorkbookOverviewResult | null = await bridge.getWorkbookOverview({
          targetWorkbookName: args.targetWorkbookName,
          targetWorkbookFullName: args.targetWorkbookFullName,
        });
        if (res) {
          return {
            ok: true,
            toolName: 'excel_get_workbook_overview',
            blueprintSummary: res.blueprintSummary,
            content: [
              { type: 'text', text: res.blueprintSummary },
              { type: 'json', json: res },
            ],
          };
        }
        return {
          ok: false,
          toolName: 'excel_get_workbook_overview',
          content: [],
          error: '获取工作簿概览失败或当前未打开任何有效工作簿',
        };
      } catch (e: any) {
        return {
          ok: false,
          toolName: 'excel_get_workbook_overview',
          content: [],
          error: e?.message || String(e),
        };
      }
    },
  },

  excel_trace_formula_dependencies: {
    name: 'excel_trace_formula_dependencies',
    displayName: '公式血缘与依赖穿透审计',
    description: '基于 Excel 原生 COM 精准追踪指定单元格公式的上游依赖源 (Precedents) 与下游引用者 (Dependents)，并诊断 #N/A、#DIV/0!、循环引用等异常。',
    bridgeAction: 'trace_dependencies',
    isReadOnly: true,
    category: 'audit',
    inputSchema: {
      type: 'object',
      properties: {
        cellAddress: {
          type: 'string',
          description: '单元格坐标，例如 "E5" 或 "A1"',
        },
        sheetName: {
          type: 'string',
          description: '单元格所在的工作表名称，例如 "汇总看板"',
        },
        targetWorkbookName: {
          type: 'string',
          description: '可选指定的目标工作簿文件名',
        },
      },
      required: ['cellAddress'],
    },
    validate: (args) => {
      if (!args || !args.cellAddress) {
        return { valid: false, error: '参数 cellAddress 为必填项' };
      }
      if (!isValidCellAddress(args.cellAddress)) {
        return { valid: false, error: `单元格坐标格式无效: "${args.cellAddress}"，必须为合法的 Excel 坐标（如 A1、E5）` };
      }
      return { valid: true };
    },
    execute: async (args) => {
      try {
        const res: FormulaAuditResult | null = await bridge.traceDependencies({
          targetWorkbookName: args.targetWorkbookName,
          targetWorkbookFullName: args.targetWorkbookFullName,
          sheetName: args.sheetName,
          cellAddress: args.cellAddress,
        });
        if (res) {
          return {
            ok: true,
            toolName: 'excel_trace_formula_dependencies',
            blueprintSummary: res.auditSummary,
            content: [
              { type: 'text', text: res.auditSummary },
              { type: 'json', json: res },
            ],
          };
        }
        return {
          ok: false,
          toolName: 'excel_trace_formula_dependencies',
          content: [],
          error: '公式依赖审计未返回有效结果（该单元格可能为空或不存在公式）',
        };
      } catch (e: any) {
        return {
          ok: false,
          toolName: 'excel_trace_formula_dependencies',
          content: [],
          error: e?.message || String(e),
        };
      }
    },
  },

  excel_preview_external_data: {
    name: 'excel_preview_external_data',
    displayName: '外部数据源安全抽样预览',
    description: '通过安全白名单只读接入外部 CSV、JSON 或 HTTP API 数据并进行预览抽样与格式推断，绝不污染存量工作簿。',
    bridgeAction: 'preview_external_data',
    isReadOnly: true,
    category: 'external_data',
    inputSchema: {
      type: 'object',
      properties: {
        sourceType: {
          type: 'string',
          enum: ['csv', 'json', 'http'],
          description: '数据源类型',
        },
        pathOrUrl: {
          type: 'string',
          description: '本地文件路径或已在白名单内的 HTTP URL',
        },
      },
      required: ['sourceType', 'pathOrUrl'],
    },
    validate: (args) => {
      if (!args || !args.sourceType || !args.pathOrUrl) {
        return { valid: false, error: 'sourceType 与 pathOrUrl 均为必填项' };
      }
      if (!['csv', 'json', 'http'].includes(args.sourceType)) {
        return { valid: false, error: `不支持的数据源类型: "${args.sourceType}"，仅支持 csv, json, http` };
      }
      return { valid: true };
    },
    execute: async (args) => {
      try {
        const res: ExternalDataPreviewResult = await bridge.previewExternalData({
          sourceType: args.sourceType,
          filePath: args.pathOrUrl,
          url: args.sourceType === 'http' ? args.pathOrUrl : undefined,
          previewRowCount: 5,
        });
        return {
          ok: res.ok,
          toolName: 'excel_preview_external_data',
          blueprintSummary: res.sanitizedSummary,
          content: [
            {
              type: 'text',
              text: res.sanitizedSummary || `已预览外部数据: 共 ${res.totalCols} 列，预估 ${res.totalRowsEstimate} 行`,
            },
            { type: 'json', json: res },
          ],
          error: res.error,
        };
      } catch (e: any) {
        return {
          ok: false,
          toolName: 'excel_preview_external_data',
          content: [],
          error: e?.message || String(e),
        };
      }
    },
  },
};

/**
 * 获取所有已注册的工具描述清单
 */
export function getAllToolDescriptors(): ToolDescriptor[] {
  return Object.values(TOOL_REGISTRY);
}

/**
 * 根据工具名称查找工具描述
 */
export function findToolDescriptor(toolName: string): ToolDescriptor | undefined {
  return TOOL_REGISTRY[toolName];
}
