export interface SnapshotItem {
  id: string;
  timestamp: string;
  timeDisplay: string;
  fileName: string;
  originalPath: string;
  promptSummary: string;
  vbaPreview: string;
}

export interface CellSampleItem {
  row: number;
  col: number;
  address: string;
  value: any;
  displayText: string;
  formula?: string | null;
  valueType: 'string' | 'number' | 'boolean' | 'empty' | 'error';
  isTextTruncated?: boolean;
}

export interface SelectionContextData {
  workbookName: string;
  workbookFullName: string;
  sheetName: string;
  address: string;
  totalRows: number;
  totalColumns: number;
  startRow: number;
  startColumn: number;
  endRow: number;
  endColumn: number;
  isSingleArea: boolean;
  sampleRowCount: number;
  sampleColumnCount: number;
  sampleAddress: string;
  candidateHeaders: string[];
  sampleRows: CellSampleItem[][];
  formulaStatus: 'has_formula' | 'no_formula' | 'sample_mixed' | 'unknown' | string;
  mergeStatus: 'has_merged' | 'no_merged' | 'sample_mixed' | 'unknown' | string;
  visibilityStatus: 'sample_scanned_only' | 'unknown' | string;
  isRowTruncated: boolean;
  isColumnTruncated: boolean;
  maxTextLengthLimit: number;
  unscannedNotes: string;
}

export interface SelectionSendOptions {
  includeStructure: boolean; // 总是 true
  includeSamples: boolean; // 默认 false
  includeFormulas: boolean; // 默认 false
  firstRowAsHeader: boolean; // 默认 true
}

export interface SelectionSnapshot {
  context: SelectionContextData;
  options: SelectionSendOptions;
  formattedText: string;
  summary: {
    fieldsIncluded: string[];
    sampleRowRange: string;
    sampleColRange: string;
    totalChars: number;
    isTruncated: boolean;
    truncatedNotes: string;
  };
}

export interface WorkbookInfo {
  name: string;
  fullName: string;
  isSaved: boolean;
  activeSheetName?: string;
  usedRangeAddress?: string;
  sheets: string[];
  snapshots: SnapshotItem[];
}

export interface ScriptItem {
  id?: string;
  name: string;
  displayName?: string;
  fileName: string;
  filePath?: string;
  createdAt: string;
  updatedAt?: string;
  description: string;
  category?: string;
  sourceType?: 'file' | 'paste' | 'legacy' | string;
  originalFileName?: string;
  encoding?: string;
  code: string;
  originalCodeHash?: string;
  entryPoint?: string;
  rawBytesBase64?: string;
  lastExecutionResult?: string;
  lastExecutedAt?: string;
  isVerified?: boolean;
}

export interface WorkbookReadback {
  targetWorkbookName: string;
  targetWorkbookFullName: string;
  targetSheetName: string;
  usedRangeAddress: string;
  rowCount: number;
  columnCount: number;
  startCell: string;
  endCell: string;
  sampleValues: string[];
  hasFormulas: boolean;
  hasBorders: boolean;
  hasInteriorColor: boolean;
  sheetCount: number;
  targetVerified: boolean;
  otherWorkbooksAffected?: boolean;
  affectedWorkbooksWarning?: string;
}

export interface VbaExecutionData {
  summary: string;
  error?: string;
  elapsedMs: number;
  snapshot?: SnapshotItem;
  rawModelResponse?: string;
  originalVbaCode?: string;
  executedVbaCode?: string;
  wrapperCode?: string;
  originalCodeHash?: string;
  executedCodeHash?: string;
  isSourceIdentical?: boolean;
  vbaCode?: string; // backwards compatibility
  transformSteps?: string[];
  readback?: WorkbookReadback;
  targetWorkbookName?: string;
  targetWorkbookFullName?: string;
  verificationStatus?: 'verified' | 'unconfirmed' | 'failed';
  verificationNote?: string;
  precheckStatus?: 'passed' | 'warning' | 'failed' | 'unavailable' | 'scope_risk_intercepted' | 'workbook_locked' | 'skipped';
  executionPhase?:
    | 'macro_completed'
    | 'hang_suspected_interrupt_sent'
    | 'hang_interrupted_recovered'
    | 'hang_unconfirmed_locked'
    | 'runtime_hang'
    | 'runtime_error'
    | 'syntax_failed'
    | 'compile_failed'
    | 'injection_failed'
    | 'invocation_failed'
    | 'extract_failed'
    | 'intercepted_before_run'
    | 'blocked_by_lock';
  failureStage?: string;
  rawErrorCode?: string;
  errorTriggerPoint?: string;
  vbaErrNumber?: number;
  vbaErrDescription?: string;
  comHResult?: string;
  hostExecutionPhase?: string;
  isPartiallyModified?: boolean;
  retryCount?: number;
  llmCost?: {
    promptTokens?: number;
    completionTokens?: number;
    reasoningTokens?: number;
    contentTokens?: number;
    totalTokens?: number;
  };
  hangRecovery?: string;
  apiAudit?: {
    maxTokensStatus: string;
    thinkingBudgetStatus: string;
    historyCountSent: number;
    totalHistoryAvailable: number;
    historyStrategy: string;
    isHistoryCompressedOrStripped: boolean;
    actualPayloadSummary?: Record<string, any>;
  };
}

export interface BridgeResponse<T = any> {
  ok: boolean;
  action: string;
  requestId?: string;
  message?: string;
  data?: T;
  error?: string;
}

type MessageHandler = (res: BridgeResponse) => void;

class NativeBridgeClient {
  private handlers = new Map<string, MessageHandler>();
  private workbookChangeListeners: ((info: WorkbookInfo) => void)[] = [];
  private actionListeners = new Map<string, (() => void)[]>();
  private requestCounter = 0;

  constructor() {
    if (typeof window !== 'undefined' && (window as any).chrome?.webview) {
      (window as any).chrome.webview.addEventListener('message', (event: any) => {
        try {
          const raw: BridgeResponse = typeof event.data === 'string' ? JSON.parse(event.data) : event.data;

          // 若为 C# 宿主主动推送的工作簿变更（无特定 requestId），通知全局监听器
          if (raw.action === 'get_workbook_info' && !raw.requestId && raw.ok && raw.data) {
            this.workbookChangeListeners.forEach((fn) => fn(raw.data));
          }

          // 处理宿主主动触发的界面指令（如 open_scripts, open_settings）
          if (raw.action && this.actionListeners.has(raw.action)) {
            this.actionListeners.get(raw.action)?.forEach((fn) => {
              try {
                fn();
              } catch (err) {
                console.error('[NativeBridgeClient] Error in action listener:', err);
              }
            });
          }

          // 1. 优先按唯一 requestId 寻址并立即清理
          if (raw.requestId && this.handlers.has(raw.requestId)) {
            const handler = this.handlers.get(raw.requestId);
            this.handlers.delete(raw.requestId);
            handler?.(raw);
            return;
          }

          // 2. 兼容旧协议或无 requestId 的按 action 回退寻址并清理
          if (raw.action && this.handlers.has(raw.action)) {
            const handler = this.handlers.get(raw.action);
            this.handlers.delete(raw.action);
            handler?.(raw);
          }
        } catch (e) {
          console.error('[NativeBridgeClient] Parse message error:', e);
        }
      });
    }
  }

  public isNative(): boolean {
    return typeof window !== 'undefined' && Boolean((window as any).chrome?.webview);
  }

  public onWorkbookChange(fn: (info: WorkbookInfo) => void): () => void {
    this.workbookChangeListeners.push(fn);
    return () => {
      const idx = this.workbookChangeListeners.indexOf(fn);
      if (idx !== -1) this.workbookChangeListeners.splice(idx, 1);
    };
  }

  public onAction(action: string, fn: () => void): () => void {
    if (!this.actionListeners.has(action)) {
      this.actionListeners.set(action, []);
    }
    this.actionListeners.get(action)!.push(fn);
    return () => {
      const list = this.actionListeners.get(action);
      if (list) {
        const idx = list.indexOf(fn);
        if (idx !== -1) list.splice(idx, 1);
      }
    };
  }

  public async send<T = any>(action: string, payload: Record<string, any> = {}, timeoutMs: number = 120000): Promise<BridgeResponse<T>> {
    if (!this.isNative()) {
      return this.mockResponse<T>(action, payload);
    }

    const requestId = `req_${Date.now()}_${++this.requestCounter}_${Math.random().toString(36).substring(2, 7)}`;

    return new Promise((resolve) => {
      const timer = setTimeout(() => {
        this.handlers.delete(requestId);
        resolve({
          ok: false,
          action,
          requestId,
          error: `宿主响应超时（${timeoutMs / 1000}秒未收到响应），请检查 Excel 进程状态后重试。`,
        } as BridgeResponse<T>);
      }, timeoutMs);

      const msg = JSON.stringify({ action, requestId, ...payload });
      this.handlers.set(requestId, (res) => {
        clearTimeout(timer);
        resolve(res);
      });
      (window as any).chrome.webview.postMessage(msg);
    });
  }

  // 纯浏览器单测 Mock 模式
  private async mockResponse<T>(action: string, payload: any): Promise<BridgeResponse<T>> {
    await new Promise((r) => setTimeout(r, 200));

    if (action === 'get_workbook_info') {
      return {
        ok: true,
        action,
        data: {
          name: '2026年9月部门预算与绩效表.xlsx',
          fullName: 'C:\\Users\\35651\\Documents\\2026年9月部门预算与绩效表.xlsx',
          isSaved: true,
          activeSheetName: '汇总看板',
          usedRangeAddress: 'A1:H25',
          sheets: ['汇总看板', '部门明细', '人员考评'],
          snapshots: [
            {
              id: 'snap_mock_1',
              timestamp: new Date().toISOString(),
              timeDisplay: '10:45:10',
              fileName: 'snap_1.xlsx',
              originalPath: 'C:\\test.xlsx',
              promptSummary: '自动标红超出预算的金额',
              vbaPreview: 'Sub HighlightOverBudget()...',
            },
          ],
        } as any,
      };
    }

    if (action === 'execute_vba') {
      return {
        ok: true,
        action,
        message: '宏已运行（耗时 382 ms）',
        data: {
          summary: '宏已运行（耗时 382 ms）',
          elapsedMs: 382,
          originalVbaCode: payload.code,
          executedVbaCode: payload.code,
          vbaCode: payload.code,
          transformSteps: ['规整过程入口: RunTask -> LeeTaskEntry', '显式绑定目标工作簿'],
          targetWorkbookName: payload.targetWorkbookName || '2026年9月部门预算与绩效表.xlsx',
          readback: {
            targetWorkbookName: payload.targetWorkbookName || '2026年9月部门预算与绩效表.xlsx',
            targetWorkbookFullName: payload.targetWorkbookFullName || 'C:\\test.xlsx',
            targetSheetName: '汇总看板',
            usedRangeAddress: 'D1:L9',
            rowCount: 9,
            columnCount: 9,
            startCell: 'D1',
            endCell: 'L9',
            sampleValues: ['1×1=1', '1×2=2', '1×3=3'],
            hasFormulas: false,
            hasBorders: true,
            hasInteriorColor: true,
            sheetCount: 3,
            targetVerified: true,
          },
          snapshot: {
            id: 'snap_' + Date.now(),
            timestamp: new Date().toISOString(),
            timeDisplay: new Date().toTimeString().slice(0, 8),
            fileName: 'snap_new.xlsx',
            originalPath: 'C:\\test.xlsx',
            promptSummary: payload.prompt || '操作快照',
            vbaPreview: payload.code.slice(0, 60),
          },
        } as any,
      };
    }

    if (action === 'list_scripts') {
      return {
        ok: true,
        action,
        data: [
          {
            name: '数据区域自动边框与隔行变色',
            fileName: '数据区域自动边框与隔行变色.bas',
            filePath: 'C:\\AppData\\LeeExcel\\Scripts\\数据区域自动边框与隔行变色.bas',
            createdAt: '2026-09-25 10:15',
            description: '自动寻找当前使用区域并添加细边框与浅绿隔行底色',
            code: 'Sub FormatTable()\n    Dim rng As Range\n    Set rng = ActiveSheet.UsedRange\n    rng.Borders.LineStyle = xlContinuous\nEnd Sub',
          },
        ] as any,
      };
    }

    if (action === 'get_selection_context') {
      return {
        ok: true,
        action,
        data: {
          workbookName: '2026年9月部门预算与绩效表.xlsx',
          workbookFullName: 'C:\\Users\\35651\\Documents\\2026年9月部门预算与绩效表.xlsx',
          sheetName: '汇总看板',
          address: '$A$1:$F$10',
          totalRows: 10,
          totalColumns: 6,
          startRow: 1,
          startColumn: 1,
          endRow: 10,
          endColumn: 6,
          isSingleArea: true,
          sampleRowCount: 3,
          sampleColumnCount: 6,
          sampleAddress: '$A$1:$F$3',
          candidateHeaders: ['部门', '科目', '预算金额', '实际支出', '执行率', '责任人'],
          sampleRows: [
            [
              { row: 1, col: 1, address: '$A$1', value: '部门', displayText: '部门', valueType: 'string', isTextTruncated: false },
              { row: 1, col: 2, address: '$B$1', value: '科目', displayText: '科目', valueType: 'string', isTextTruncated: false },
              { row: 1, col: 3, address: '$C$1', value: '预算金额', displayText: '预算金额', valueType: 'string', isTextTruncated: false },
              { row: 1, col: 4, address: '$D$1', value: '实际支出', displayText: '实际支出', valueType: 'string', isTextTruncated: false },
              { row: 1, col: 5, address: '$E$1', value: '执行率', displayText: '执行率', valueType: 'string', isTextTruncated: false },
              { row: 1, col: 6, address: '$F$1', value: '责任人', displayText: '责任人', valueType: 'string', isTextTruncated: false }
            ],
            [
              { row: 2, col: 1, address: '$A$2', value: '研发部', displayText: '研发部', valueType: 'string', isTextTruncated: false },
              { row: 2, col: 2, address: '$B$2', value: '云资源租赁', displayText: '云资源租赁', valueType: 'string', isTextTruncated: false },
              { row: 2, col: 3, address: '$C$2', value: 500000, displayText: '500000', valueType: 'number', isTextTruncated: false },
              { row: 2, col: 4, address: '$D$2', value: 420000, displayText: '420000', valueType: 'number', isTextTruncated: false },
              { row: 2, col: 5, address: '$E$2', value: 0.84, displayText: '0.84', formula: '=D2/C2', valueType: 'number', isTextTruncated: false },
              { row: 2, col: 6, address: '$F$2', value: '张工', displayText: '张工', valueType: 'string', isTextTruncated: false }
            ],
            [
              { row: 3, col: 1, address: '$A$3', value: '市场部', displayText: '市场部', valueType: 'string', isTextTruncated: false },
              { row: 3, col: 2, address: '$B$3', value: '广告宣发', displayText: '广告宣发', valueType: 'string', isTextTruncated: false },
              { row: 3, col: 3, address: '$C$3', value: 300000, displayText: '300000', valueType: 'number', isTextTruncated: false },
              { row: 3, col: 4, address: '$D$3', value: 310000, displayText: '310000', valueType: 'number', isTextTruncated: false },
              { row: 3, col: 5, address: '$E$3', value: 1.033, displayText: '1.033', formula: '=D3/C3', valueType: 'number', isTextTruncated: false },
              { row: 3, col: 6, address: '$F$3', value: '李经理', displayText: '李经理', valueType: 'string', isTextTruncated: false }
            ]
          ],
          formulaStatus: 'sample_mixed',
          mergeStatus: 'no_merged',
          visibilityStatus: 'sample_scanned_only',
          isRowTruncated: true,
          isColumnTruncated: false,
          maxTextLengthLimit: 100,
          unscannedNotes: '选区共 10 行 × 6 列。本次仅安全抽样前 3 行 × 前 6 列。未扫描其余单元格内容，未扫描全表筛选/隐藏行状态。'
        } as any,
      };
    }

    return {
      ok: true,
      action,
      message: 'Mock 操作成功',
      data: {} as any,
    };
  }

  public async getSelectionContext(
    sampleRows: number = 5,
    sampleCols: number = 15
  ): Promise<{ ok: boolean; data?: SelectionContextData; error?: string; errorType?: string }> {
    const res = await this.send<SelectionContextData>('get_selection_context', {
      sampleRows: sampleRows.toString(),
      sampleCols: sampleCols.toString(),
    });
    if (res.ok && res.data) {
      return { ok: true, data: res.data };
    }
    return {
      ok: false,
      error: res.error || '获取选区失败',
      errorType: (res.data as any)?.errorType || 'exception',
    };
  }
}

export const bridge = new NativeBridgeClient();

