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
  capturedAt: number;
  attachmentId?: string;
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

export interface ScriptParameterDef {
  name: string;
  type: string; // 'String' | 'Long' | 'Double' | 'Boolean' | 'Date' | 'Worksheet' | 'Range' | 'Workbook'
  typeName?: string;
  rawType?: string;
  description?: string;
  isOptional?: boolean;
  defaultValue?: string;
  isTargetWorkbook?: boolean;
  isSupported?: boolean;
  unsupportedReason?: string;
}

export interface VbaEntryPointInfo {
  name: string;
  kind: 'Sub' | 'Function';
  visibility: 'Public' | 'Private' | 'Friend' | 'Default';
  isExecutable: boolean;
  isSupported?: boolean;
  unsupportedReason?: string;
  hasNativeWbParam?: boolean;
  parameters: ScriptParameterDef[];
  rawSignature?: string;
  lineIndex?: number;
}

export interface DedupAnalysisResult {
  ok: boolean;
  error?: string;
  failureStage?: string;
  targetWorkbookName?: string;
  targetSheetName?: string;
  rangeAddress?: string;
  hasHeader: boolean;
  keyColumnIndices: number[];
  keyColumnNames: string[];
  selectedRowCount: number;
  selectedColumnCount: number;
  dataRowCount: number;
  uniqueCount: number;
  duplicateCount: number;
  duplicateGroupCount: number;
  excludedCount: number;
  excludedReason?: string;
  analysisElapsedMs: number;
}

export interface DedupExecutionResult {
  ok: boolean;
  mode: 'highlight' | 'export_unique';
  message?: string;
  error?: string;
  snapshotId?: string;
  snapshotCreated?: boolean;
  dataRowCount: number;
  uniqueCount: number;
  duplicateCount: number;
  duplicateGroupCount: number;
  excludedCount: number;
  modifiedRowCount: number;
  resultSheetName?: string;
  highlightedRange?: string;
  analysisElapsedMs: number;
  writeElapsedMs: number;
  totalElapsedMs: number;
}

export interface CompareColMappingDto {
  leftColIndex: number;
  leftColName?: string;
  rightColIndex: number;
  rightColName?: string;
}

export interface CellDiffItem {
  colName: string;
  leftValue: any;
  leftType: string;
  rightValue: any;
  rightType: string;
}

export interface ReconcileRowDetail {
  category: 'SAME' | 'DIFF' | 'LEFT_ONLY' | 'RIGHT_ONLY' | 'LEFT_DUP' | 'RIGHT_DUP' | 'LEFT_INVALID' | 'RIGHT_INVALID' | string;
  leftRowIndex: number;
  rightRowIndex: number;
  keyDisplay: string;
  diffs?: CellDiffItem[];
}

export interface ReconcileAnalysisResult {
  ok: boolean;
  error?: string;
  failureStage?: string;
  leftWorkbookName?: string;
  leftSheetName?: string;
  leftRangeAddress?: string;
  leftHasHeader: boolean;
  leftKeyCols: number[];
  leftKeyNames: string[];
  rightWorkbookName?: string;
  rightSheetName?: string;
  rightRangeAddress?: string;
  rightHasHeader: boolean;
  rightKeyCols: number[];
  rightKeyNames: string[];
  compareCols: CompareColMappingDto[];
  leftRowCount: number;
  leftColCount: number;
  leftDataRowCount: number;
  rightRowCount: number;
  rightColCount: number;
  rightDataRowCount: number;
  matchedBothSameCount: number;
  matchedBothDiffCount: number;
  leftOnlyCount: number;
  rightOnlyCount: number;
  leftDuplicateKeyCount: number;
  rightDuplicateKeyCount: number;
  leftInvalidKeyCount: number;
  rightInvalidKeyCount: number;
  diffCellCount: number;
  analysisElapsedMs: number;
  dataFingerprint?: string;
  details?: ReconcileRowDetail[];
}

export interface ReconcileExecutionResult {
  ok: boolean;
  message?: string;
  error?: string;
  targetWorkbookName?: string;
  resultSheetName?: string;
  snapshotId?: string;
  snapshotCreated?: boolean;
  matchedBothSameCount: number;
  matchedBothDiffCount: number;
  leftOnlyCount: number;
  rightOnlyCount: number;
  leftDuplicateKeyCount: number;
  rightDuplicateKeyCount: number;
  leftInvalidKeyCount: number;
  rightInvalidKeyCount: number;
  diffCellCount: number;
  exportedRowCount: number;
  analysisElapsedMs: number;
  writeElapsedMs: number;
  totalElapsedMs: number;
}

export interface ConsolidationSourceDef {
  filePath: string;
  sheetName?: string;
  rangeAddress?: string;
  hasHeader: boolean;
  headerRowIndex?: number;
  displayIdentifier?: string;
}

export interface ConsolidationColMapping {
  sourceFileIndex: number;
  sourceColName: string;
  targetColName: string;
}

export interface ConsolidationSourceProfile {
  fileIndex: number;
  filePath: string;
  displayIdentifier: string;
  sheetName: string;
  rangeAddress: string;
  totalRows: number;
  totalCols: number;
  headerRow: number;
  includedDataRows: number;
  excludedBlankRows: number;
  columns: string[];
  fileSha256: string;
}

export interface ConsolidationAnalysisResult {
  ok: boolean;
  error?: string;
  failureStage?: string;
  sources: ConsolidationSourceProfile[];
  alignedColumns: string[];
  metadataSourceFileCol: string;
  metadataSourceSheetCol: string;
  missingColumnsPerSource: Record<number, string[]>;
  unmappedSourceColumns: Record<number, string[]>;
  columnTypeProfiles: Record<string, string[]>;
  previewRows: Array<Record<string, any>>;
  totalSourceCount: number;
  totalInputRows: number;
  totalHeaderRows: number;
  totalIncludedDataRows: number;
  totalExcludedBlankRows: number;
  expectedFinalOutputRows: number;
  accountingIdentityConfirmed: boolean;
  dataFingerprint?: string;
  analysisElapsedMs: number;
}

export interface ConsolidationExecutionResult {
  ok: boolean;
  error?: string;
  failureStage?: string;
  outputFilePath: string;
  outputSheetName: string;
  finalDataRows: number;
  finalDataCols: number;
  totalIncludedDataRows: number;
  totalExcludedBlankRows: number;
  accountingIdentityConfirmed: boolean;
  verifiedReadback: boolean;
  elapsedMs: number;
}

export interface RunRecordDto {
  id: string;
  executedAt: string;
  status: 'success' | 'failed' | 'blocked';
  phase?: string;
  targetWorkbookName?: string;
  codeHash?: string;
  snapshotId?: string;
  snapshotExists?: boolean;
  snapshotReason?: string;
  elapsedMs?: number;
  summary?: string;
  entryPoint?: string;
  parameterTypes?: string[];
  parameterSummary?: Record<string, string>;
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
  isFavorite?: boolean;
  tags?: string[];
  runHistory?: RunRecordDto[];
  parameters?: ScriptParameterDef[];
  candidateEntryPoints?: VbaEntryPointInfo[];
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
  metaSaveError?: string;
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
  metaSaveError?: string;
  data?: T;
  error?: string;
}

export interface BatchFileItem {
  originalFilePath: string;
  originalFileName: string;
  fileExtension: string;
  fileSizeBytes: number;
  fileHash: string;
  isSupported: boolean;
  unsupportedReason?: string;
}

export interface BatchJobDefinition {
  jobId: string;
  createdAt: string;
  files: BatchFileItem[];
  macroCode: string;
  macroHash: string;
  entryPoint?: string;
  parametersJson?: string;
  outputDir: string;
  stopOnError: boolean;
  fixedReferenceRiskNotice: string;
}

export type BatchTaskFileStatus = 'pending' | 'running' | 'success' | 'failed' | 'blocked' | 'cancelled';

export interface BatchFileTaskResult {
  fileIndex: number;
  originalFilePath: string;
  originalFileName: string;
  workingCopyPath?: string;
  snapshotPath?: string;
  snapshotId?: string;
  finalOutputPath?: string;
  status: BatchTaskFileStatus;
  failureStage?: string | null;
  error?: string;
  elapsedMs: number;
  executionSummary?: string;
  isPartiallyModified: boolean;
  temporaryCleaned: boolean;
  readback?: any;
}

export type BatchJobStatus = 'running' | 'completed' | 'stopped_on_error' | 'cancelled';

export interface BatchJobSummary {
  jobId: string;
  status: BatchJobStatus;
  totalFiles: number;
  successCount: number;
  failedCount: number;
  blockedCount: number;
  cancelledCount: number;
  pendingCount: number;
  startTime: string;
  endTime?: string;
  cancelRequestedAt?: string;
  riskNotice: string;
  fileResults: BatchFileTaskResult[];
}

type MessageHandler = (res: BridgeResponse) => void;

class NativeBridgeClient {
  private handlers = new Map<string, MessageHandler>();
  private workbookChangeListeners: ((info: WorkbookInfo) => void)[] = [];
  private actionListeners = new Map<string, ((data?: any) => void)[]>();
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

          // 处理宿主主动触发的界面指令（如 open_scripts, open_settings, prompt_run_macro）
          if (raw.action && this.actionListeners.has(raw.action)) {
            const payload = (raw as any).scriptId !== undefined ? (raw as any).scriptId : ((raw as any).payload || raw);
            this.actionListeners.get(raw.action)?.forEach((fn) => {
              try {
                fn(payload);
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

  public onAction(action: string, fn: (data?: any) => void): () => void {
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
          unscannedNotes: '选区共 10 行 × 6 列。本次仅安全抽样前 3 行 × 前 6 列。未扫描其余单元格内容，未扫描全表筛选/隐藏行状态。',
          capturedAt: Date.now(),
          attachmentId: 'att_mock_001'
        } as any,
      };
    }

    if (action === 'get_target_sheets') {
      return {
        ok: true,
        action,
        data: {
          sheets: ['汇总看板', 'Sheet1', '明细数据'],
          activeSheet: '汇总看板',
        } as any,
      };
    }

    if (action === 'inspect_macro_signature') {
      return {
        ok: true,
        action,
        data: {
          entryPoints: [
            {
              name: 'FormatRange',
              isSub: true,
              isPublic: true,
              isSupported: true,
              parameters: [
                { name: 'targetRange', typeName: 'Range', isOptional: false, isSupported: true },
                { name: 'highlightColor', typeName: 'Long', isOptional: true, defaultValue: '65535', isSupported: true }
              ]
            }
          ]
        } as any,
      };
    }

    if (action === 'list_managed_charts') {
      return {
        ok: true,
        action,
        data: [
          {
            chartId: 'mock1234',
            chartName: '__EM_CHART_mock1234',
            title: '销售趋势柱状图',
            chartType: 'column',
            sheetName: payload?.sheetName || 'Sheet1',
            left: 200,
            top: 50,
            width: 480,
            height: 300,
          },
        ] as any,
      };
    }

    if (action === 'execute_quick_chart') {
      return {
        ok: true,
        action,
        data: {
          ok: true,
          chartId: 'mock1234',
          chartName: '__EM_CHART_mock1234',
          action: 'create_new',
          summaryText: '已成功在工作表【汇总看板】新建图表【销售趋势图】（类型: column，包含 1 个数值系列）。',
          snapshotId: 'snap_mock_chart',
          readback: {
            chartId: 'mock1234',
            chartName: '__EM_CHART_mock1234',
            sheetName: '汇总看板',
            chartType: 'column',
            actualChartTypeNum: 51,
            title: '销售趋势图',
            seriesCount: 1,
            seriesNames: ['销售额'],
            categoryAddress: '$A$2:$A$10',
            valuesAddresses: ['$B$2:$B$10'],
            left: 200,
            top: 50,
            width: 480,
            height: 300,
            isReplaced: false,
          },
        } as any,
      };
    }

    if (action === 'export_macro_package') {
      return {
        ok: true,
        action,
        data: {
          ok: true,
          packageFilePath: 'C:\\Users\\MockUser\\Documents\\macros_mock.exmpack',
          outputPath: 'C:\\Users\\MockUser\\Documents\\macros_mock.exmpack',
          exportedCount: 1,
          totalBytes: 2048,
          packageSize: 2048,
          warnings: [],
          manifest: {
            schemaVersion: '1.0',
            packageId: 'mock_pkg_1',
            name: 'MockPackage',
            version: '1.0.0',
            description: 'Mock 导出的宏包',
            exportedAt: new Date().toISOString(),
            exportedBy: 'ExcelMind AI',
            entries: [],
          },
        } as any,
      };
    }

    if (action === 'preview_macro_package') {
      return {
        ok: true,
        action,
        data: {
          ok: true,
          packageFilePath: payload.packageFilePath || 'mock.exmpack',
          packagePath: payload.packageFilePath || 'mock.exmpack',
          macroCount: 1,
          entryCount: 1,
          totalUncompressedBytes: 1500,
          totalUncompressedSize: 1500,
          warnings: [],
          conflictingNames: [],
          nameConflicts: [],
          isTamperedOrCorrupt: false,
          manifest: {
            schemaVersion: '1.0',
            packageId: 'mock_pkg_1',
            name: '示例宏包',
            version: '1.0.0',
            description: '示例无凭据宏包',
            exportedAt: new Date().toISOString(),
            exportedBy: 'ExcelMind AI',
            entries: [
              {
                macroId: 'mock_m_1',
                packageRelativePath: 'scripts/DemoMacro.bas',
                displayName: 'DemoMacro',
                category: '通用',
                description: '演示宏',
                entryPoint: 'Main',
                parameterDefs: [],
                sourceByteLength: 256,
                sha256: 'mockhash123',
              },
            ],
          },
        } as any,
      };
    }

    if (action === 'import_macro_package') {
      return {
        ok: true,
        action,
        data: {
          ok: true,
          importedCount: 1,
          importedIds: ['mock_imported_1'],
          renamedMacros: [],
          nameMapping: {},
          summary: '已成功导入 1 个宏到宏库（零宏自动执行）。',
          warnings: [],
        } as any,
      };
    }

    if (action === 'browse_save_file') {
      return {
        ok: true,
        action,
        data: 'C:\\Users\\MockUser\\Documents\\macros_export.exmpack',
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
    paramsOrRows?: {
      sampleRows?: number;
      sampleCols?: number;
      targetWorkbookName?: string;
      targetWorkbookFullName?: string;
      targetSheetName?: string;
      targetAddress?: string;
      attachmentId?: string;
    } | number,
    sampleColsParam?: number
  ): Promise<{ ok: boolean; data?: SelectionContextData; error?: string; errorType?: string }> {
    const payload: Record<string, string> = {};

    if (typeof paramsOrRows === 'object' && paramsOrRows !== null) {
      payload.sampleRows = (paramsOrRows.sampleRows ?? 5).toString();
      payload.sampleCols = (paramsOrRows.sampleCols ?? 15).toString();
      if (paramsOrRows.targetWorkbookName) payload.targetWorkbookName = paramsOrRows.targetWorkbookName;
      if (paramsOrRows.targetWorkbookFullName) payload.targetWorkbookFullName = paramsOrRows.targetWorkbookFullName;
      if (paramsOrRows.targetSheetName) payload.targetSheetName = paramsOrRows.targetSheetName;
      if (paramsOrRows.targetAddress) payload.targetAddress = paramsOrRows.targetAddress;
      if (paramsOrRows.attachmentId) payload.attachmentId = paramsOrRows.attachmentId;
    } else {
      payload.sampleRows = (paramsOrRows ?? 5).toString();
      payload.sampleCols = (sampleColsParam ?? 15).toString();
    }

    const res = await this.send<SelectionContextData>('get_selection_context', payload);
    if (res.ok && res.data) {
      return { ok: true, data: res.data };
    }
    return {
      ok: false,
      error: res.error || '获取选区失败',
      errorType: (res.data as any)?.errorType || 'exception',
    };
  }

  public async updateScriptTags(id: string, tags: string[]): Promise<{ ok: boolean; error?: string }> {
    const res = await this.send('update_script_tags', {
      id,
      tags: JSON.stringify(tags || []),
    });
    return { ok: res.ok, error: res.error };
  }

  public async updateScriptFavorite(id: string, isFavorite: boolean): Promise<{ ok: boolean; error?: string }> {
    const res = await this.send('update_script_favorite', {
      id,
      isFavorite: isFavorite ? 'true' : 'false',
    });
    return { ok: res.ok, error: res.error };
  }

  public async getTargetSheets(params: {
    targetWorkbookName?: string;
    targetWorkbookFullName?: string;
  }): Promise<{ ok: boolean; sheets: string[]; activeSheet?: string; error?: string }> {
    const res = await this.send<{ sheets: string[]; activeSheet?: string }>('get_target_sheets', {
      targetWorkbookName: params.targetWorkbookName || '',
      targetWorkbookFullName: params.targetWorkbookFullName || '',
    });
    if (res.ok && res.data) {
      return { ok: true, sheets: res.data.sheets || [], activeSheet: res.data.activeSheet };
    }
    return { ok: false, sheets: [], error: res.error || '获取工作表列表失败' };
  }

  public async inspectMacroSignature(vbaCode: string): Promise<{ ok: boolean; entryPoints: VbaEntryPointInfo[]; error?: string }> {
    const res = await this.send<any>('inspect_macro_signature', {
      vbaCode: vbaCode || '',
    });
    if (res.ok && res.data) {
      const entryPoints = Array.isArray(res.data) ? res.data : (res.data.entryPoints || []);
      return { ok: true, entryPoints };
    }
    return { ok: false, entryPoints: [], error: res.error || '解析宏签名失败' };
  }

  public async analyzeDedup(params: {
    targetWorkbookName?: string;
    targetWorkbookFullName?: string;
    sheetName?: string;
    rangeAddress: string;
    hasHeader: boolean;
    keyColumnIndices: number[];
  }): Promise<{ ok: boolean; data?: DedupAnalysisResult; error?: string }> {
    const res = await this.send<DedupAnalysisResult>('analyze_dedup', {
      targetWorkbookName: params.targetWorkbookName || '',
      targetWorkbookFullName: params.targetWorkbookFullName || '',
      sheetName: params.sheetName || '',
      rangeAddress: params.rangeAddress || '',
      hasHeader: params.hasHeader ? 'true' : 'false',
      keyColumnIndices: JSON.stringify(params.keyColumnIndices || []),
    });
    if (res.ok && res.data) {
      return { ok: true, data: res.data };
    }
    return { ok: false, error: res.error || '去重分析失败' };
  }

  public async applyDedup(params: {
    targetWorkbookName?: string;
    targetWorkbookFullName?: string;
    sheetName?: string;
    rangeAddress: string;
    hasHeader: boolean;
    keyColumnIndices: number[];
    mode: 'highlight' | 'export_unique';
    expectedFingerprint?: string;
  }): Promise<{ ok: boolean; data?: DedupExecutionResult; message?: string; error?: string }> {
    const res = await this.send<DedupExecutionResult>('apply_dedup', {
      targetWorkbookName: params.targetWorkbookName || '',
      targetWorkbookFullName: params.targetWorkbookFullName || '',
      sheetName: params.sheetName || '',
      rangeAddress: params.rangeAddress || '',
      hasHeader: params.hasHeader ? 'true' : 'false',
      keyColumnIndices: JSON.stringify(params.keyColumnIndices || []),
      mode: params.mode || 'highlight',
      expectedFingerprint: params.expectedFingerprint || '',
    });
    if (res.ok && res.data) {
      return { ok: true, data: res.data, message: res.data.message || res.message };
    }
    return { ok: false, error: res.error || '执行去重失败' };
  }

  public async analyzeReconcile(params: {
    leftWorkbookName?: string;
    leftSheetName?: string;
    leftRangeAddress: string;
    leftHasHeader: boolean;
    leftKeyCols: number[];
    rightWorkbookName?: string;
    rightSheetName?: string;
    rightRangeAddress: string;
    rightHasHeader: boolean;
    rightKeyCols: number[];
    compareCols: CompareColMappingDto[];
  }): Promise<{ ok: boolean; data?: ReconcileAnalysisResult; error?: string }> {
    const res = await this.send<ReconcileAnalysisResult>('analyze_reconcile', {
      leftWorkbookName: params.leftWorkbookName || '',
      leftSheetName: params.leftSheetName || '',
      leftRangeAddress: params.leftRangeAddress || '',
      leftHasHeader: params.leftHasHeader ? 'true' : 'false',
      leftKeyCols: JSON.stringify(params.leftKeyCols || []),
      rightWorkbookName: params.rightWorkbookName || '',
      rightSheetName: params.rightSheetName || '',
      rightRangeAddress: params.rightRangeAddress || '',
      rightHasHeader: params.rightHasHeader ? 'true' : 'false',
      rightKeyCols: JSON.stringify(params.rightKeyCols || []),
      compareCols: JSON.stringify(params.compareCols || []),
    });
    if (res.ok && res.data) {
      return { ok: true, data: res.data };
    }
    return { ok: false, error: res.error || '对账分析失败' };
  }

  public async applyReconcile(params: {
    leftWorkbookName?: string;
    leftSheetName?: string;
    leftRangeAddress: string;
    leftHasHeader: boolean;
    leftKeyCols: number[];
    rightWorkbookName?: string;
    rightSheetName?: string;
    rightRangeAddress: string;
    rightHasHeader: boolean;
    rightKeyCols: number[];
    compareCols: CompareColMappingDto[];
    outputWorkbookName?: string;
    expectedFingerprint?: string;
  }): Promise<{ ok: boolean; data?: ReconcileExecutionResult; message?: string; error?: string }> {
    const res = await this.send<ReconcileExecutionResult>('apply_reconcile', {
      leftWorkbookName: params.leftWorkbookName || '',
      leftSheetName: params.leftSheetName || '',
      leftRangeAddress: params.leftRangeAddress || '',
      leftHasHeader: params.leftHasHeader ? 'true' : 'false',
      leftKeyCols: JSON.stringify(params.leftKeyCols || []),
      rightWorkbookName: params.rightWorkbookName || '',
      rightSheetName: params.rightSheetName || '',
      rightRangeAddress: params.rightRangeAddress || '',
      rightHasHeader: params.rightHasHeader ? 'true' : 'false',
      rightKeyCols: JSON.stringify(params.rightKeyCols || []),
      compareCols: JSON.stringify(params.compareCols || []),
      outputWorkbookName: params.outputWorkbookName || '',
      expectedFingerprint: params.expectedFingerprint || '',
    });
    if (res.ok && res.data) {
      return { ok: true, data: res.data, message: res.data.message || res.message };
    }
    return { ok: false, error: res.error || '执行对账导出失败' };
  }

  public async validateBatchJob(params: {
    filePaths: string[];
    code: string;
    entryPoint?: string;
    parameters?: string;
    outputDir: string;
    stopOnError?: boolean;
  }): Promise<{ ok: boolean; data?: BatchJobDefinition; error?: string; message?: string }> {
    const res = await this.send<BatchJobDefinition>('validate_batch_job', {
      filePaths: JSON.stringify(params.filePaths || []),
      code: params.code || '',
      entryPoint: params.entryPoint || '',
      parameters: params.parameters || '',
      outputDir: params.outputDir || '',
      stopOnError: params.stopOnError !== false ? 'true' : 'false',
    });
    if (res.ok && res.data) {
      return { ok: true, data: res.data, message: res.message };
    }
    return { ok: false, error: res.error || '固化校验批量任务失败' };
  }

  public async startBatchJob(params: {
    jobId: string;
    filePaths?: string[];
    code?: string;
    entryPoint?: string;
    parameters?: string;
    outputDir?: string;
    stopOnError?: boolean;
    sync?: boolean;
  }): Promise<{ ok: boolean; data?: BatchJobSummary; error?: string; message?: string }> {
    const payload: Record<string, string> = {
      jobId: params.jobId,
    };
    if (params.filePaths && params.filePaths.length > 0) {
      payload.filePaths = JSON.stringify(params.filePaths);
    }
    if (params.code) payload.code = params.code;
    if (params.entryPoint) payload.entryPoint = params.entryPoint;
    if (params.parameters) payload.parameters = params.parameters;
    if (params.outputDir) payload.outputDir = params.outputDir;
    if (params.stopOnError !== undefined) payload.stopOnError = params.stopOnError ? 'true' : 'false';
    if (params.sync) payload.sync = 'true';

    const res = await this.send<BatchJobSummary>('start_batch_job', payload);
    if (res.ok && res.data) {
      return { ok: true, data: res.data, message: res.message };
    }
    return { ok: false, error: res.error || '启动批量任务失败' };
  }

  public async getBatchJobStatus(jobId: string): Promise<{ ok: boolean; data?: BatchJobSummary; error?: string }> {
    const res = await this.send<BatchJobSummary>('get_batch_job_status', {
      jobId: jobId || '',
    });
    if (res.ok && res.data) {
      return { ok: true, data: res.data };
    }
    return { ok: false, error: res.error || '获取批量任务状态失败' };
  }

  public async cancelBatchJob(jobId: string): Promise<{ ok: boolean; message?: string; error?: string }> {
    const res = await this.send<{ message?: string }>('cancel_batch_job', {
      jobId: jobId || '',
    });
    if (res.ok) {
      return { ok: true, message: res.message || '已登记取消请求' };
    }
    return { ok: false, error: res.error || '取消批量任务失败' };
  }

  public async openOutputFolder(params: { outputDir?: string; jobId?: string }): Promise<{ ok: boolean; data?: string; error?: string; message?: string }> {
    const res = await this.send<string>('open_output_folder', {
      outputDir: params.outputDir || '',
      jobId: params.jobId || '',
    });
    if (res.ok) {
      return { ok: true, data: res.data, message: res.message };
    }
    return { ok: false, error: res.error || '打开输出目录失败' };
  }

  public async browseFiles(filter?: string): Promise<{ ok: boolean; data?: string[]; error?: string }> {
    const res = await this.send<string[]>('browse_files', {
      filter: filter || '',
    });
    if (res.ok && res.data) {
      return { ok: true, data: res.data };
    }
    return { ok: false, error: res.error || '选择文件失败' };
  }

  public async browseFolder(description?: string): Promise<{ ok: boolean; data?: string; error?: string }> {
    const res = await this.send<string>('browse_folder', {
      description: description || '',
    });
    if (res.ok && res.data) {
      return { ok: true, data: res.data };
    }
    return { ok: false, error: res.error || '选择文件夹失败' };
  }

  public async analyzeConsolidation(params: {
    sources: ConsolidationSourceDef[];
    outputColumnOrder?: string[];
    columnMappings?: ConsolidationColMapping[];
  }): Promise<{ ok: boolean; data?: ConsolidationAnalysisResult; error?: string }> {
    const res = await this.send<ConsolidationAnalysisResult>('analyze_consolidation', {
      sources: JSON.stringify(params.sources || []),
      outputColumnOrder: params.outputColumnOrder ? JSON.stringify(params.outputColumnOrder) : '',
      columnMappings: params.columnMappings ? JSON.stringify(params.columnMappings) : '',
    });
    if (res.ok && res.data) {
      return { ok: true, data: res.data };
    }
    return { ok: false, error: res.error || '多文件列名对齐分析失败' };
  }

  public async applyConsolidation(params: {
    sources: ConsolidationSourceDef[];
    outputFilePath: string;
    outputColumnOrder?: string[];
    columnMappings?: ConsolidationColMapping[];
    expectedFingerprint?: string;
    includeMetadataCols?: boolean;
  }): Promise<{ ok: boolean; data?: ConsolidationExecutionResult; error?: string }> {
    const res = await this.send<ConsolidationExecutionResult>('apply_consolidation', {
      sources: JSON.stringify(params.sources || []),
      outputFilePath: params.outputFilePath || '',
      outputColumnOrder: params.outputColumnOrder ? JSON.stringify(params.outputColumnOrder) : '',
      columnMappings: params.columnMappings ? JSON.stringify(params.columnMappings) : '',
      expectedFingerprint: params.expectedFingerprint || '',
      includeMetadataCols: params.includeMetadataCols !== false ? 'true' : 'false',
    });
    if (res.ok && res.data) {
      return { ok: true, data: res.data };
    }
    return { ok: false, error: res.error || '多文件列名对齐汇总写入失败' };
  }

  public async listWorkflows(): Promise<{ ok: boolean; data?: WorkflowDefinition[]; error?: string }> {
    const res = await this.send<WorkflowDefinition[]>('list_workflows');
    if (res.ok && res.data) {
      return { ok: true, data: res.data };
    }
    return { ok: false, error: res.error || '获取工作流列表失败' };
  }

  public async getWorkflow(workflowId: string): Promise<{ ok: boolean; data?: { definition: WorkflowDefinition; runs: WorkflowRunRecord[] }; error?: string }> {
    const res = await this.send<{ definition: WorkflowDefinition; runs: WorkflowRunRecord[] }>('get_workflow', { workflowId });
    if (res.ok && res.data) {
      return { ok: true, data: res.data };
    }
    return { ok: false, error: res.error || '获取工作流详情失败' };
  }

  public async saveWorkflow(workflow: WorkflowDefinition): Promise<{ ok: boolean; data?: WorkflowDefinition; message?: string; error?: string }> {
    const res = await this.send<WorkflowDefinition>('save_workflow', {
      workflow: JSON.stringify(workflow)
    });
    if (res.ok && res.data) {
      return { ok: true, data: res.data, message: res.message };
    }
    return { ok: false, error: res.error || '保存工作流失败' };
  }

  public async deleteWorkflow(workflowId: string): Promise<{ ok: boolean; message?: string; error?: string }> {
    const res = await this.send<{ message: string }>('delete_workflow', { workflowId });
    if (res.ok) {
      return { ok: true, message: res.message };
    }
    return { ok: false, error: res.error || '删除工作流失败' };
  }

  public async executeWorkflow(params: {
    workflow?: WorkflowDefinition;
    workflowId?: string;
    targetWorkbookName?: string;
    targetWorkbookFullName?: string;
  }): Promise<{ ok: boolean; data?: WorkflowRunRecord; message?: string; error?: string }> {
    const payload: any = {
      targetWorkbookName: params.targetWorkbookName || '',
      targetWorkbookFullName: params.targetWorkbookFullName || ''
    };
    if (params.workflowId) payload.workflowId = params.workflowId;
    if (params.workflow) payload.workflow = JSON.stringify(params.workflow);

    const res = await this.send<WorkflowRunRecord>('execute_workflow', payload);
    if (res.ok && res.data) {
      return { ok: true, data: res.data, message: res.message };
    }
    return { ok: false, data: res.data, error: res.error || '执行工作流未成功' };
  }

  public async cancelWorkflow(): Promise<{ ok: boolean; message?: string }> {
    const res = await this.send<{ message: string }>('cancel_workflow');
    return { ok: res.ok, message: res.message };
  }

  public async listManagedCharts(params: {
    targetWorkbookName?: string;
    targetWorkbookFullName?: string;
    sheetName: string;
  }): Promise<ManagedChartInfo[]> {
    const res = await this.send<ManagedChartInfo[]>('list_managed_charts', {
      targetWorkbookName: params.targetWorkbookName || '',
      targetWorkbookFullName: params.targetWorkbookFullName || '',
      sheetName: params.sheetName || '',
    });
    return res.ok && res.data ? res.data : [];
  }

  public async executeQuickChart(params: QuickChartParams): Promise<QuickChartResult> {
    const res = await this.send<QuickChartResult>('execute_quick_chart', {
      params: JSON.stringify(params),
    });
    if (res.ok && res.data) {
      return res.data;
    }
    return {
      ok: false,
      error: res.error || '执行快捷图表生成失败',
      action: params.action,
    };
  }

  public async previewExternalData(params: ExternalDataPreviewParams): Promise<ExternalDataPreviewResult> {
    const res = await this.send<ExternalDataPreviewResult>('preview_external_data', {
      params: JSON.stringify(params),
    });
    if (res.ok && res.data) {
      return res.data;
    }
    return {
      ok: false,
      error: res.error || '外部数据预览失败',
      columns: [],
      detectedTypes: [],
      sampleRows: [],
      totalRowsEstimate: 0,
      totalCols: 0,
      isTruncated: false,
    };
  }

  public async importExternalData(params: ExternalDataImportParams): Promise<ExternalDataImportResult> {
    const res = await this.send<ExternalDataImportResult>('import_external_data', {
      params: JSON.stringify(params),
    });
    if (res.ok && res.data) {
      return res.data;
    }
    return {
      ok: false,
      error: res.error || '外部数据导入失败',
      failureStage: 'precheck',
      importedRowCount: 0,
      importedColCount: 0,
      elapsedMs: 0,
    };
  }

  public async listDataSourceConfigs(): Promise<DataSourceConfigDto[]> {
    const res = await this.send<DataSourceConfigDto[]>('list_data_source_configs');
    return res.ok && res.data ? res.data : [];
  }

  public async saveDataSourceConfig(config: DataSourceConfigDto, secret?: string): Promise<{ ok: boolean; message?: string; error?: string }> {
    const payload: any = {
      params: JSON.stringify(config),
    };
    if (secret) payload.secret = secret;
    const res = await this.send<any>('save_data_source_config', payload);
    return {
      ok: res.ok,
      message: res.message,
      error: res.error,
    };
  }

  public async browseSaveFile(params?: { title?: string; defaultName?: string; filter?: string }): Promise<{ ok: boolean; data?: string; error?: string }> {
    const res = await this.send<string>('browse_save_file', {
      title: params?.title || '保存宏包',
      defaultName: params?.defaultName || 'macros.exmpack',
      filter: params?.filter || 'ExcelMind 宏包 (*.exmpack)|*.exmpack|所有文件 (*.*)|*.*',
    });
    if (res.ok && res.data) {
      return { ok: true, data: res.data };
    }
    return { ok: false, error: res.error || '未选择保存路径' };
  }

  public async exportMacroPackage(params: MacroPackageExportParams): Promise<MacroPackageExportResult> {
    const res = await this.send<MacroPackageExportResult>('export_macro_package', {
      macroIds: JSON.stringify(params.macroIds || []),
      targetFilePath: params.outputPath || params.targetFilePath || '',
      outputPath: params.outputPath || params.targetFilePath || '',
      packageName: params.packageName || '',
      packageVersion: params.packageVersion || '1.0.0',
      packageDescription: params.packageDescription || '',
      ignoreSensitiveWarnings: params.ignoreWarnings || params.ignoreSensitiveWarnings ? 'true' : 'false',
      ignoreWarnings: params.ignoreWarnings || params.ignoreSensitiveWarnings ? 'true' : 'false',
      params: JSON.stringify({
        macroIds: params.macroIds || [],
        targetFilePath: params.outputPath || params.targetFilePath || '',
        packageName: params.packageName || '',
        packageVersion: params.packageVersion || '1.0.0',
        packageDescription: params.packageDescription || '',
        ignoreSensitiveWarnings: params.ignoreWarnings || params.ignoreSensitiveWarnings || false,
      }),
    });
    if (res.ok && res.data) {
      return res.data;
    }
    return {
      ok: false,
      error: res.error || '导出宏包失败',
      exportedCount: 0,
      totalBytes: 0,
      packageSize: 0,
      warnings: (res.data as any)?.warnings || (res.data as any)?.sensitiveWarnings || [],
    };
  }

  public async previewMacroPackage(packagePath: string): Promise<MacroPackagePreviewResult> {
    const res = await this.send<MacroPackagePreviewResult>('preview_macro_package', {
      packageFilePath: packagePath || '',
      filePath: packagePath || '',
      path: packagePath || '',
      params: JSON.stringify({
        packageFilePath: packagePath || '',
      }),
    });
    if (res.ok && res.data) {
      return res.data;
    }
    return {
      ok: false,
      error: res.error || '解析宏包预览失败',
      macroCount: 0,
      entryCount: 0,
      totalUncompressedBytes: 0,
      totalUncompressedSize: 0,
      warnings: (res.data as any)?.warnings || (res.data as any)?.sensitiveWarnings || [],
      conflictingNames: [],
      nameConflicts: [],
      isTamperedOrCorrupt: Boolean((res.data as any)?.isTamperedOrCorrupt),
    };
  }

  public async importMacroPackage(params: MacroPackageImportParams): Promise<MacroPackageImportResult> {
    const res = await this.send<MacroPackageImportResult>('import_macro_package', {
      packageFilePath: params.packageFilePath || params.packagePath || '',
      packagePath: params.packageFilePath || params.packagePath || '',
      selectedMacroIds: params.selectedMacroIds ? JSON.stringify(params.selectedMacroIds) : '',
      conflictResolution: params.conflictResolution || params.nameConflictResolution || 'rename_both',
      nameConflictResolution: params.conflictResolution || params.nameConflictResolution || 'rename_both',
      ignoreWarnings: params.ignoreWarnings ? 'true' : 'false',
      params: JSON.stringify({
        packageFilePath: params.packageFilePath || params.packagePath || '',
        selectedMacroIds: params.selectedMacroIds || [],
        conflictResolution: params.conflictResolution || params.nameConflictResolution || 'rename_both',
        ignoreWarnings: Boolean(params.ignoreWarnings),
      }),
    });
    if (res.ok && res.data) {
      return res.data;
    }
    return {
      ok: false,
      error: res.error || '导入宏包失败',
      importedCount: 0,
      importedMacros: [],
      importedIds: [],
      renamedMacros: [],
      warnings: (res.data as any)?.warnings || [],
    };
  }

  public async previewDiagnostics(customLogContent?: string): Promise<DiagnosticsPreviewResult> {
    const res = await this.send<DiagnosticsPreviewResult>('preview_diagnostics', {
      params: JSON.stringify({
        customLogContent: customLogContent || '',
      }),
    });
    if (res.ok && res.data) {
      return res.data;
    }
    return {
      ok: false,
      error: res.error || '生成诊断预览失败',
      summary: {} as any,
      includedFiles: [],
      includedCategories: [],
      excludedCategories: [],
      sanitizedLogPreview: [],
      estimatedTotalBytes: 0,
      sanitizedLogLinesCount: 0,
      omittedSensitiveLinesCount: 0,
    };
  }

  public async exportDiagnostics(targetZipPath?: string, customLogContent?: string): Promise<DiagnosticsExportResult> {
    const res = await this.send<DiagnosticsExportResult>('export_diagnostics', {
      targetZipPath: targetZipPath || '',
      filePath: targetZipPath || '',
      params: JSON.stringify({
        targetZipPath: targetZipPath || '',
        customLogContent: customLogContent || '',
      }),
    });
    if (res.ok && res.data) {
      return res.data;
    }
    return {
      ok: false,
      error: res.error || '导出诊断包失败',
      zipFilePath: '',
      zipSizeBytes: 0,
      sha256: '',
      includedFiles: [],
      totalLogLines: 0,
      omittedSensitiveLinesCount: 0,
    };
  }
}

export interface QuickChartParams {
  targetWorkbookName?: string;
  targetWorkbookFullName?: string;
  sourceSheet: string;
  sourceRange: string;
  hasHeaders?: boolean;
  dataStartRow?: number;
  dataEndRow?: number;
  categoryColIndex: number;
  categoryColName?: string;
  seriesColIndices: number[];
  seriesNames?: string[];
  chartType: 'column' | 'line' | 'pie';
  title?: string;
  targetSheet?: string;
  placementMode?: 'cell' | 'coordinates';
  targetCell?: string;
  left?: number;
  top?: number;
  width?: number;
  height?: number;
  action?: 'create_new' | 'replace_existing';
  targetChartId?: string;
  errorHandling?: 'reject_on_invalid' | 'coerce_zero';
}

export interface ChartReadbackDto {
  chartId: string;
  chartName: string;
  sheetName: string;
  chartType: string;
  actualChartTypeNum: number;
  title?: string;
  seriesCount: number;
  seriesNames: string[];
  categoryAddress?: string;
  valuesAddresses: string[];
  left: number;
  top: number;
  width: number;
  height: number;
  isReplaced: boolean;
}

export interface QuickChartResult {
  ok: boolean;
  chartId?: string;
  chartName?: string;
  action?: string;
  summaryText?: string;
  snapshotId?: string;
  readback?: ChartReadbackDto;
  error?: string;
  failureStage?: string;
  recoveryNotice?: string;
}

export interface ManagedChartInfo {
  chartId: string;
  chartName: string;
  title?: string;
  chartType: string;
  sheetName: string;
  left: number;
  top: number;
  width: number;
  height: number;
}

export interface ChartStepParams {
  sourceSheet?: string;
  sourceRange?: string;
  hasHeaders?: boolean;
  dataStartRow?: number;
  dataEndRow?: number;
  categoryColIndex?: number;
  categoryColName?: string;
  seriesColIndices?: number[];
  seriesNames?: string[];
  chartType?: 'column' | 'line' | 'pie';
  title?: string;
  targetSheet?: string;
  placementMode?: 'cell' | 'coordinates';
  targetCell?: string;
  left?: number;
  top?: number;
  width?: number;
  height?: number;
  action?: 'create_new' | 'replace_existing';
  targetChartId?: string;
  errorHandling?: 'reject_on_invalid' | 'coerce_zero';
}

export interface DedupStepParams {
  targetSheet?: string;
  rangeAddress?: string;
  hasHeader?: boolean;
  keyColumns?: number[];
  mode?: 'highlight' | 'export_unique';
  outputSheetName?: string;
  expectedFingerprint?: string;
}

export interface ReconcileStepParams {
  leftSheet?: string;
  leftAddress?: string;
  leftHasHeader?: boolean;
  leftKeyCols?: number[];
  rightSheet?: string;
  rightAddress?: string;
  rightHasHeader?: boolean;
  rightKeyCols?: number[];
  compareCols?: CompareColMappingDto[];
  outputSheetName?: string;
  expectedFingerprint?: string;
}

export interface SavedMacroStepParams {
  scriptId?: string;
  entryPoint?: string;
  macroSha256?: string;
  targetSheet?: string;
  targetRange?: string;
  parameters?: Array<{ name: string; type: string; value: any }>;
  declaredOutputSheet?: string;
  declaredOutputRange?: string;
}

export interface StepOutputReference {
  targetWorkbookName: string;
  sheetName: string;
  rangeAddress: string;
  hasHeader: boolean;
  headerRowIndex: number;
  rowCount: number;
  columnCount: number;
  dataFingerprint?: string;
}

export interface WorkflowStepDefinition {
  stepIndex: 1 | 2;
  stepName: string;
  toolType: 'dedup' | 'reconcile' | 'saved_macro' | 'chart';
  inputSource: 'initial_selection' | 'prev_step_output';
  dedupParams?: DedupStepParams;
  reconcileParams?: ReconcileStepParams;
  macroParams?: SavedMacroStepParams;
  chartParams?: ChartStepParams;
}

export interface WorkflowDefinition {
  workflowId?: string;
  definitionVersion?: number;
  name: string;
  description?: string;
  targetWorkbookName?: string;
  targetWorkbookPath?: string;
  isPreset?: boolean;
  steps: WorkflowStepDefinition[];
  createdAt?: string;
  updatedAt?: string;
}

export interface StepRunResult {
  stepIndex: number;
  stepName: string;
  toolType: string;
  status: 'success' | 'failed' | 'skipped' | 'blocked';
  failureStage?: string;
  summaryText: string;
  error?: string;
  outputRef?: StepOutputReference;
  elapsedMs: number;
}

export interface WorkflowRunRecord {
  runId: string;
  workflowId: string;
  definitionVersion: number;
  executedAt: string;
  targetWorkbookName: string;
  targetWorkbookPath: string;
  snapshotId: string;
  snapshotExists: boolean;
  status: 'completed' | 'stopped_on_step1' | 'stopped_on_step2' | 'blocked' | 'cancelled';
  phase: 'precheck' | 'snapshot' | 'step1' | 'step2' | 'completed';
  completedSteps: number;
  failedStepIndex: number;
  failureMessage?: string;
  recoveryNotice?: string;
  stepResults: StepRunResult[];
  totalElapsedMs: number;
}

export const bridge = new NativeBridgeClient();

export interface CsvParseOptions {
  encoding?: string;
  delimiter?: string;
  hasHeader?: boolean;
  quote?: string;
}

export interface JsonParseOptions {
  arrayPath?: string;
}

export interface HttpGetOptions {
  url: string;
  headers?: Record<string, string>;
  timeoutSeconds?: number;
  allowRedirect?: boolean;
  whitelistRules?: string[];
  credentialKey?: string;
}

export interface ExternalDataPreviewParams {
  sourceType: 'csv' | 'json' | 'http_get';
  filePath?: string;
  csvOptions?: CsvParseOptions;
  jsonOptions?: JsonParseOptions;
  httpOptions?: HttpGetOptions;
  previewRowCount?: number;
}

export interface ExternalDataPreviewResult {
  ok: boolean;
  error?: string;
  sourceType?: string;
  previewId?: string;
  dataFingerprint?: string;
  columns: string[];
  detectedTypes: string[];
  unsupportedColumns?: string[];
  sampleRows: string[][];
  totalRowsEstimate: number;
  totalCols: number;
  isTruncated: boolean;
  sanitizedSummary?: string;
  nullHandlingRules?: Record<string, string>;
  capacityLimits?: Record<string, string>;
}

export interface ExternalDataImportParams {
  targetWorkbookFullName?: string;
  targetWorkbookName?: string;
  targetSheetName?: string;
  sourceType: 'csv' | 'json' | 'http_get';
  filePath?: string;
  previewId?: string;
  expectedFingerprint?: string;
  csvOptions?: CsvParseOptions;
  jsonOptions?: JsonParseOptions;
  httpOptions?: HttpGetOptions;
  selectedColumns?: string[];
}

export interface ExternalDataImportResult {
  ok: boolean;
  error?: string;
  failureStage?: string;
  targetWorkbookFullName?: string;
  sheetName?: string;
  importedRowCount: number;
  importedColCount: number;
  snapshotId?: string;
  elapsedMs: number;
  recoveryNotice?: string;
  sanitizedSource?: string;
}

export interface DataSourceConfigDto {
  id?: string;
  name: string;
  sourceType: 'csv' | 'json' | 'http_get';
  pathOrUrl: string;
  csvOptions?: CsvParseOptions;
  jsonOptions?: JsonParseOptions;
  httpOptions?: HttpGetOptions;
  hasCredential?: boolean;
  createdAt?: string;
  updatedAt?: string;
}

export interface SensitiveScanWarning {
  macroDisplayName?: string;
  macroId?: string;
  macroName?: string;
  fieldName?: string;
  location?: string;
  matchedPattern?: string;
  patternType?: string;
  snippet: string;
  warningMessage?: string;
}

export interface MacroPackageEntry {
  macroId: string;
  packageRelativePath: string;
  displayName: string;
  category: string;
  description: string;
  entryPoint: string;
  parameterDefs: ScriptParameterDef[];
  sourceByteLength: number;
  sha256: string;
}

export interface MacroPackageManifest {
  schemaVersion: string;
  packageId: string;
  name: string;
  version: string;
  description: string;
  exportedAt: string;
  exportedBy: string;
  entries: MacroPackageEntry[];
}

export interface MacroPackageExportParams {
  macroIds: string[];
  outputPath?: string;
  targetFilePath?: string;
  packageName?: string;
  packageVersion?: string;
  packageDescription?: string;
  ignoreWarnings?: boolean;
  ignoreSensitiveWarnings?: boolean;
}

export interface MacroPackageExportResult {
  ok: boolean;
  error?: string;
  outputPath?: string;
  packageFilePath?: string;
  exportedCount: number;
  totalBytes?: number;
  packageSize?: number;
  warnings: SensitiveScanWarning[];
  sensitiveWarnings?: SensitiveScanWarning[];
  hasSensitiveWarnings?: boolean;
  manifest?: MacroPackageManifest;
}

export interface MacroPackagePreviewResult {
  ok: boolean;
  error?: string;
  packagePath?: string;
  packageFilePath?: string;
  manifest?: MacroPackageManifest;
  macroCount?: number;
  entryCount?: number;
  totalUncompressedSize?: number;
  totalUncompressedBytes?: number;
  warnings: SensitiveScanWarning[];
  sensitiveWarnings?: SensitiveScanWarning[];
  hasSensitiveWarnings?: boolean;
  conflictingNames?: string[];
  nameConflicts?: string[];
  isTamperedOrCorrupt?: boolean;
}

export interface MacroPackageImportParams {
  packagePath?: string;
  packageFilePath?: string;
  selectedMacroIds?: string[];
  conflictResolution?: 'rename_both' | 'skip' | string;
  nameConflictResolution?: 'rename_both' | 'skip' | string;
  ignoreWarnings?: boolean;
}

export interface MacroPackageImportResult {
  ok: boolean;
  error?: string;
  importedCount: number;
  importedMacros?: ScriptItem[];
  importedIds?: string[];
  renamedMacros?: string[];
  nameMapping?: Record<string, string>;
  summary?: string;
  warnings?: SensitiveScanWarning[];
}

export interface DiagnosticsEnvironmentSummary {
  appName: string;
  appVersion: string;
  buildCommit: string;
  clrVersion: string;
  osVersion: string;
  osArchitecture: string;
  processArchitecture: string;
  excelVersion: string;
  excelBitness: string;
  webView2Version: string;
  taskPaneStatus: string;
  hasActiveWorkbook: boolean;
  activeWorkbookMaskedName: string;
  lastFailureStage: string;
  lastErrorSummary: string;
  generatedAtUtc: string;
}

export interface DiagnosticsPreviewResult {
  ok: boolean;
  error?: string;
  summary: DiagnosticsEnvironmentSummary;
  includedFiles: string[];
  includedCategories: string[];
  excludedCategories: string[];
  sanitizedLogPreview: string[];
  estimatedTotalBytes: number;
  sanitizedLogLinesCount: number;
  omittedSensitiveLinesCount: number;
}

export interface DiagnosticsExportResult {
  ok: boolean;
  error?: string;
  zipFilePath: string;
  zipSizeBytes: number;
  sha256: string;
  includedFiles: string[];
  totalLogLines: number;
  omittedSensitiveLinesCount: number;
}



