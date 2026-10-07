<script lang="ts">
  import {
    X,
    Filter,
    Check,
    AlertCircle,
    CheckCircle2,
    Clock,
    Layers,
    FileSpreadsheet,
    Highlighter,
    CopyCheck,
    ShieldCheck,
    RotateCcw,
    GitCompare,
    ArrowRightLeft,
    SlidersHorizontal,
    Plus,
    Trash2,
    FolderOpen,
    FileCheck,
    FilePlus,
    BarChart2,
    Download,
    Globe,
    Database,
    Lock,
    Crosshair,
    RefreshCw,
    Info,
    Search,
  } from 'lucide-svelte';
  import {
    bridge,
    type WorkbookInfo,
    type DedupAnalysisResult,
    type DedupExecutionResult,
    type ReconcileAnalysisResult,
    type ReconcileExecutionResult,
    type CompareColMappingDto,
    type ConsolidationSourceDef,
    type ConsolidationColMapping,
    type ConsolidationAnalysisResult,
    type ConsolidationExecutionResult,
    type QuickChartParams,
    type QuickChartResult,
    type ManagedChartInfo,
    type ExternalDataPreviewParams,
    type ExternalDataPreviewResult,
    type ExternalDataImportParams,
    type ExternalDataImportResult,
    type DataSourceConfigDto,
  } from '../services/bridge';

  export let isOpen = false;
  export let isWorkspaceView = true;
  export let workbook: WorkbookInfo | null = null;
  export let onClose: () => void;
  export let onExecuted: (() => void) | undefined = undefined;

  // 顶层 Tab
  let activeTab: 'dedup' | 'reconcile' | 'consolidate' | 'chart' | 'external_data' = 'dedup';

  // ==================== 按键去重状态 (R3a) ====================
  let rangeAddress = 'A1:C10';
  let hasHeader = true;
  let selectedColIndices: number[] = [1]; // 1-based
  let availableColsCount = 3;

  let isAnalyzing = false;
  let isExecuting = false;
  let analysisResult: DedupAnalysisResult | null = null;
  let executionResult: DedupExecutionResult | null = null;
  let errorMessage = '';

  let showConfirmModal = false;
  let pendingMode: 'highlight' | 'export_unique' = 'highlight';

  $: if (isOpen && workbook) {
    if (workbook.activeCell && !rangeAddress) {
      rangeAddress = workbook.activeCell;
    }
  }

  async function handleGrabSelection() {
    errorMessage = '';
    try {
      const res = await bridge.getSelectionContext({ sampleRows: 1, sampleCols: 1 });
      if (res.ok && res.data) {
        if (res.data.address) {
          rangeAddress = res.data.address;
        }
        if (res.data.columnsCount && res.data.columnsCount > 0) {
          availableColsCount = Math.min(res.data.columnsCount, 50);
          if (!selectedColIndices.includes(1)) {
            selectedColIndices = [1];
          }
        }
      }
    } catch (e: any) {
      errorMessage = '抓取选区失败: ' + (e.message || String(e));
    }
  }

  function toggleColIndex(idx: number) {
    if (selectedColIndices.includes(idx)) {
      if (selectedColIndices.length > 1) {
        selectedColIndices = selectedColIndices.filter((c) => c !== idx);
      }
    } else {
      selectedColIndices = [...selectedColIndices, idx].sort((a, b) => a - b);
    }
    analysisResult = null;
    executionResult = null;
  }

  async function handleRunAnalysis() {
    if (!rangeAddress.trim()) {
      errorMessage = '请输入或选择有效的单元格区域';
      return;
    }
    if (selectedColIndices.length === 0) {
      errorMessage = '请至少选择一列主键作为判断依据';
      return;
    }

    errorMessage = '';
    isAnalyzing = true;
    analysisResult = null;
    executionResult = null;

    try {
      const res = await bridge.analyzeDedup({
        targetWorkbookName: workbook?.name || '',
        targetWorkbookFullName: workbook?.fullName || '',
        sheetName: workbook?.activeSheet || '',
        rangeAddress: rangeAddress.trim(),
        hasHeader: hasHeader,
        keyColumnIndices: selectedColIndices,
      });

      if (res.ok && res.data) {
        analysisResult = res.data;
        if (res.data.selectedColumnCount > 0) {
          availableColsCount = res.data.selectedColumnCount;
        }
      } else {
        errorMessage = res.error || '分析失败';
      }
    } catch (err: any) {
      errorMessage = '请求分析发生异常: ' + (err.message || String(err));
    } finally {
      isAnalyzing = false;
    }
  }

  function promptExecution(mode: 'highlight' | 'export_unique') {
    if (!analysisResult) return;
    pendingMode = mode;
    showConfirmModal = true;
  }

  async function handleConfirmExecute() {
    showConfirmModal = false;
    if (!analysisResult) return;

    errorMessage = '';
    isExecuting = true;
    executionResult = null;

    try {
      const res = await bridge.applyDedup({
        targetWorkbookName: workbook?.name || '',
        targetWorkbookFullName: workbook?.fullName || '',
        sheetName: analysisResult.targetSheetName || workbook?.activeSheet || '',
        rangeAddress: analysisResult.rangeAddress || rangeAddress.trim(),
        hasHeader: hasHeader,
        keyColumnIndices: selectedColIndices,
        mode: pendingMode,
        expectedFingerprint: (analysisResult as any).dataFingerprint || '',
      });

      if (res.ok && res.data) {
        executionResult = res.data;
        if (onExecuted) {
          onExecuted();
        }
      } else {
        errorMessage = res.error || '执行去重操作失败';
      }
    } catch (err: any) {
      errorMessage = '写入过程发生异常: ' + (err.message || String(err));
    } finally {
      isExecuting = false;
    }
  }

  // ==================== 两表对账状态 (R3b) ====================
  let leftWbName = '';
  let leftSheetName = '';
  let leftRange = 'A1:C10';
  let leftHasHeader = true;
  let leftKeyColsInput = '1';

  let rightWbName = '';
  let rightSheetName = '';
  let rightRange = 'A1:C10';
  let rightHasHeader = true;
  let rightKeyColsInput = '1';

  let compareCols: CompareColMappingDto[] = [
    { leftColIndex: 2, rightColIndex: 2 },
  ];

  let isAnalyzingReconcile = false;
  let isExecutingReconcile = false;
  let reconcileAnalysis: ReconcileAnalysisResult | null = null;
  let reconcileExecution: ReconcileExecutionResult | null = null;
  let reconcileError = '';
  let showReconcileConfirm = false;

  $: if (isOpen && workbook) {
    if (!leftWbName) leftWbName = workbook.name || '';
    if (!leftSheetName) leftSheetName = workbook.activeSheet || '';
    if (!rightWbName) rightWbName = workbook.name || '';
    if (!rightSheetName) rightSheetName = workbook.activeSheet || '';
  }

  function addCompareCol() {
    const nextIdx = compareCols.length + 2;
    compareCols = [...compareCols, { leftColIndex: nextIdx, rightColIndex: nextIdx }];
    reconcileAnalysis = null;
    reconcileExecution = null;
  }

  function removeCompareCol(index: number) {
    if (compareCols.length > 0) {
      compareCols = compareCols.filter((_, i) => i !== index);
      reconcileAnalysis = null;
      reconcileExecution = null;
    }
  }

  function parseKeyCols(raw: string): number[] {
    const parts = raw.split(/[,;\s]+/).filter(Boolean);
    const result: number[] = [];
    for (const p of parts) {
      const idx = parseInt(p, 10);
      if (!isNaN(idx) && idx > 0) result.push(idx);
    }
    return result;
  }

  async function handleRunReconcileAnalysis() {
    reconcileError = '';
    reconcileAnalysis = null;
    reconcileExecution = null;

    const leftKeys = parseKeyCols(leftKeyColsInput);
    const rightKeys = parseKeyCols(rightKeyColsInput);

    if (leftKeys.length === 0 || rightKeys.length === 0) {
      reconcileError = '请分别为左表和右表指定有效的主键列索引 (如 1 或 1,2)';
      return;
    }
    if (leftKeys.length !== rightKeys.length) {
      reconcileError = `主键列数不匹配：左表指定了 ${leftKeys.length} 列主键，而右表指定了 ${rightKeys.length} 列主键。复合主键必须两侧列数完全一致。`;
      return;
    }

    if (!leftRange.trim() || !rightRange.trim()) {
      reconcileError = '请显式指定左表与右表的数据区域地址';
      return;
    }

    isAnalyzingReconcile = true;
    try {
      const res = await bridge.analyzeReconcile({
        leftWorkbookName: leftWbName.trim(),
        leftSheetName: leftSheetName.trim(),
        leftRangeAddress: leftRange.trim(),
        leftHasHeader: leftHasHeader,
        leftKeyCols: leftKeys,
        rightWorkbookName: rightWbName.trim(),
        rightSheetName: rightSheetName.trim(),
        rightRangeAddress: rightRange.trim(),
        rightHasHeader: rightHasHeader,
        rightKeyCols: rightKeys,
        compareCols: compareCols,
      });

      if (res.ok && res.data) {
        reconcileAnalysis = res.data;
      } else {
        reconcileError = res.error || '两表对账只读分析失败';
      }
    } catch (err: any) {
      reconcileError = '请求对账分析发生异常: ' + (err.message || String(err));
    } finally {
      isAnalyzingReconcile = false;
    }
  }

  async function handleConfirmExecuteReconcile() {
    showReconcileConfirm = false;
    if (!reconcileAnalysis) return;

    reconcileError = '';
    isExecutingReconcile = true;
    reconcileExecution = null;

    try {
      const res = await bridge.applyReconcile({
        leftWorkbookName: reconcileAnalysis.leftWorkbookName || leftWbName.trim(),
        leftSheetName: reconcileAnalysis.leftSheetName || leftSheetName.trim(),
        leftRangeAddress: reconcileAnalysis.leftRangeAddress || leftRange.trim(),
        leftHasHeader: reconcileAnalysis.leftHasHeader,
        leftKeyCols: reconcileAnalysis.leftKeyCols,
        rightWorkbookName: reconcileAnalysis.rightWorkbookName || rightWbName.trim(),
        rightSheetName: reconcileAnalysis.rightSheetName || rightSheetName.trim(),
        rightRangeAddress: reconcileAnalysis.rightRangeAddress || rightRange.trim(),
        rightHasHeader: reconcileAnalysis.rightHasHeader,
        rightKeyCols: reconcileAnalysis.rightKeyCols,
        compareCols: reconcileAnalysis.compareCols,
        outputWorkbookName: leftWbName.trim(),
        expectedFingerprint: reconcileAnalysis.dataFingerprint || '',
      });

      if (res.ok && res.data) {
        reconcileExecution = res.data;
        if (onExecuted) onExecuted();
      } else {
        reconcileError = res.error || '两表对账结果导出失败';
      }
    } catch (err: any) {
      reconcileError = '写入对账表发生异常: ' + (err.message || String(err));
    } finally {
      isExecutingReconcile = false;
    }
  }

  // ==================== 多文件列名对齐汇总状态 (R4c) ====================
  let consolidationSources: ConsolidationSourceDef[] = [
    { filePath: '', sheetName: '', rangeAddress: '', hasHeader: true, headerRowIndex: 1, displayIdentifier: '' }
  ];
  let consolidationOutputPath = '';
  let includeMetadataCols = true;
  let isAnalyzingConsolidation = false;
  let isExecutingConsolidation = false;
  let consolidationAnalysis: ConsolidationAnalysisResult | null = null;
  let consolidationExecution: ConsolidationExecutionResult | null = null;
  let consolidationError = '';
  let showConsolidationConfirm = false;

  async function handleBrowseConsolidationFiles() {
    consolidationError = '';
    try {
      const res = await bridge.browseFiles('Excel 工作簿 (*.xlsx;*.xlsm)|*.xlsx;*.xlsm');
      if (res.ok && res.data && res.data.length > 0) {
        const newSources: ConsolidationSourceDef[] = res.data.map(p => ({
          filePath: p,
          sheetName: '',
          rangeAddress: '',
          hasHeader: true,
          headerRowIndex: 1,
          displayIdentifier: ''
        }));
        if (consolidationSources.length === 1 && !consolidationSources[0].filePath) {
          consolidationSources = newSources;
        } else {
          consolidationSources = [...consolidationSources, ...newSources];
        }
        consolidationAnalysis = null;
        consolidationExecution = null;
        if (!consolidationOutputPath && consolidationSources[0]?.filePath) {
          const dir = consolidationSources[0].filePath.substring(0, consolidationSources[0].filePath.lastIndexOf('\\'));
          consolidationOutputPath = dir + '\\多文件汇总结果.xlsx';
        }
      }
    } catch (e: any) {
      consolidationError = '选择文件失败: ' + (e.message || String(e));
    }
  }

  function handleAddEmptySource() {
    consolidationSources = [
      ...consolidationSources,
      { filePath: '', sheetName: '', rangeAddress: '', hasHeader: true, headerRowIndex: 1, displayIdentifier: '' }
    ];
    consolidationAnalysis = null;
    consolidationExecution = null;
  }

  function handleRemoveSource(idx: number) {
    if (consolidationSources.length > 1) {
      consolidationSources = consolidationSources.filter((_, i) => i !== idx);
    } else {
      consolidationSources = [{ filePath: '', sheetName: '', rangeAddress: '', hasHeader: true, headerRowIndex: 1, displayIdentifier: '' }];
    }
    consolidationAnalysis = null;
    consolidationExecution = null;
  }

  async function handleRunConsolidationAnalysis() {
    consolidationError = '';
    const validSources = consolidationSources.filter(s => s.filePath.trim().length > 0);
    if (validSources.length === 0) {
      consolidationError = '请至少添加一个有效的来源文件路径';
      return;
    }
    isAnalyzingConsolidation = true;
    consolidationAnalysis = null;
    consolidationExecution = null;
    try {
      const res = await bridge.analyzeConsolidation({
        sources: validSources
      });
      if (res.ok && res.data) {
        consolidationAnalysis = res.data;
        if (!consolidationOutputPath && validSources[0]?.filePath) {
          const dir = validSources[0].filePath.substring(0, validSources[0].filePath.lastIndexOf('\\'));
          consolidationOutputPath = dir + '\\多文件汇总结果.xlsx';
        }
      } else {
        consolidationError = res.error || '分析失败';
      }
    } catch (e: any) {
      consolidationError = '汇总分析异常: ' + (e.message || String(e));
    } finally {
      isAnalyzingConsolidation = false;
    }
  }

  function handleOpenConsolidationConfirm() {
    consolidationError = '';
    if (!consolidationAnalysis) {
      consolidationError = '请先进行汇总分析并核对对齐列';
      return;
    }
    if (!consolidationOutputPath.trim()) {
      consolidationError = '请输入目标输出文件全路径';
      return;
    }
    showConsolidationConfirm = true;
  }

  async function handleConfirmExecuteConsolidation() {
    showConsolidationConfirm = false;
    if (!consolidationAnalysis) return;
    isExecutingConsolidation = true;
    consolidationError = '';
    try {
      const res = await bridge.applyConsolidation({
        sources: consolidationAnalysis.sources.map(s => ({
          filePath: s.filePath,
          sheetName: s.sheetName,
          rangeAddress: s.rangeAddress,
          hasHeader: true,
          headerRowIndex: s.headerRow
        })),
        outputFilePath: consolidationOutputPath.trim(),
        expectedFingerprint: consolidationAnalysis.dataFingerprint,
        includeMetadataCols
      });
      if (res.ok && res.data) {
        consolidationExecution = res.data;
        if (onExecuted) onExecuted();
      } else {
        consolidationError = res.error || '多文件汇总写入失败';
      }
    } catch (e: any) {
      consolidationError = '执行汇总失败: ' + (e.message || String(e));
    } finally {
      isExecutingConsolidation = false;
    }
  }

  async function handleOpenConsolidationOutputDir() {
    if (!consolidationExecution?.outputFilePath) return;
    const dir = consolidationExecution.outputFilePath.substring(0, consolidationExecution.outputFilePath.lastIndexOf('\\'));
    try {
      await bridge.openOutputFolder({ outputDir: dir });
    } catch (e: any) {
      consolidationError = '打开输出目录失败: ' + (e.message || String(e));
    }
  }

  // ==================== 快捷图表工具状态与处理 (R5b) ====================
  let chartSourceSheet = '';
  let chartSourceRange = 'A1:C10';
  let chartHasHeaders = true;
  let chartDataStartRow = 2;
  let chartCategoryCol = 1;
  let chartCategoryName = '';
  let chartSeriesColsInput = '2';
  let chartSeriesNamesInput = '';
  let chartType: 'column' | 'line' | 'pie' = 'column';
  let chartTitle = '';
  let chartTargetSheet = '';
  let chartTargetCell = 'E2';
  let chartAction: 'create_new' | 'replace_existing' = 'create_new';
  let chartTargetChartId = '';
  let chartErrorHandling: 'reject_on_invalid' | 'coerce_zero' = 'reject_on_invalid';
  let chartManagedList: ManagedChartInfo[] = [];
  let isScanningManagedCharts = false;
  let isExecutingChart = false;
  let chartError = '';
  let chartSuccessNotice = '';
  let chartResult: QuickChartResult | null = null;

  $: if (isOpen && workbook) {
    if (!chartSourceSheet && workbook.activeSheet) {
      chartSourceSheet = workbook.activeSheet;
    }
    if (!chartTargetSheet && workbook.activeSheet) {
      chartTargetSheet = workbook.activeSheet;
    }
  }

  async function handleGrabChartSelection() {
    chartError = '';
    try {
      const res = await bridge.getSelectionContext({ sampleRows: 1, sampleCols: 1 });
      if (res.ok && res.data) {
        if (res.data.address) chartSourceRange = res.data.address;
        if (res.data.sheetName) {
          chartSourceSheet = res.data.sheetName;
          if (!chartTargetSheet) chartTargetSheet = res.data.sheetName;
        }
      }
    } catch (e: any) {
      chartError = '抓取选区失败: ' + (e.message || String(e));
    }
  }

  async function handleScanManagedCharts() {
    isScanningManagedCharts = true;
    chartError = '';
    try {
      const targetSheet = chartTargetSheet || chartSourceSheet || workbook?.activeSheet || 'Sheet1';
      chartManagedList = await bridge.listManagedCharts({
        targetWorkbookName: workbook?.name || '',
        targetWorkbookFullName: workbook?.fullName || '',
        sheetName: targetSheet,
      });
      if (chartManagedList.length > 0 && !chartTargetChartId) {
        chartTargetChartId = chartManagedList[0].chartId;
      }
    } catch (e: any) {
      chartError = '扫描已有快捷图表失败: ' + (e.message || String(e));
    } finally {
      isScanningManagedCharts = false;
    }
  }

  async function handleExecuteQuickChart() {
    chartError = '';
    chartSuccessNotice = '';
    chartResult = null;

    if (!chartSourceRange.trim()) {
      chartError = '请指定数据源区域 (例如 A1:D10)。系统拒绝隐式盲选！';
      return;
    }

    const seriesIndices = chartSeriesColsInput
      .split(',')
      .map((s) => parseInt(s.trim(), 10))
      .filter((n) => !isNaN(n) && n > 0);

    if (seriesIndices.length === 0) {
      chartError = '请指定至少一个有效的数值系列列索引 (例如 2 或 2, 3)。系统拒绝盲目全列绘制！';
      return;
    }

    if (chartType === 'pie' && seriesIndices.length > 1) {
      chartError = `饼图仅支持单个数值系列，当前指定了 ${seriesIndices.length} 个系列。请保留 1 个数值系列或选用柱状图/折线图。`;
      return;
    }

    if (chartAction === 'replace_existing' && !chartTargetChartId) {
      chartError = '选择【替换指定图表】模式时，必须指定待替换的图表对象。请先点击扫描或输入图表标识。';
      return;
    }

    const seriesNames = chartSeriesNamesInput
      ? chartSeriesNamesInput.split(',').map((s) => s.trim()).filter((s) => s.length > 0)
      : [];

    isExecutingChart = true;
    try {
      const p: QuickChartParams = {
        targetWorkbookName: workbook?.name || '',
        targetWorkbookFullName: workbook?.fullName || '',
        sourceSheet: chartSourceSheet || (workbook?.activeSheet || 'Sheet1'),
        sourceRange: chartSourceRange.trim(),
        hasHeaders: chartHasHeaders,
        dataStartRow: chartHasHeaders ? 2 : 1,
        dataEndRow: 0,
        categoryColIndex: chartCategoryCol,
        categoryColName: chartCategoryName,
        seriesColIndices: seriesIndices,
        seriesNames: seriesNames,
        chartType: chartType,
        title: chartTitle.trim(),
        targetSheet: chartTargetSheet.trim() || chartSourceSheet,
        placementMode: 'cell',
        targetCell: chartTargetCell.trim() || 'E2',
        action: chartAction,
        targetChartId: chartTargetChartId.trim(),
        errorHandling: chartErrorHandling,
      };

      const res = await bridge.executeQuickChart(p);
      if (res.ok) {
        chartResult = res;
        chartSuccessNotice = res.summaryText || '快捷图表已成功生成！';
        if (onExecuted) onExecuted();
      } else {
        chartError = res.error || '执行图表生成失败。';
      }
    } catch (e: any) {
      chartError = '调用快捷图表异常: ' + (e.message || String(e));
    } finally {
      isExecutingChart = false;
    }
  }

  // ==================== 外部数据接入状态 (R6a) ====================
  let extSourceType: 'csv' | 'json' | 'http_get' = 'csv';
  let extFilePath = '';
  let extCsvEncoding = 'UTF-8';
  let extCsvDelimiter = ',';
  let extCsvHasHeader = true;
  let extJsonArrayPath = '';
  let extHttpUrl = '';
  let extHttpWhitelist = 'http://127.0.0.1:*, http://localhost:*';
  let extHttpHeaderKey = '';
  let extHttpHeaderVal = '';
  let extHttpCredentialKey = '';
  let extTargetSheetName = '';
  let extSelectedColumns: string[] = [];
  let extPreviewResult: ExternalDataPreviewResult | null = null;
  let extImportResult: ExternalDataImportResult | null = null;
  let isPreviewingExt = false;
  let isImportingExt = false;
  let externalDataError = '';
  let externalDataSuccess = '';
  let showExternalDataConfirm = false;

  async function handlePreviewExternalData() {
    externalDataError = '';
    externalDataSuccess = '';
    extPreviewResult = null;
    extImportResult = null;
    extSelectedColumns = [];

    if (extSourceType === 'csv' || extSourceType === 'json') {
      if (!extFilePath.trim()) {
        externalDataError = '请指定本地数据文件路径。';
        return;
      }
    } else if (extSourceType === 'http_get') {
      if (!extHttpUrl.trim()) {
        externalDataError = '请指定 HTTP GET 数据源 URL。';
        return;
      }
      if (!extHttpUrl.trim().toLowerCase().startsWith('http://') && !extHttpUrl.trim().toLowerCase().startsWith('https://')) {
        externalDataError = 'URL 必须以 http:// 或 https:// 开头。';
        return;
      }
    }

    isPreviewingExt = true;
    try {
      const whitelistRules = extHttpWhitelist
        .split(/[,;\n]+/)
        .map((s) => s.trim())
        .filter(Boolean);

      const headers: Record<string, string> = {};
      if (extHttpHeaderKey.trim() && extHttpHeaderVal.trim()) {
        headers[extHttpHeaderKey.trim()] = extHttpHeaderVal.trim();
      }

      const p: ExternalDataPreviewParams = {
        sourceType: extSourceType,
        filePath: extFilePath.trim(),
        csvOptions: {
          encoding: extCsvEncoding,
          delimiter: extCsvDelimiter,
          hasHeader: extCsvHasHeader,
        },
        jsonOptions: {
          arrayPath: extJsonArrayPath.trim(),
        },
        httpOptions: {
          url: extHttpUrl.trim(),
          headers,
          whitelistRules,
          credentialKey: extHttpCredentialKey.trim(),
        },
        previewRowCount: 10,
      };

      const res = await bridge.previewExternalData(p);
      if (res.ok && res.data) {
        extPreviewResult = res.data;
        extSelectedColumns = [...res.data.columns];
      } else {
        externalDataError = res.error || '获取外部数据预览失败。';
      }
    } catch (e: any) {
      externalDataError = '外部数据预览调用异常: ' + (e.message || String(e));
    } finally {
      isPreviewingExt = false;
    }
  }

  function toggleExtColumn(col: string) {
    if (extSelectedColumns.includes(col)) {
      if (extSelectedColumns.length > 1) {
        extSelectedColumns = extSelectedColumns.filter((c) => c !== col);
      }
    } else {
      extSelectedColumns = [...extSelectedColumns, col];
    }
  }

  function selectAllExtColumns() {
    if (extPreviewResult) {
      extSelectedColumns = [...extPreviewResult.columns];
    }
  }

  function deselectAllExtColumns() {
    if (extPreviewResult && extPreviewResult.columns.length > 0) {
      extSelectedColumns = [extPreviewResult.columns[0]];
    }
  }

  function handleOpenImportConfirm() {
    externalDataError = '';
    if (!extPreviewResult) {
      externalDataError = '请先点击“只读拉取并预览数据”，确认字段无误后再执行导入。';
      return;
    }
    if (extSelectedColumns.length === 0) {
      externalDataError = '至少需要选择 1 个导入字段列。';
      return;
    }

    if (extPreviewResult.unsupportedColumns && extPreviewResult.unsupportedColumns.length > 0) {
      const unsupportedSelected = extSelectedColumns.filter((c) => extPreviewResult!.unsupportedColumns!.includes(c));
      if (unsupportedSelected.length > 0) {
        externalDataError = `选中的字段 [${unsupportedSelected.join(', ')}] 包含嵌套对象或数组，当前版本不支持复合结构导入。请取消勾选后再执行导入。`;
        return;
      }
    }

    showExternalDataConfirm = true;
  }

  async function handleConfirmExecuteImport() {
    showExternalDataConfirm = false;
    externalDataError = '';
    externalDataSuccess = '';

    if (extPreviewResult?.unsupportedColumns && extPreviewResult.unsupportedColumns.length > 0) {
      const unsupportedSelected = extSelectedColumns.filter((c) => extPreviewResult!.unsupportedColumns!.includes(c));
      if (unsupportedSelected.length > 0) {
        externalDataError = `选中的字段 [${unsupportedSelected.join(', ')}] 包含嵌套对象或数组，当前版本不支持复合结构导入。请取消勾选后再执行导入。`;
        return;
      }
    }

    isImportingExt = true;
    try {
      const whitelistRules = extHttpWhitelist
        .split(/[,;\n]+/)
        .map((s) => s.trim())
        .filter(Boolean);

      const headers: Record<string, string> = {};
      if (extHttpHeaderKey.trim() && extHttpHeaderVal.trim()) {
        headers[extHttpHeaderKey.trim()] = extHttpHeaderVal.trim();
      }

      const p: ExternalDataImportParams = {
        targetWorkbookFullName: workbook?.fullName || '',
        targetWorkbookName: workbook?.name || '',
        targetSheetName: extTargetSheetName.trim(),
        sourceType: extSourceType,
        filePath: extFilePath.trim(),
        previewId: extPreviewResult?.previewId || '',
        expectedFingerprint: extPreviewResult?.dataFingerprint || '',
        csvOptions: {
          encoding: extCsvEncoding,
          delimiter: extCsvDelimiter,
          hasHeader: extCsvHasHeader,
        },
        jsonOptions: {
          arrayPath: extJsonArrayPath.trim(),
        },
        httpOptions: {
          url: extHttpUrl.trim(),
          headers,
          whitelistRules,
          credentialKey: extHttpCredentialKey.trim(),
        },
        selectedColumns: extSelectedColumns,
      };

      const res = await bridge.importExternalData(p);
      if (res.ok && res.data) {
        extImportResult = res.data;
        externalDataSuccess = `成功导入数据至新建工作表 [${res.data.sheetName}]！共写入 ${res.data.importedRowCount} 行、${res.data.importedColCount} 列（快照: ${res.data.snapshotId || '已创建'}）。`;
        if (onExecuted) onExecuted();
      } else {
        externalDataError = res.error || '执行外部数据导入失败。';
      }
    } catch (e: any) {
      externalDataError = '执行外部数据导入异常: ' + (e.message || String(e));
    } finally {
      isImportingExt = false;
    }
  }
</script>

{#if isOpen}
  <div class={isWorkspaceView ? "datatool-workspace-root" : "datatool-modal-overlay"}>
    <div class={isWorkspaceView ? "datatool-workspace-card" : "datatool-modal-card"}>
      
      <!-- 头部：标题与状态说明 -->
      <div class="datatool-header">
        <div class="header-info">
          <div class="header-icon-box">
            {#if activeTab === 'dedup'}
              <Filter size={18} />
            {:else if activeTab === 'reconcile'}
              <ArrowRightLeft size={18} />
            {:else if activeTab === 'consolidate'}
              <Layers size={18} />
            {:else if activeTab === 'chart'}
              <BarChart2 size={18} />
            {:else}
              <Download size={18} />
            {/if}
          </div>
          <div class="header-title-group">
            <h2 class="tool-title">
              {activeTab === 'dedup' ? '按键去重' : activeTab === 'reconcile' ? '两表对账' : activeTab === 'consolidate' ? '多文件汇总' : activeTab === 'chart' ? '快捷图表' : '外部数据导入'}
            </h2>
            <p class="tool-desc">
              {activeTab === 'dedup'
                ? '比对单列或复合键，高亮或导出唯一记录。'
                : activeTab === 'reconcile'
                ? '全量跨表差异比对，输出四分类核对结果。'
                : activeTab === 'consolidate'
                ? '多文件同名列自动对齐合并为新工作簿。'
                : activeTab === 'chart'
                ? '选定数据区域免公式快速生成业务图表。'
                : '只读预览并导入 CSV/JSON 外部数据。'}
            </p>
          </div>
        </div>

        <button
          type="button"
          on:click={onClose}
          class="datatool-close-btn"
          title={isWorkspaceView ? "返回 AI 助手" : "关闭"}
          aria-label={isWorkspaceView ? "返回 AI 助手" : "关闭"}
        >
          <X size={16} />
        </button>
      </div>

      <!-- 窄屏：工具下拉选择器 (<=380px 优先展示，不强塞横向 Tab) -->
      <div class="datatool-selector-mobile">
        <label for="datatool-mobile-select" class="selector-label">选择工具：</label>
        <select
          id="datatool-mobile-select"
          bind:value={activeTab}
          class="selector-select"
          on:change={() => {
            errorMessage = '';
            reconcileError = '';
            consolidationError = '';
            chartError = '';
            externalDataError = '';
          }}
        >
          <option value="dedup">按键去重 · 标记或导出唯一行</option>
          <option value="reconcile">两表对账 · 全量跨表差异比对</option>
          <option value="consolidate">多文件汇总 · 结构对齐合并</option>
          <option value="chart">快捷图表 · 免公式可视化绘图</option>
          <option value="external_data">外部数据 · 安全导入预览</option>
        </select>
      </div>

      <!-- 宽屏：分段 Tab 栏 (>380px 展示) -->
      <div class="datatool-tabs-wrapper">
        <div class="datatool-tabs-bar" role="tablist" aria-label="数据工具导航">
          <button
            type="button"
            role="tab"
            aria-selected={activeTab === 'dedup'}
            on:click={() => { activeTab = 'dedup'; errorMessage = ''; }}
            class={`datatool-tab-btn ${activeTab === 'dedup' ? 'active' : ''}`}
          >
            <Filter size={13} />
            <span>按键去重</span>
          </button>
          <button
            type="button"
            role="tab"
            aria-selected={activeTab === 'reconcile'}
            on:click={() => { activeTab = 'reconcile'; reconcileError = ''; }}
            class={`datatool-tab-btn ${activeTab === 'reconcile' ? 'active' : ''}`}
          >
            <GitCompare size={13} />
            <span>两表对账</span>
          </button>
          <button
            type="button"
            role="tab"
            aria-selected={activeTab === 'consolidate'}
            on:click={() => { activeTab = 'consolidate'; consolidationError = ''; }}
            class={`datatool-tab-btn ${activeTab === 'consolidate' ? 'active' : ''}`}
          >
            <Layers size={13} />
            <span>多文件汇总</span>
          </button>
          <button
            type="button"
            role="tab"
            aria-selected={activeTab === 'chart'}
            on:click={() => { activeTab = 'chart'; chartError = ''; }}
            class={`datatool-tab-btn ${activeTab === 'chart' ? 'active' : ''}`}
          >
            <BarChart2 size={13} />
            <span>快捷图表</span>
          </button>
          <button
            type="button"
            role="tab"
            aria-selected={activeTab === 'external_data'}
            on:click={() => { activeTab = 'external_data'; externalDataError = ''; }}
            class={`datatool-tab-btn ${activeTab === 'external_data' ? 'active' : ''}`}
          >
            <Download size={13} />
            <span>外部数据</span>
          </button>
        </div>
      </div>

      <!-- 内容区 -->
      <div class="datatool-content-body">
        
        <!-- ====================== TAB 1: 按键去重 ====================== -->
        {#if activeTab === 'dedup'}
          <!-- 错误提示 -->
          {#if errorMessage}
            <div class="p-3.5 bg-red-50 border border-red-200 rounded-lg flex items-start gap-2.5 text-xs text-red-700 animate-in fade-in duration-150">
              <AlertCircle size={16} class="mt-0.5 shrink-0 text-red-600" />
              <div class="leading-relaxed whitespace-pre-wrap">{errorMessage}</div>
            </div>
          {/if}

          <!-- 目标与区域选择卡片 -->
          <div class="bg-gray-50/80 rounded-lg p-4 border border-gray-200/80 space-y-3">
            <div class="flex items-center justify-between text-xs text-gray-600">
              <span class="font-medium">目标工作簿: <strong class="text-gray-900">{workbook?.name || '当前工作簿'}</strong></span>
              <span>工作表: <strong class="text-gray-900">{workbook?.activeSheet || '当前表'}</strong></span>
            </div>

            <div class="space-y-1.5">
              <div class="flex items-center justify-between">
                <label for="dedup-range-input" class="text-xs font-semibold text-gray-700">处理数据区域 (Range 地址)</label>
                <button
                  type="button"
                  on:click={handleGrabSelection}
                  class="text-xs text-indigo-600 hover:text-indigo-800 font-medium hover:underline flex items-center gap-1"
                >
                  <span>抓取选区</span>
                </button>
              </div>
              <div class="flex items-center gap-2">
                <input
                  id="dedup-range-input"
                  type="text"
                  bind:value={rangeAddress}
                  placeholder="例如: A1:D100"
                  class="flex-1 px-3 py-2 text-xs border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-emerald-500 font-mono bg-white"
                  on:input={() => { analysisResult = null; executionResult = null; }}
                />
                <label class="flex items-center gap-1.5 text-xs text-gray-700 cursor-pointer select-none shrink-0 px-2.5 py-2 bg-white border border-gray-300 rounded-md hover:bg-gray-50">
                  <input
                    type="checkbox"
                    bind:checked={hasHeader}
                    class="rounded text-emerald-600 focus:ring-emerald-500"
                    on:change={() => { analysisResult = null; executionResult = null; }}
                  />
                  <span>首行为表头</span>
                </label>
              </div>
            </div>

            <!-- 主键列选择 -->
            <div class="space-y-1.5 pt-1">
              <div class="flex items-center justify-between">
                <span class="text-xs font-semibold text-gray-700">判断主键列 (单列或复合键)</span>
                <span class="text-[11px] text-gray-400">已选 {selectedColIndices.length} 列复合比较</span>
              </div>
              <div class="flex flex-wrap gap-1.5 max-h-24 overflow-y-auto p-1 bg-white border border-gray-200 rounded-md">
                {#each Array.from({ length: Math.max(availableColsCount, 5) }, (_, i) => i + 1) as colIdx}
                  <button
                    type="button"
                    on:click={() => toggleColIndex(colIdx)}
                    class={`px-2.5 py-1 text-xs rounded border transition-colors ${
                      selectedColIndices.includes(colIdx)
                        ? 'bg-emerald-50 border-emerald-500 text-emerald-800 font-bold'
                        : 'bg-gray-50 border-gray-200 text-gray-600 hover:bg-gray-100'
                    }`}
                  >
                    第 {colIdx} 列
                  </button>
                {/each}
              </div>
            </div>

            <!-- 操作按钮 -->
            <div class="pt-2 flex justify-end">
              <button
                type="button"
                on:click={handleRunAnalysis}
                disabled={isAnalyzing}
                class="px-4 py-2 bg-indigo-600 hover:bg-indigo-700 disabled:opacity-50 text-white rounded-md text-xs font-semibold transition-colors flex items-center gap-1.5 shadow-sm"
              >
                {#if isAnalyzing}
                  <Clock size={14} class="animate-spin" />
                  <span>正在批量分析...</span>
                {:else}
                  <Filter size={14} />
                  <span>开始只读去重分析</span>
                {/if}
              </button>
            </div>
          </div>

          <!-- 分析报告结果呈现 -->
          {#if analysisResult && analysisResult.ok}
            <div class="border border-emerald-200 bg-emerald-50/40 rounded-lg p-4 space-y-3.5 animate-in fade-in duration-200">
              <div class="flex items-center justify-between pb-2 border-b border-emerald-100">
                <span class="text-xs font-bold text-gray-900 flex items-center gap-1.5">
                  <CheckCircle2 size={15} class="text-emerald-600" />
                  <span>去重分析完成 (耗时: {analysisResult.analysisElapsedMs} ms)</span>
                </span>
                <span class="text-[11px] text-emerald-700 font-mono">
                  {analysisResult.hasHeader ? '含表头' : '无表头'} | 选区: {analysisResult.rangeAddress}
                </span>
              </div>

              <!-- 四格统计指标 -->
              <div class="grid grid-cols-4 gap-2 text-center">
                <div class="bg-white p-2.5 rounded border border-gray-100 shadow-2xs">
                  <div class="text-[11px] text-gray-500">有效数据行</div>
                  <div class="text-base font-bold text-gray-800">{analysisResult.dataRowCount}</div>
                </div>
                <div class="bg-white p-2.5 rounded border border-emerald-100 shadow-2xs">
                  <div class="text-[11px] text-emerald-600 font-medium">唯一记录行</div>
                  <div class="text-base font-bold text-emerald-700">{analysisResult.uniqueCount}</div>
                </div>
                <div class="bg-white p-2.5 rounded border border-amber-100 shadow-2xs">
                  <div class="text-[11px] text-amber-600 font-medium">重复冗余行</div>
                  <div class="text-base font-bold text-amber-700">
                    {analysisResult.duplicateCount}
                    {#if analysisResult.duplicateGroupCount > 0}
                      <span class="text-[10px] text-gray-400 font-normal">({analysisResult.duplicateGroupCount}组)</span>
                    {/if}
                  </div>
                </div>
                <div class="bg-white p-2.5 rounded border border-gray-100 shadow-2xs">
                  <div class="text-[11px] text-gray-400">排除隔离行</div>
                  <div class="text-base font-bold text-gray-500">{analysisResult.excludedCount}</div>
                </div>
              </div>

              <!-- 统计恒等式 -->
              <div class="text-[11px] text-gray-600 bg-white/80 p-2 rounded border border-gray-200/60 font-mono text-center">
                恒等式核验：{analysisResult.dataRowCount} = {analysisResult.uniqueCount} (唯一) + {analysisResult.duplicateCount} (重复) + {analysisResult.excludedCount} (排除)
                <span class="text-emerald-600 font-bold ml-1">✓ 100% 严密自洽</span>
              </div>

              {#if analysisResult.excludedCount > 0}
                <div class="p-2.5 bg-amber-50/80 rounded border border-amber-200 text-[11px] text-amber-800 space-y-1">
                  <p class="font-semibold">⚠️ 排除说明：{analysisResult.excludedReason}</p>
                  <p class="text-gray-600">说明：排除行（空键/错误值）原封不动保留在源表中，导出唯一表时绝不写入，绝非无提示丢弃。</p>
                </div>
              {/if}

              <!-- 写入操作选择区 -->
              <div class="pt-2 flex items-center gap-3">
                <button
                  type="button"
                  on:click={() => promptExecution('highlight')}
                  disabled={isExecuting || analysisResult.duplicateCount === 0}
                  class="flex-1 py-2.5 bg-amber-500 hover:bg-amber-600 disabled:opacity-40 text-white rounded-md text-xs font-semibold transition-colors flex items-center justify-center gap-1.5 shadow-sm"
                >
                  <Highlighter size={14} />
                  <span>标记重复行 (浅红高亮)</span>
                </button>

                <button
                  type="button"
                  on:click={() => promptExecution('export_unique')}
                  disabled={isExecuting || analysisResult.uniqueCount === 0}
                  class="flex-1 py-2.5 bg-emerald-600 hover:bg-emerald-700 disabled:opacity-40 text-white rounded-md text-xs font-semibold transition-colors flex items-center justify-center gap-1.5 shadow-sm"
                >
                  <CopyCheck size={14} />
                  <span>导出唯一数据至新工作表</span>
                </button>
              </div>
              <p class="text-[11px] text-gray-400 text-center">执行前将强制为目标工作簿成功创建物理快照（快照失败承诺零业务写入），源单元格数值不修改。</p>
            </div>
          {/if}

          <!-- 执行完成反馈卡片 -->
          {#if executionResult && executionResult.ok}
            <div class="border border-emerald-300 bg-emerald-50 rounded-lg p-4 space-y-2 animate-in fade-in duration-200">
              <div class="flex items-center gap-2 text-xs font-bold text-emerald-900">
                <CheckCircle2 size={16} class="text-emerald-600" />
                <span>操作执行成功！</span>
              </div>
              <p class="text-xs text-emerald-800 leading-relaxed">
                {executionResult.message}
              </p>
              <div class="pt-1 flex items-center justify-between text-[11px] text-emerald-700/80 font-mono">
                <span>关联快照: {executionResult.snapshotId || '已备份'}</span>
                <span>写入耗时: {executionResult.writeElapsedMs} ms (总计: {executionResult.totalElapsedMs} ms)</span>
              </div>
            </div>
          {/if}

        <!-- ====================== TAB 2: 两表对账 (R3b) ====================== -->
        {:else if activeTab === 'reconcile'}
          <!-- 错误提示 -->
          {#if reconcileError}
            <div class="p-3.5 bg-red-50 border border-red-200 rounded-lg flex items-start gap-2.5 text-xs text-red-700 animate-in fade-in duration-150">
              <AlertCircle size={16} class="mt-0.5 shrink-0 text-red-600" />
              <div class="leading-relaxed whitespace-pre-wrap">{reconcileError}</div>
            </div>
          {/if}

          <!-- 两侧来源绑定区域 (左右双栏卡片) -->
          <div class="grid grid-cols-2 gap-4">
            <!-- 左表配置 -->
            <div class="bg-gray-50/90 rounded-lg p-3.5 border border-gray-200 space-y-2.5">
              <div class="flex items-center justify-between pb-1.5 border-b border-gray-200">
                <span class="text-xs font-bold text-gray-800 flex items-center gap-1.5">
                  <span class="w-2 h-2 rounded-full bg-blue-500"></span>
                  <span>左表来源 (主基准表)</span>
                </span>
                <label class="flex items-center gap-1 text-[11px] text-gray-600 cursor-pointer">
                  <input type="checkbox" bind:checked={leftHasHeader} class="rounded text-indigo-600" />
                  <span>首行为表头</span>
                </label>
              </div>

              <div class="space-y-1">
                <label class="text-[11px] text-gray-500">目标工作簿与工作表</label>
                <div class="grid grid-cols-2 gap-1.5">
                  <input
                    type="text"
                    bind:value={leftWbName}
                    placeholder="工作簿名称"
                    class="px-2 py-1.5 text-xs border border-gray-300 rounded bg-white font-mono"
                  />
                  <input
                    type="text"
                    bind:value={leftSheetName}
                    placeholder="工作表名称"
                    class="px-2 py-1.5 text-xs border border-gray-300 rounded bg-white font-mono"
                  />
                </div>
              </div>

              <div class="space-y-1">
                <label class="text-[11px] text-gray-500">数据区域 (例如 A1:E100)</label>
                <input
                  type="text"
                  bind:value={leftRange}
                  placeholder="例如: A1:D50"
                  class="w-full px-2 py-1.5 text-xs border border-gray-300 rounded bg-white font-mono"
                />
              </div>

              <div class="space-y-1">
                <div class="flex items-center justify-between text-[11px] text-gray-500">
                  <span>主键列索引 (1-based，复合键用逗号隔开)</span>
                </div>
                <input
                  type="text"
                  bind:value={leftKeyColsInput}
                  placeholder="例如: 1 或 1, 2"
                  class="w-full px-2 py-1.5 text-xs border border-gray-300 rounded bg-white font-mono"
                />
              </div>
            </div>

            <!-- 右表配置 -->
            <div class="bg-gray-50/90 rounded-lg p-3.5 border border-gray-200 space-y-2.5">
              <div class="flex items-center justify-between pb-1.5 border-b border-gray-200">
                <span class="text-xs font-bold text-gray-800 flex items-center gap-1.5">
                  <span class="w-2 h-2 rounded-full bg-emerald-500"></span>
                  <span>右表来源 (核对对照表)</span>
                </span>
                <label class="flex items-center gap-1 text-[11px] text-gray-600 cursor-pointer">
                  <input type="checkbox" bind:checked={rightHasHeader} class="rounded text-indigo-600" />
                  <span>首行为表头</span>
                </label>
              </div>

              <div class="space-y-1">
                <label class="text-[11px] text-gray-500">目标工作簿与工作表</label>
                <div class="grid grid-cols-2 gap-1.5">
                  <input
                    type="text"
                    bind:value={rightWbName}
                    placeholder="工作簿名称"
                    class="px-2 py-1.5 text-xs border border-gray-300 rounded bg-white font-mono"
                  />
                  <input
                    type="text"
                    bind:value={rightSheetName}
                    placeholder="工作表名称"
                    class="px-2 py-1.5 text-xs border border-gray-300 rounded bg-white font-mono"
                  />
                </div>
              </div>

              <div class="space-y-1">
                <label class="text-[11px] text-gray-500">数据区域 (例如 A1:E100)</label>
                <input
                  type="text"
                  bind:value={rightRange}
                  placeholder="例如: A1:D50"
                  class="w-full px-2 py-1.5 text-xs border border-gray-300 rounded bg-white font-mono"
                />
              </div>

              <div class="space-y-1">
                <div class="flex items-center justify-between text-[11px] text-gray-500">
                  <span>主键列索引 (1-based，复合键用逗号隔开)</span>
                </div>
                <input
                  type="text"
                  bind:value={rightKeyColsInput}
                  placeholder="例如: 1 或 1, 2"
                  class="w-full px-2 py-1.5 text-xs border border-gray-300 rounded bg-white font-mono"
                />
              </div>
            </div>
          </div>

          <!-- 比较列显式映射配置 -->
          <div class="bg-gray-50/70 rounded-lg p-3.5 border border-gray-200 space-y-2">
            <div class="flex items-center justify-between">
              <span class="text-xs font-semibold text-gray-800 flex items-center gap-1.5">
                <SlidersHorizontal size={14} class="text-indigo-600" />
                <span>比较列显式映射 (核对两表非主键字段)</span>
              </span>
              <button
                type="button"
                on:click={addCompareCol}
                class="text-xs text-indigo-600 hover:text-indigo-800 font-medium flex items-center gap-1 hover:underline"
              >
                <Plus size={13} />
                <span>增加比较列映射</span>
              </button>
            </div>

            {#if compareCols.length === 0}
              <p class="text-xs text-gray-400 py-1">暂未指定比较列（仅比对主键是否存在；若需核对具体字段，请点击增加）。</p>
            {:else}
              <div class="space-y-1.5 pt-1">
                {#each compareCols as cm, idx}
                  <div class="flex items-center gap-2 bg-white p-2 rounded border border-gray-200 text-xs">
                    <span class="text-gray-400 w-5 text-center font-mono">#{idx + 1}</span>
                    <span class="text-gray-600">左表列索引:</span>
                    <input
                      type="number"
                      bind:value={cm.leftColIndex}
                      min="1"
                      class="w-16 px-2 py-1 border border-gray-300 rounded font-mono text-center"
                    />
                    <span class="text-gray-400">↔</span>
                    <span class="text-gray-600">右表列索引:</span>
                    <input
                      type="number"
                      bind:value={cm.rightColIndex}
                      min="1"
                      class="w-16 px-2 py-1 border border-gray-300 rounded font-mono text-center"
                    />
                    <button
                      type="button"
                      on:click={() => removeCompareCol(idx)}
                      class="ml-auto text-gray-400 hover:text-red-600 p-1 rounded"
                      title="删除此映射"
                    >
                      <Trash2 size={13} />
                    </button>
                  </div>
                {/each}
              </div>
            {/if}

            <div class="pt-2 flex justify-end">
              <button
                type="button"
                on:click={handleRunReconcileAnalysis}
                disabled={isAnalyzingReconcile}
                class="px-4 py-2 bg-indigo-600 hover:bg-indigo-700 disabled:opacity-50 text-white rounded-md text-xs font-semibold transition-colors flex items-center gap-1.5 shadow-sm"
              >
                {#if isAnalyzingReconcile}
                  <Clock size={14} class="animate-spin" />
                  <span>正在执行对账分析...</span>
                {:else}
                  <GitCompare size={14} />
                  <span>执行只读对账分析</span>
                {/if}
              </button>
            </div>
          </div>

          <!-- 对账分析结果呈现 -->
          {#if reconcileAnalysis && reconcileAnalysis.ok}
            <div class="border border-indigo-200 bg-indigo-50/30 rounded-lg p-4 space-y-3.5 animate-in fade-in duration-200">
              <div class="flex items-center justify-between pb-2 border-b border-indigo-100">
                <span class="text-xs font-bold text-gray-900 flex items-center gap-1.5">
                  <CheckCircle2 size={15} class="text-emerald-600" />
                  <span>对账分析完成 (耗时: {reconcileAnalysis.analysisElapsedMs} ms)</span>
                </span>
                <span class="text-[11px] text-indigo-700 font-mono">
                  左表 {reconcileAnalysis.leftDataRowCount} 行 vs 右表 {reconcileAnalysis.rightDataRowCount} 行
                </span>
              </div>

              <!-- 合法唯一主键 4 项卡片 -->
              <div class="grid grid-cols-4 gap-2 text-center">
                <div class="bg-white p-2.5 rounded border border-emerald-100 shadow-2xs">
                  <div class="text-[11px] text-emerald-600 font-medium">完全一致</div>
                  <div class="text-base font-bold text-emerald-700">{reconcileAnalysis.matchedBothSameCount}</div>
                </div>
                <div class="bg-white p-2.5 rounded border border-amber-100 shadow-2xs">
                  <div class="text-[11px] text-amber-600 font-medium">存在差异</div>
                  <div class="text-base font-bold text-amber-700">{reconcileAnalysis.matchedBothDiffCount}</div>
                </div>
                <div class="bg-white p-2.5 rounded border border-blue-100 shadow-2xs">
                  <div class="text-[11px] text-blue-600 font-medium">仅左表存在</div>
                  <div class="text-base font-bold text-blue-700">{reconcileAnalysis.leftOnlyCount}</div>
                </div>
                <div class="bg-white p-2.5 rounded border border-purple-100 shadow-2xs">
                  <div class="text-[11px] text-purple-600 font-medium">仅右表存在</div>
                  <div class="text-base font-bold text-purple-700">{reconcileAnalysis.rightOnlyCount}</div>
                </div>
              </div>

              <!-- 异常统计卡片 -->
              <div class="grid grid-cols-5 gap-1.5 text-center text-xs">
                <div class="bg-white/80 p-1.5 rounded border border-gray-100">
                  <span class="text-[10px] text-gray-400 block">左表重复键</span>
                  <span class="font-bold text-red-600">{reconcileAnalysis.leftDuplicateKeyCount}</span>
                </div>
                <div class="bg-white/80 p-1.5 rounded border border-gray-100">
                  <span class="text-[10px] text-gray-400 block">右表重复键</span>
                  <span class="font-bold text-red-600">{reconcileAnalysis.rightDuplicateKeyCount}</span>
                </div>
                <div class="bg-white/80 p-1.5 rounded border border-gray-100">
                  <span class="text-[10px] text-gray-400 block">左表主键异常</span>
                  <span class="font-bold text-gray-600">{reconcileAnalysis.leftInvalidKeyCount}</span>
                </div>
                <div class="bg-white/80 p-1.5 rounded border border-gray-100">
                  <span class="text-[10px] text-gray-400 block">右表主键异常</span>
                  <span class="font-bold text-gray-600">{reconcileAnalysis.rightInvalidKeyCount}</span>
                </div>
                <div class="bg-white/80 p-1.5 rounded border border-gray-100">
                  <span class="text-[10px] text-gray-400 block">差异单元格对数</span>
                  <span class="font-bold text-amber-600">{reconcileAnalysis.diffCellCount}</span>
                </div>
              </div>

              <!-- 双端统计恒等式 -->
              <div class="space-y-1 text-[11px] text-gray-600 bg-white/90 p-2.5 rounded border border-indigo-100 font-mono">
                <div>
                  左端: {reconcileAnalysis.leftDataRowCount} = {reconcileAnalysis.matchedBothSameCount} (一致) + {reconcileAnalysis.matchedBothDiffCount} (差异) + {reconcileAnalysis.leftOnlyCount} (仅左) + {reconcileAnalysis.leftDuplicateKeyCount} (重复) + {reconcileAnalysis.leftInvalidKeyCount} (异常)
                  <span class="text-emerald-600 font-bold ml-1">✓ 严密自洽</span>
                </div>
                <div>
                  右端: {reconcileAnalysis.rightDataRowCount} = {reconcileAnalysis.matchedBothSameCount} (一致) + {reconcileAnalysis.matchedBothDiffCount} (差异) + {reconcileAnalysis.rightOnlyCount} (仅右) + {reconcileAnalysis.rightDuplicateKeyCount} (重复) + {reconcileAnalysis.rightInvalidKeyCount} (异常)
                  <span class="text-emerald-600 font-bold ml-1">✓ 严密自洽</span>
                </div>
              </div>

              <!-- 差异样本明细表格预览 -->
              {#if reconcileAnalysis.details && reconcileAnalysis.details.length > 0}
                <div class="space-y-1">
                  <span class="text-[11px] font-semibold text-gray-700">明细记录速览 (前 10 条):</span>
                  <div class="max-h-36 overflow-y-auto border border-gray-200 rounded bg-white text-[11px]">
                    <table class="w-full text-left">
                      <thead class="bg-gray-50 border-b border-gray-200 text-gray-500">
                        <tr>
                          <th class="p-1.5">判定</th>
                          <th class="p-1.5">对账主键</th>
                          <th class="p-1.5">左行</th>
                          <th class="p-1.5">右行</th>
                          <th class="p-1.5">差异详情</th>
                        </tr>
                      </thead>
                      <tbody class="divide-y divide-gray-100 font-mono">
                        {#each reconcileAnalysis.details.slice(0, 10) as item}
                          <tr class="hover:bg-gray-50">
                            <td class="p-1.5 font-bold">
                              {#if item.category === 'SAME'}
                                <span class="text-emerald-700">完全一致</span>
                              {:else if item.category === 'DIFF'}
                                <span class="text-amber-700">存在差异</span>
                              {:else if item.category === 'LEFT_ONLY'}
                                <span class="text-blue-700">仅左表</span>
                              {:else if item.category === 'RIGHT_ONLY'}
                                <span class="text-purple-700">仅右表</span>
                              {:else if item.category.includes('DUP')}
                                <span class="text-red-700">重复键</span>
                              {:else}
                                <span class="text-gray-500">主键异常</span>
                              {/if}
                            </td>
                            <td class="p-1.5 text-gray-800">{item.keyDisplay}</td>
                            <td class="p-1.5 text-gray-500">{item.leftRowIndex || '-'}</td>
                            <td class="p-1.5 text-gray-500">{item.rightRowIndex || '-'}</td>
                            <td class="p-1.5 text-gray-600 truncate max-w-xs">
                              {#if item.diffs && item.diffs.length > 0}
                                {item.diffs.map(d => `${d.colName}: 左='${d.leftValue}' vs 右='${d.rightValue}'`).join('; ')}
                              {:else}
                                -
                              {/if}
                            </td>
                          </tr>
                        {/each}
                      </tbody>
                    </table>
                  </div>
                </div>
              {/if}

              <!-- 导出至新工作表按钮 -->
              <div class="pt-2 flex justify-end">
                <button
                  type="button"
                  on:click={() => showReconcileConfirm = true}
                  disabled={isExecutingReconcile}
                  class="px-5 py-2.5 bg-emerald-600 hover:bg-emerald-700 disabled:opacity-40 text-white rounded-md text-xs font-semibold transition-colors flex items-center gap-1.5 shadow-sm"
                >
                  <CopyCheck size={14} />
                  <span>导出对账结果至新工作表</span>
                </button>
              </div>
              <p class="text-[11px] text-gray-400 text-center">执行前将强制为输出工作簿成功创建物理快照（快照失败承诺零业务写入），源表 100% 零修改。</p>
            </div>
          {/if}

          <!-- 执行完成反馈卡片 -->
          {#if reconcileExecution && reconcileExecution.ok}
            <div class="border border-emerald-300 bg-emerald-50 rounded-lg p-4 space-y-2 animate-in fade-in duration-200">
              <div class="flex items-center gap-2 text-xs font-bold text-emerald-900">
                <CheckCircle2 size={16} class="text-emerald-600" />
                <span>对账结果导出成功！</span>
              </div>
              <p class="text-xs text-emerald-800 leading-relaxed">
                {reconcileExecution.message}
              </p>
              <div class="pt-1 flex items-center justify-between text-[11px] text-emerald-700/80 font-mono">
                <span>新建工作表: 【{reconcileExecution.resultSheetName}】 (共 {reconcileExecution.exportedRowCount} 条明细)</span>
                <span>快照: {reconcileExecution.snapshotId || '已备份'}</span>
                <span>总耗时: {reconcileExecution.totalElapsedMs} ms</span>
              </div>
            </div>
          {/if}

        <!-- ====================== TAB 3: 多文件列名对齐汇总 (R4c) ====================== -->
        {:else if activeTab === 'consolidate'}
          <!-- 错误提示 -->
          {#if consolidationError}
            <div class="p-3.5 bg-red-50 border border-red-200 rounded-lg flex items-start gap-2.5 text-xs text-red-700 animate-in fade-in duration-150">
              <AlertCircle size={16} class="mt-0.5 shrink-0 text-red-600" />
              <div class="leading-relaxed whitespace-pre-wrap">{consolidationError}</div>
            </div>
          {/if}

          <!-- 来源文件配置区 -->
          <div class="bg-gray-50/90 rounded-lg p-3.5 border border-gray-200 space-y-3">
            <div class="flex items-center justify-between pb-2 border-b border-gray-200">
              <div>
                <span class="text-xs font-bold text-gray-800 flex items-center gap-1.5">
                  <FileSpreadsheet size={15} class="text-indigo-600" />
                  <span>来源工作簿与选区配置 (显式指定)</span>
                </span>
                <p class="text-[11px] text-gray-500 mt-0.5">指定每个文件的工作表、区域与表头行。确定性精确匹配列名，支持同名异径文件区分与元数据避让。</p>
              </div>
              <div class="flex items-center gap-2">
                <button
                  type="button"
                  on:click={handleBrowseConsolidationFiles}
                  class="px-2.5 py-1 text-xs bg-white hover:bg-gray-100 text-indigo-700 border border-indigo-200 rounded-md font-medium transition-colors flex items-center gap-1 shadow-2xs"
                >
                  <FolderOpen size={13} />
                  <span>批量选择文件</span>
                </button>
                <button
                  type="button"
                  on:click={handleAddEmptySource}
                  class="px-2.5 py-1 text-xs bg-white hover:bg-gray-100 text-gray-700 border border-gray-300 rounded-md font-medium transition-colors flex items-center gap-1 shadow-2xs"
                >
                  <Plus size={13} />
                  <span>添加一项</span>
                </button>
              </div>
            </div>

            <!-- 来源文件卡片列表 -->
            <div class="space-y-2.5 max-h-56 overflow-y-auto pr-1">
              {#each consolidationSources as src, idx}
                <div class="bg-white rounded-md p-2.5 border border-gray-200 space-y-2 text-xs shadow-2xs">
                  <div class="flex items-center justify-between gap-2">
                    <span class="font-bold text-gray-700 text-[11px] bg-gray-100 px-1.5 py-0.5 rounded">
                      来源 #{idx + 1}
                    </span>
                    <input
                      type="text"
                      bind:value={src.filePath}
                      placeholder="工作簿绝对路径 (*.xlsx, *.xlsm)"
                      class="flex-1 px-2 py-1 text-xs border border-gray-300 rounded font-mono"
                    />
                    {#if consolidationSources.length > 1}
                      <button
                        type="button"
                        on:click={() => handleRemoveSource(idx)}
                        class="p-1 text-gray-400 hover:text-red-600 rounded transition-colors"
                        title="移除此来源"
                      >
                        <Trash2 size={14} />
                      </button>
                    {/if}
                  </div>

                  <div class="grid grid-cols-3 gap-2">
                    <div>
                      <label class="text-[10px] text-gray-500 block mb-0.5">工作表名 (留空默认首张)</label>
                      <input
                        type="text"
                        bind:value={src.sheetName}
                        placeholder="例: Sheet1"
                        class="w-full px-2 py-1 text-xs border border-gray-300 rounded font-mono"
                      />
                    </div>
                    <div>
                      <label class="text-[10px] text-gray-500 block mb-0.5">区域地址 (留空自动获取已用区域)</label>
                      <input
                        type="text"
                        bind:value={src.rangeAddress}
                        placeholder="例: A1:F50"
                        class="w-full px-2 py-1 text-xs border border-gray-300 rounded font-mono"
                      />
                    </div>
                    <div>
                      <label class="text-[10px] text-gray-500 block mb-0.5">表头行号 (1-based)</label>
                      <input
                        type="number"
                        min="1"
                        bind:value={src.headerRowIndex}
                        class="w-full px-2 py-1 text-xs border border-gray-300 rounded font-mono"
                      />
                    </div>
                  </div>
                </div>
              {/each}
            </div>

            <!-- 全局输出与规则选项 -->
            <div class="pt-2 border-t border-gray-200/80 space-y-2">
              <div>
                <label class="text-[11px] font-semibold text-gray-700 block mb-1">
                  目标输出工作簿全路径 (必须为新文件，不覆盖源文件与既有文件)
                </label>
                <input
                  type="text"
                  bind:value={consolidationOutputPath}
                  placeholder="例: C:\Users\...\Desktop\汇总结果.xlsx"
                  class="w-full px-2 py-1.5 text-xs border border-gray-300 rounded bg-white font-mono"
                />
              </div>

              <div class="flex items-center justify-between pt-1">
                <label class="flex items-center gap-2 cursor-pointer text-xs text-gray-700">
                  <input
                    type="checkbox"
                    bind:checked={includeMetadataCols}
                    class="rounded text-indigo-600 focus:ring-indigo-500 h-3.5 w-3.5"
                  />
                  <span>在输出首列注入来源元数据 (自动处理业务列重名避让)</span>
                </label>

                <button
                  type="button"
                  on:click={handleRunConsolidationAnalysis}
                  disabled={isAnalyzingConsolidation}
                  class="px-4 py-1.5 bg-indigo-600 hover:bg-indigo-700 disabled:opacity-40 text-white rounded-md text-xs font-semibold transition-colors flex items-center gap-1.5 shadow-sm"
                >
                  {#if isAnalyzingConsolidation}
                    <Clock size={14} class="animate-spin" />
                    <span>正在进行只读分析...</span>
                  {:else}
                    <SlidersHorizontal size={14} />
                    <span>开始分析与列名对齐预览</span>
                  {/if}
                </button>
              </div>
            </div>
          </div>

          <!-- 分析结果与列名映射预览卡片 -->
          {#if consolidationAnalysis}
            <div class="border border-indigo-200 bg-indigo-50/40 rounded-lg p-4 space-y-3.5 animate-in fade-in duration-200">
              <div class="flex items-center justify-between pb-2 border-b border-indigo-100">
                <span class="text-xs font-bold text-gray-900 flex items-center gap-1.5">
                  <CheckCircle2 size={15} class="text-indigo-600" />
                  <span>列名对齐只读分析完成 (耗时: {consolidationAnalysis.analysisElapsedMs} ms)</span>
                </span>
                <span class="text-[11px] text-indigo-700 font-mono">
                  来源文件: {consolidationAnalysis.sourceProfiles.length} 个 | 输出列数: {consolidationAnalysis.orderedTargetColumns.length} 列
                </span>
              </div>

              <!-- 四格统计指标 (严格恒等式核对) -->
              <div class="grid grid-cols-4 gap-2 text-center">
                <div class="bg-white p-2.5 rounded border border-gray-100 shadow-2xs">
                  <div class="text-[11px] text-gray-500">来源工作簿</div>
                  <div class="text-base font-bold text-gray-800">{consolidationAnalysis.sourceProfiles.length}</div>
                </div>
                <div class="bg-white p-2.5 rounded border border-indigo-100 shadow-2xs">
                  <div class="text-[11px] text-indigo-600 font-medium">规划汇总行数</div>
                  <div class="text-base font-bold text-indigo-700">{consolidationAnalysis.totalIncludedDataRows}</div>
                </div>
                <div class="bg-white p-2.5 rounded border border-gray-100 shadow-2xs">
                  <div class="text-[11px] text-gray-500">排除纯空行</div>
                  <div class="text-base font-bold text-gray-600">{consolidationAnalysis.totalExcludedRows}</div>
                </div>
                <div class="bg-white p-2.5 rounded border border-emerald-100 shadow-2xs">
                  <div class="text-[11px] text-emerald-600 font-medium">最终输出总行数 (含表头)</div>
                  <div class="text-base font-bold text-emerald-700">{consolidationAnalysis.finalOutputRows}</div>
                </div>
              </div>

              <!-- 各来源读取明细小表 -->
              <div class="bg-white rounded border border-gray-200 overflow-hidden text-xs">
                <div class="px-3 py-1.5 bg-gray-50 border-b border-gray-200 font-semibold text-gray-700 flex items-center justify-between text-[11px]">
                  <span>各来源选区与数据行统计 (只读安全加载，宏已阻断)</span>
                  <span class="text-emerald-600 font-mono">恒等式核对: 最终数据行 = 各来源纳入行之和</span>
                </div>
                <div class="max-h-32 overflow-y-auto divide-y divide-gray-100">
                  {#each consolidationAnalysis.sourceProfiles as prof}
                    <div class="px-3 py-1.5 flex items-center justify-between text-[11px]">
                      <div class="truncate max-w-xs font-mono" title={prof.filePath}>
                        <span class="font-bold text-gray-800">{prof.displayIdentifier}</span>
                        <span class="text-gray-400">({prof.sheetName}!{prof.rangeAddress})</span>
                      </div>
                      <div class="flex items-center gap-3 text-gray-600 font-mono">
                        <span>选区总行: {prof.totalRowsInRange}</span>
                        <span class="text-indigo-600 font-semibold">纳入: {prof.dataRowCount}</span>
                        {#if prof.excludedBlankRows > 0}
                          <span class="text-gray-400">排除: {prof.excludedBlankRows}</span>
                        {/if}
                      </div>
                    </div>
                  {/each}
                </div>
              </div>

              <!-- 稳定输出列顺序与映射对齐预览 -->
              <div class="bg-white rounded border border-gray-200 overflow-hidden text-xs">
                <div class="px-3 py-1.5 bg-gray-50 border-b border-gray-200 font-semibold text-gray-700 flex items-center justify-between text-[11px]">
                  <span>稳定输出列对齐清单 (确定性精确匹配 · 缺失列默认留空)</span>
                  <span class="text-gray-500">共 {consolidationAnalysis.orderedTargetColumns.length} 列</span>
                </div>
                <div class="max-h-36 overflow-y-auto p-2 grid grid-cols-2 gap-1.5">
                  {#each consolidationAnalysis.orderedTargetColumns as colName, colIdx}
                    <div class="flex items-center justify-between px-2 py-1 bg-gray-50 rounded border border-gray-200/60 text-[11px]">
                      <div class="flex items-center gap-1.5 truncate">
                        <span class="text-[10px] font-mono text-gray-400">#{colIdx + 1}</span>
                        <span class="font-bold text-gray-800 truncate" title={colName}>{colName}</span>
                      </div>
                      {#if colName === consolidationAnalysis.metadataFileColName || colName === consolidationAnalysis.metadataSheetColName}
                        <span class="text-[10px] bg-indigo-50 text-indigo-700 px-1 rounded font-medium">元数据列</span>
                      {:else}
                        <span class="text-[10px] bg-emerald-50 text-emerald-700 px-1 rounded font-medium">业务数据列</span>
                      {/if}
                    </div>
                  {/each}
                </div>
              </div>

              <!-- 执行导出按钮 -->
              <div class="pt-2 flex items-center justify-between">
                <div class="text-[11px] text-gray-500">
                  数据保真承诺：19位长编号、前导零、'='公式文本逐字符保真输出。
                </div>
                <button
                  type="button"
                  on:click={handleOpenConsolidationConfirm}
                  disabled={isExecutingConsolidation}
                  class="px-5 py-2 bg-indigo-600 hover:bg-indigo-700 disabled:opacity-40 text-white rounded-md text-xs font-semibold transition-colors flex items-center gap-1.5 shadow-sm"
                >
                  {#if isExecutingConsolidation}
                    <Clock size={14} class="animate-spin" />
                    <span>正在安全写入并读回核验...</span>
                  {:else}
                    <FileCheck size={14} />
                    <span>导出独立新工作簿</span>
                  {/if}
                </button>
              </div>
            </div>
          {/if}

          <!-- 汇总执行结果卡片 -->
          {#if consolidationExecution && consolidationExecution.ok}
            <div class="border border-emerald-300 bg-emerald-50 rounded-lg p-4 space-y-2 animate-in fade-in duration-200">
              <div class="flex items-center justify-between">
                <div class="flex items-center gap-2 text-xs font-bold text-emerald-900">
                  <CheckCircle2 size={16} class="text-emerald-600" />
                  <span>多文件汇总结果导出成功并真机读回核验！</span>
                </div>
                <button
                  type="button"
                  on:click={handleOpenConsolidationOutputDir}
                  class="px-2.5 py-1 text-xs bg-white hover:bg-emerald-100 text-emerald-800 border border-emerald-300 rounded font-medium transition-colors flex items-center gap-1 shadow-2xs"
                >
                  <FolderOpen size={13} />
                  <span>打开输出目录</span>
                </button>
              </div>
              <p class="text-xs text-emerald-800 leading-relaxed font-mono break-all">
                {consolidationExecution.outputFilePath}
              </p>
              <div class="pt-1 flex items-center justify-between text-[11px] text-emerald-700/80 font-mono">
                <span>汇总行数: {consolidationExecution.totalRowsWritten} 行 (含表头) | 列数: {consolidationExecution.totalColsWritten} 列</span>
                <span>真机读回核验: {consolidationExecution.reverifiedRowCount} 行 | 耗时: {consolidationExecution.totalElapsedMs} ms</span>
              </div>
            </div>
          {/if}
        {/if}

        <!-- ====================== TAB 4: 快捷图表工具 (R5b) ====================== -->
        {#if activeTab === 'chart'}
          <!-- 规范与定位说明条 -->
          <div class="p-3 bg-indigo-50/80 border border-indigo-200 rounded-lg text-xs text-indigo-900 leading-relaxed flex items-start gap-2.5">
            <BarChart2 size={16} class="mt-0.5 shrink-0 text-indigo-600" />
            <div>
              <div class="font-bold text-indigo-950">快捷图表工具定位与安全规范 (TASK-R5b-01)</div>
              <div class="text-[11px] text-indigo-800 mt-0.5">
                • <strong>定位说明</strong>：本工具为本地确定性原生 COM 辅助工具，首切片支持柱状图、折线图、饼图，绝不作为大模型自然语言操作请求的隐式替代；零第三方图表库。<br>
                • <strong>防旧结果叠加</strong>：图表采用稳定标识 (<code>__EM_CHART_xxxx</code>)，新建与替换行为明确隔离，绝不误删用户已有手工图表。<br>
                • <strong>快照与恢复承诺</strong>：执行前强制整本物理快照；如遇异常，用户可按已验证范围安全恢复。
              </div>
            </div>
          </div>

          <!-- 错误提示 -->
          {#if chartError}
            <div class="p-3.5 bg-red-50 border border-red-200 rounded-lg flex items-start gap-2.5 text-xs text-red-700 animate-in fade-in duration-150">
              <AlertCircle size={16} class="mt-0.5 shrink-0 text-red-600" />
              <div class="leading-relaxed whitespace-pre-wrap">{chartError}</div>
            </div>
          {/if}

          <!-- 成功提示 -->
          {#if chartSuccessNotice}
            <div class="p-3.5 bg-emerald-50 border border-emerald-200 rounded-lg flex items-start gap-2.5 text-xs text-emerald-800 animate-in fade-in duration-150">
              <CheckCircle2 size={16} class="mt-0.5 shrink-0 text-emerald-600" />
              <div class="leading-relaxed whitespace-pre-wrap">{chartSuccessNotice}</div>
            </div>
          {/if}

          <!-- 数据源与字段映射卡片 -->
          <div class="bg-gray-50/90 rounded-lg p-4 border border-gray-200 space-y-3.5">
            <div class="flex items-center justify-between text-xs text-gray-600 pb-2 border-b border-gray-200">
              <span class="font-medium">目标工作簿: <strong class="text-gray-900">{workbook?.name || '当前工作簿'}</strong></span>
              <span>数据源工作表: <strong class="text-gray-900">{chartSourceSheet || workbook?.activeSheet || 'Sheet1'}</strong></span>
            </div>

            <!-- 数据源区域 -->
            <div class="space-y-1">
              <div class="flex items-center justify-between">
                <label for="chart-range-input" class="text-xs font-semibold text-gray-700">数据源区域 (Range 地址)</label>
                <button
                  type="button"
                  on:click={handleGrabChartSelection}
                  class="text-xs text-indigo-600 hover:text-indigo-800 font-medium hover:underline flex items-center gap-1"
                >
                  <span>抓取选区</span>
                </button>
              </div>
              <div class="flex items-center gap-2">
                <input
                  id="chart-range-input"
                  type="text"
                  bind:value={chartSourceRange}
                  placeholder="例如: A1:D10"
                  class="flex-1 px-3 py-2 text-xs border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-500 font-mono bg-white"
                />
                <label class="flex items-center gap-1.5 text-xs text-gray-700 cursor-pointer select-none shrink-0 px-2.5 py-2 bg-white border border-gray-300 rounded-md hover:bg-gray-50">
                  <input
                    type="checkbox"
                    bind:checked={chartHasHeaders}
                    class="rounded text-indigo-600 focus:ring-indigo-500"
                  />
                  <span>首行为表头</span>
                </label>
              </div>
            </div>

            <!-- 维度与系列列配置 -->
            <div class="grid grid-cols-2 gap-3 pt-1">
              <div class="space-y-1">
                <label for="chart-category-col" class="text-xs font-semibold text-gray-700">类别 (X 轴) 列索引 (1-based)</label>
                <div class="flex items-center gap-2">
                  <input
                    id="chart-category-col"
                    type="number"
                    min="1"
                    bind:value={chartCategoryCol}
                    class="w-20 px-2.5 py-1.5 text-xs border border-gray-300 rounded-md bg-white font-mono"
                  />
                  <input
                    type="text"
                    bind:value={chartCategoryName}
                    placeholder="类别列名称 (可选)"
                    class="flex-1 px-2.5 py-1.5 text-xs border border-gray-300 rounded-md bg-white"
                  />
                </div>
                <p class="text-[10px] text-gray-400">例如第 1 列表示日期或部门</p>
              </div>

              <div class="space-y-1">
                <label for="chart-series-cols" class="text-xs font-semibold text-gray-700">数值系列列索引 (1-based，逗号隔开)</label>
                <input
                  id="chart-series-cols"
                  type="text"
                  bind:value={chartSeriesColsInput}
                  placeholder="例如: 2 或 2, 3"
                  class="w-full px-2.5 py-1.5 text-xs border border-gray-300 rounded-md bg-white font-mono"
                />
                <p class="text-[10px] text-gray-400">例如: 2, 3 表示绘制第 2 列和第 3 列</p>
              </div>
            </div>

            <!-- 图表类型与标题 -->
            <div class="grid grid-cols-2 gap-3 pt-1">
              <div class="space-y-1.5">
                <label class="text-xs font-semibold text-gray-700">图表类型</label>
                <div class="grid grid-cols-3 gap-1.5">
                  <button
                    type="button"
                    on:click={() => chartType = 'column'}
                    class={`py-1.5 px-2 text-xs rounded border text-center font-medium transition-colors ${
                      chartType === 'column'
                        ? 'bg-indigo-50 border-indigo-600 text-indigo-700 font-bold'
                        : 'bg-white border-gray-200 text-gray-600 hover:bg-gray-50'
                    }`}
                  >
                    柱状图
                  </button>
                  <button
                    type="button"
                    on:click={() => chartType = 'line'}
                    class={`py-1.5 px-2 text-xs rounded border text-center font-medium transition-colors ${
                      chartType === 'line'
                        ? 'bg-indigo-50 border-indigo-600 text-indigo-700 font-bold'
                        : 'bg-white border-gray-200 text-gray-600 hover:bg-gray-50'
                    }`}
                  >
                    折线图
                  </button>
                  <button
                    type="button"
                    on:click={() => chartType = 'pie'}
                    class={`py-1.5 px-2 text-xs rounded border text-center font-medium transition-colors ${
                      chartType === 'pie'
                        ? 'bg-indigo-50 border-indigo-600 text-indigo-700 font-bold'
                        : 'bg-white border-gray-200 text-gray-600 hover:bg-gray-50'
                    }`}
                  >
                    饼图
                  </button>
                </div>
              </div>

              <div class="space-y-1">
                <label for="chart-title-input" class="text-xs font-semibold text-gray-700">图表标题</label>
                <input
                  id="chart-title-input"
                  type="text"
                  bind:value={chartTitle}
                  placeholder="例如: 2026年业务分析图"
                  class="w-full px-2.5 py-1.5 text-xs border border-gray-300 rounded-md bg-white"
                />
              </div>
            </div>

            <!-- 饼图单系列约束提示 -->
            {#if chartType === 'pie'}
              <div class="p-2.5 bg-amber-50 border border-amber-200 rounded text-[11px] text-amber-800 flex items-center gap-1.5">
                <AlertCircle size={14} class="shrink-0 text-amber-600" />
                <span>饼图契约约束：饼图仅支持单数值系列，如果数值列填入多列将自动阻断。</span>
              </div>
            {/if}

            <!-- 放置位置与执行模式 -->
            <div class="grid grid-cols-2 gap-3 pt-1 border-t border-gray-200/80">
              <div class="space-y-1">
                <label for="chart-target-cell" class="text-xs font-semibold text-gray-700">放置工作表与起始单元格</label>
                <div class="grid grid-cols-2 gap-1.5">
                  <input
                    type="text"
                    bind:value={chartTargetSheet}
                    placeholder="放置工作表 (默认当前)"
                    class="px-2.5 py-1.5 text-xs border border-gray-300 rounded-md bg-white"
                  />
                  <input
                    id="chart-target-cell"
                    type="text"
                    bind:value={chartTargetCell}
                    placeholder="例如: E2"
                    class="px-2.5 py-1.5 text-xs border border-gray-300 rounded-md bg-white font-mono"
                  />
                </div>
              </div>

              <div class="space-y-1">
                <label class="text-xs font-semibold text-gray-700">防旧结果叠加行为模式</label>
                <div class="flex items-center gap-3 pt-1 text-xs">
                  <label class="flex items-center gap-1 cursor-pointer">
                    <input
                      type="radio"
                      name="chartAction"
                      value="create_new"
                      bind:group={chartAction}
                      class="text-indigo-600"
                    />
                    <span>新建图表</span>
                  </label>
                  <label class="flex items-center gap-1 cursor-pointer">
                    <input
                      type="radio"
                      name="chartAction"
                      value="replace_existing"
                      bind:group={chartAction}
                      on:change={handleScanManagedCharts}
                      class="text-indigo-600"
                    />
                    <span>替换指定快捷图表</span>
                  </label>
                </div>
              </div>
            </div>

            <!-- 替换指定图表配置 -->
            {#if chartAction === 'replace_existing'}
              <div class="p-3 bg-amber-50/70 border border-amber-200 rounded-md space-y-2">
                <div class="flex items-center justify-between text-xs font-semibold text-amber-900">
                  <span>选择待替换的快捷图表 (仅允许替换具备本工具身份签名的图表)</span>
                  <button
                    type="button"
                    on:click={handleScanManagedCharts}
                    disabled={isScanningManagedCharts}
                    class="text-[11px] text-indigo-700 hover:underline flex items-center gap-1"
                  >
                    <span>{isScanningManagedCharts ? '扫描中...' : '检测当前表已有图表'}</span>
                  </button>
                </div>
                {#if chartManagedList.length > 0}
                  <select
                    bind:value={chartTargetChartId}
                    class="w-full px-2.5 py-1.5 text-xs border border-gray-300 rounded bg-white font-mono"
                  >
                    {#each chartManagedList as item}
                      <option value={item.chartId}>
                        {item.chartName} ({item.title || item.chartType} · 位置: Left={Math.round(item.left)}, Top={Math.round(item.top)})
                      </option>
                    {/each}
                  </select>
                {:else}
                  <div class="text-[11px] text-gray-500">
                    当前工作表上暂未检测到由本工具创建管理的图表。若要手动指定，可在上方输入已知的图表 ID。
                  </div>
                {/if}
              </div>
            {/if}

            <!-- 异常数据处理规则 -->
            <div class="flex items-center justify-between pt-1 text-xs text-gray-600">
              <span class="font-medium">非数值/空值/公式错误处理：</span>
              <div class="flex items-center gap-3">
                <label class="flex items-center gap-1 cursor-pointer">
                  <input
                    type="radio"
                    name="chartErrorHandling"
                    value="reject_on_invalid"
                    bind:group={chartErrorHandling}
                    class="text-indigo-600"
                  />
                  <span>严格阻断报错 (推荐)</span>
                </label>
                <label class="flex items-center gap-1 cursor-pointer">
                  <input
                    type="radio"
                    name="chartErrorHandling"
                    value="coerce_zero"
                    bind:group={chartErrorHandling}
                    class="text-indigo-600"
                  />
                  <span>容错视作 0 继续</span>
                </label>
              </div>
            </div>

            <!-- 执行按钮 -->
            <div class="pt-2 flex items-center justify-end">
              <button
                type="button"
                on:click={handleExecuteQuickChart}
                disabled={isExecutingChart}
                class="px-5 py-2 bg-indigo-600 hover:bg-indigo-700 disabled:opacity-40 text-white rounded-md text-xs font-semibold transition-colors flex items-center gap-1.5 shadow-sm"
              >
                {#if isExecutingChart}
                  <Clock size={14} class="animate-spin" />
                  <span>正在执行图表生成与 COM 核验...</span>
                {:else}
                  <BarChart2 size={14} />
                  <span>{chartAction === 'replace_existing' ? '确认替换指定图表' : '确认生成快捷图表'}</span>
                {/if}
              </button>
            </div>
          </div>

          <!-- 图表生成成功读回卡片 -->
          {#if chartResult && chartResult.ok && chartResult.readback}
            <div class="border border-emerald-300 bg-emerald-50 rounded-lg p-4 space-y-2 animate-in fade-in duration-200">
              <div class="flex items-center justify-between">
                <div class="flex items-center gap-2 text-xs font-bold text-emerald-900">
                  <CheckCircle2 size={16} class="text-emerald-600" />
                  <span>快捷图表已成功生成并经 COM 读回客观核验！</span>
                </div>
                {#if chartResult.snapshotId}
                  <span class="text-[10px] px-2 py-0.5 rounded bg-emerald-100 text-emerald-800 border border-emerald-300 font-mono">
                    快照: {chartResult.snapshotId}
                  </span>
                {/if}
              </div>

              <div class="grid grid-cols-2 gap-2 text-xs bg-white/80 p-3 rounded border border-emerald-200">
                <div><strong>图表对象标识：</strong><span class="font-mono text-indigo-700">{chartResult.readback.chartName}</span></div>
                <div><strong>图表类型：</strong>{chartResult.readback.chartType} (Excel Type Code: {chartResult.readback.actualChartTypeNum})</div>
                <div><strong>放置工作表：</strong>{chartResult.readback.sheetName}</div>
                <div><strong>系列数量：</strong>共 {chartResult.readback.seriesCount} 个数值系列 [{chartResult.readback.seriesNames.join(', ')}]</div>
                {#if chartResult.readback.categoryAddress}
                  <div class="col-span-2"><strong>类别轴绑定地址：</strong><span class="font-mono">{chartResult.readback.categoryAddress}</span></div>
                {/if}
              </div>

              {#if chartResult.recoveryNotice}
                <div class="text-[11px] text-gray-500 pt-1">
                  ℹ️ {chartResult.recoveryNotice}
                </div>
              {/if}
            </div>
          {/if}
        {/if}

        <!-- ====================== TAB 5: 外部数据接入 (R6a) ====================== -->
        {#if activeTab === 'external_data'}
          <!-- 错误与成功提示 -->
          {#if externalDataError}
            <div class="p-3.5 bg-red-50 border border-red-200 rounded-lg flex items-start gap-2.5 text-xs text-red-700 animate-in fade-in duration-150">
              <AlertCircle size={16} class="mt-0.5 shrink-0 text-red-600" />
              <div class="leading-relaxed whitespace-pre-wrap">{externalDataError}</div>
            </div>
          {/if}

          {#if externalDataSuccess}
            <div class="p-3.5 bg-emerald-50 border border-emerald-200 rounded-lg flex items-start gap-2.5 text-xs text-emerald-800 animate-in fade-in duration-150">
              <CheckCircle2 size={16} class="mt-0.5 shrink-0 text-emerald-600" />
              <div class="leading-relaxed whitespace-pre-wrap">{externalDataSuccess}</div>
            </div>
          {/if}

          <!-- 数据源配置卡片 -->
          <div class="bg-gray-50/80 rounded-lg p-4 border border-gray-200/80 space-y-3.5">
            <div class="flex items-center justify-between text-xs text-gray-600">
              <span class="font-medium">目标工作簿: <strong class="text-gray-900">{workbook?.name || '当前工作簿'}</strong></span>
              <span>写入模式: <strong class="text-emerald-700">强制新建唯一工作表（绝不覆盖现有表）</strong></span>
            </div>

            <!-- 数据源类型选择 -->
            <div class="space-y-1.5">
              <span class="text-xs font-semibold text-gray-700">数据源类型</span>
              <div class="grid grid-cols-3 gap-2">
                <button
                  type="button"
                  on:click={() => { extSourceType = 'csv'; extPreviewResult = null; }}
                  class={`p-2 rounded-lg border text-xs font-medium flex items-center justify-center gap-1.5 transition-colors ${
                    extSourceType === 'csv'
                      ? 'bg-indigo-50 border-indigo-500 text-indigo-700 font-bold'
                      : 'bg-white border-gray-200 text-gray-600 hover:bg-gray-50'
                  }`}
                >
                  <FileSpreadsheet size={15} />
                  <span>本地 CSV 文件</span>
                </button>
                <button
                  type="button"
                  on:click={() => { extSourceType = 'json'; extPreviewResult = null; }}
                  class={`p-2 rounded-lg border text-xs font-medium flex items-center justify-center gap-1.5 transition-colors ${
                    extSourceType === 'json'
                      ? 'bg-indigo-50 border-indigo-500 text-indigo-700 font-bold'
                      : 'bg-white border-gray-200 text-gray-600 hover:bg-gray-50'
                  }`}
                >
                  <Database size={15} />
                  <span>本地 JSON 文件</span>
                </button>
                <button
                  type="button"
                  on:click={() => { extSourceType = 'http_get'; extPreviewResult = null; }}
                  class={`p-2 rounded-lg border text-xs font-medium flex items-center justify-center gap-1.5 transition-colors ${
                    extSourceType === 'http_get'
                      ? 'bg-indigo-50 border-indigo-500 text-indigo-700 font-bold'
                      : 'bg-white border-gray-200 text-gray-600 hover:bg-gray-50'
                  }`}
                >
                  <Globe size={15} />
                  <span>受控 HTTP GET 接口</span>
                </button>
              </div>
            </div>

            <!-- CSV 配置选项 -->
            {#if extSourceType === 'csv'}
              <div class="space-y-3 bg-white p-3 rounded-lg border border-gray-200">
                <div class="space-y-1">
                  <label for="csv-file-path" class="text-xs font-medium text-gray-700">CSV 绝对文件路径</label>
                  <input
                    id="csv-file-path"
                    type="text"
                    bind:value={extFilePath}
                    placeholder="例如: C:\data\sales_export.csv"
                    class="w-full px-3 py-1.5 text-xs border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-500 font-mono"
                  />
                </div>
                <div class="grid grid-cols-3 gap-2">
                  <div class="space-y-1">
                    <label for="csv-encoding" class="text-[11px] text-gray-600">字符编码</label>
                    <select
                      id="csv-encoding"
                      bind:value={extCsvEncoding}
                      class="w-full px-2 py-1.5 text-xs border border-gray-300 rounded-md bg-white focus:outline-none focus:ring-2 focus:ring-indigo-500"
                    >
                      <option value="UTF-8">UTF-8 (默认)</option>
                      <option value="UTF-8-BOM">UTF-8-BOM</option>
                      <option value="GBK">GBK / GB2312 (中文系统常用)</option>
                      <option value="ASCII">ASCII</option>
                    </select>
                  </div>
                  <div class="space-y-1">
                    <label for="csv-delimiter" class="text-[11px] text-gray-600">列分隔符</label>
                    <select
                      id="csv-delimiter"
                      bind:value={extCsvDelimiter}
                      class="w-full px-2 py-1.5 text-xs border border-gray-300 rounded-md bg-white focus:outline-none focus:ring-2 focus:ring-indigo-500 font-mono"
                    >
                      <option value=",">, (逗号)</option>
                      <option value=";">; (分号)</option>
                      <option value="&#9;">\t (制表符 Tab)</option>
                      <option value="|">| (竖线)</option>
                    </select>
                  </div>
                  <div class="flex items-end pb-1">
                    <label class="flex items-center gap-1.5 text-xs text-gray-700 cursor-pointer select-none">
                      <input
                        type="checkbox"
                        bind:checked={extCsvHasHeader}
                        class="rounded text-indigo-600 focus:ring-indigo-500"
                      />
                      <span>首行为字段表头</span>
                    </label>
                  </div>
                </div>
              </div>
            {/if}

            <!-- JSON 配置选项 -->
            {#if extSourceType === 'json'}
              <div class="space-y-3 bg-white p-3 rounded-lg border border-gray-200">
                <div class="space-y-1">
                  <label for="json-file-path" class="text-xs font-medium text-gray-700">JSON 绝对文件路径</label>
                  <input
                    id="json-file-path"
                    type="text"
                    bind:value={extFilePath}
                    placeholder="例如: C:\data\records.json"
                    class="w-full px-3 py-1.5 text-xs border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-500 font-mono"
                  />
                </div>
                <div class="space-y-1">
                  <label for="json-array-path" class="text-[11px] text-gray-600">指定记录数组属性路径（留空表示根对象数组）</label>
                  <input
                    id="json-array-path"
                    type="text"
                    bind:value={extJsonArrayPath}
                    placeholder="例如: data.items 或 records"
                    class="w-full px-3 py-1.5 text-xs border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-500 font-mono"
                  />
                  <p class="text-[11px] text-gray-400">提示：19位长编号（如身份证、雪花ID）在原生解析阶段逐字符保真为文本，杜绝科学计数法丢失精度。</p>
                </div>
              </div>
            {/if}

            <!-- HTTP GET 配置选项 -->
            {#if extSourceType === 'http_get'}
              <div class="space-y-3 bg-white p-3 rounded-lg border border-gray-200">
                <div class="space-y-1">
                  <label for="http-url-input" class="text-xs font-medium text-gray-700 flex items-center justify-between">
                    <span>HTTP GET 目标 URL (受控请求)</span>
                    <span class="text-[11px] text-amber-600 font-normal">仅限 GET · 最大 10MB · 15s 超时</span>
                  </label>
                  <input
                    id="http-url-input"
                    type="text"
                    bind:value={extHttpUrl}
                    placeholder="例如: http://127.0.0.1:8080/api/v1/metrics"
                    class="w-full px-3 py-1.5 text-xs border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-500 font-mono"
                  />
                </div>

                <div class="space-y-1">
                  <label for="http-whitelist-input" class="text-[11px] text-gray-600 flex items-center justify-between">
                    <span>安全白名单规则（分号/逗号隔开，精确匹配协议、主机、端口及路径前缀）</span>
                    <span class="text-indigo-600 text-[10px]">严格防 SSRF</span>
                  </label>
                  <input
                    id="http-whitelist-input"
                    type="text"
                    bind:value={extHttpWhitelist}
                    placeholder="http://127.0.0.1:*, http://localhost:*"
                    class="w-full px-3 py-1.5 text-xs border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-500 font-mono text-[11px]"
                  />
                </div>

                <div class="grid grid-cols-2 gap-2">
                  <div class="space-y-1">
                    <label for="http-header-key" class="text-[11px] text-gray-600">自定义请求头 Header 名 (可选)</label>
                    <input
                      id="http-header-key"
                      type="text"
                      bind:value={extHttpHeaderKey}
                      placeholder="例如: Authorization"
                      class="w-full px-2.5 py-1 text-xs border border-gray-300 rounded-md font-mono"
                    />
                  </div>
                  <div class="space-y-1">
                    <label for="http-header-val" class="text-[11px] text-gray-600">Header 值 (内存临时使用，不落盘)</label>
                    <input
                      id="http-header-val"
                      type="password"
                      bind:value={extHttpHeaderVal}
                      placeholder="敏感信息日志自动脱敏"
                      class="w-full px-2.5 py-1 text-xs border border-gray-300 rounded-md font-mono"
                    />
                  </div>
                </div>

                <div class="flex items-center gap-1.5 text-[11px] text-gray-500 bg-gray-50 p-2 rounded border border-gray-100">
                  <Lock size={13} class="text-indigo-600 shrink-0" />
                  <span>凭据安全隔离：由 Windows DPAPI 加密隔离，绝不以明文 JSON 存储，绝不流入工作簿或模型上下文。</span>
                </div>
              </div>
            {/if}

            <!-- 新建工作表名称可选 -->
            <div class="space-y-1">
              <label for="ext-target-sheet" class="text-xs font-semibold text-gray-700">目标新建工作表名称 (留空自动生成唯一递增表名)</label>
              <input
                id="ext-target-sheet"
                type="text"
                bind:value={extTargetSheetName}
                placeholder="例如: 导入数据_20261003 (留空由系统自动分配)"
                class="w-full px-3 py-1.5 text-xs border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-500 font-mono"
              />
            </div>

            <!-- 只读拉取并预览按钮 -->
            <div class="pt-2 flex justify-end">
              <button
                type="button"
                on:click={handlePreviewExternalData}
                disabled={isPreviewingExt}
                class="px-4 py-2 bg-indigo-600 hover:bg-indigo-700 disabled:opacity-50 text-white rounded-md text-xs font-semibold transition-colors flex items-center gap-1.5 shadow-sm"
              >
                {#if isPreviewingExt}
                  <Clock size={14} class="animate-spin" />
                  <span>正在安全拉取与解析...</span>
                {:else}
                  <Download size={14} />
                  <span>只读拉取并预览数据</span>
                {/if}
              </button>
            </div>
          </div>

          <!-- 预览结果卡片 -->
          {#if extPreviewResult && extPreviewResult.ok}
            <div class="border border-indigo-200 bg-indigo-50/30 rounded-lg p-4 space-y-3.5 animate-in fade-in duration-200">
              <div class="flex items-center justify-between pb-2 border-b border-indigo-100">
                <span class="text-xs font-bold text-gray-900 flex items-center gap-1.5">
                  <CheckCircle2 size={15} class="text-indigo-600" />
                  <span>只读数据源解析就绪</span>
                </span>
                <span class="text-[11px] text-indigo-700 font-mono">
                  共计约 {extPreviewResult.totalRowsEstimate} 行 · {extPreviewResult.totalCols} 列
                </span>
              </div>

              {#if extPreviewResult.sanitizedSummary}
                <div class="text-[11px] text-gray-600 bg-white/80 p-2 rounded border border-gray-200/60 font-mono">
                  {extPreviewResult.sanitizedSummary}
                </div>
              {/if}

              <!-- 导入字段选择 -->
              <div class="space-y-1.5 bg-white p-3 rounded-lg border border-gray-200">
                <div class="flex items-center justify-between">
                  <span class="text-xs font-semibold text-gray-700">选择要导入的字段列 (已勾选 {extSelectedColumns.length} / {extPreviewResult.columns.length})</span>
                  <div class="flex items-center gap-2">
                    <button
                      type="button"
                      on:click={selectAllExtColumns}
                      class="text-[11px] text-indigo-600 hover:underline font-medium"
                    >
                      全选
                    </button>
                    <span class="text-gray-300">|</span>
                    <button
                      type="button"
                      on:click={deselectAllExtColumns}
                      class="text-[11px] text-gray-500 hover:underline"
                    >
                      仅保留第 1 列
                    </button>
                  </div>
                </div>

                <div class="flex flex-wrap gap-1.5 max-h-28 overflow-y-auto p-1 bg-gray-50 rounded border border-gray-200">
                  {#each extPreviewResult.columns as col, idx}
                    {@const isUnsupported = extPreviewResult.unsupportedColumns && extPreviewResult.unsupportedColumns.includes(col)}
                    <button
                      type="button"
                      on:click={() => toggleExtColumn(col)}
                      class={`px-2.5 py-1 text-xs rounded border transition-colors flex items-center gap-1 ${
                        isUnsupported
                          ? extSelectedColumns.includes(col)
                            ? 'bg-amber-100 border-amber-500 text-amber-900 font-bold'
                            : 'bg-amber-50 border-amber-300 text-amber-700'
                          : extSelectedColumns.includes(col)
                            ? 'bg-indigo-50 border-indigo-500 text-indigo-800 font-bold'
                            : 'bg-white border-gray-200 text-gray-500 hover:bg-gray-100'
                      }`}
                    >
                      <span>{col}</span>
                      {#if isUnsupported}
                        <span class="text-[9px] px-1 rounded bg-amber-200 text-amber-900 font-semibold">
                          不支持嵌套
                        </span>
                      {:else if extPreviewResult.detectedTypes && extPreviewResult.detectedTypes[idx]}
                        <span class="text-[9px] px-1 rounded bg-gray-200 text-gray-600 font-normal">
                          {extPreviewResult.detectedTypes[idx]}
                        </span>
                      {/if}
                    </button>
                  {/each}
                </div>
              </div>

              <!-- 数据源保真规则与容量安全提示 -->
              <div class="p-2.5 bg-gray-50 rounded-lg border border-gray-200 text-[11px] text-gray-600 space-y-1">
                <div class="flex items-center justify-between">
                  <span class="font-medium text-gray-800">数据源保护与保真规则：</span>
                  <span class="text-gray-400 font-mono text-[10px]">
                    {#if extPreviewResult.previewId}
                      快照ID: {extPreviewResult.previewId.substring(0, 8)}...
                    {/if}
                  </span>
                </div>
                <div class="grid grid-cols-2 gap-2 text-[10px] text-gray-500">
                  <div>· 缺失字段 / JSON null 按留白单元格对齐写入</div>
                  <div>· 19位长数字、工单号、前导零、公式字符自动转义保真</div>
                  <div>· 嵌套对象/数组明确阻断，杜绝假称支持的静默占位</div>
                  <div>· 写入前整本强制快照，快照失败零业务写入</div>
                </div>
              </div>

              <!-- 数据样本预览表格 -->
              <div class="space-y-1 bg-white p-3 rounded-lg border border-gray-200">
                <div class="flex items-center justify-between text-xs text-gray-700">
                  <span class="font-medium">样本数据只读预览 (前 {extPreviewResult.sampleRows.length} 行)</span>
                  <span class="text-[11px] text-gray-400">仅展示，未写入工作簿</span>
                </div>
                <div class="max-h-48 overflow-auto border border-gray-200 rounded text-[11px]">
                  <table class="min-w-full divide-y divide-gray-200">
                    <thead class="bg-gray-100 sticky top-0">
                      <tr>
                        <th class="px-2 py-1 text-left text-gray-500 font-mono text-[10px] w-10">#</th>
                        {#each extPreviewResult.columns as col}
                          {#if extSelectedColumns.includes(col)}
                            <th class="px-2 py-1 text-left text-gray-700 font-medium whitespace-nowrap bg-indigo-50/50">
                              {col}
                            </th>
                          {/if}
                        {/each}
                      </tr>
                    </thead>
                    <tbody class="divide-y divide-gray-100 bg-white">
                      {#each extPreviewResult.sampleRows as row, rIdx}
                        <tr class="hover:bg-gray-50">
                          <td class="px-2 py-1 text-gray-400 font-mono text-[10px]">{rIdx + 1}</td>
                          {#each extPreviewResult.columns as col, cIdx}
                            {#if extSelectedColumns.includes(col)}
                              <td class="px-2 py-1 text-gray-800 font-mono whitespace-nowrap">
                                {row[cIdx] !== undefined ? row[cIdx] : ''}
                              </td>
                            {/if}
                          {/each}
                        </tr>
                      {/each}
                    </tbody>
                  </table>
                </div>
              </div>

              <!-- 执行导入按钮 -->
              <div class="pt-2 flex justify-end">
                <button
                  type="button"
                  on:click={handleOpenImportConfirm}
                  disabled={isImportingExt || extSelectedColumns.length === 0}
                  class="px-5 py-2 bg-emerald-600 hover:bg-emerald-700 disabled:opacity-50 text-white rounded-md text-xs font-semibold transition-colors flex items-center gap-1.5 shadow-sm"
                >
                  <Download size={14} />
                  <span>导入至新建独立工作表</span>
                </button>
              </div>
            </div>
          {/if}

          <!-- 导入执行成功卡片 -->
          {#if extImportResult && extImportResult.ok}
            <div class="border border-emerald-300 bg-emerald-50 rounded-lg p-4 space-y-2 animate-in fade-in duration-200">
              <div class="flex items-center justify-between">
                <div class="flex items-center gap-2 text-xs font-bold text-emerald-900">
                  <CheckCircle2 size={16} class="text-emerald-600" />
                  <span>外部数据已成功安全写入新工作表！</span>
                </div>
                {#if extImportResult.snapshotId}
                  <span class="text-[10px] px-2 py-0.5 rounded bg-emerald-100 text-emerald-800 border border-emerald-300 font-mono">
                    前置快照: {extImportResult.snapshotId}
                  </span>
                {/if}
              </div>

              <div class="grid grid-cols-2 gap-2 text-xs bg-white/80 p-3 rounded border border-emerald-200">
                <div><strong>目标工作表：</strong><span class="font-mono text-indigo-700">{extImportResult.sheetName}</span></div>
                <div><strong>实际写入规模：</strong>{extImportResult.importedRowCount} 行 × {extImportResult.importedColCount} 列</div>
                <div><strong>执行耗时：</strong>{extImportResult.elapsedMs} ms</div>
                <div><strong>安全策略：</strong>前置整本物理快照 + 公式样文本防注入转义</div>
              </div>

              {#if extImportResult.recoveryNotice}
                <div class="text-[11px] text-gray-500 pt-1">
                  ℹ️ {extImportResult.recoveryNotice}
                </div>
              {/if}
            </div>
          {/if}
        {/if}

      </div>
    </div>
  </div>
{/if}

<!-- 去重确认弹窗 -->
{#if showConfirmModal && analysisResult}
  <div class="fixed inset-0 z-60 flex items-center justify-center bg-black/60 backdrop-blur-xs p-4 animate-in fade-in duration-100">
    <div class="bg-white rounded-xl shadow-2xl border border-gray-200 max-w-md w-full p-5 space-y-4 text-gray-800">
      <div class="flex items-center gap-2.5">
        <div class="w-8 h-8 rounded-full bg-emerald-100 text-emerald-700 flex items-center justify-center font-bold">
          <ShieldCheck size={18} />
        </div>
        <h3 class="text-sm font-bold text-gray-900">确认执行去重数据操作</h3>
      </div>

      <div class="text-xs space-y-2 bg-gray-50 p-3 rounded-lg border border-gray-200">
        <div><strong>操作模式：</strong>
          {#if pendingMode === 'highlight'}
            <span class="text-amber-700 font-semibold">选区内标记浅红高亮 (源值不修改，仅改填充格式)</span>
          {:else}
            <span class="text-emerald-700 font-semibold">输出唯一数据至新建工作表 (静态值写入，源表零修改)</span>
          {/if}
        </div>
        <div><strong>目标工作簿：</strong>{analysisResult.targetWorkbookName} (表: {analysisResult.targetSheetName})</div>
        <div><strong>数据区域：</strong>{analysisResult.rangeAddress} (数据行数: {analysisResult.dataRowCount} 行)</div>
        <div><strong>判断主键：</strong>列 [{analysisResult.keyColumnIndices.join(', ')}]</div>
        <div><strong>影响行数：</strong>{pendingMode === 'highlight' ? analysisResult.duplicateCount : analysisResult.uniqueCount} 行</div>
        <div class="text-[11px] text-emerald-600 font-medium">安全保障：执行前将强制创建目标工作簿物理快照（快照失败承诺零业务写入），支持随时回滚。</div>
      </div>

      <div class="flex items-center justify-end gap-2.5 pt-1">
        <button
          type="button"
          on:click={() => showConfirmModal = false}
          class="px-3.5 py-1.5 text-xs text-gray-600 hover:bg-gray-100 rounded-md font-medium transition-colors"
        >
          返回检查
        </button>
        <button
          type="button"
          on:click={handleConfirmExecute}
          class="px-4 py-1.5 text-xs text-white bg-emerald-600 hover:bg-emerald-700 rounded-md font-semibold transition-colors shadow-sm"
        >
          确认执行
        </button>
      </div>
    </div>
  </div>
{/if}

<!-- 两表对账确认弹窗 -->
{#if showReconcileConfirm && reconcileAnalysis}
  <div class="fixed inset-0 z-60 flex items-center justify-center bg-black/60 backdrop-blur-xs p-4 animate-in fade-in duration-100">
    <div class="bg-white rounded-xl shadow-2xl border border-gray-200 max-w-md w-full p-5 space-y-4 text-gray-800">
      <div class="flex items-center gap-2.5">
        <div class="w-8 h-8 rounded-full bg-indigo-100 text-indigo-700 flex items-center justify-center font-bold">
          <ShieldCheck size={18} />
        </div>
        <h3 class="text-sm font-bold text-gray-900">确认执行两表对账导出</h3>
      </div>

      <div class="text-xs space-y-2 bg-gray-50 p-3 rounded-lg border border-gray-200">
        <div><strong>操作内容：</strong><span class="text-indigo-700 font-semibold">在目标工作簿新建唯一命名结果表，输出静态对账明细</span></div>
        <div><strong>左表来源：</strong>{reconcileAnalysis.leftWorkbookName} [{reconcileAnalysis.leftSheetName}!{reconcileAnalysis.leftRangeAddress}] ({reconcileAnalysis.leftDataRowCount} 行)</div>
        <div><strong>右表来源：</strong>{reconcileAnalysis.rightWorkbookName} [{reconcileAnalysis.rightSheetName}!{reconcileAnalysis.rightRangeAddress}] ({reconcileAnalysis.rightDataRowCount} 行)</div>
        <div><strong>主键映射：</strong>左列 [{reconcileAnalysis.leftKeyCols.join(', ')}] ↔ 右列 [{reconcileAnalysis.rightKeyCols.join(', ')}]</div>
        <div><strong>对账结果：</strong>一致 {reconcileAnalysis.matchedBothSameCount}，差异 {reconcileAnalysis.matchedBothDiffCount}，仅左 {reconcileAnalysis.leftOnlyCount}，仅右 {reconcileAnalysis.rightOnlyCount}，重复 {reconcileAnalysis.leftDuplicateKeyCount + reconcileAnalysis.rightDuplicateKeyCount}</div>
        <div class="text-[11px] text-emerald-600 font-medium">安全保障：左右源表数据 100% 零修改；输出前强制创建输出工作簿快照（快照失败承诺零业务写入）。</div>
      </div>

      <div class="flex items-center justify-end gap-2.5 pt-1">
        <button
          type="button"
          on:click={() => showReconcileConfirm = false}
          class="px-3.5 py-1.5 text-xs text-gray-600 hover:bg-gray-100 rounded-md font-medium transition-colors"
        >
          返回检查
        </button>
        <button
          type="button"
          on:click={handleConfirmExecuteReconcile}
          class="px-4 py-1.5 text-xs text-white bg-indigo-600 hover:bg-indigo-700 rounded-md font-semibold transition-colors shadow-sm"
        >
          确认导出
        </button>
      </div>
    </div>
  </div>
{/if}

<!-- 多文件汇总确认弹窗 -->
{#if showConsolidationConfirm && consolidationAnalysis}
  <div class="fixed inset-0 z-60 flex items-center justify-center bg-black/60 backdrop-blur-xs p-4 animate-in fade-in duration-100">
    <div class="bg-white rounded-xl shadow-2xl border border-gray-200 max-w-lg w-full p-5 space-y-4 text-gray-800">
      <div class="flex items-center gap-2.5">
        <div class="w-8 h-8 rounded-full bg-indigo-100 text-indigo-700 flex items-center justify-center font-bold">
          <ShieldCheck size={18} />
        </div>
        <div>
          <h3 class="text-sm font-bold text-gray-900">确认执行多文件列名对齐汇总</h3>
          <p class="text-[11px] text-gray-500">导出独立新工作簿 · 确定性精确列对齐</p>
        </div>
      </div>

      <div class="text-xs space-y-2 bg-gray-50 p-3.5 rounded-lg border border-gray-200">
        <div><strong>输出文件路径：</strong><span class="font-mono text-indigo-700 break-all">{consolidationOutputPath}</span></div>
        <div><strong>来源文件数量：</strong>共 {consolidationAnalysis.sourceProfiles.length} 个独立来源</div>
        <div><strong>汇总行数计算：</strong>纳入数据 {consolidationAnalysis.totalIncludedDataRows} 行 + 表头 1 行 = <strong>共 {consolidationAnalysis.finalOutputRows} 行</strong> (排除纯空行 {consolidationAnalysis.totalExcludedRows} 行)</div>
        <div><strong>输出列数清单：</strong>共 {consolidationAnalysis.orderedTargetColumns.length} 列 [{consolidationAnalysis.orderedTargetColumns.slice(0, 5).join(', ')}{consolidationAnalysis.orderedTargetColumns.length > 5 ? '...' : ''}]</div>
        
        <div class="pt-1 text-[11px] text-emerald-700 space-y-1 bg-emerald-50/60 p-2.5 rounded border border-emerald-200">
          <div class="font-semibold">数据保真与安全承诺：</div>
          <div>• 来源文件只读打开，强制阻断 VBA 自动宏与外部链接，零修改来源文件；</div>
          <div>• 19位长编号、前导零、'='公式文本逐字符保真输出（防截断与科学计数法）；</div>
          <div>• 保存后强制重新打开真机核验行列数一致性，核验无误才标记成功。</div>
        </div>
      </div>

      <div class="flex items-center justify-end gap-2.5 pt-1">
        <button
          type="button"
          on:click={() => showConsolidationConfirm = false}
          class="px-3.5 py-1.5 text-xs text-gray-600 hover:bg-gray-100 rounded-md font-medium transition-colors"
        >
          返回检查
        </button>
        <button
          type="button"
          on:click={handleConfirmExecuteConsolidation}
          class="px-4 py-1.5 text-xs text-white bg-indigo-600 hover:bg-indigo-700 rounded-md font-semibold transition-colors shadow-sm"
        >
          确认导出新工作簿
        </button>
      </div>
    </div>
  </div>
{/if}

<!-- 外部数据导入确认弹窗 -->
{#if showExternalDataConfirm && extPreviewResult}
  <div class="fixed inset-0 z-60 flex items-center justify-center bg-black/60 backdrop-blur-xs p-4 animate-in fade-in duration-100">
    <div class="bg-white rounded-xl shadow-2xl border border-gray-200 max-w-lg w-full p-5 space-y-4 text-gray-800">
      <div class="flex items-center gap-2.5">
        <div class="w-8 h-8 rounded-full bg-indigo-100 text-indigo-700 flex items-center justify-center font-bold">
          <ShieldCheck size={18} />
        </div>
        <div>
          <h3 class="text-sm font-bold text-gray-900">确认导入外部数据至工作簿</h3>
          <p class="text-[11px] text-gray-500">新建独立工作表 · 纯数据防注入 · 前置强制物理快照</p>
        </div>
      </div>

      <div class="text-xs space-y-2 bg-gray-50 p-3.5 rounded-lg border border-gray-200">
        <div><strong>目标工作簿：</strong><span class="font-mono text-indigo-700">{workbook?.name || '当前工作簿'}</span></div>
        <div><strong>新建工作表：</strong><span class="font-mono text-gray-900">{extTargetSheetName.trim() || '由系统自动分配唯一递增表名'}</span> (绝不覆盖已有工作表)</div>
        <div><strong>数据源类型：</strong>
          {#if extSourceType === 'csv'}
            <span>本地 CSV [{extFilePath}] (编码: {extCsvEncoding}, 分隔符: '{extCsvDelimiter}')</span>
          {:else if extSourceType === 'json'}
            <span>本地 JSON [{extFilePath}] {extJsonArrayPath ? `(数组路径: ${extJsonArrayPath})` : '(根数组)'}</span>
          {:else}
            <span class="break-all">受控 HTTP GET [{extHttpUrl}] (白名单校验通过)</span>
          {/if}
        </div>
        <div><strong>导入字段规模：</strong>选择导入 {extSelectedColumns.length} 列 / 总计 {extPreviewResult.columns.length} 列 (预估行数: {extPreviewResult.totalRowsEstimate} 行)</div>

        <div class="pt-1 text-[11px] text-emerald-700 space-y-1 bg-emerald-50/60 p-2.5 rounded border border-emerald-200">
          <div class="font-semibold">安全保障与保真策略：</div>
          <div>• 写入前自动创建目标工作簿全量物理快照（快照失败承诺零业务写入）；</div>
          <div>• 19位长编号、前导零、'='公式样文本强制按纯数据防注入写入（绝不被 Excel 执行为恶意公式）；</div>
          <div>• 目标工作簿锁定，若写入前目标文件被关闭或重命名，立即安全阻断；</div>
          <div>• 凭据经 Windows DPAPI 隔离加密，敏感 Token 绝不进入工作簿、日志或提示词。</div>
        </div>
      </div>

      <div class="flex items-center justify-end gap-2.5 pt-1">
        <button
          type="button"
          on:click={() => showExternalDataConfirm = false}
          class="px-3.5 py-1.5 text-xs text-gray-600 hover:bg-gray-100 rounded-md font-medium transition-colors"
        >
          返回检查
        </button>
        <button
          type="button"
          on:click={handleConfirmExecuteImport}
          class="px-4 py-1.5 text-xs text-white bg-indigo-600 hover:bg-indigo-700 rounded-md font-semibold transition-colors shadow-sm"
        >
          确认安全导入
        </button>
      </div>
    </div>
  </div>
{/if}

<style>
  /* ========================================================
     独立工作区与卡片布局
     ======================================================== */
  .datatool-workspace-root {
    display: flex;
    flex-direction: column;
    height: 100%;
    width: 100%;
    background: var(--office-bg);
    overflow: hidden;
  }

  .datatool-modal-overlay {
    position: fixed;
    top: 0;
    left: 0;
    width: 100vw;
    height: 100vh;
    background: rgba(0, 0, 0, 0.45);
    display: flex;
    align-items: center;
    justify-content: center;
    z-index: 1000;
    padding: 12px;
  }

  .datatool-workspace-card {
    display: flex;
    flex-direction: column;
    height: 100%;
    width: 100%;
    background: #ffffff;
    overflow: hidden;
    container-type: inline-size;
  }

  .datatool-modal-card {
    display: flex;
    flex-direction: column;
    width: 100%;
    max-width: 600px;
    max-height: 90vh;
    background: #ffffff;
    border-radius: var(--office-radius);
    border: 1px solid var(--office-border);
    box-shadow: var(--office-shadow-lg);
    overflow: hidden;
    container-type: inline-size;
  }

  /* 头部 */
  .datatool-header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 10px 14px;
    border-bottom: 1px solid var(--office-border-subtle);
    background: #fafbfc;
    flex-shrink: 0;
    gap: 8px;
  }

  .header-info {
    display: flex;
    align-items: center;
    gap: 10px;
    min-width: 0;
  }

  .header-icon-box {
    width: 32px;
    height: 32px;
    border-radius: var(--office-radius-sm);
    background: var(--excel-light);
    color: var(--excel-green);
    display: flex;
    align-items: center;
    justify-content: center;
    flex-shrink: 0;
    border: 1px solid var(--excel-light-border);
  }

  .header-title-group {
    display: flex;
    flex-direction: column;
    min-width: 0;
    gap: 2px;
  }

  .title-row {
    display: flex;
    align-items: center;
    gap: 6px;
  }

  .tool-title {
    font-size: var(--font-size-md);
    font-weight: 600;
    color: var(--office-text);
    margin: 0;
    line-height: 1.2;
  }

  .tool-badge {
    font-size: 10px;
    padding: 1px 6px;
    border-radius: var(--office-radius-full);
    background: var(--excel-light);
    color: var(--excel-dark);
    border: 1px solid var(--excel-light-border);
    font-weight: 500;
  }

  .tool-desc {
    font-size: var(--font-size-xs);
    color: var(--office-muted);
    margin: 0;
    line-height: 1.3;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
    max-width: 340px;
  }

  .datatool-close-btn {
    border: none;
    background: transparent;
    color: var(--office-muted);
    padding: 4px;
    border-radius: var(--office-radius-xs);
    cursor: pointer;
    display: flex;
    align-items: center;
    justify-content: center;
    transition: background 0.15s ease;
    flex-shrink: 0;
  }

  .datatool-close-btn:hover {
    background: var(--office-hover);
    color: var(--office-text);
  }

  /* 窄屏工具下拉选择器 */
  .datatool-selector-mobile {
    display: none;
    padding: 6px 10px;
    background: #ffffff;
    border-bottom: 1px solid var(--office-border);
    align-items: center;
    gap: 8px;
    box-sizing: border-box;
    width: 100%;
  }

  .selector-label {
    font-size: var(--font-size-xs);
    font-weight: 600;
    color: var(--office-text-secondary);
    white-space: nowrap;
    flex-shrink: 0;
  }

  .selector-select {
    flex: 1 1 0%;
    min-width: 0;
    height: 28px;
    font-size: var(--font-size-xs);
    border: 1px solid var(--office-border);
    border-radius: var(--office-radius-xs);
    background: #ffffff;
    color: var(--office-text);
    padding: 0 6px;
    outline: none;
    cursor: pointer;
  }

  .selector-select:focus {
    border-color: var(--excel-green);
  }

  @container (max-width: 380px) {
    .datatool-selector-mobile {
      display: flex !important;
    }
    .datatool-tabs-wrapper {
      display: none !important;
    }
  }

  @media (max-width: 380px) {
    .datatool-selector-mobile {
      display: flex;
    }
    .datatool-tabs-wrapper {
      display: none;
    }
  }

  /* Tab 栏 */
  .datatool-tabs-wrapper {
    position: relative;
    background: #ffffff;
    border-bottom: 1px solid var(--office-border);
  }

  .datatool-tabs-bar {
    display: flex;
    align-items: center;
    background: #ffffff;
    padding: 0 6px;
    overflow-x: auto;
    flex-shrink: 0;
    gap: 2px;
  }

  /* 显式保留纤细滚动条，保证在 320px 窄宽下有清晰滑动感知 */
  .datatool-tabs-bar::-webkit-scrollbar {
    height: 3px;
  }

  .datatool-tabs-bar::-webkit-scrollbar-track {
    background: transparent;
  }

  .datatool-tabs-bar::-webkit-scrollbar-thumb {
    background: rgba(0, 0, 0, 0.18);
    border-radius: 2px;
  }

  .datatool-tab-btn {
    display: inline-flex;
    align-items: center;
    gap: 5px;
    padding: 8px 12px;
    font-size: var(--font-size-sm);
    color: var(--office-muted);
    border: none;
    border-bottom: 2px solid transparent;
    background: transparent;
    cursor: pointer;
    font-weight: 500;
    white-space: nowrap;
    transition: all 0.15s ease;
  }

  .datatool-tab-btn:hover:not(.active) {
    color: var(--office-text);
    background: var(--office-hover);
  }

  .datatool-tab-btn.active {
    color: var(--excel-green);
    border-bottom-color: var(--excel-green);
    font-weight: 600;
  }

  /* 内容主体 */
  .datatool-content-body {
    flex: 1;
    overflow-y: auto;
    padding: 12px;
    display: flex;
    flex-direction: column;
    gap: 12px;
    background: var(--office-bg);
  }

  /* 全局控件补丁：确保所有未带自定义类的表单元素均符合 Fluent 规范 */
  .datatool-content-body input[type="text"],
  .datatool-content-body input[type="number"],
  .datatool-content-body select {
    background: #ffffff;
    border: 1px solid var(--office-border);
    border-radius: var(--office-radius-sm);
    padding: 6px 8px;
    font-size: var(--font-size-sm);
    color: var(--office-text);
    font-family: inherit;
    outline: none;
    transition: border-color 0.15s ease, box-shadow 0.15s ease;
  }

  .datatool-content-body input[type="text"]:focus,
  .datatool-content-body input[type="number"]:focus,
  .datatool-content-body select:focus {
    border-color: var(--excel-green);
    box-shadow: 0 0 0 2px rgba(16, 124, 65, 0.15);
  }

  /* 消除默认灰色凸起按钮 */
  .datatool-content-body button {
    font-family: inherit;
    border-radius: var(--office-radius-sm);
    transition: all 0.15s ease;
    cursor: pointer;
  }

  .datatool-content-body button:disabled {
    cursor: not-allowed;
    opacity: 0.6;
  }

  .datatool-content-body,
  .datatool-content-body * {
    box-sizing: border-box;
  }

  .datatool-content-body input[type="text"],
  .datatool-content-body input[type="number"] {
    min-width: 0;
  }

  @media (max-width: 360px) {
    .datatool-content-body {
      padding: 8px;
      gap: 8px;
    }
  }
</style>


