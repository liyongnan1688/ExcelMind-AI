<script lang="ts">
  import {
    X,
    Play,
    Trash2,
    Code2,
    ChevronDown,
    ChevronUp,
    FileCode,
    Upload,
    FileText,
    Copy,
    Edit3,
    Check,
    AlertCircle,
    AlertTriangle,
    Search,
    Download,
    Tag,
    Clock,
    Hash,
    PlusCircle,
    Layers,
    Star,
    MessageSquare,
    Package,
    ShieldAlert,
  } from 'lucide-svelte';
  import {
    bridge,
    type ScriptItem,
    type WorkbookInfo,
    type VbaEntryPointInfo,
    type VbaParamSignature,
    type ScriptParameterDef,
    type SensitiveScanWarning,
    type MacroPackageEntry,
    type MacroPackageManifest,
    type MacroPackageExportResult,
    type MacroPackagePreviewResult,
    type MacroPackageImportResult,
  } from '../services/bridge';

  export let isOpen = false;
  export let workbook: WorkbookInfo | null = null;
  export let onClose: () => void;
  export let onRunScript: (script: ScriptItem, entryPoint?: string, parameters?: Record<string, any>) => void;
  export let onCancelWithQuestion: ((data: { script: ScriptItem; entryPoint: string; questionText: string }) => void) | undefined = undefined;

  // 主视图状态
  let activeTab: 'list' | 'import' = 'list';

  // 宏列表状态
  let scripts: ScriptItem[] = [];
  let isLoading = false;
  let expandedIndex: number | null = null;
  let searchTerm = '';
  let selectedCategory = '全部';
  let copiedScriptId: string | null = null;

  // 运行前入口与目标确认弹窗状态 (TASK-R2b & TASK-R2c)
  let runModalScript: ScriptItem | null = null;
  let runModalSelectedEntryPoint = '';
  let runModalEntryPoints: VbaEntryPointInfo[] = [];
  let runModalShowCode = false;
  let runModalTargetName = '';
  let runModalTargetFullName = '';
  let runModalTargetError = '';
  let runModalTargetSheets: string[] = [];
  let runModalActiveSheet = '';
  let runModalParamValues: Record<string, any> = {};
  let runModalParamErrors: Record<string, string> = {};
  let runModalMetaConflictError = '';
  // 维护参数草稿缓存：scriptKey_entryPoint -> { paramName: val }，避免取消时清空已填内容
  let paramDrafts: Record<string, Record<string, any>> = {};
  let showFavoritesOnly = false;

  // 重命名弹窗状态
  let renameModalScript: ScriptItem | null = null;
  let renameNewDisplayName = '';
  let renameError = '';

  // 导入宏表单状态
  let importSourceType: 'file' | 'paste' | 'package' = 'file';
  let importFileName = '';
  let importRawBuffer: ArrayBuffer | null = null;
  let importRawBytesBase64 = '';
  let importEncoding = 'UTF-8';
  let importCode = '';
  let importDisplayName = '';
  let importDescription = '';
  let importCategory = '';
  let importEntryPoint = '';
  let importCodeHash = '';
  let importErrorMessage = '';
  let importSuccessMessage = '';
  let isSaving = false;

  // 宏包导出弹窗与表单状态 (TASK-R6b-01)
  let showExportPackageModal = false;
  let exportPackageSelectedMacroIds: string[] = [];
  let exportPackageName = 'ExcelMind_Macros';
  let exportPackageVersion = '1.0.0';
  let exportPackageDescription = '';
  let exportPackageOutputPath = '';
  let exportPackageIsScanning = false;
  let exportPackageWarnings: SensitiveScanWarning[] = [];
  let exportPackageErrorMessage = '';
  let exportPackageSuccessResult: MacroPackageExportResult | null = null;

  // 宏包导入视图状态 (TASK-R6b-01)
  let packageFilePath = '';
  let packagePreviewLoading = false;
  let packagePreviewResult: MacroPackagePreviewResult | null = null;
  let packageImportSelectedMacroIds: string[] = [];
  let packageImportLoading = false;
  let packageImportResult: MacroPackageImportResult | null = null;
  let packageImportError = '';

  // 覆盖确认弹窗
  let showOverwriteConfirm = false;
  let pendingOverwritePayload: any = null;

  export function openTab(tab: 'list' | 'import') {
    activeTab = tab;
  }

  export async function promptRunMacro(scriptId: string) {
    if (!scriptId) return;
    activeTab = 'list';
    if (scripts.length === 0) {
      await refreshScripts();
    }
    const targetId = scriptId.trim().toLowerCase();
    let found = scripts.find(
      (s) =>
        (s.id && s.id.toLowerCase() === targetId) ||
        (s.fileName && s.fileName.toLowerCase() === targetId) ||
        (s.fileName && s.fileName.replace(/\.bas$/i, '').toLowerCase() === targetId)
    );
    if (!found) {
      await refreshScripts();
      found = scripts.find(
        (s) =>
          (s.id && s.id.toLowerCase() === targetId) ||
          (s.fileName && s.fileName.toLowerCase() === targetId) ||
          (s.fileName && s.fileName.replace(/\.bas$/i, '').toLowerCase() === targetId)
      );
    }

    if (found) {
      handleOpenRunModal(found);
    } else {
      alert(`未在本地宏库中找到标识为【${scriptId}】的宏，该宏可能已被重命名或物理移除。`);
    }
  }

  export async function refreshScripts() {
    isLoading = true;
    const res = await bridge.send<ScriptItem[]>('list_scripts');
    isLoading = false;
    if (res.ok && res.data) {
      scripts = res.data;
    }
  }

  $: if (isOpen) {
    refreshScripts();
  }

  // 提取分类列表
  $: categories = ['全部', ...Array.from(new Set(scripts.map((s) => s.category || '未分类').filter(Boolean)))];

  // 提取所有唯一标签集合
  $: allTags = Array.from(
    new Set(
      scripts
        .flatMap((s) => s.tags || [])
        .map((t) => t.trim())
        .filter(Boolean)
    )
  );

  let selectedTag = '全部';

  // 标签编辑状态
  let newTagInput: Record<string, string> = {};
  let tagOperationError: string = '';

  async function handleAddTag(script: ScriptItem) {
    tagOperationError = '';
    const sid = script.id || script.fileName;
    const tagText = (newTagInput[sid] || '').trim();
    if (!tagText) return;

    const currentTags = script.tags || [];
    if (currentTags.includes(tagText)) {
      newTagInput[sid] = '';
      return;
    }

    const updatedTags = [...currentTags, tagText];
    const res = await bridge.updateScriptTags(sid, updatedTags);
    if (res.ok) {
      script.tags = updatedTags;
      scripts = [...scripts];
      newTagInput[sid] = '';
    } else {
      tagOperationError = res.error || '添加标签失败';
    }
  }

  async function handleRemoveTag(script: ScriptItem, tagToRemove: string) {
    tagOperationError = '';
    const sid = script.id || script.fileName;
    const currentTags = script.tags || [];
    const updatedTags = currentTags.filter((t) => t !== tagToRemove);

    const res = await bridge.updateScriptTags(sid, updatedTags);
    if (res.ok) {
      script.tags = updatedTags;
      scripts = [...scripts];
      if (selectedTag === tagToRemove && !allTags.includes(tagToRemove)) {
        selectedTag = '全部';
      }
    } else {
      tagOperationError = res.error || '删除标签失败';
    }
  }

  async function handleToggleFavorite(script: ScriptItem) {
    const sid = script.id || script.fileName;
    const nextVal = !script.isFavorite;
    const res = await bridge.updateScriptFavorite(sid, nextVal);
    if (res.ok) {
      script.isFavorite = nextVal;
      scripts = [...scripts];
    } else {
      alert('更新收藏状态失败: ' + (res.error || '未知错误'));
    }
  }

  // 过滤后的宏列表（支持名称、描述、分类、源码、标签、收藏的多维度组合检索）
  $: filteredScripts = scripts.filter((s) => {
    if (showFavoritesOnly && !s.isFavorite) {
      return false;
    }

    const sTags = s.tags || [];
    const term = searchTerm.trim().toLowerCase();

    // 组合检索：名称、描述、分类、源码、标签均可被关键词检索
    const matchSearch =
      !term ||
      (s.displayName || s.name || '').toLowerCase().includes(term) ||
      (s.description || '').toLowerCase().includes(term) ||
      (s.category || '').toLowerCase().includes(term) ||
      (s.code || '').toLowerCase().includes(term) ||
      sTags.some((tag) => tag.toLowerCase().includes(term));

    // 分类筛选
    const matchCategory =
      selectedCategory === '全部' ||
      (s.category || '未分类') === selectedCategory;

    // 标签筛选
    const matchTag =
      selectedTag === '全部' ||
      sTags.includes(selectedTag);

    return matchSearch && matchCategory && matchTag;
  });

  // 实时分析导入源码中的过程与入口
  interface VbaProcedure {
    name: string;
    type: 'Sub' | 'Function';
    modifier: string;
    params: string;
    isRunnable: boolean;
    paramType: 'none' | 'workbook' | 'other';
    reason: string;
  }

  function parseVbaProcedures(code: string): VbaProcedure[] {
    const list: VbaProcedure[] = [];
    if (!code) return list;

    const regex = /(?:^|\r?\n)\s*(?:(Public|Private|Friend)\s+)?(Sub|Function)\s+([a-zA-Z0-9_\u4e00-\u9fa5]+)\s*(?:\(([^)]*)\))?/gi;
    let m: RegExpExecArray | null;
    while ((m = regex.exec(code)) !== null) {
      const modifier = (m[1] || 'Public').trim();
      const type = (m[2].toLowerCase() === 'sub' ? 'Sub' : 'Function') as 'Sub' | 'Function';
      const name = m[3];
      const params = (m[4] || '').trim();

      let isRunnable = false;
      let paramType: 'none' | 'workbook' | 'other' = 'none';
      let reason = '';

      if (modifier.toLowerCase() === 'private') {
        isRunnable = false;
        paramType = 'other';
        reason = 'Private 私有过程无法由外部作为宏独立调用';
      } else if (type === 'Function') {
        isRunnable = false;
        paramType = 'other';
        reason = 'Function 函数用于返回值，不能作为独立宏入口';
      } else if (!params || /^\s*'.*$/.test(params)) {
        isRunnable = true;
        paramType = 'none';
        reason = '无参公开 Sub，可直接调用';
      } else if (/^(?:targetWb|wb|workbook)\s+As\s+(?:Workbook|Object)$/i.test(params)) {
        isRunnable = true;
        paramType = 'workbook';
        reason = '接收目标工作簿参数，宿主原生支持直调';
      } else {
        isRunnable = false;
        paramType = 'other';
        reason = `包含必填参数 (${params})，暂不支持直接运行`;
      }

      list.push({ name, type, modifier, params, isRunnable, paramType, reason });
    }
    return list;
  }

  $: detectedProcedures = parseVbaProcedures(importCode);
  $: runnableEntryPoints = detectedProcedures.filter((p) => p.isRunnable);

  // 当识别到可运行入口变化时，更新默认选定
  $: {
    if (runnableEntryPoints.length > 0) {
      const mainSub = runnableEntryPoints.find((p) => p.name.toLowerCase() === 'main');
      if (!importEntryPoint || !runnableEntryPoints.some((p) => p.name === importEntryPoint)) {
        importEntryPoint = mainSub ? mainSub.name : runnableEntryPoints[0].name;
      }
    } else {
      importEntryPoint = '';
    }
  }

  // 计算源码哈希
  async function updateCodeHash(code: string) {
    if (!code) {
      importCodeHash = '';
      return;
    }
    try {
      const msgBuffer = new TextEncoder().encode(code);
      const hashBuffer = await crypto.subtle.digest('SHA-256', msgBuffer);
      const hashArray = Array.from(new Uint8Array(hashBuffer));
      importCodeHash = hashArray.map((b) => b.toString(16).padStart(2, '0')).join('');
    } catch {
      importCodeHash = '';
    }
  }

  $: updateCodeHash(importCode);

  // 文件导入处理
  function handleFileInput(e: Event) {
    const input = e.target as HTMLInputElement;
    if (!input.files || input.files.length === 0) return;
    const file = input.files[0];
    processSelectedFile(file);
    input.value = '';
  }

  function processSelectedFile(file: File) {
    importErrorMessage = '';
    importSuccessMessage = '';

    const ext = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();

    // 格式合法性严格门禁
    if (ext === '.vbs') {
      importErrorMessage = '❌ .vbs 文件为 Windows Script Host 脚本，不属于 Excel VBA 宏，不支持直接导入。';
      return;
    }

    if (ext === '.cls' || ext === '.frm' || ext === '.xlsm') {
      importErrorMessage = `⚠️ 当前第一版仅支持标准 VBA 模块（.bas, .vba, .txt），暂不支持类模块(${ext})或整本工程导入。`;
      return;
    }

    if (ext !== '.bas' && ext !== '.vba' && ext !== '.txt') {
      importErrorMessage = '⚠️ 仅支持导入 .bas、.vba、.txt 格式的 VBA 宏源码文件。';
      return;
    }

    importFileName = file.name;
    const baseName = file.name.substring(0, file.name.lastIndexOf('.')) || file.name;
    if (!importDisplayName || importDisplayName === '新建宏' || importDisplayName === '自定义宏') {
      importDisplayName = baseName;
    }

    const reader = new FileReader();
    reader.onload = (event) => {
      const buffer = event.target?.result as ArrayBuffer;
      if (!buffer) return;
      importRawBuffer = buffer;

      // 保存 raw bytes base64 副本
      const uint8 = new Uint8Array(buffer);
      let binary = '';
      for (let i = 0; i < uint8.length; i++) {
        binary += String.fromCharCode(uint8[i]);
      }
      importRawBytesBase64 = btoa(binary);

      // 默认尝试 UTF-8 解码，如果乱码或非标准文本用户可自行重选编码
      decodeCurrentBuffer(buffer, importEncoding);
    };
    reader.readAsArrayBuffer(file);
  }

  function decodeCurrentBuffer(buffer: ArrayBuffer, enc: string) {
    try {
      const decoder = new TextDecoder(enc);
      importCode = decoder.decode(buffer);
    } catch {
      try {
        const fallbackDecoder = new TextDecoder('utf-8');
        importCode = fallbackDecoder.decode(buffer);
      } catch (err: any) {
        importErrorMessage = '编码解码失败: ' + err.message;
      }
    }
  }

  function handleEncodingChange(newEnc: string) {
    importEncoding = newEnc;
    if (importRawBuffer) {
      decodeCurrentBuffer(importRawBuffer, newEnc);
    }
  }

  // 保存宏
  async function handleSaveMacro(forceOverwrite = false) {
    importErrorMessage = '';
    importSuccessMessage = '';

    if (!importDisplayName.trim()) {
      importErrorMessage = '请填写宏显示名称';
      return;
    }

    if (!importCode.trim()) {
      importErrorMessage = 'VBA 源码内容不能为空';
      return;
    }

    isSaving = true;

    const payload = {
      displayName: importDisplayName.trim(),
      code: importCode,
      description: importDescription.trim(),
      category: importCategory.trim() || '未分类',
      sourceType: importSourceType,
      originalFileName: importFileName,
      encoding: importEncoding,
      entryPoint: importEntryPoint,
      rawBytesBase64: importRawBytesBase64,
      overwrite: forceOverwrite ? 'true' : 'false',
    };

    const res = await bridge.send<{ id: string; displayName: string }>('save_script', payload);
    isSaving = false;

    if (res.ok) {
      importSuccessMessage = res.message || '宏已成功保存到本地宏库！';
      showOverwriteConfirm = false;
      pendingOverwritePayload = null;
      await refreshScripts();
      // 切换至列表
      setTimeout(() => {
        activeTab = 'list';
        importSuccessMessage = '';
      }, 800);
    } else {
      if (res.error === 'DUPLICATE_NAME') {
        pendingOverwritePayload = payload;
        showOverwriteConfirm = true;
      } else {
        importErrorMessage = '保存失败: ' + (res.message || res.error);
      }
    }
  }

  // TASK-R2c-01: 校验单个参数输入
  function validateParamValue(param: VbaParamSignature, val: any): string {
    const isOptional = param.isOptional;
    const type = (param.typeName || param.type || 'Variant').toLowerCase();

    // 可选参数若未填则视为空通过
    if (isOptional && (val === undefined || val === null || val === '')) {
      return '';
    }

    if (!isOptional && (val === undefined || val === null || val === '')) {
      return '此参数为必需参数，不能为空';
    }

    if (val === undefined || val === null || val === '') {
      return '';
    }

    const strVal = String(val).trim();

    if (type === 'long') {
      if (!/^-?\d+$/.test(strVal)) return '必须为有效的整数 (Long)';
    } else if (type === 'double') {
      if (!/^-?\d+(\.\d+)?$/.test(strVal)) return '必须为有效的数值 (Double)';
    } else if (type === 'boolean') {
      if (typeof val !== 'boolean' && strVal.toLowerCase() !== 'true' && strVal.toLowerCase() !== 'false') {
        return '必须为布尔值 (True/False)';
      }
    } else if (type === 'date') {
      if (!/^\d{4}-\d{2}-\d{2}( \d{2}:\d{2}:\d{2})?$/.test(strVal)) {
        return '日期格式必须为 yyyy-MM-dd 或 yyyy-MM-dd HH:mm:ss';
      }
      const d = new Date(strVal.replace(' ', 'T'));
      if (isNaN(d.getTime())) return '输入的日期不存在或无效';
    } else if (type === 'range') {
      if (typeof val === 'object' && val !== null) {
        if (!val.sheet) return '必须指定目标工作表';
        if (!val.address || !/^\$?[A-Za-z]+\$?[0-9]*(:\$?[A-Za-z]+\$?[0-9]*)?$/i.test(String(val.address).trim())) {
          return '单元格区域地址格式无效（例: A1:D20）';
        }
      } else {
        if (!/^\$?[A-Za-z]+\$?[0-9]*(:\$?[A-Za-z]+\$?[0-9]*)?$/i.test(strVal)) {
          return '单元格区域地址格式无效（例: A1:D20）';
        }
      }
    } else if (type === 'worksheet') {
      if (!strVal) return '必须选择有效的工作表';
      if (runModalTargetSheets.length > 0 && !runModalTargetSheets.includes(strVal)) {
        return `当前目标工作簿中不存在工作表【${strVal}】`;
      }
    }

    return '';
  }

  // TASK-R2c-01: 核对元数据与实际过程签名一致性
  function checkMetadataConflict(proc: VbaEntryPointInfo, metaParams?: ScriptParameterDef[]): string {
    if (!metaParams || metaParams.length === 0) return '';
    const userParams = (proc.parameters || []).filter((p) => (p.typeName || p.type)?.toLowerCase() !== 'workbook');
    if (userParams.length !== metaParams.length) {
      return `元数据声明的参数数量(${metaParams.length})与过程【${proc.name}】实际入参数量(${userParams.length})不一致，已严格阻断执行以防误操作。`;
    }
    for (let i = 0; i < userParams.length; i++) {
      const p = userParams[i];
      const m = metaParams[i];
      if (p.name.toLowerCase() !== m.name.toLowerCase()) {
        return `参数顺序或名称不匹配：第 ${i + 1} 个参数元数据为【${m.name}】，源码签名中为【${p.name}】。`;
      }
      const pType = p.typeName || p.type;
      if (m.type && pType && m.type.toLowerCase() !== pType.toLowerCase()) {
        return `参数【${p.name}】类型冲突：元数据声明为【${m.type}】，源码签名实际为【${pType}】。`;
      }
    }
    return '';
  }

  // 初始化选定入口的参数与元数据校验
  function initParamsForSelectedEntryPoint(procName: string) {
    if (!runModalScript) return;
    const proc = runModalEntryPoints.find((p) => p.name === procName);
    if (!proc) {
      runModalMetaConflictError = '';
      runModalParamValues = {};
      runModalParamErrors = {};
      return;
    }

    // 核对元数据冲突
    runModalMetaConflictError = checkMetadataConflict(proc, runModalScript.parameters);

    const draftKey = `${runModalScript.id || runModalScript.fileName}_${procName}`;
    const draft = paramDrafts[draftKey] || {};
    const newValues: Record<string, any> = {};
    const newErrors: Record<string, string> = {};

    for (const p of proc.parameters || []) {
      const pName = p.name;
      const type = (p.typeName || p.type || '').toLowerCase();
      if (type === 'workbook') continue; // 宿主已确认目标自动绑定

      if (draft[pName] !== undefined) {
        newValues[pName] = draft[pName];
      } else {
        // 查找元数据默认值或签名默认值
        const metaP = (runModalScript.parameters || []).find((m) => m.name.toLowerCase() === pName.toLowerCase());
        if (metaP?.defaultValue !== undefined) {
          newValues[pName] = metaP.defaultValue;
        } else if (p.defaultValue) {
          newValues[pName] = p.defaultValue.replace(/^"(.*)"$/, '$1');
        } else if (type === 'boolean') {
          newValues[pName] = false;
        } else if (type === 'worksheet') {
          newValues[pName] = runModalActiveSheet || (runModalTargetSheets[0] || 'Sheet1');
        } else if (type === 'range') {
          newValues[pName] = {
            sheet: runModalActiveSheet || (runModalTargetSheets[0] || 'Sheet1'),
            address: 'A1:C10',
          };
        } else {
          newValues[pName] = '';
        }
      }

      const err = validateParamValue(p, newValues[pName]);
      if (err) newErrors[pName] = err;
    }

    runModalParamValues = newValues;
    runModalParamErrors = newErrors;
  }

  // 参数变更处理（同步写入草稿，避免取消时清空已填内容）
  function handleParamChange(param: VbaParamSignature, val: any) {
    runModalParamValues[param.name] = val;
    const err = validateParamValue(param, val);
    if (err) {
      runModalParamErrors[param.name] = err;
    } else {
      delete runModalParamErrors[param.name];
    }
    runModalParamErrors = { ...runModalParamErrors };

    if (runModalScript && runModalSelectedEntryPoint) {
      const draftKey = `${runModalScript.id || runModalScript.fileName}_${runModalSelectedEntryPoint}`;
      paramDrafts[draftKey] = { ...(paramDrafts[draftKey] || {}), [param.name]: val };
    }
  }

  // 运行前入口过程可执行判定
  function isProcRunnable(proc?: VbaEntryPointInfo): boolean {
    if (!proc) return false;
    return Boolean(proc.isSupported !== undefined ? proc.isSupported : proc.isExecutable);
  }

  // 获取具体不支持的原因（明确显示具体参数名、类型及原因，不使用笼统错误）
  function getProcUnsupportedReason(proc?: VbaEntryPointInfo): string {
    if (!proc) return '';
    if (proc.unsupportedReason) return proc.unsupportedReason;
    if (proc.kind === 'Function') {
      return 'Function 过程为有返回值的函数，不能作为独立宏直接运行。请将其包装在 Sub 过程中调用。';
    }
    if (proc.visibility?.toLowerCase() === 'private') {
      return `过程【${proc.name}】声明为 Private 私有过程，无法从外部直接调用。`;
    }
    if (proc.parameters && proc.parameters.length > 0) {
      const unsupported = proc.parameters.filter((p) => p.isSupported === false);
      if (unsupported.length > 0) {
        return (
          '过程包含暂不支持的参数签名: ' +
          unsupported
            .map(
              (p) =>
                `参数【${p.name}】(类型: ${p.rawType || p.typeName || p.type || '未声明'}) 不支持: ${p.unsupportedReason || '类型不在支持范围内'}`
            )
            .join('; ')
        );
      }
    }
    return '包含暂不支持的参数签名或过程格式';
  }

  // 运行前入口确认流程 (异步加载真实工作表与过程签名)
  async function handleOpenRunModal(script: ScriptItem) {
    runModalScript = script;
    runModalShowCode = false;
    runModalTargetError = '';
    runModalMetaConflictError = '';
    runModalTargetName = workbook?.name || '';
    runModalTargetFullName = workbook?.fullName || '';
    runModalTargetSheets = [];
    runModalActiveSheet = '';

    // 1. 读取目标工作簿的真实工作表列表
    try {
      const sheetRes = await bridge.getTargetSheets({
        targetWorkbookName: workbook?.name,
        targetWorkbookFullName: workbook?.fullName,
      });
      if (sheetRes.ok && sheetRes.sheets) {
        runModalTargetSheets = sheetRes.sheets;
        runModalActiveSheet = sheetRes.activeSheet || sheetRes.sheets[0] || '';
      }
    } catch {
      // 容灾回退
    }

    // 2. 解析源码签名（优先调用宿主 inspect_macro_signature）
    let entryPoints: VbaEntryPointInfo[] = [];
    try {
      const res = await bridge.inspectMacroSignature(script.code);
      if (res.ok && res.entryPoints) {
        entryPoints = res.entryPoints;
      }
    } catch {
      // 容灾
    }

    if (entryPoints.length === 0 && script.candidateEntryPoints && script.candidateEntryPoints.length > 0) {
      entryPoints = script.candidateEntryPoints;
    }
    runModalEntryPoints = entryPoints;

    // 3. 选择初始入口
    const supportedPoints = entryPoints.filter((p) => isProcRunnable(p));
    if (supportedPoints.length > 0) {
      if (script.entryPoint && supportedPoints.some((p) => p.name === script.entryPoint)) {
        runModalSelectedEntryPoint = script.entryPoint;
      } else {
        const mainP = supportedPoints.find((p) => p.name.toLowerCase() === 'main');
        runModalSelectedEntryPoint = mainP ? mainP.name : supportedPoints[0].name;
      }
    } else {
      runModalSelectedEntryPoint = entryPoints.length > 0 ? entryPoints[0].name : '';
    }

    // 4. 初始化所选入口的参数与元数据核对
    if (runModalSelectedEntryPoint) {
      initParamsForSelectedEntryPoint(runModalSelectedEntryPoint);
    }
  }

  $: currentSelectedProc = runModalEntryPoints.find((p) => p.name === runModalSelectedEntryPoint);
  $: currentSelectedProcRunnable = currentSelectedProc ? isProcRunnable(currentSelectedProc) : false;
  $: isCurrentEntryPointRunnable = currentSelectedProc
    ? currentSelectedProcRunnable &&
      (!workbook || workbook.isSaved) &&
      !runModalMetaConflictError &&
      Object.keys(runModalParamErrors).length === 0
    : false;
  $: cannotRunReason = currentSelectedProc && !currentSelectedProcRunnable
    ? getProcUnsupportedReason(currentSelectedProc)
    : '';

  function handleConfirmRun() {
    if (!runModalScript) return;
    runModalTargetError = '';

    // 1. 检查入口支持度
    if (currentSelectedProc && !isProcRunnable(currentSelectedProc)) {
      runModalTargetError = `所选过程【${currentSelectedProc.name}】暂不支持直接运行：${cannotRunReason}。系统不会猜测入口或修改您的源码。请取消后在对话框使用文字提问。`;
      return;
    }

    // 2. 检查元数据冲突
    if (runModalMetaConflictError) {
      runModalTargetError = runModalMetaConflictError;
      return;
    }

    // 3. 校验必需参数
    if (currentSelectedProc && currentSelectedProc.parameters) {
      const errors: Record<string, string> = {};
      for (const p of currentSelectedProc.parameters) {
        const pType = (p.typeName || p.type || '').toLowerCase();
        if (pType === 'workbook') continue;
        const err = validateParamValue(p, runModalParamValues[p.name]);
        if (err) errors[p.name] = err;
      }
      if (Object.keys(errors).length > 0) {
        runModalParamErrors = errors;
        runModalTargetError = '存在未填写或格式非法的参数，请修改后再执行！';
        return;
      }
    }

    // 4. 检查当前目标工作簿是否存在且有效
    const currentName = workbook?.name || '';
    const currentFullName = workbook?.fullName || '';
    if (!currentName || currentName === '未检测到活动工作簿') {
      runModalTargetError = '未检测到有效目标工作簿，已阻断执行。请先在 Excel 中打开或选择目标工作簿后再执行！';
      return;
    }

    // 5. 检查目标工作簿是否已保存（快照前置门禁：不能把纯名称当作磁盘路径，也不能为了运行而跳过快照）
    if (workbook && !workbook.isSaved) {
      runModalTargetError = '当前目标工作簿尚未保存到磁盘文件，无法生成整本物理快照副本。系统严格执行快照安全门禁，禁止跳过快照直接执行代码。请在 Excel 中保存工作簿后重试！';
      return;
    }

    // 6. 检查目标工作簿是否在确认期间发生变动（防静默换目标）
    if (runModalTargetName && (currentName !== runModalTargetName || currentFullName !== runModalTargetFullName)) {
      runModalTargetError = `目标工作簿在确认期间已发生变更（原目标：【${runModalTargetName}】，当前：【${currentName}】）。为防误操作已停止执行，请重新核对目标！`;
      runModalTargetName = currentName;
      runModalTargetFullName = currentFullName;
      return;
    }

    const s = runModalScript;
    const ep = runModalSelectedEntryPoint;
    const params = { ...runModalParamValues };
    runModalScript = null;
    onClose();
    onRunScript(s, ep, params);
  }

  // TASK-R2c-01: 取消并纯文字提问（零执行、零快照、保留草稿、切换到对话模式预填草稿）
  function handleCancelAndAsk() {
    if (!runModalScript) return;
    const currentProc = runModalEntryPoints.find((p) => p.name === runModalSelectedEntryPoint);
    const procName = currentProc?.name || runModalSelectedEntryPoint || '未指定入口';

    let paramDetails = '';
    if (currentProc && currentProc.parameters && currentProc.parameters.length > 0) {
      paramDetails = currentProc.parameters
        .map((p) => {
          const val = runModalParamValues[p.name];
          const valStr =
            val !== undefined && val !== null
              ? typeof val === 'object'
                ? JSON.stringify(val)
                : String(val)
              : '(未填写)';
          return `• ${p.name} (${p.typeName || p.type || 'Variant'}${p.isOptional ? ', 可选' : ', 必需'}): ${valStr}`;
        })
        .join('\n');
    } else {
      paramDetails = '（无参过程）';
    }

    let notes = '';
    if (currentProc && !isProcRunnable(currentProc)) {
      notes += `\n- 不支持原因：${cannotRunReason || getProcUnsupportedReason(currentProc)}`;
    }
    if (runModalMetaConflictError) {
      notes += `\n- 元数据冲突：${runModalMetaConflictError}`;
    }

    const questionText = `关于宏【${runModalScript.displayName || runModalScript.name}】的过程【${procName}】：\n- 目标工作簿：${workbook?.name || '未选定'}${notes}\n- 当前参数：\n${paramDetails}\n\n请教：`;

    const scriptObj = runModalScript;
    runModalScript = null;
    onClose();

    if (onCancelWithQuestion) {
      onCancelWithQuestion({
        script: scriptObj,
        entryPoint: procName,
        questionText,
      });
    }
  }

  // 重命名处理
  function openRenameModal(script: ScriptItem) {
    renameModalScript = script;
    renameNewDisplayName = script.displayName || script.name;
    renameError = '';
  }

  async function handleConfirmRename() {
    if (!renameModalScript) return;
    if (!renameNewDisplayName.trim()) {
      renameError = '显示名称不能为空';
      return;
    }
    const res = await bridge.send('rename_script', {
      id: renameModalScript.id || renameModalScript.fileName,
      newDisplayName: renameNewDisplayName.trim(),
    });
    if (res.ok) {
      renameModalScript = null;
      await refreshScripts();
    } else {
      renameError = res.error || res.message || '重命名失败';
    }
  }

  // 编辑并另存为新版本
  function handleEditAsNew(script: ScriptItem) {
    activeTab = 'import';
    importSourceType = 'paste';
    importFileName = script.originalFileName || '';
    importEncoding = script.encoding || 'UTF-8';
    importCode = script.code;
    importDisplayName = (script.displayName || script.name) + '_v2';
    importDescription = script.description || '';
    importCategory = script.category || '';
    importEntryPoint = script.entryPoint || '';
    importErrorMessage = '';
    importSuccessMessage = '';
  }

  // 删除宏
  async function handleDelete(script: ScriptItem) {
    const disp = script.displayName || script.name;
    if (!confirm(`确认删除宏【${disp}】吗？此操作不可撤销。`)) return;
    const res = await bridge.send('delete_script', { id: script.id || script.fileName });
    if (res.ok) {
      await refreshScripts();
    } else {
      alert('删除失败: ' + (res.error || res.message));
    }
  }

  // 复制代码
  async function handleCopyCode(script: ScriptItem) {
    try {
      await navigator.clipboard.writeText(script.code);
      copiedScriptId = script.id || script.fileName;
      setTimeout(() => {
        copiedScriptId = null;
      }, 1500);
    } catch {
      alert('复制到剪贴板失败，请手动选取复制代码');
    }
  }

  // 导出源码文件
  function handleExportFile(script: ScriptItem) {
    try {
      const fileName = (script.displayName || script.name || 'Macro').replace(/[/\\?%*:|"<>]/g, '_') + '.bas';
      const blob = new Blob([script.code], { type: 'text/plain;charset=utf-8' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = fileName;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    } catch (err: any) {
      alert('导出失败: ' + err.message);
    }
  }

  // ===================== 宏包导出与导入逻辑 (TASK-R6b-01) =====================
  function openExportPackageModal(selectedMacros?: ScriptItem[]) {
    if (selectedMacros && selectedMacros.length > 0) {
      exportPackageSelectedMacroIds = selectedMacros.map((s) => s.id || s.fileName);
      exportPackageName = (selectedMacros[0].displayName || selectedMacros[0].name || 'Macros').replace(/[^\w]/g, '_');
    } else {
      exportPackageSelectedMacroIds = scripts.map((s) => s.id || s.fileName);
      exportPackageName = 'ExcelMind_Macros';
    }
    exportPackageVersion = '1.0.0';
    exportPackageDescription = '';
    exportPackageOutputPath = '';
    exportPackageWarnings = [];
    exportPackageErrorMessage = '';
    exportPackageSuccessResult = null;
    showExportPackageModal = true;
  }

  async function handleBrowseExportSavePath() {
    const defaultName = `${exportPackageName || 'macros'}.exmpack`;
    const res = await bridge.browseSaveFile({
      title: '导出宏包为 .exmpack 文件',
      defaultName,
      filter: 'ExcelMind 宏包 (*.exmpack)|*.exmpack|所有文件 (*.*)|*.*',
    });
    if (res.ok && res.data) {
      exportPackageOutputPath = res.data;
    }
  }

  async function handleExecuteExportPackage(forceIgnoreWarnings = false) {
    if (exportPackageSelectedMacroIds.length === 0) {
      exportPackageErrorMessage = '请至少选择一个待导出的宏。';
      return;
    }
    exportPackageIsScanning = true;
    exportPackageErrorMessage = '';
    try {
      const res = await bridge.exportMacroPackage({
        macroIds: exportPackageSelectedMacroIds,
        packageName: exportPackageName.trim() || 'ExcelMind_Macros',
        packageVersion: exportPackageVersion.trim() || '1.0.0',
        packageDescription: exportPackageDescription.trim(),
        outputPath: exportPackageOutputPath.trim(),
        ignoreWarnings: forceIgnoreWarnings,
      });

      if (res.ok) {
        exportPackageSuccessResult = res;
        exportPackageWarnings = [];
      } else {
        if (res.warnings && res.warnings.length > 0 && !forceIgnoreWarnings) {
          exportPackageWarnings = res.warnings;
          exportPackageErrorMessage = res.error || '检测到疑似敏感凭据，请核对后确认是否继续导出。';
        } else {
          exportPackageErrorMessage = res.error || '导出宏包失败';
        }
      }
    } catch (err: any) {
      exportPackageErrorMessage = '导出宏包异常: ' + (err.message || String(err));
    } finally {
      exportPackageIsScanning = false;
    }
  }

  async function handleBrowsePackageFile() {
    packageImportError = '';
    packageImportResult = null;
    const res = await bridge.browseFiles('ExcelMind 宏包 (*.exmpack)|*.exmpack|所有文件 (*.*)|*.*');
    if (res.ok && res.data && res.data.length > 0) {
      packageFilePath = res.data[0];
      await loadPackagePreview(packageFilePath);
    }
  }

  async function loadPackagePreview(filePath: string) {
    if (!filePath) return;
    packagePreviewLoading = true;
    packageImportError = '';
    packagePreviewResult = null;
    packageImportResult = null;
    try {
      const res = await bridge.previewMacroPackage(filePath);
      if (res.ok && res.manifest) {
        packagePreviewResult = res;
        packageImportSelectedMacroIds = (res.manifest.entries || []).map((e) => e.macroId);
      } else {
        packageImportError = res.error || '解析宏包失败';
      }
    } catch (err: any) {
      packageImportError = '解析宏包异常: ' + (err.message || String(err));
    } finally {
      packagePreviewLoading = false;
    }
  }

  async function handleConfirmImportPackage() {
    if (!packageFilePath || !packagePreviewResult) return;
    if (packageImportSelectedMacroIds.length === 0) {
      packageImportError = '请至少选择一个待导入的宏。';
      return;
    }
    packageImportLoading = true;
    packageImportError = '';
    try {
      const res = await bridge.importMacroPackage({
        packagePath: packageFilePath,
        selectedMacroIds: packageImportSelectedMacroIds,
        conflictResolution: 'rename_both',
      });
      if (res.ok) {
        packageImportResult = res;
        await refreshScripts();
      } else {
        packageImportError = res.error || '导入宏包失败';
      }
    } catch (err: any) {
      packageImportError = '导入宏包异常: ' + (err.message || String(err));
    } finally {
      packageImportLoading = false;
    }
  }
</script>

{#if isOpen}
  <div class="drawer-backdrop" on:click|self={onClose} on:keydown={(e) => e.key === 'Escape' && onClose()} role="presentation">
    <div class="drawer-content" role="dialog" aria-modal="true" aria-label="宏管理与脚本库">
      <!-- 抽屉顶部标题栏与 Tab 切换 -->
      <div class="drawer-header">
        <div class="header-left">
          <div class="header-title">
            <Code2 size={18} color="#107C41" />
            <span>宏工具</span>
          </div>
          <div class="header-tabs">
            <button
              class="tab-btn {activeTab === 'list' ? 'active' : ''}"
              on:click={() => (activeTab = 'list')}
            >
              宏库 <span class="badge-count">{scripts.length}</span>
            </button>
            <button
              class="tab-btn {activeTab === 'import' ? 'active' : ''}"
              on:click={() => (activeTab = 'import')}
            >
              <PlusCircle size={13} />
              导入
            </button>
          </div>
        </div>
        <button class="btn-icon" on:click={onClose} aria-label="关闭"><X size={16} /></button>
      </div>

      <!-- 抽屉主体内容 -->
      <div class="drawer-body">
        {#if activeTab === 'list'}
          <!-- 列表顶部工具栏：搜索与分类过滤 -->
          <div class="list-toolbar">
            <div class="search-box">
              <Search size={14} color="#8a8886" />
              <input
                type="text"
                placeholder="搜索宏名称、描述、分类或代码..."
                bind:value={searchTerm}
              />
              {#if searchTerm}
                <button class="btn-clear-search" on:click={() => (searchTerm = '')}><X size={12} /></button>
              {/if}
            </div>

            {#if categories.length > 2}
              <div class="category-pills">
                {#each categories as cat}
                  <button
                    class="pill-btn {selectedCategory === cat ? 'active' : ''}"
                    on:click={() => (selectedCategory = cat)}
                  >
                    {cat}
                  </button>
                {/each}
              </div>
            {/if}

            {#if allTags.length > 0}
              <div class="tags-filter-bar">
                <span class="tags-filter-label"><Tag size={11} /> 标签:</span>
                <button
                  class="pill-btn tag-filter-pill {selectedTag === '全部' ? 'active' : ''}"
                  on:click={() => (selectedTag = '全部')}
                >
                  全部
                </button>
                {#each allTags as tag}
                  <button
                    class="pill-btn tag-filter-pill {selectedTag === tag ? 'active' : ''}"
                    on:click={() => (selectedTag = tag)}
                  >
                    #{tag}
                  </button>
                {/each}
              </div>
            {/if}

            <div class="favorite-filter-bar">
              <button
                class="pill-btn fav-filter-pill {showFavoritesOnly ? 'active' : ''}"
                on:click={() => (showFavoritesOnly = !showFavoritesOnly)}
                title="只展示已添加到 Excel 功能区动态菜单的收藏宏"
              >
                <Star size={11} fill={showFavoritesOnly ? '#ffaa00' : 'none'} color={showFavoritesOnly ? '#ffaa00' : 'currentColor'} />
                <span>仅看收藏宏 ({scripts.filter((s) => s.isFavorite).length})</span>
              </button>
            </div>

            <div class="package-export-bar">
              <button
                class="pill-btn pkg-export-pill"
                on:click={() => openExportPackageModal()}
                title="将宏库中的宏打包导出为无凭据标准宏包 (.exmpack)"
              >
                <Package size={11} color="#107C41" />
                <span>导出宏包 (.exmpack)</span>
              </button>
            </div>
          </div>

          <!-- 列表区域 -->
          {#if isLoading}
            <div class="empty-state">
              <Clock size={24} color="#8a8886" class="spin" />
              <p>正在加载宏库...</p>
            </div>
          {:else if filteredScripts.length === 0}
            <div class="empty-state">
              <FileCode size={36} color="#c8c6c4" />
              <p class="empty-title">{searchTerm || selectedTag !== '全部' || showFavoritesOnly ? '未找到匹配的宏' : '暂无已保存的宏'}</p>
              <span class="empty-hint">
                {searchTerm || selectedTag !== '全部' || showFavoritesOnly ? '尝试更换搜索词或清空筛选条件' : '点击上方【导入宏】按钮，可从本地 .bas/.vba 文件或直接粘贴源码导入。'}
              </span>
              {#if !searchTerm && selectedTag === '全部' && !showFavoritesOnly}
                <button class="btn btn-sm btn-primary" on:click={() => (activeTab = 'import')}>
                  <Upload size={12} />
                  <span>立即导入宏</span>
                </button>
              {/if}
            </div>
          {:else}
            <div class="script-list">
              {#each filteredScripts as s, idx}
                <div class="script-card">
                  <div class="card-header">
                    <div class="card-title-row">
                      <span class="macro-name" title={s.displayName || s.name}>
                        {s.displayName || s.name}
                      </span>
                      <button
                        class="btn-icon-inline"
                        on:click={() => openRenameModal(s)}
                        title="重命名显示名称（不改变源码）"
                      >
                        <Edit3 size={12} />
                      </button>
                    </div>

                    <div class="badge-row">
                      {#if s.isFavorite}
                        <span class="badge badge-favorite" title="已设为功能区收藏宏">
                          ⭐ 收藏
                        </span>
                      {/if}
                      {#if s.category}
                        <span class="badge badge-category">
                          <Tag size={10} />
                          {s.category}
                        </span>
                      {/if}
                      <span class="badge badge-source">
                        {s.sourceType === 'file' ? '本地导入' : (s.sourceType === 'legacy' ? '历史脚本' : (s.sourceType === 'corrupted_meta' ? '元数据受损' : '手工粘贴'))}
                      </span>
                      {#if s.lastExecutionResult}
                        <span class="badge {s.lastExecutionResult.includes('成功') ? 'badge-success' : (s.lastExecutionResult === '未运行' ? 'badge-neutral' : 'badge-danger')}" title={s.lastExecutionResult}>
                          {s.lastExecutionResult}
                        </span>
                      {/if}
                    </div>

                    {#if s.tags && s.tags.length > 0}
                      <div class="card-tags-row">
                        {#each s.tags as tag}
                          <button
                            class="script-tag-pill {selectedTag === tag ? 'active' : ''}"
                            on:click|stopPropagation={() => (selectedTag = selectedTag === tag ? '全部' : tag)}
                            title="按标签【{tag}】过滤"
                          >
                            #{tag}
                          </button>
                        {/each}
                      </div>
                    {/if}

                    {#if s.description}
                      <div class="macro-desc" title={s.description}>
                        {s.description}
                      </div>
                    {/if}

                    <div class="card-meta">
                      <span>保存于 {s.createdAt || s.updatedAt}</span>
                      {#if s.originalFileName}
                        <span class="meta-dot">·</span>
                        <span class="meta-file" title="原文件名: {s.originalFileName}">原文件: {s.originalFileName}</span>
                      {/if}
                      {#if s.entryPoint}
                        <span class="meta-dot">·</span>
                        <span class="meta-entry">入口: {s.entryPoint}</span>
                      {/if}
                    </div>

                    <!-- 操作栏 -->
                    <div class="card-actions">
                      <button
                        class="btn-icon btn-sm-icon {s.isFavorite ? 'btn-fav-active' : ''}"
                        on:click={() => handleToggleFavorite(s)}
                        title={s.isFavorite ? '取消收藏' : '添加至功能区收藏宏'}
                      >
                        <Star size={13} fill={s.isFavorite ? '#ffaa00' : 'none'} color={s.isFavorite ? '#ffaa00' : 'currentColor'} />
                      </button>

                      <button
                        class="btn btn-sm btn-primary"
                        on:click={() => handleOpenRunModal(s)}
                        title="选择入口并对目标工作簿运行"
                      >
                        <Play size={12} />
                        <span>运行</span>
                      </button>

                      <button
                        class="btn btn-sm btn-subtle"
                        on:click={() => (expandedIndex = expandedIndex === idx ? null : idx)}
                        title={expandedIndex === idx ? '收起详情与源码' : '查看标签、运行历史与源码'}
                      >
                        <FileText size={12} />
                        <span>{expandedIndex === idx ? '收起' : '详情'}</span>
                        {#if expandedIndex === idx}
                          <ChevronUp size={12} />
                        {:else}
                          <ChevronDown size={12} />
                        {/if}
                      </button>

                      <button
                        class="btn btn-sm btn-subtle"
                        on:click={() => handleEditAsNew(s)}
                        title="以当前宏为模板编辑并另存为新版本"
                      >
                        <Layers size={12} />
                        <span>另存新版</span>
                      </button>

                      <button
                        class="btn-icon btn-sm-icon btn-danger-icon"
                        on:click={() => handleDelete(s)}
                        title="删除该宏"
                      >
                        <Trash2 size={13} color="#C42B1C" />
                      </button>
                    </div>
                  </div>

                  <!-- 展开查看标签管理、运行历史与完整源码 -->
                  {#if expandedIndex === idx}
                    <div class="expanded-drawer-panel">
                      <!-- 标签管理区 -->
                      <div class="section-block tags-management-block">
                        <div class="section-title">
                          <Tag size={12} color="#0078D4" />
                          <span>标签管理</span>
                          <span class="sub-hint">（增删标签仅更新元数据，不修改 .bas 源码）</span>
                        </div>
                        <div class="tags-editor">
                          {#if s.tags && s.tags.length > 0}
                            <div class="tag-chips">
                              {#each s.tags as tag}
                                <span class="tag-chip">
                                  <span>#{tag}</span>
                                  <button
                                    class="btn-tag-remove"
                                    on:click={() => handleRemoveTag(s, tag)}
                                    title="移除标签 {tag}"
                                  >
                                    <X size={10} />
                                  </button>
                                </span>
                              {/each}
                            </div>
                          {:else}
                            <span class="empty-tag-hint">暂无标签</span>
                          {/if}
                          <div class="add-tag-box">
                            <input
                              type="text"
                              placeholder="新标签名"
                              class="input-tag"
                              bind:value={newTagInput[s.id || s.fileName]}
                              on:keydown={(e) => e.key === 'Enter' && handleAddTag(s)}
                            />
                            <button
                              class="btn btn-xs btn-subtle btn-add-tag"
                              on:click={() => handleAddTag(s)}
                              disabled={!(newTagInput[s.id || s.fileName] || '').trim()}
                            >
                              +添加
                            </button>
                          </div>
                        </div>
                        {#if tagOperationError}
                          <div class="tag-error-text">{tagOperationError}</div>
                        {/if}
                      </div>

                      <!-- 运行历史溯源（最近10次） -->
                      <div class="section-block run-history-block">
                        <div class="section-title">
                          <Clock size={12} color="#107C41" />
                          <span>运行历史溯源 (最近 10 次)</span>
                        </div>
                        {#if s.runHistory && s.runHistory.length > 0}
                          <div class="history-list">
                            {#each [...s.runHistory].reverse() as rh}
                              <div class="history-item {rh.status}">
                                <div class="history-main">
                                  <div class="history-header">
                                    <span class="history-status-badge status-{rh.status}">
                                      {rh.status === 'success' ? '✔ 成功' : (rh.status === 'blocked' ? '⊘ 阻断' : '✖ 失败')}
                                    </span>
                                    <span class="history-phase">阶段: {rh.phase || '执行'}</span>
                                    {#if rh.status !== 'blocked' && rh.elapsedMs !== undefined && rh.elapsedMs >= 0}
                                      <span class="history-elapsed">耗时: {rh.elapsedMs}ms</span>
                                    {/if}
                                    <span class="history-time">{rh.executedAt}</span>
                                  </div>
                                  <div class="history-details">
                                    <span class="history-wb" title="目标工作簿">目标: {rh.targetWorkbookName || '活动工作簿'}</span>
                                    {#if rh.entryPoint}
                                      <span class="history-entry">入口: {rh.entryPoint}</span>
                                    {/if}
                                    {#if rh.codeHash}
                                      <span class="history-hash" title="源码SHA256哈希: {rh.codeHash}">
                                        哈希: {rh.codeHash.substring(0, 8)}...
                                      </span>
                                    {/if}
                                  </div>
                                  <div class="history-snapshot-row">
                                    {#if rh.snapshotId}
                                      {#if rh.snapshotExists}
                                        <span class="snapshot-badge snapshot-valid" title="快照文件存在于备份目录">
                                          🛡️ 快照可用 ({rh.snapshotId})
                                        </span>
                                      {:else}
                                        <span class="snapshot-badge snapshot-missing" title="快照记录存在但文件已从备份目录物理移除">
                                          ⚠️ 快照文件已丢失 ({rh.snapshotId})
                                        </span>
                                      {/if}
                                    {:else}
                                      <span class="snapshot-badge snapshot-none" title={rh.snapshotReason || '未备份'}>
                                        ⚪ 无快照: {rh.snapshotReason || '未创建备份'}
                                      </span>
                                    {/if}
                                  </div>
                                  {#if rh.summary}
                                    <div class="history-summary" title={rh.summary}>
                                      {rh.summary}
                                    </div>
                                  {/if}
                                </div>
                              </div>
                            {/each}
                          </div>
                        {:else}
                          <div class="history-empty">
                            <span>暂无运行记录。对工作簿运行此宏后将自动在此溯源。</span>
                          </div>
                        {/if}
                      </div>

                      <!-- 完整源码展示 -->
                      <div class="code-container">
                        <div class="code-meta-bar">
                          <div class="code-hash">
                            <Hash size={11} />
                            <span>SHA256: {s.originalCodeHash ? s.originalCodeHash.substring(0, 16) + '...' : '未生成'}</span>
                          </div>
                          <div class="code-toolbar">
                            <button
                              class="btn-text-sm"
                              on:click={() => handleCopyCode(s)}
                              title="复制全部 VBA 源码"
                            >
                              {#if copiedScriptId === (s.id || s.fileName)}
                                <Check size={12} color="#107C41" />
                                <span style="color: #107C41;">已复制</span>
                              {:else}
                                <Copy size={12} />
                                <span>复制</span>
                              {/if}
                            </button>
                            <button
                              class="btn-text-sm"
                              on:click={() => handleExportFile(s)}
                              title="导出为 .bas 源码文件"
                            >
                              <Download size={12} />
                              <span>导出.bas</span>
                            </button>
                            <button
                              class="btn-text-sm"
                              on:click={() => openExportPackageModal([s])}
                              title="将此宏导出为无凭据宏包 (.exmpack)"
                            >
                              <Package size={12} />
                              <span>导出宏包</span>
                            </button>
                          </div>
                        </div>
                        <pre><code>{s.code}</code></pre>
                      </div>
                    </div>
                  {/if}
                </div>
              {/each}
            </div>
          {/if}

        {:else if activeTab === 'import'}
          <!-- 导入宏视图 -->
          <div class="import-view">
            <!-- 导入方式选择 -->
            <div class="import-mode-toggle">
              <button
                class="mode-btn {importSourceType === 'file' ? 'active' : ''}"
                on:click={() => (importSourceType = 'file')}
              >
                <Upload size={14} />
                <span>从本地文件导入 (.bas)</span>
              </button>
              <button
                class="mode-btn {importSourceType === 'paste' ? 'active' : ''}"
                on:click={() => (importSourceType = 'paste')}
              >
                <FileText size={14} />
                <span>粘贴 VBA 源码</span>
              </button>
              <button
                class="mode-btn {importSourceType === 'package' ? 'active' : ''}"
                on:click={() => { importSourceType = 'package'; packageImportError = ''; packageImportResult = null; }}
              >
                <Package size={14} />
                <span>从宏包导入 (.exmpack)</span>
              </button>
            </div>

            {#if importSourceType === 'file' || importSourceType === 'paste'}
              <!-- 本地文件选取区域 -->
              {#if importSourceType === 'file'}
              <div class="file-picker-card">
                <input
                  type="file"
                  id="vbaFileInput"
                  accept=".bas,.vba,.txt"
                  on:change={handleFileInput}
                  style="display: none;"
                />
                <label for="vbaFileInput" class="file-dropzone">
                  <Upload size={24} color="#107C41" />
                  <span class="dropzone-text">点击选择本地 VBA 模块文件</span>
                  <span class="dropzone-hint">支持标准模块 .bas、纯文本 .vba、.txt（拒绝 .vbs，暂不支持类模块 .cls/.frm）</span>
                </label>

                {#if importFileName}
                  <div class="file-info-badge">
                    <FileCode size={14} color="#107C41" />
                    <span class="file-name">{importFileName}</span>
                    <div class="encoding-selector">
                      <span class="encoding-label">编码:</span>
                      <select
                        bind:value={importEncoding}
                        on:change={(e) => handleEncodingChange(e.currentTarget.value)}
                      >
                        <option value="UTF-8">UTF-8</option>
                        <option value="gbk">GBK / GB2312 (中文ANSI)</option>
                        <option value="windows-1252">Windows-1252 (西欧)</option>
                        <option value="utf-16le">UTF-16LE</option>
                      </select>
                    </div>
                  </div>
                {/if}
              </div>
            {/if}

            <!-- 粘贴 VBA 源码区域 -->
            {#if importSourceType === 'paste'}
              <div class="form-group">
                <label for="pasteCodeInput">
                  <span>VBA 完整源码</span>
                  <span class="field-hint">可包含 Option Explicit、注释及辅助 Sub/Function 过程，无需强制重命名为 Main</span>
                </label>
                <textarea
                  id="pasteCodeInput"
                  rows={8}
                  placeholder="' 在此粘贴完整 VBA 源码&#10;Sub CustomMacro()&#10;    Range(&quot;A1&quot;).Value = &quot;Hello&quot;&#10;End Sub"
                  bind:value={importCode}
                ></textarea>
              </div>
            {/if}

            <!-- 错误或警告提示 -->
            {#if importErrorMessage}
              <div class="alert-box alert-error">
                <AlertCircle size={15} />
                <span>{importErrorMessage}</span>
              </div>
            {/if}

            {#if importSuccessMessage}
              <div class="alert-box alert-success">
                <Check size={15} />
                <span>{importSuccessMessage}</span>
              </div>
            {/if}

            <!-- 识别到的过程与入口预览 -->
            {#if detectedProcedures.length > 0}
              <div class="procedure-panel">
                <div class="panel-header">
                  <span>识别到的过程结构 ({detectedProcedures.length})</span>
                </div>
                <div class="proc-list">
                  {#each detectedProcedures as proc}
                    <div class="proc-item {proc.isRunnable ? 'proc-runnable' : 'proc-non-runnable'}">
                      <div class="proc-info">
                        <span class="proc-type">{proc.modifier} {proc.type}</span>
                        <span class="proc-name">{proc.name}({proc.params})</span>
                      </div>
                      <span class="proc-status">
                        {proc.isRunnable ? '可作为入口' : proc.reason}
                      </span>
                    </div>
                  {/each}
                </div>
              </div>
            {/if}

            <!-- 宏元数据配置表单 -->
            <div class="meta-form">
              <div class="form-group">
                <label for="importNameInput">
                  <span>宏显示名称 <span class="required">*</span></span>
                  <span class="field-hint">仅用于管理列表展示，绝不篡改 VBA 内部 Sub/Function 名称</span>
                </label>
                <input
                  id="importNameInput"
                  type="text"
                  placeholder="例如：按部门汇总销售报表"
                  bind:value={importDisplayName}
                />
              </div>

              <div class="form-row">
                <div class="form-group flex-1">
                  <label for="importCategoryInput">分类 / 标签</label>
                  <input
                    id="importCategoryInput"
                    type="text"
                    placeholder="如：日常报表、数据清洗"
                    bind:value={importCategory}
                  />
                </div>

                <div class="form-group flex-1">
                  <label for="importEntrySelect">默认执行入口</label>
                  {#if runnableEntryPoints.length > 0}
                    <select id="importEntrySelect" bind:value={importEntryPoint}>
                      {#each runnableEntryPoints as p}
                        <option value={p.name}>{p.name} ({p.paramType === 'workbook' ? '接收工作簿参数' : '无参Sub'})</option>
                      {/each}
                    </select>
                  {:else}
                    <input
                      id="importEntrySelect"
                      type="text"
                      disabled
                      value="未检测到公开无参入口 (允许保存为未验证源码)"
                    />
                  {/if}
                </div>
              </div>

              <div class="form-group">
                <label for="importDescInput">功能描述（可选）</label>
                <input
                  id="importDescInput"
                  type="text"
                  placeholder="简要说明此宏的功能作用及使用说明..."
                  bind:value={importDescription}
                />
              </div>

              {#if importCodeHash}
                <div class="hash-preview">
                  <Hash size={12} color="#8a8886" />
                  <span>原文 SHA256 哈希: <code>{importCodeHash}</code></span>
                </div>
              {/if}

              <!-- 保存安全提示 -->
              <div class="save-security-note">
                <AlertCircle size={13} color="#0078D4" />
                <span>保存仅在本地宏库持久化，<strong>不向工作簿注入代码、不执行宏、不创建快照</strong>。有错误的代码亦可作为未验证源码保存。</span>
              </div>

              <!-- 操作按钮 -->
              <div class="form-actions">
                <button
                  class="btn btn-primary"
                  disabled={isSaving || !importCode.trim() || !importDisplayName.trim()}
                  on:click={() => handleSaveMacro(false)}
                >
                  <Check size={14} />
                  <span>{isSaving ? '正在保存...' : '保存到我的宏'}</span>
                </button>
                <button
                  class="btn btn-subtle"
                  disabled={isSaving}
                  on:click={() => (activeTab = 'list')}
                >
                  取消
                </button>
              </div>
            </div>

            <!-- 源码实时预览卡片 (从文件导入时) -->
            {#if importSourceType === 'file' && importCode}
              <div class="code-preview-section">
                <div class="preview-header">
                  <span>源码原文预览 ({importCode.length} 字符)</span>
                </div>
                <pre class="preview-code"><code>{importCode}</code></pre>
              </div>
            {/if}
          {:else if importSourceType === 'package'}
            <!-- 宏包导入视图 (TASK-R6b-01) -->
            <div class="package-import-view">
              {#if !packagePreviewResult && !packageImportResult}
                <div class="file-picker-card">
                  <button class="file-dropzone w-full" on:click={handleBrowsePackageFile} type="button">
                    <Package size={28} color="#107C41" />
                    <span class="dropzone-text">点击选择本地 .exmpack 宏包文件</span>
                    <span class="dropzone-hint">标准 ZIP 容器，包含 manifest.json 与 scripts/*.bas 源码，纯离线无凭据导入</span>
                  </button>
                </div>

                <div class="info-callout">
                  <div class="info-title">🛡️ 隔离解析与容量保护保证</div>
                  <div class="info-text">
                    宏包将在隔离临时解包目录中进行预检与哈希验证，自动拦截绝对路径、盘符/UNC路径、路径穿越、重复条目、压缩炸弹与意外文件类型。<br/>
                    导入时为每个宏生成全新本地稳定 ID，同名宏自动重命名共存，绝不静默覆盖已有宏与运行历史。<br/>
                    <strong>导入全程零宏执行</strong>，不自动添加任何运行任务。
                  </div>
                </div>
              {/if}

              {#if packagePreviewLoading}
                <div class="empty-state">
                  <Clock size={24} color="#8a8886" class="spin" />
                  <p>正在隔离解析宏包并校验文件哈希...</p>
                </div>
              {/if}

              {#if packageImportError}
                <div class="alert-box alert-error">
                  <AlertCircle size={15} />
                  <span>{packageImportError}</span>
                </div>
              {/if}

              {#if packagePreviewResult && packagePreviewResult.manifest && !packageImportResult}
                <!-- 宏包信息概览 -->
                <div class="manifest-card">
                  <div class="manifest-header">
                    <div class="manifest-title">
                      <Package size={16} color="#107C41" />
                      <strong>{packagePreviewResult.manifest.name || '未命名宏包'}</strong>
                      <span class="pkg-version-badge">v{packagePreviewResult.manifest.version || '1.0.0'}</span>
                    </div>
                    <button class="btn btn-xs btn-subtle" on:click={handleBrowsePackageFile}>更换宏包</button>
                  </div>
                  <div class="manifest-meta-grid">
                    <div><span class="meta-label">导出方:</span> {packagePreviewResult.manifest.exportedBy || 'ExcelMind AI'}</div>
                    <div><span class="meta-label">导出时间:</span> {packagePreviewResult.manifest.exportedAt || '未知'}</div>
                    <div><span class="meta-label">包含宏数:</span> {packagePreviewResult.entryCount || packagePreviewResult.macroCount} 个</div>
                    <div><span class="meta-label">总解压大小:</span> {packagePreviewResult.totalUncompressedSize || packagePreviewResult.totalUncompressedBytes} 字节</div>
                  </div>
                  {#if packagePreviewResult.manifest.description}
                    <div class="manifest-desc">{packagePreviewResult.manifest.description}</div>
                  {/if}
                </div>

                <!-- 敏感内容警告 -->
                {#if packagePreviewResult.hasSensitiveWarnings && packagePreviewResult.warnings && packagePreviewResult.warnings.length > 0}
                  <div class="alert-box alert-warning">
                    <div class="alert-title">
                      <AlertTriangle size={15} color="#B25900" />
                      <strong>疑似敏感凭据风险提示</strong>
                    </div>
                    <p class="alert-sub">在宏包源码、描述或参数中检出疑似凭据/密钥信息。此检查为本地静态规则，未检出不等于绝对无敏感内容。本功能不担保宏正文绝对安全性，导入后绝不会自动运行宏。请人工核对：</p>
                    <ul class="warning-detail-list">
                      {#each packagePreviewResult.warnings as w}
                        <li>• 【{w.macroDisplayName || w.macroName}】{w.fieldName || w.location}: <code>{w.snippet}</code></li>
                      {/each}
                    </ul>
                  </div>
                {/if}

                <!-- 同名宏冲突提示 -->
                {#if packagePreviewResult.conflictingNames && packagePreviewResult.conflictingNames.length > 0}
                  <div class="alert-box alert-info">
                    <AlertCircle size={15} color="#005A9E" />
                    <span>宏库中已存在同名宏（{packagePreviewResult.conflictingNames.join('、')}）。导入时将自动保留双方并显式命名为 <code>{'{宏名}'} (导入)</code>，不覆盖已有宏的本地 ID、标签或运行历史。</span>
                  </div>
                {/if}

                <!-- 包含宏条目清单 -->
                <div class="package-entries-block">
                  <div class="entries-header">
                    <span>包含的宏条目 (勾选待导入项)</span>
                    <span class="entries-count">已勾选 {packageImportSelectedMacroIds.length} / {packagePreviewResult.manifest.entries.length}</span>
                  </div>
                  <div class="entries-list">
                    {#each packagePreviewResult.manifest.entries as entry}
                      <label class="entry-item">
                        <input
                          type="checkbox"
                          value={entry.macroId}
                          checked={packageImportSelectedMacroIds.includes(entry.macroId)}
                          on:change={(e) => {
                            if (e.currentTarget.checked) {
                              packageImportSelectedMacroIds = [...packageImportSelectedMacroIds, entry.macroId];
                            } else {
                              packageImportSelectedMacroIds = packageImportSelectedMacroIds.filter(id => id !== entry.macroId);
                            }
                          }}
                        />
                        <div class="entry-info">
                          <div class="entry-name-row">
                            <strong>{entry.displayName}</strong>
                            <span class="entry-badge">{entry.category || '通用'}</span>
                            <span class="entry-size">{entry.sourceByteLength} 字节</span>
                          </div>
                          <div class="entry-detail-row">
                            <span>入口: <code>{entry.entryPoint || '无指定入口'}</code></span>
                            <span>参数: {entry.parameterDefs ? entry.parameterDefs.length : 0} 个</span>
                            <span title="SHA-256 哈希: {entry.sha256}">SHA: {entry.sha256 ? entry.sha256.substring(0, 12) + '...' : '无'}</span>
                          </div>
                          {#if entry.description}
                            <div class="entry-desc">{entry.description}</div>
                          {/if}
                        </div>
                      </label>
                    {/each}
                  </div>
                </div>

                <!-- 导入操作栏 -->
                <div class="form-actions">
                  <button
                    class="btn btn-primary"
                    disabled={packageImportLoading || packageImportSelectedMacroIds.length === 0}
                    on:click={handleConfirmImportPackage}
                  >
                    <Check size={14} />
                    <span>{packageImportLoading ? '正在校验并写入宏库...' : `确认导入选定宏 (${packageImportSelectedMacroIds.length} 个)`}</span>
                  </button>
                  <button
                    class="btn btn-subtle"
                    disabled={packageImportLoading}
                    on:click={() => { packagePreviewResult = null; packageFilePath = ''; }}
                  >
                    取消
                  </button>
                </div>
              {/if}

              {#if packageImportResult}
                <div class="import-result-card">
                  <div class="result-header">
                    <Check size={20} color="#107C41" />
                    <strong>宏包导入成功！</strong>
                  </div>
                  <p class="result-summary">{packageImportResult.summary}</p>
                  {#if packageImportResult.renamedMacros && packageImportResult.renamedMacros.length > 0}
                    <div class="renamed-box">
                      <span class="renamed-title">同名宏重命名记录:</span>
                      <ul>
                        {#each packageImportResult.renamedMacros as rm}
                          <li>• {rm}</li>
                        {/each}
                      </ul>
                    </div>
                  {/if}
                  <div class="zero-execution-notice">
                    <AlertCircle size={13} color="#107C41" />
                    <span>导入全程<strong>零宏执行</strong>，未自动添加到任务、工作流或功能区收藏。</span>
                  </div>
                  <div class="result-actions">
                    <button class="btn btn-primary" on:click={() => (activeTab = 'list')}>查看宏库</button>
                    <button class="btn btn-subtle" on:click={() => { packageImportResult = null; packagePreviewResult = null; packageFilePath = ''; }}>导入其他宏包</button>
                  </div>
                </div>
              {/if}
            </div>
          {/if}
          </div>
        {/if}
      </div>
    </div>
  </div>
{/if}

<!-- 运行前目标工作簿与入口确认弹窗 -->
{#if runModalScript}
  <div class="modal-backdrop" role="presentation">
    <div class="modal-dialog modal-dialog-lg" role="dialog" aria-modal="true" aria-label="确认运行宏">
      <div class="modal-header">
        <div class="modal-title">
          <Play size={16} color="#107C41" />
          <span>确认运行宏: 【{runModalScript.displayName || runModalScript.name}】</span>
        </div>
        <button class="btn-icon" on:click={() => (runModalScript = null)}><X size={14} /></button>
      </div>

      <div class="modal-body modal-body-scroll">
        <div class="info-row">
          <span class="info-label">目标工作簿:</span>
          <span class="info-val highlight">{workbook?.name || '未检测到活动工作簿'}</span>
        </div>
        {#if workbook?.isSaved && workbook?.fullName}
          <div class="info-row">
            <span class="info-label">磁盘路径:</span>
            <span class="info-val subtext" title={workbook.fullName}>{workbook.fullName}</span>
          </div>
        {:else if workbook && !workbook.isSaved}
          <div class="info-row">
            <span class="info-label">保存状态:</span>
            <span class="info-val subtext" style="color: #b25900; font-weight: 500;">⚠️ 尚未保存至磁盘（内存新建工作簿）</span>
          </div>
        {/if}

        <!-- 未保存工作簿快照门禁警示 -->
        {#if workbook && !workbook.isSaved}
          <div class="alert-box alert-warning" style="margin-top: 4px;">
            <AlertTriangle size={14} color="#b25900" />
            <span><strong>快照前置门禁：</strong>当前目标工作簿尚未保存到磁盘文件，无法生成整本物理快照副本。系统严格执行快照安全门禁，禁止跳过快照直接执行。请在 Excel 中先保存工作簿再运行。</span>
          </div>
        {/if}

        <div class="info-row">
          <span class="info-label">执行入口过程:</span>
          {#if runModalEntryPoints.length > 1}
            <select
              value={runModalSelectedEntryPoint}
              on:change={(e) => {
                runModalSelectedEntryPoint = e.currentTarget.value;
                initParamsForSelectedEntryPoint(e.currentTarget.value);
              }}
              class="entry-select"
            >
              {#each runModalEntryPoints as p}
                <option value={p.name}>
                  {p.name}
                  {#if !isProcRunnable(p)}
                    (不支持: {getProcUnsupportedReason(p)})
                  {:else if p.parameters.length === 0}
                    (无参过程)
                  {:else if p.parameters.length === 1 && (p.parameters[0].typeName || p.parameters[0].type) === 'Workbook'}
                    (宿主绑定工作簿)
                  {:else}
                    (参数: {p.parameters.map((x) => x.typeName || x.type).join(', ')})
                  {/if}
                </option>
              {/each}
            </select>
          {:else if runModalEntryPoints.length === 1}
            <span class="info-val">
              {runModalEntryPoints[0].name}
              {#if !isProcRunnable(runModalEntryPoints[0])}
                (不可独立调用: {getProcUnsupportedReason(runModalEntryPoints[0])})
              {:else if runModalEntryPoints[0].parameters.length === 0}
                (无参入口)
              {:else if runModalEntryPoints[0].parameters.length === 1 && (runModalEntryPoints[0].parameters[0].typeName || runModalEntryPoints[0].parameters[0].type) === 'Workbook'}
                (工作簿参数入口)
              {:else}
                (带显式参数: {runModalEntryPoints[0].parameters.map((x) => (x.typeName || x.type) + ' ' + x.name).join(', ')})
              {/if}
            </span>
          {:else}
            <span class="info-val error-text">未检测到有效过程入口</span>
          {/if}
        </div>

        <!-- 元数据与签名冲突警示 -->
        {#if runModalMetaConflictError}
          <div class="alert-box alert-warning" style="margin-top: 4px;">
            <AlertTriangle size={14} color="#b25900" />
            <span><strong>契约冲突阻断：</strong>{runModalMetaConflictError}</span>
          </div>
        {/if}

        <!-- 不受支持签名提示 -->
        {#if currentSelectedProc && !currentSelectedProcRunnable}
          <div class="alert-box alert-error" style="margin-top: 4px;">
            <AlertCircle size={14} />
            <span>当前所选入口暂不支持直接运行：{cannotRunReason}。系统不会猜测入口、不自动改写源码。请使用文字提问咨询。</span>
          </div>
        {/if}

        <!-- 参数填写表单区 (TASK-R2c-01 显式参数) -->
        {#if currentSelectedProc && currentSelectedProcRunnable && currentSelectedProc.parameters && currentSelectedProc.parameters.filter((p) => (p.typeName || p.type)?.toLowerCase() !== 'workbook').length > 0}
          <div class="params-section">
            <div class="params-section-title">
              <span>参数填写（用户主动输入，宿主独立适配）</span>
            </div>

            <div class="params-grid">
              {#each currentSelectedProc.parameters.filter((p) => (p.typeName || p.type)?.toLowerCase() !== 'workbook') as param}
                {@const metaP = (runModalScript.parameters || []).find((m) => m.name.toLowerCase() === param.name.toLowerCase())}
                {@const pTypeName = param.typeName || param.type || 'Variant'}
                <div class="param-card {runModalParamErrors[param.name] ? 'has-error' : ''}">
                  <div class="param-header">
                    <span class="param-name">{param.name}</span>
                    <span class="param-badge-type">{pTypeName}</span>
                    {#if param.isOptional}
                      <span class="param-badge-optional">可选</span>
                    {:else}
                      <span class="param-badge-required">必需</span>
                    {/if}
                  </div>

                  {#if metaP?.description}
                    <div class="param-meta-desc">{metaP.description}</div>
                  {/if}

                  <div class="param-control">
                    {#if pTypeName.toLowerCase() === 'boolean'}
                      <select
                        class="form-control"
                        value={String(runModalParamValues[param.name] ?? false)}
                        on:change={(e) => handleParamChange(param, e.currentTarget.value === 'true')}
                      >
                        <option value="false">False (假)</option>
                        <option value="true">True (真)</option>
                      </select>
                    {:else if pTypeName.toLowerCase() === 'worksheet'}
                      {#if runModalTargetSheets.length > 0}
                        <select
                          class="form-control"
                          value={runModalParamValues[param.name] ?? ''}
                          on:change={(e) => handleParamChange(param, e.currentTarget.value)}
                        >
                          <option value="">-- 请选择目标工作表 --</option>
                          {#each runModalTargetSheets as sheetName}
                            <option value={sheetName}>{sheetName}</option>
                          {/each}
                        </select>
                      {:else}
                        <input
                          type="text"
                          class="form-control"
                          placeholder="目标工作表名称"
                          value={runModalParamValues[param.name] ?? ''}
                          on:input={(e) => handleParamChange(param, e.currentTarget.value)}
                        />
                      {/if}
                      <span class="field-hint">从已锁定目标工作簿中选择工作表</span>
                    {:else if pTypeName.toLowerCase() === 'range'}
                      <div class="range-compound-group">
                        <div class="range-sheet-part">
                          <span class="inner-label">工作表</span>
                          {#if runModalTargetSheets.length > 0}
                            <select
                              class="form-control"
                              value={runModalParamValues[param.name]?.sheet ?? (runModalActiveSheet || runModalTargetSheets[0] || '')}
                              on:change={(e) => {
                                const prev = typeof runModalParamValues[param.name] === 'object' && runModalParamValues[param.name] !== null ? runModalParamValues[param.name] : { address: 'A1:C10' };
                                handleParamChange(param, { ...prev, sheet: e.currentTarget.value });
                              }}
                            >
                              {#each runModalTargetSheets as sheetName}
                                <option value={sheetName}>{sheetName}</option>
                              {/each}
                            </select>
                          {:else}
                            <input
                              type="text"
                              class="form-control"
                              placeholder="表名"
                              value={runModalParamValues[param.name]?.sheet ?? (runModalActiveSheet || '')}
                              on:input={(e) => {
                                const prev = typeof runModalParamValues[param.name] === 'object' && runModalParamValues[param.name] !== null ? runModalParamValues[param.name] : { address: 'A1:C10' };
                                handleParamChange(param, { ...prev, sheet: e.currentTarget.value });
                              }}
                            />
                          {/if}
                        </div>
                        <div class="range-addr-part">
                          <span class="inner-label">区域地址</span>
                          <input
                            type="text"
                            class="form-control"
                            placeholder="例: A1:D20"
                            value={runModalParamValues[param.name]?.address ?? (typeof runModalParamValues[param.name] === 'string' ? runModalParamValues[param.name] : 'A1:C10')}
                            on:input={(e) => {
                              const currentSheet = runModalParamValues[param.name]?.sheet || runModalActiveSheet || runModalTargetSheets[0] || 'Sheet1';
                              handleParamChange(param, { sheet: currentSheet, address: e.currentTarget.value });
                            }}
                          />
                        </div>
                      </div>
                      <span class="field-hint">明确选择工作表与区域地址，执行前清晰展示</span>
                    {:else if pTypeName.toLowerCase() === 'date'}
                      <input
                        type="text"
                        class="form-control"
                        placeholder="yyyy-MM-dd (例: 2026-10-02)"
                        value={runModalParamValues[param.name] ?? ''}
                        on:input={(e) => handleParamChange(param, e.currentTarget.value)}
                      />
                      <span class="field-hint">固定日期格式，避免地区格式歧义</span>
                    {:else if pTypeName.toLowerCase() === 'long'}
                      <input
                        type="number"
                        step="1"
                        class="form-control"
                        placeholder="输入整数 (Long)"
                        value={runModalParamValues[param.name] ?? ''}
                        on:input={(e) => handleParamChange(param, e.currentTarget.value)}
                      />
                    {:else if pTypeName.toLowerCase() === 'double'}
                      <input
                        type="number"
                        step="any"
                        class="form-control"
                        placeholder="输入数值 (Double)"
                        value={runModalParamValues[param.name] ?? ''}
                        on:input={(e) => handleParamChange(param, e.currentTarget.value)}
                      />
                    {:else}
                      <input
                        type="text"
                        class="form-control"
                        placeholder="输入文本内容..."
                        value={runModalParamValues[param.name] ?? ''}
                        on:input={(e) => handleParamChange(param, e.currentTarget.value)}
                      />
                    {/if}
                  </div>

                  {#if runModalParamErrors[param.name]}
                    <div class="param-error-text">{runModalParamErrors[param.name]}</div>
                  {/if}
                </div>
              {/each}
            </div>
          </div>
        {/if}

        <!-- 源码查看入口 (折叠展开) -->
        <div class="source-view-toggle">
          <button class="btn-text-sm" on:click={() => (runModalShowCode = !runModalShowCode)}>
            <Code2 size={13} />
            <span>{runModalShowCode ? '收起完整源码' : '查看完整源码'}</span>
            {#if runModalShowCode}
              <ChevronUp size={12} />
            {:else}
              <ChevronDown size={12} />
            {/if}
          </button>
          {#if runModalScript.originalCodeHash}
            <span class="code-hash-mini" title="SHA256哈希">SHA256: {runModalScript.originalCodeHash.substring(0, 16)}...</span>
          {/if}
        </div>

        {#if runModalShowCode}
          <div class="modal-code-preview">
            <pre><code>{runModalScript.code}</code></pre>
          </div>
        {/if}

        {#if runModalTargetError}
          <div class="alert-box alert-error" style="margin-top: 6px;">
            <AlertTriangle size={14} />
            <span>{runModalTargetError}</span>
          </div>
        {/if}

        <div class="security-banner">
          <AlertCircle size={14} color="#0078D4" />
          <div class="banner-text">
            <strong>安全须知与执行声明：</strong>
            <p>1. 确认后将<strong>自动创建整本物理副本快照</strong>（需目标具备磁盘路径）并关联本次调用记录；</p>
            <p>2. <strong>包装器独立适配，正文零改写</strong>：参数经独立宿主包装器安全传递，原源码逐字节不变，杜绝代码逃逸；</p>
            <p>3. 宿主绝不猜测入口、绝不修改源码、不请求大模型。</p>
          </div>
        </div>
      </div>

      <div class="modal-footer">
        <button
          class="btn btn-primary"
          disabled={!isCurrentEntryPointRunnable || !workbook?.name || workbook?.name === '未检测到活动工作簿' || !!runModalMetaConflictError || Object.keys(runModalParamErrors).length > 0}
          on:click={handleConfirmRun}
        >
          <Play size={13} />
          <span>立即执行</span>
        </button>
        <button
          class="btn btn-secondary"
          on:click={handleCancelAndAsk}
          type="button"
          title="切换至对话模式并预填咨询文本，不发送、不执行、零快照"
        >
          <MessageSquare size={13} />
          <span>取消并纯文字提问</span>
        </button>
        <button class="btn btn-subtle" on:click={() => (runModalScript = null)}>
          取消
        </button>
      </div>
    </div>
  </div>
{/if}

<!-- 重命名弹窗 -->
{#if renameModalScript}
  <div class="modal-backdrop" role="presentation">
    <div class="modal-dialog sm-dialog" role="dialog" aria-modal="true" aria-label="重命名宏">
      <div class="modal-header">
        <div class="modal-title">
          <Edit3 size={15} color="#107C41" />
          <span>修改宏显示名称</span>
        </div>
        <button class="btn-icon" on:click={() => (renameModalScript = null)}><X size={14} /></button>
      </div>

      <div class="modal-body">
        <div class="form-group">
          <label for="renameInput">
            <span>新显示名称</span>
            <span class="field-hint">仅修改管理界面名称，源码中的过程名和代码哈希 100% 保持不变</span>
          </label>
          <input
            id="renameInput"
            type="text"
            bind:value={renameNewDisplayName}
            on:keydown={(e) => e.key === 'Enter' && handleConfirmRename()}
          />
        </div>
        {#if renameError}
          <div class="alert-box alert-error">
            <AlertCircle size={14} />
            <span>{renameError}</span>
          </div>
        {/if}
      </div>

      <div class="modal-footer">
        <button class="btn btn-primary" on:click={handleConfirmRename}>
          保存
        </button>
        <button class="btn btn-subtle" on:click={() => (renameModalScript = null)}>
          取消
        </button>
      </div>
    </div>
  </div>
{/if}

<!-- 覆盖同名确认弹窗 -->
{#if showOverwriteConfirm}
  <div class="modal-backdrop" role="presentation">
    <div class="modal-dialog sm-dialog" role="dialog" aria-modal="true" aria-label="同名覆盖确认">
      <div class="modal-header">
        <div class="modal-title">
          <AlertTriangle size={16} color="#B25900" />
          <span>已存在同名宏</span>
        </div>
        <button class="btn-icon" on:click={() => (showOverwriteConfirm = false)}><X size={14} /></button>
      </div>

      <div class="modal-body">
        <p>宏库中已存在名为<strong>【{importDisplayName}】</strong>的宏条目。</p>
        <p class="subtext">是否确认覆盖并更新该已有条目？覆盖后历史条目的源码将被替换。</p>
      </div>

      <div class="modal-footer">
        <button class="btn btn-danger" on:click={() => handleSaveMacro(true)}>
          确认覆盖
        </button>
        <button class="btn btn-subtle" on:click={() => (showOverwriteConfirm = false)}>
          取消并修改名称
        </button>
      </div>
    </div>
  </div>
{/if}

<!-- 导出无凭据宏包弹窗 (TASK-R6b-01) -->
{#if showExportPackageModal}
  <div class="modal-backdrop" role="presentation">
    <div class="modal-dialog modal-dialog-lg" role="dialog" aria-modal="true" aria-label="导出无凭据宏包">
      <div class="modal-header">
        <div class="modal-title">
          <Package size={16} color="#107C41" />
          <span>导出无凭据宏包 (.exmpack)</span>
        </div>
        <button class="btn-icon" on:click={() => (showExportPackageModal = false)}><X size={14} /></button>
      </div>

      <div class="modal-body modal-body-scroll">
        <div class="info-callout">
          <div class="info-title">🛡️ 元数据字段白名单与无凭据承诺</div>
          <div class="info-text">
            系统严格采用元数据白名单导出（仅包含显示名称、分类、描述、入口过程及参数契约）。<br/>
            <strong>绝不导出：</strong>API 配置、DPAPI 凭据、聊天记录、运行历史溯源、工作流运行记录、目标工作簿路径、业务数据、快照备份或本机环境信息。<br/>
            源码按原始二进制字节写入，manifest SHA-256 仅供完整性核验。未检出敏感内容不等于绝对安全。
          </div>
        </div>

        <div class="form-group">
          <label for="pkgNameInput">
            <span>宏包名称 <span class="required">*</span></span>
          </label>
          <input id="pkgNameInput" type="text" bind:value={exportPackageName} placeholder="例如：财务数据处理工具集" />
        </div>

        <div class="form-row">
          <div class="form-group flex-1">
            <label for="pkgVerInput">版本号</label>
            <input id="pkgVerInput" type="text" bind:value={exportPackageVersion} placeholder="1.0.0" />
          </div>
          <div class="form-group flex-2">
            <label for="pkgPathInput">导出保存路径</label>
            <div class="input-with-button">
              <input id="pkgPathInput" type="text" bind:value={exportPackageOutputPath} placeholder="留空默认保存至文档目录" />
              <button class="btn btn-subtle" type="button" on:click={handleBrowseExportSavePath}>浏览...</button>
            </div>
          </div>
        </div>

        <div class="form-group">
          <label for="pkgDescInput">宏包描述（可选）</label>
          <textarea id="pkgDescInput" rows={2} bind:value={exportPackageDescription} placeholder="说明此宏包所含宏的用途与使用场景..."></textarea>
        </div>

        <!-- 待导出宏清单多选 -->
        <div class="package-entries-block">
          <div class="entries-header">
            <span>选择待打包的宏</span>
            <span class="entries-count">已选 {exportPackageSelectedMacroIds.length} / {scripts.length}</span>
          </div>
          <div class="entries-list">
            {#each scripts as s}
              <label class="entry-item">
                <input
                  type="checkbox"
                  value={s.id || s.fileName}
                  checked={exportPackageSelectedMacroIds.includes(s.id || s.fileName)}
                  on:change={(e) => {
                    const sid = s.id || s.fileName;
                    if (e.currentTarget.checked) {
                      exportPackageSelectedMacroIds = [...exportPackageSelectedMacroIds, sid];
                    } else {
                      exportPackageSelectedMacroIds = exportPackageSelectedMacroIds.filter(id => id !== sid);
                    }
                  }}
                />
                <div class="entry-info">
                  <div class="entry-name-row">
                    <strong>{s.displayName || s.name}</strong>
                    <span class="entry-badge">{s.category || '通用'}</span>
                  </div>
                  <div class="entry-detail-row">
                    <span>入口: <code>{s.entryPoint || '无指定'}</code></span>
                    <span>哈希: {s.originalCodeHash ? s.originalCodeHash.substring(0, 10) + '...' : '未生成'}</span>
                  </div>
                </div>
              </label>
            {/each}
          </div>
        </div>

        <!-- 疑似敏感内容告警 -->
        {#if exportPackageWarnings && exportPackageWarnings.length > 0}
          <div class="alert-box alert-warning">
            <div class="alert-title">
              <AlertTriangle size={15} color="#B25900" />
              <strong>检出疑似敏感内容提醒</strong>
            </div>
            <p class="alert-sub">在待导出的宏中检测到以下疑似 API Key、密码或凭据信息。系统<strong>绝不自动篡改或打码源码正文</strong>，请用户核实是否包含私密凭据：</p>
            <ul class="warning-detail-list">
              {#each exportPackageWarnings as w}
                <li>• 【{w.macroDisplayName}】{w.fieldName}: <code>{w.snippet}</code></li>
              {/each}
            </ul>
          </div>
        {/if}

        {#if exportPackageErrorMessage}
          <div class="alert-box alert-error">
            <AlertCircle size={15} />
            <span>{exportPackageErrorMessage}</span>
          </div>
        {/if}

        {#if exportPackageSuccessResult}
          <div class="alert-box alert-success">
            <Check size={16} color="#107C41" />
            <span>✔ 宏包导出成功！已将 {exportPackageSuccessResult.exportedCount} 个宏打包至 <code>{exportPackageSuccessResult.outputPath || exportPackageSuccessResult.packageFilePath}</code> ({exportPackageSuccessResult.packageSize || exportPackageSuccessResult.totalBytes} 字节)。</span>
          </div>
        {/if}
      </div>

      <div class="modal-footer">
        {#if exportPackageWarnings && exportPackageWarnings.length > 0}
          <button
            class="btn btn-warning"
            disabled={exportPackageIsScanning}
            on:click={() => handleExecuteExportPackage(true)}
          >
            {exportPackageIsScanning ? '正在导出...' : '已人工核对无凭据，确认导出'}
          </button>
        {:else if !exportPackageSuccessResult}
          <button
            class="btn btn-primary"
            disabled={exportPackageIsScanning || exportPackageSelectedMacroIds.length === 0}
            on:click={() => handleExecuteExportPackage(false)}
          >
            <Package size={14} />
            <span>{exportPackageIsScanning ? '正在扫描并打包...' : `打包并导出 (${exportPackageSelectedMacroIds.length} 个宏)`}</span>
          </button>
        {/if}
        <button class="btn btn-subtle" on:click={() => (showExportPackageModal = false)}>
          {exportPackageSuccessResult ? '完成' : '取消'}
        </button>
      </div>
    </div>
  </div>
{/if}

<style>
  .drawer-backdrop {
    position: fixed;
    top: 0;
    left: 0;
    width: 100vw;
    height: 100vh;
    background: rgba(0, 0, 0, 0.42);
    backdrop-filter: blur(2px);
    z-index: 100;
    display: flex;
    justify-content: flex-end;
  }

  .drawer-content {
    width: 100%;
    height: 100vh;
    background: var(--office-card);
    box-shadow: -4px 0 24px rgba(0, 0, 0, 0.18);
    display: flex;
    flex-direction: column;
    animation: slideIn 0.22s cubic-bezier(0.16, 1, 0.3, 1);
  }

  @keyframes slideIn {
    from { transform: translateX(100%); }
    to { transform: translateX(0); }
  }

  .drawer-header {
    height: 52px;
    padding: 0 14px;
    border-bottom: 1px solid var(--office-border);
    display: flex;
    align-items: center;
    justify-content: space-between;
    flex-shrink: 0;
    background: #ffffff;
  }

  .header-left {
    display: flex;
    align-items: center;
    gap: 16px;
  }

  .header-title {
    display: flex;
    align-items: center;
    gap: 6px;
    font-size: var(--font-size-md);
    font-weight: 600;
    color: var(--office-text);
  }

  .header-tabs {
    display: flex;
    align-items: center;
    gap: 4px;
    background: var(--office-bg);
    padding: 3px;
    border-radius: var(--office-radius-sm);
  }

  .tab-btn {
    display: inline-flex;
    align-items: center;
    gap: 5px;
    padding: 4px 10px;
    font-size: var(--font-size-xs);
    font-weight: 500;
    color: var(--office-muted);
    background: transparent;
    border: none;
    border-radius: var(--office-radius-xs);
    cursor: pointer;
    transition: all 0.15s ease;
  }

  .tab-btn:hover {
    color: var(--office-text);
  }

  .tab-btn.active {
    color: var(--excel-green);
    background: #ffffff;
    font-weight: 600;
    box-shadow: 0 1px 3px rgba(0, 0, 0, 0.08);
  }

  .badge-count {
    background: var(--office-border-subtle);
    padding: 1px 5px;
    border-radius: var(--office-radius-full);
    font-size: 10px;
    color: var(--office-muted);
  }

  .drawer-body {
    flex: 1;
    overflow-y: auto;
    padding: 14px;
    background: var(--office-card-subtle);
  }

  /* 列表工具栏 */
  .list-toolbar {
    display: flex;
    flex-direction: column;
    gap: 8px;
    margin-bottom: 12px;
  }

  .search-box {
    display: flex;
    align-items: center;
    gap: 6px;
    background: #ffffff;
    border: 1px solid var(--office-border);
    border-radius: var(--office-radius);
    padding: 6px 10px;
    transition: border-color 0.15s;
  }

  .search-box:focus-within {
    border-color: var(--excel-green);
    box-shadow: 0 0 0 1px var(--excel-green);
  }

  .search-box input {
    flex: 1;
    border: none;
    outline: none;
    font-size: var(--font-size-sm);
    color: var(--office-text);
  }

  .btn-clear-search {
    border: none;
    background: transparent;
    cursor: pointer;
    color: var(--office-dim);
    padding: 2px;
  }

  .category-pills {
    display: flex;
    align-items: center;
    gap: 6px;
    overflow-x: auto;
    padding-bottom: 2px;
  }

  .pill-btn {
    border: 1px solid var(--office-border);
    background: #ffffff;
    color: var(--office-muted);
    font-size: var(--font-size-xs);
    padding: 2px 8px;
    border-radius: var(--office-radius-full);
    cursor: pointer;
    white-space: nowrap;
  }

  .pill-btn.active {
    background: var(--excel-light);
    border-color: var(--excel-light-border);
    color: var(--excel-dark);
    font-weight: 600;
  }

  .tags-filter-bar {
    display: flex;
    align-items: center;
    gap: 5px;
    overflow-x: auto;
    padding-bottom: 2px;
  }

  .tags-filter-label {
    display: inline-flex;
    align-items: center;
    gap: 3px;
    font-size: 11px;
    color: var(--office-dim);
    white-space: nowrap;
  }

  .tag-filter-pill {
    font-size: 11px;
    padding: 1px 7px;
  }

  /* 宏卡片列表 */
  .script-list {
    display: flex;
    flex-direction: column;
    gap: 10px;
  }

  .script-card {
    background: #ffffff;
    border: 1px solid var(--office-border);
    border-radius: var(--office-radius);
    overflow: hidden;
    transition: all 0.15s ease;
  }

  .script-card:hover {
    border-color: var(--office-border-strong);
    box-shadow: var(--office-shadow);
  }

  .card-header {
    padding: 10px 12px;
    display: flex;
    flex-direction: column;
    gap: 6px;
  }

  .card-title-row {
    display: flex;
    align-items: center;
    gap: 6px;
  }

  .macro-name {
    font-size: var(--font-size-sm);
    font-weight: 600;
    color: var(--office-text);
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
    line-height: 1.3;
  }

  .btn-icon-inline {
    border: none;
    background: transparent;
    color: var(--office-dim);
    cursor: pointer;
    padding: 2px;
    display: inline-flex;
    align-items: center;
    border-radius: var(--office-radius-xs);
  }

  .btn-icon-inline:hover {
    color: var(--excel-green);
    background: var(--excel-light);
  }

  .badge-row {
    display: flex;
    align-items: center;
    flex-wrap: wrap;
    gap: 5px;
  }

  .badge {
    display: inline-flex;
    align-items: center;
    gap: 3px;
    font-size: 10px;
    padding: 1px 6px;
    border-radius: var(--office-radius-full);
    line-height: 1.3;
  }

  .badge-category {
    background: var(--office-blue-light);
    color: var(--office-blue-dark);
    border: 1px solid var(--office-blue-border);
  }

  .badge-source {
    background: var(--office-card-subtle);
    color: var(--office-muted);
    border: 1px solid var(--office-border-subtle);
  }

  .badge-success {
    background: var(--excel-light);
    color: var(--excel-dark);
    border: 1px solid var(--excel-light-border);
  }

  .badge-danger {
    background: var(--office-danger-light);
    color: var(--office-danger);
    border: 1px solid var(--office-danger-border);
    max-width: 140px;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
  }

  .badge-neutral {
    background: var(--office-bg);
    color: var(--office-muted);
    border: 1px solid var(--office-border);
  }

  .macro-desc {
    font-size: var(--font-size-xs);
    color: var(--office-muted);
    line-height: var(--line-height-normal);
  }

  .card-meta {
    font-size: var(--font-size-xs);
    color: var(--office-dim);
    display: flex;
    align-items: center;
    gap: 5px;
    flex-wrap: wrap;
  }

  .meta-dot {
    color: #c8c6c4;
  }

  .card-actions {
    display: flex;
    align-items: center;
    gap: 6px;
    margin-top: 4px;
    padding-top: 6px;
    border-top: 1px solid var(--office-border-subtle);
  }

  .btn-danger-icon:hover {
    background: var(--office-danger-light);
  }

  .card-tags-row {
    display: flex;
    align-items: center;
    flex-wrap: wrap;
    gap: 4px;
    margin-top: 2px;
  }

  .script-tag-pill {
    display: inline-flex;
    align-items: center;
    font-size: 11px;
    padding: 1px 6px;
    border-radius: var(--office-radius-full);
    background: #e8f3ff;
    color: #0078d4;
    border: 1px solid #c7e0f4;
    cursor: pointer;
    transition: all 0.15s ease;
  }

  .script-tag-pill:hover {
    background: #c7e0f4;
  }

  .script-tag-pill.active {
    background: #0078d4;
    color: #ffffff;
    border-color: #0078d4;
    font-weight: 600;
  }

  /* 展开的详情面板 */
  .expanded-drawer-panel {
    border-top: 1px solid var(--office-border-subtle);
    background: var(--office-bg);
    display: flex;
    flex-direction: column;
  }

  .section-block {
    padding: 8px 12px;
    border-bottom: 1px solid var(--office-border-subtle);
    display: flex;
    flex-direction: column;
    gap: 6px;
  }

  .section-title {
    display: flex;
    align-items: center;
    gap: 5px;
    font-size: 11px;
    font-weight: 600;
    color: var(--office-text);
  }

  .sub-hint {
    font-size: 10px;
    color: var(--office-dim);
    font-weight: normal;
  }

  /* 标签编辑区 */
  .tags-editor {
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    gap: 6px;
  }

  .tag-chips {
    display: flex;
    flex-wrap: wrap;
    gap: 4px;
    align-items: center;
  }

  .tag-chip {
    display: inline-flex;
    align-items: center;
    gap: 3px;
    background: #ffffff;
    border: 1px solid var(--office-border);
    padding: 1px 6px;
    border-radius: var(--office-radius-full);
    font-size: 11px;
    color: var(--office-text);
  }

  .btn-tag-remove {
    border: none;
    background: transparent;
    cursor: pointer;
    padding: 1px;
    display: inline-flex;
    align-items: center;
    color: var(--office-dim);
    border-radius: 50%;
  }

  .btn-tag-remove:hover {
    background: var(--office-danger-light);
    color: var(--office-danger);
  }

  .empty-tag-hint {
    font-size: 11px;
    color: var(--office-dim);
  }

  .add-tag-box {
    display: inline-flex;
    align-items: center;
    gap: 4px;
  }

  .input-tag {
    width: 90px;
    font-size: 11px;
    padding: 2px 6px;
    border: 1px solid var(--office-border);
    border-radius: var(--office-radius-xs);
    outline: none;
    background: #ffffff;
  }

  .input-tag:focus {
    border-color: var(--excel-green);
  }

  .btn-xs {
    padding: 2px 6px;
    font-size: 11px;
  }

  .btn-add-tag {
    font-weight: 500;
  }

  .tag-error-text {
    font-size: 11px;
    color: var(--office-danger);
  }

  /* 运行历史溯源区 */
  .history-list {
    display: flex;
    flex-direction: column;
    gap: 6px;
    max-height: 200px;
    overflow-y: auto;
  }

  .history-item {
    background: #ffffff;
    border: 1px solid var(--office-border-subtle);
    border-left: 3px solid #8a8886;
    border-radius: var(--office-radius-xs);
    padding: 6px 8px;
    font-size: 11px;
    display: flex;
    flex-direction: column;
    gap: 3px;
  }

  .history-item.success {
    border-left-color: var(--excel-green);
  }

  .history-item.blocked {
    border-left-color: #d83b01;
  }

  .history-item.failed {
    border-left-color: var(--office-danger);
  }

  .history-header {
    display: flex;
    align-items: center;
    gap: 6px;
    flex-wrap: wrap;
  }

  .history-status-badge {
    padding: 0 4px;
    border-radius: var(--office-radius-xs);
    font-weight: 600;
    font-size: 10px;
  }

  .status-success {
    background: var(--excel-light);
    color: var(--excel-dark);
  }

  .status-blocked {
    background: #fdf3eb;
    color: #a80000;
  }

  .status-failed {
    background: var(--office-danger-light);
    color: var(--office-danger);
  }

  .history-phase {
    color: var(--office-muted);
    font-size: 10px;
  }

  .history-elapsed {
    color: var(--office-muted);
    font-size: 10px;
  }

  .history-time {
    margin-left: auto;
    color: var(--office-dim);
    font-size: 10px;
  }

  .history-details {
    display: flex;
    align-items: center;
    gap: 6px;
    color: var(--office-muted);
    font-size: 10px;
    flex-wrap: wrap;
  }

  .history-wb {
    font-weight: 500;
    color: var(--office-text);
  }

  .history-entry {
    background: var(--office-bg);
    padding: 0 4px;
    border-radius: 2px;
  }

  .history-hash {
    font-family: var(--font-family-code);
    color: var(--office-dim);
  }

  .history-snapshot-row {
    display: flex;
    align-items: center;
    gap: 4px;
    font-size: 10px;
  }

  .snapshot-badge {
    padding: 1px 4px;
    border-radius: 2px;
    display: inline-flex;
    align-items: center;
    gap: 3px;
  }

  .snapshot-valid {
    background: #f0f7f3;
    color: #107c41;
  }

  .snapshot-missing {
    background: #fff8e5;
    color: #795b00;
  }

  .snapshot-none {
    background: var(--office-card-subtle);
    color: var(--office-dim);
  }

  .history-summary {
    color: var(--office-text);
    background: var(--office-card-subtle);
    padding: 2px 6px;
    border-radius: 2px;
    font-size: 10px;
    line-height: 1.3;
    word-break: break-all;
  }

  .history-empty {
    padding: 8px 0;
    font-size: 11px;
    color: var(--office-dim);
  }

  /* 展开的源码查看区 */
  .code-container {
    background: #1e1e1e;
    color: #d4d4d4;
    border-top: 1px solid #333333;
    overflow: hidden;
  }

  .code-meta-bar {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 5px 10px;
    background: #252526;
    border-bottom: 1px solid #333333;
    font-size: 11px;
    color: #858585;
  }

  .code-hash {
    display: flex;
    align-items: center;
    gap: 4px;
    font-family: var(--font-family-code);
  }

  .code-toolbar {
    display: flex;
    align-items: center;
    gap: 8px;
  }

  .btn-text-sm {
    border: none;
    background: transparent;
    color: #cccccc;
    display: inline-flex;
    align-items: center;
    gap: 3px;
    font-size: 11px;
    cursor: pointer;
    padding: 2px 4px;
    border-radius: 3px;
  }

  .btn-text-sm:hover {
    background: #37373d;
    color: #ffffff;
  }

  .code-container pre {
    margin: 0;
    padding: 8px 12px;
    max-height: 220px;
    overflow-y: auto;
    font-family: var(--font-family-code);
    font-size: var(--font-size-xs);
    line-height: var(--line-height-normal);
    white-space: pre-wrap;
    word-break: break-all;
  }

  /* 导入宏视图 */
  .import-view {
    display: flex;
    flex-direction: column;
    gap: 12px;
  }

  .import-mode-toggle {
    display: flex;
    border: 1px solid var(--office-border);
    border-radius: var(--office-radius);
    background: #ffffff;
    padding: 3px;
    gap: 4px;
  }

  .mode-btn {
    flex: 1;
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 6px;
    padding: 6px 12px;
    font-size: var(--font-size-xs);
    font-weight: 500;
    border: none;
    background: transparent;
    color: var(--office-muted);
    border-radius: var(--office-radius-xs);
    cursor: pointer;
    transition: all 0.15s ease;
  }

  .mode-btn.active {
    background: var(--excel-light);
    color: var(--excel-dark);
    font-weight: 600;
  }

  .file-dropzone {
    border: 2px dashed var(--office-border-strong);
    background: #ffffff;
    border-radius: var(--office-radius);
    padding: 20px 14px;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    gap: 6px;
    cursor: pointer;
    transition: all 0.15s ease;
    text-align: center;
  }

  .file-dropzone:hover {
    border-color: var(--excel-green);
    background: var(--excel-light);
  }

  .dropzone-text {
    font-size: var(--font-size-sm);
    font-weight: 600;
    color: var(--office-text);
  }

  .dropzone-hint {
    font-size: var(--font-size-xs);
    color: var(--office-dim);
    max-width: 320px;
  }

  .file-info-badge {
    margin-top: 8px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    background: #ffffff;
    border: 1px solid var(--excel-light-border);
    padding: 6px 10px;
    border-radius: var(--office-radius);
  }

  .file-name {
    font-size: var(--font-size-xs);
    font-weight: 600;
    color: var(--excel-dark);
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
    max-width: 180px;
  }

  .encoding-selector {
    display: flex;
    align-items: center;
    gap: 4px;
    font-size: var(--font-size-xs);
  }

  .encoding-label {
    color: var(--office-muted);
  }

  .encoding-selector select {
    font-size: var(--font-size-xs);
    padding: 2px 4px;
    border: 1px solid var(--office-border);
    border-radius: var(--office-radius-xs);
    background: var(--office-card-subtle);
  }

  /* 表单控件 */
  .meta-form {
    background: #ffffff;
    border: 1px solid var(--office-border);
    border-radius: var(--office-radius);
    padding: 12px;
    display: flex;
    flex-direction: column;
    gap: 10px;
  }

  .form-group {
    display: flex;
    flex-direction: column;
    gap: 4px;
  }

  .form-group label {
    display: flex;
    align-items: center;
    justify-content: space-between;
    font-size: var(--font-size-xs);
    font-weight: 600;
    color: var(--office-text);
  }

  .required {
    color: var(--office-danger);
  }

  .field-hint {
    font-size: 10px;
    font-weight: normal;
    color: var(--office-dim);
  }

  .form-group input,
  .form-group select,
  .form-group textarea {
    border: 1px solid var(--office-border);
    border-radius: var(--office-radius-xs);
    padding: 6px 8px;
    font-size: var(--font-size-sm);
    color: var(--office-text);
    outline: none;
    transition: border-color 0.15s;
    font-family: inherit;
  }

  .form-group textarea {
    font-family: var(--font-family-code);
    font-size: var(--font-size-xs);
    line-height: var(--line-height-normal);
    resize: vertical;
  }

  .form-group input:focus,
  .form-group select:focus,
  .form-group textarea:focus {
    border-color: var(--excel-green);
  }

  .form-row {
    display: flex;
    gap: 10px;
  }

  .flex-1 {
    flex: 1;
  }

  .hash-preview {
    display: flex;
    align-items: center;
    gap: 4px;
    font-size: 10px;
    color: var(--office-muted);
    background: var(--office-card-subtle);
    padding: 4px 8px;
    border-radius: var(--office-radius-xs);
  }

  .hash-preview code {
    font-family: var(--font-family-code);
    color: var(--office-text);
  }

  .save-security-note {
    display: flex;
    align-items: flex-start;
    gap: 6px;
    background: var(--office-blue-light);
    border: 1px solid var(--office-blue-border);
    border-radius: var(--office-radius-xs);
    padding: 6px 8px;
    font-size: var(--font-size-xs);
    color: var(--office-blue-dark);
    line-height: var(--line-height-normal);
  }

  .form-actions {
    display: flex;
    align-items: center;
    gap: 8px;
    margin-top: 4px;
  }

  /* 过程列表预览面板 */
  .procedure-panel {
    background: #ffffff;
    border: 1px solid var(--office-border);
    border-radius: var(--office-radius);
    padding: 8px 12px;
    display: flex;
    flex-direction: column;
    gap: 6px;
  }

  .panel-header {
    font-size: var(--font-size-xs);
    font-weight: 600;
    color: var(--office-text);
  }

  .proc-list {
    display: flex;
    flex-direction: column;
    gap: 4px;
  }

  .proc-item {
    display: flex;
    align-items: center;
    justify-content: space-between;
    font-size: 11px;
    padding: 3px 6px;
    border-radius: var(--office-radius-xs);
  }

  .proc-runnable {
    background: var(--excel-light);
    color: var(--excel-dark);
  }

  .proc-non-runnable {
    background: var(--office-bg);
    color: var(--office-muted);
  }

  .proc-type {
    font-weight: 600;
    margin-right: 4px;
  }

  .proc-status {
    font-size: 10px;
  }

  /* 源码预览 */
  .code-preview-section {
    background: #ffffff;
    border: 1px solid var(--office-border);
    border-radius: var(--office-radius);
    overflow: hidden;
  }

  .preview-header {
    padding: 6px 10px;
    background: var(--office-card-subtle);
    border-bottom: 1px solid var(--office-border);
    font-size: var(--font-size-xs);
    font-weight: 600;
    color: var(--office-text);
  }

  .preview-code {
    margin: 0;
    padding: 8px 12px;
    max-height: 160px;
    overflow-y: auto;
    font-family: var(--font-family-code);
    font-size: var(--font-size-xs);
    line-height: var(--line-height-normal);
    background: #1e1e1e;
    color: #d4d4d4;
  }

  /* 提示框 */
  .alert-box {
    display: flex;
    align-items: center;
    gap: 6px;
    padding: 7px 10px;
    border-radius: var(--office-radius-xs);
    font-size: var(--font-size-xs);
    line-height: var(--line-height-normal);
  }

  .alert-error {
    background: var(--office-danger-light);
    color: var(--office-danger);
    border: 1px solid var(--office-danger-border);
  }

  .alert-success {
    background: var(--excel-light);
    color: var(--excel-dark);
    border: 1px solid var(--excel-light-border);
  }

  /* 弹窗通用样式 */
  .modal-backdrop {
    position: fixed;
    top: 0;
    left: 0;
    width: 100vw;
    height: 100vh;
    background: rgba(0, 0, 0, 0.45);
    backdrop-filter: blur(2px);
    z-index: 120;
    display: flex;
    align-items: center;
    justify-content: center;
    padding: 16px;
    box-sizing: border-box;
  }

  @media (max-width: 360px) {
    .modal-backdrop {
      padding: 6px;
    }
  }

  .modal-dialog {
    background: #ffffff;
    border-radius: var(--office-radius);
    box-shadow: 0 8px 30px rgba(0, 0, 0, 0.2);
    width: 100%;
    max-width: 420px;
    overflow: hidden;
    animation: popIn 0.18s ease;
    box-sizing: border-box;
  }

  .modal-dialog-lg {
    max-width: 480px;
    max-height: 90vh;
    display: flex;
    flex-direction: column;
  }

  .modal-body-scroll {
    overflow-y: auto;
    max-height: calc(90vh - 110px);
  }

  .alert-warning {
    background: #fff8e5;
    color: #795b00;
    border: 1px solid #ffdc7d;
  }

  /* 参数表单区域 (TASK-R2c-01) */
  .params-section {
    background: #f8f9fa;
    border: 1px solid var(--office-border);
    border-radius: var(--office-radius);
    padding: 10px;
    display: flex;
    flex-direction: column;
    gap: 8px;
  }

  .params-section-title {
    font-size: 11px;
    font-weight: 600;
    color: var(--office-text);
    border-bottom: 1px solid var(--office-border-subtle);
    padding-bottom: 4px;
  }

  .params-grid {
    display: flex;
    flex-direction: column;
    gap: 8px;
  }

  .param-card {
    background: #ffffff;
    border: 1px solid var(--office-border);
    border-radius: var(--office-radius-xs);
    padding: 8px;
    display: flex;
    flex-direction: column;
    gap: 4px;
  }

  .param-card.has-error {
    border-color: var(--office-danger);
    background: #fff9f9;
  }

  .param-header {
    display: flex;
    align-items: center;
    gap: 6px;
    font-size: 11px;
  }

  .param-name {
    font-weight: 600;
    color: var(--office-text);
    font-family: var(--font-family-code);
  }

  .param-badge-type {
    background: #eef4f9;
    color: #0068b7;
    padding: 1px 5px;
    border-radius: 2px;
    font-size: 10px;
    font-family: var(--font-family-code);
  }

  .param-badge-required {
    background: #fde7e9;
    color: #a80000;
    padding: 1px 4px;
    border-radius: 2px;
    font-size: 10px;
  }

  .param-badge-optional {
    background: var(--office-bg);
    color: var(--office-dim);
    padding: 1px 4px;
    border-radius: 2px;
    font-size: 10px;
  }

  .param-meta-desc {
    font-size: 11px;
    color: var(--office-muted);
  }

  .param-control {
    display: flex;
    flex-direction: column;
    gap: 3px;
  }

  .param-control .form-control {
    width: 100%;
    padding: 4px 6px;
    border: 1px solid var(--office-border);
    border-radius: var(--office-radius-xs);
    font-size: var(--font-size-xs);
    outline: none;
    box-sizing: border-box;
  }

  .param-control .form-control:focus {
    border-color: var(--excel-green);
  }

  .range-compound-group {
    display: flex;
    gap: 6px;
  }

  .range-sheet-part {
    flex: 1.2;
    display: flex;
    flex-direction: column;
    gap: 2px;
  }

  .range-addr-part {
    flex: 1;
    display: flex;
    flex-direction: column;
    gap: 2px;
  }

  .inner-label {
    font-size: 10px;
    color: var(--office-dim);
  }

  .param-error-text {
    font-size: 10px;
    color: var(--office-danger);
    font-weight: 500;
  }

  .btn-secondary {
    background: #eef4f9;
    color: #0068b7;
    border: 1px solid #b8d9f2;
  }

  .btn-secondary:hover:not(:disabled) {
    background: #dbeef9;
  }

  .sm-dialog {
    max-width: 360px;
  }

  @keyframes popIn {
    from { transform: scale(0.95); opacity: 0; }
    to { transform: scale(1); opacity: 1; }
  }

  .modal-header {
    height: 44px;
    padding: 0 14px;
    border-bottom: 1px solid var(--office-border);
    display: flex;
    align-items: center;
    justify-content: space-between;
  }

  .modal-title {
    display: flex;
    align-items: center;
    gap: 6px;
    font-size: var(--font-size-sm);
    font-weight: 600;
    color: var(--office-text);
  }

  .modal-body {
    padding: 14px;
    display: flex;
    flex-direction: column;
    gap: 10px;
  }

  .info-row {
    display: flex;
    flex-direction: column;
    gap: 2px;
    font-size: var(--font-size-xs);
  }

  .info-label {
    color: var(--office-muted);
    font-weight: 500;
  }

  .info-val {
    color: var(--office-text);
    font-weight: 600;
  }

  .info-val.highlight {
    color: var(--excel-green);
    font-size: var(--font-size-sm);
  }

  .info-val.subtext {
    font-weight: normal;
    color: var(--office-dim);
    font-family: var(--font-family-code);
    font-size: 11px;
    word-break: break-all;
  }

  .entry-select {
    padding: 4px 6px;
    border: 1px solid var(--office-border);
    border-radius: var(--office-radius-xs);
    font-size: var(--font-size-sm);
    outline: none;
  }

  .favorite-filter-bar {
    display: flex;
    align-items: center;
    gap: 5px;
    padding-bottom: 2px;
  }

  .fav-filter-pill {
    display: inline-flex;
    align-items: center;
    gap: 4px;
    font-size: 11px;
    padding: 1px 8px;
  }

  .fav-filter-pill.active {
    background: #fff4ce;
    border-color: #fce100;
    color: #795b00;
    font-weight: 600;
  }

  .badge-favorite {
    background: #fff4ce;
    color: #795b00;
    border: 1px solid #fce100;
    font-weight: 600;
  }

  .btn-fav-active {
    color: #ffaa00;
  }

  .source-view-toggle {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 6px 0 2px 0;
    border-top: 1px solid var(--office-border-subtle);
  }

  .code-hash-mini {
    font-family: var(--font-family-code);
    font-size: 10px;
    color: var(--office-dim);
  }

  .modal-code-preview {
    background: #1e1e1e;
    color: #d4d4d4;
    padding: 6px 8px;
    border-radius: var(--office-radius-xs);
    max-height: 120px;
    overflow-y: auto;
  }

  .modal-code-preview pre {
    margin: 0;
    font-family: var(--font-family-code);
    font-size: 11px;
    white-space: pre-wrap;
    word-break: break-all;
  }

  .security-banner {
    display: flex;
    align-items: flex-start;
    gap: 6px;
    background: var(--office-blue-light);
    border: 1px solid var(--office-blue-border);
    border-radius: var(--office-radius-xs);
    padding: 8px;
    font-size: var(--font-size-xs);
    color: var(--office-blue-dark);
    line-height: var(--line-height-normal);
  }

  .banner-text span {
    display: block;
    margin-top: 2px;
  }

  .banner-text p {
    margin: 2px 0;
    font-size: 11px;
    line-height: 1.35;
  }

  .modal-footer {
    padding: 10px 14px;
    border-top: 1px solid var(--office-border);
    background: var(--office-card-subtle);
    display: flex;
    justify-content: flex-end;
    gap: 8px;
  }

  /* 基础按钮与空状态 */
  .empty-state {
    padding: 40px 16px;
    text-align: center;
    color: var(--office-muted);
    display: flex;
    flex-direction: column;
    align-items: center;
    gap: 8px;
  }

  .empty-title {
    font-size: var(--font-size-sm);
    font-weight: 600;
    color: var(--office-text);
    margin: 0;
  }

  .empty-hint {
    font-size: var(--font-size-xs);
    color: var(--office-dim);
    max-width: 260px;
    line-height: var(--line-height-normal);
  }

  .btn {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    padding: 6px 12px;
    border-radius: var(--office-radius-xs);
    font-size: var(--font-size-sm);
    font-weight: 600;
    cursor: pointer;
    border: none;
    transition: all 0.15s ease;
  }

  .btn-sm {
    padding: 4px 8px;
    font-size: var(--font-size-xs);
  }

  .btn-primary {
    background: var(--excel-green);
    color: #ffffff;
  }

  .btn-primary:hover:not(:disabled) {
    background: var(--excel-hover-bg);
  }

  .btn-primary:disabled {
    opacity: 0.55;
    cursor: not-allowed;
  }

  .btn-subtle {
    background: var(--office-bg);
    color: var(--office-text);
    border: 1px solid var(--office-border);
  }

  .btn-subtle:hover {
    background: var(--office-hover);
  }

  .btn-danger {
    background: var(--office-danger);
    color: #ffffff;
  }

  .btn-danger:hover {
    background: var(--office-danger-dark);
  }

  .btn-icon {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    border: none;
    background: transparent;
    color: var(--office-muted);
    cursor: pointer;
    padding: 4px;
    border-radius: var(--office-radius-xs);
  }

  .btn-icon:hover {
    color: var(--office-text);
    background: var(--office-hover);
  }

  .btn-sm-icon {
    padding: 3px;
  }

  .spin {
    animation: spin 1.2s linear infinite;
  }

  @keyframes spin {
    from { transform: rotate(0deg); }
    to { transform: rotate(360deg); }
  }

  /* 宏包导入与导出样式 (TASK-R6b-01) */
  .package-export-bar {
    margin-left: auto;
  }

  .pkg-export-pill {
    background: #edf7ed;
    color: #107C41;
    border: 1px solid #c8e6c9;
    font-weight: 500;
  }

  .pkg-export-pill:hover {
    background: #e8f5e9;
    border-color: #81c784;
  }

  .input-with-button {
    display: flex;
    gap: 6px;
  }

  .input-with-button input {
    flex: 1;
  }

  .info-callout {
    background: #f0f7ff;
    border: 1px solid #bae0ff;
    border-radius: var(--office-radius-sm);
    padding: 10px 12px;
    margin-bottom: 12px;
    font-size: 12px;
    line-height: 1.5;
  }

  .info-title {
    font-weight: 600;
    color: #0050b3;
    margin-bottom: 4px;
  }

  .info-text {
    color: #262626;
  }

  .manifest-card {
    background: #fcfcfc;
    border: 1px solid var(--office-border);
    border-radius: var(--office-radius-sm);
    padding: 12px;
    margin-bottom: 12px;
  }

  .manifest-header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    margin-bottom: 8px;
  }

  .manifest-title {
    display: flex;
    align-items: center;
    gap: 6px;
    font-size: 14px;
  }

  .pkg-version-badge {
    background: #e6f7ff;
    color: #0050b3;
    border: 1px solid #91d5ff;
    border-radius: 10px;
    padding: 1px 6px;
    font-size: 11px;
    font-weight: 500;
  }

  .manifest-meta-grid {
    display: grid;
    grid-template-columns: repeat(2, 1fr);
    gap: 6px 12px;
    font-size: 12px;
    color: var(--office-text);
  }

  .meta-label {
    color: var(--office-muted);
  }

  .manifest-desc {
    margin-top: 8px;
    font-size: 12px;
    color: var(--office-muted);
    border-top: 1px dashed var(--office-border);
    padding-top: 6px;
  }

  .package-entries-block {
    border: 1px solid var(--office-border);
    border-radius: var(--office-radius-sm);
    margin-bottom: 12px;
    overflow: hidden;
  }

  .entries-header {
    background: var(--office-bg);
    padding: 8px 12px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    font-size: 12px;
    font-weight: 600;
    border-bottom: 1px solid var(--office-border);
  }

  .entries-count {
    color: var(--office-muted);
    font-weight: normal;
  }

  .entries-list {
    max-height: 240px;
    overflow-y: auto;
    padding: 4px 0;
  }

  .entry-item {
    display: flex;
    align-items: flex-start;
    gap: 8px;
    padding: 8px 12px;
    cursor: pointer;
    border-bottom: 1px solid #f0f0f0;
  }

  .entry-item:last-child {
    border-bottom: none;
  }

  .entry-item:hover {
    background: var(--office-hover);
  }

  .entry-item input[type="checkbox"] {
    margin-top: 3px;
    cursor: pointer;
  }

  .entry-info {
    flex: 1;
    display: flex;
    flex-direction: column;
    gap: 2px;
    font-size: 12px;
  }

  .entry-name-row {
    display: flex;
    align-items: center;
    gap: 8px;
  }

  .entry-badge {
    background: #f0f0f0;
    color: #595959;
    padding: 0 4px;
    border-radius: 2px;
    font-size: 10px;
  }

  .entry-size {
    color: var(--office-muted);
    font-size: 11px;
    margin-left: auto;
  }

  .entry-detail-row {
    display: flex;
    align-items: center;
    gap: 12px;
    color: var(--office-muted);
    font-size: 11px;
  }

  .entry-desc {
    font-size: 11px;
    color: var(--office-muted);
    margin-top: 2px;
  }

  .alert-warning {
    background: #fffbe6;
    border: 1px solid #ffe58f;
    color: #ad4e00;
  }

  .alert-info {
    background: #e6f7ff;
    border: 1px solid #91d5ff;
    color: #0050b3;
  }

  .warning-detail-list {
    margin: 6px 0 0 0;
    padding-left: 14px;
    font-size: 11px;
    line-height: 1.4;
  }

  .warning-detail-list li {
    margin-bottom: 3px;
  }

  .import-result-card {
    background: #f6ffed;
    border: 1px solid #b7eb8f;
    border-radius: var(--office-radius-sm);
    padding: 16px;
    text-align: left;
  }

  .result-header {
    display: flex;
    align-items: center;
    gap: 8px;
    font-size: 15px;
    color: #237804;
    margin-bottom: 8px;
  }

  .result-summary {
    font-size: 13px;
    color: #262626;
    margin-bottom: 10px;
  }

  .renamed-box {
    background: #ffffff;
    border: 1px solid #d9d9d9;
    border-radius: 4px;
    padding: 8px 12px;
    margin-bottom: 10px;
    font-size: 12px;
  }

  .renamed-title {
    font-weight: 600;
    color: #595959;
    display: block;
    margin-bottom: 4px;
  }

  .zero-execution-notice {
    display: flex;
    align-items: center;
    gap: 6px;
    font-size: 12px;
    color: #107C41;
    margin-bottom: 14px;
  }

  .result-actions {
    display: flex;
    gap: 8px;
  }
</style>
