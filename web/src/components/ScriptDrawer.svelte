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
  } from 'lucide-svelte';
  import { bridge, type ScriptItem, type WorkbookInfo } from '../services/bridge';

  export let isOpen = false;
  export let workbook: WorkbookInfo | null = null;
  export let onClose: () => void;
  export let onRunScript: (script: ScriptItem, entryPoint?: string) => void;

  // 主视图状态
  let activeTab: 'list' | 'import' = 'list';

  // 宏列表状态
  let scripts: ScriptItem[] = [];
  let isLoading = false;
  let expandedIndex: number | null = null;
  let searchTerm = '';
  let selectedCategory = '全部';
  let copiedScriptId: string | null = null;

  // 运行前入口与目标确认弹窗状态
  let runModalScript: ScriptItem | null = null;
  let runModalSelectedEntryPoint = '';
  let runModalCandidateEntryPoints: VbaProcedure[] = [];

  // 重命名弹窗状态
  let renameModalScript: ScriptItem | null = null;
  let renameNewDisplayName = '';
  let renameError = '';

  // 导入宏表单状态
  let importSourceType: 'file' | 'paste' = 'file';
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

  // 覆盖确认弹窗
  let showOverwriteConfirm = false;
  let pendingOverwritePayload: any = null;

  export function openTab(tab: 'list' | 'import') {
    activeTab = tab;
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

  // 过滤后的宏列表
  $: filteredScripts = scripts.filter((s) => {
    const matchSearch =
      !searchTerm ||
      (s.displayName || s.name || '').toLowerCase().includes(searchTerm.toLowerCase()) ||
      (s.description || '').toLowerCase().includes(searchTerm.toLowerCase()) ||
      (s.category || '').toLowerCase().includes(searchTerm.toLowerCase()) ||
      (s.code || '').toLowerCase().includes(searchTerm.toLowerCase());

    const matchCategory =
      selectedCategory === '全部' ||
      (s.category || '未分类') === selectedCategory;

    return matchSearch && matchCategory;
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

  // 运行前入口确认流程
  function handleOpenRunModal(script: ScriptItem) {
    runModalScript = script;
    const procs = parseVbaProcedures(script.code);
    runModalCandidateEntryPoints = procs.filter((p) => p.isRunnable);

    if (runModalCandidateEntryPoints.length > 0) {
      if (script.entryPoint && runModalCandidateEntryPoints.some((p) => p.name === script.entryPoint)) {
        runModalSelectedEntryPoint = script.entryPoint;
      } else {
        const mainP = runModalCandidateEntryPoints.find((p) => p.name.toLowerCase() === 'main');
        runModalSelectedEntryPoint = mainP ? mainP.name : runModalCandidateEntryPoints[0].name;
      }
    } else {
      runModalSelectedEntryPoint = '';
    }
  }

  function handleConfirmRun() {
    if (!runModalScript) return;
    const s = runModalScript;
    const ep = runModalSelectedEntryPoint;
    runModalScript = null;
    onClose();
    onRunScript(s, ep);
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
              <p class="empty-title">{searchTerm ? '未找到匹配的宏' : '暂无已保存的宏'}</p>
              <span class="empty-hint">
                {searchTerm ? '尝试更换搜索词或清空筛选条件' : '点击上方【导入宏】按钮，可从本地 .bas/.vba 文件或直接粘贴源码导入。'}
              </span>
              {#if !searchTerm}
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
                      {#if s.category}
                        <span class="badge badge-category">
                          <Tag size={10} />
                          {s.category}
                        </span>
                      {/if}
                      <span class="badge badge-source">
                        {s.sourceType === 'file' ? '本地导入' : (s.sourceType === 'legacy' ? '历史脚本' : '手工粘贴')}
                      </span>
                      {#if s.lastExecutionResult}
                        <span class="badge {s.lastExecutionResult.includes('成功') ? 'badge-success' : (s.lastExecutionResult === '未运行' ? 'badge-neutral' : 'badge-danger')}" title={s.lastExecutionResult}>
                          {s.lastExecutionResult}
                        </span>
                      {/if}
                    </div>

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
                        title={expandedIndex === idx ? '收起源码' : '查看完整源码与哈希'}
                      >
                        <FileText size={12} />
                        <span>{expandedIndex === idx ? '收起' : '源码'}</span>
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

                  <!-- 展开查看完整源码 -->
                  {#if expandedIndex === idx}
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
                            title="导出为 .bas 文件"
                          >
                            <Download size={12} />
                            <span>导出</span>
                          </button>
                        </div>
                      </div>
                      <pre><code>{s.code}</code></pre>
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
                <span>从本地文件导入</span>
              </button>
              <button
                class="mode-btn {importSourceType === 'paste' ? 'active' : ''}"
                on:click={() => (importSourceType = 'paste')}
              >
                <FileText size={14} />
                <span>粘贴 VBA 源码</span>
              </button>
            </div>

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
          </div>
        {/if}
      </div>
    </div>
  </div>
{/if}

<!-- 运行前目标工作簿与入口确认弹窗 -->
{#if runModalScript}
  <div class="modal-backdrop" role="presentation">
    <div class="modal-dialog" role="dialog" aria-modal="true" aria-label="确认运行宏">
      <div class="modal-header">
        <div class="modal-title">
          <Play size={16} color="#107C41" />
          <span>确认运行宏: 【{runModalScript.displayName || runModalScript.name}】</span>
        </div>
        <button class="btn-icon" on:click={() => (runModalScript = null)}><X size={14} /></button>
      </div>

      <div class="modal-body">
        <div class="info-row">
          <span class="info-label">目标工作簿:</span>
          <span class="info-val highlight">{workbook?.name || '未检测到活动工作簿'}</span>
        </div>
        {#if workbook?.fullName}
          <div class="info-row">
            <span class="info-label">磁盘路径:</span>
            <span class="info-val subtext" title={workbook.fullName}>{workbook.fullName}</span>
          </div>
        {/if}

        <div class="info-row">
          <span class="info-label">执行入口过程:</span>
          {#if runModalCandidateEntryPoints.length > 1}
            <select bind:value={runModalSelectedEntryPoint} class="entry-select">
              {#each runModalCandidateEntryPoints as p}
                <option value={p.name}>{p.name} ({p.paramType === 'workbook' ? '接收工作簿参数' : '无参Sub'})</option>
              {/each}
            </select>
          {:else if runModalCandidateEntryPoints.length === 1}
            <span class="info-val">{runModalCandidateEntryPoints[0].name} ({runModalCandidateEntryPoints[0].paramType === 'workbook' ? '工作簿参数入口' : '无参入口'})</span>
          {:else}
            <span class="info-val error-text">未检测到标准入口，将由宿主进行安全前置校验</span>
          {/if}
        </div>

        <div class="security-banner">
          <AlertCircle size={14} color="#0078D4" />
          <div class="banner-text">
            <strong>安全保障机制：</strong>
            <span>每次运行前将<strong>自动创建整本物理副本快照</strong>并提供一键回滚入口。无需 API Key，不请求大模型。导入宏的作用范围由代码自身决定。</span>
          </div>
        </div>
      </div>

      <div class="modal-footer">
        <button class="btn btn-primary" on:click={handleConfirmRun}>
          <Play size={13} />
          <span>立即执行</span>
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
  }

  .modal-dialog {
    background: #ffffff;
    border-radius: var(--office-radius);
    box-shadow: 0 8px 30px rgba(0, 0, 0, 0.2);
    width: 100%;
    max-width: 420px;
    overflow: hidden;
    animation: popIn 0.18s ease;
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
</style>
