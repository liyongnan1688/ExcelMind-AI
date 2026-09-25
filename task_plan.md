# 任务规划 (Task Plan): Excel AI 助手 (Lee-Excel) 原生插件

## 核心目标
构建一个单进程、原生集成的 Windows 桌面 Excel AI 插件：
1. **Office 经典原生风 UI**：融入 Excel 经典绿（`#107C41`），界面紧凑、低干扰，运行于 Custom Task Pane (CTP) 内嵌的 Edge WebView2 中。
2. **大模型 API 灵活配置**：复用参考项目（`office-agents`）成熟的供应商体系，支持 DeepSeek、通义千问、智谱、Ollama、自定义 Base URL 与 API Key。
3. **真实 VBA 动态注入与运行**：大模型生成真正的 `Sub ... End Sub`，通过 Excel COM (VBIDE) 在独立隔离模块临时注入并调用，运行后瞬时清理，不破坏目标工作簿格式（保持 `.xlsx` 纯净）。
4. **多版本快照时间轴整本恢复**：运行前调用原生 `SaveCopyAs` 自动进行全文件物理备份，在面板提供时间轴历史下拉框，支持真正意义上的整本工作簿一键回滚。
5. **本地独立脚本库（我的脚本）**：支持将本次实际运行的代码命名保存到本地独立脚本目录（`%AppData%\LeeExcel\Scripts\*.bas`），可随时再次浏览与执行。
6. **始终全自动执行与折叠简报**：模型输出代码后自动备份并执行，默认折叠执行过程，仅汇报 1~2 句话真实结果；提供快捷展开，查看高亮 VBA 源码、错误堆栈和时间轴回滚。

---

## 阶段规划

### Phase 1: 环境与工程脚手架构建
- **Status:** complete
- **Tasks:**
  - [x] 探测并锁定工具链（零外部 SDK 门槛，直接使用 Windows 系统内置 .NET Framework 4.8 `csc.exe` + `nuget.exe` 拉取轻量依赖）
  - [x] 初始化 `lee-excle` 工程结构（C# 原生加载项项目 + WebView2 前端构建项目）
  - [x] 配置 Excel-DNA (64位) 模板与依赖库

### Phase 2: 核心底层能力实现 (C# COM 引擎)
- **Status:** complete
- **Tasks:**
  - [x] 实现 `SnapshotManager`：多版本物理快照创建（`SaveCopyAs`）、快照元数据管理（`snapshots.json`）、整本静默回滚（`Close(false)` + 覆盖重载）
  - [x] 实现 `VbaRunner`：VBIDE 动态注入临时模块、`app.Run` 执行、精确捕获 COM 1004 错误与异常、瞬时销毁清理模块
  - [x] 实现 `ScriptManager`：本地磁盘 `%AppData%\LeeExcel\Scripts\*.bas` 的扫描、写入（含结构化注释头）、读取与删除
  - [x] 通过独立单元测试 `test_core.ps1` 验证通过

### Phase 3: 任务窗格与 Edge WebView2 进程内集成
- **Status:** complete
- **Tasks:**
  - [x] 创建 Excel 原生自定义任务窗格 (Custom Task Pane)
  - [x] 挂载 `Microsoft.Web.WebView2.WinForms.WebView2` 控件
  - [x] 修复 `Assembly.Location` 空值与 VirtualHost 映射，彻底解决本地资源加载与 CORS 问题
  - [x] 强化 Excel 生命周期事件（Activate/Open/New/WindowActivate）

### Phase 4: Office 经典原生风前端 UI 移植与重构
- **Status:** complete
- **Tasks:**
  - [x] 实现 `office-fluent.css` 设计系统
  - [x] 移植并实现多模型 API 配置面板 (`SettingsModal.svelte` + `config.ts`)
  - [x] 实现全自动执行折叠卡片 (`ExecutionCard.svelte`：1~2 句话简报、展开代码/日志、多版本时间轴快照回滚)
  - [x] 实现“我的脚本”管理抽屉 (`ScriptDrawer.svelte`)
  - [x] 优化问答逻辑：纯文本回复时不触发红色错误卡片；增加工作簿状态定时轮询
  - [x] 完成 Vite 构建并输出到 `bin/dist/`

### Phase 5: 综合联调与最小 PoC 闭环验证
- **Status:** complete
- **Tasks:**
  - [x] 创建一键启动脚本 `run_excel.ps1`
  - [x] 验证真实桌面加载成功，任务窗格顺利渲染
  - [x] 修复非代码提问时的打断体验与工作簿状态更新

## Next Step
引导用户再次测试真实 Excel 操作指令（如“在A1到D10生成测试销售数据并设置表格边框”），验证全自动执行、代码折叠与时间轴快照回滚全闭环！
