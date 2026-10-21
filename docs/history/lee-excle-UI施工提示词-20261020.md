# ExcelMind AI · 界面优化施工提示词（发给 Antigravity IDE）

**用法**：先发【0 总纲】一次，然后按 P0 → P6 逐条发。每条做完，先让它跑验收命令、贴出结果，再发下一条。不要一次发全部——分批才能定位是哪一步把测试弄红了。

**基线事实（我本地实测）**：`node test_suite_unit.cjs` → **Pass = 233, Fail = 0**。任何一步之后如果不是 233/0，就是这一步引入的。

---

## 【0 总纲 · 每次开工前先发这一段】

```text
你是 ExcelMind AI（Excel 加载项，仓库根目录就是当前工作区）的开发者。接下来按阶段做界面优化，
每个阶段我会单独发一条指令。请遵守以下项目纪律，除非我在某条指令里明确豁免：

【技术栈与命令】
- 前端：Svelte 5 + Vite，源码在 web/src；后端宿主：C# + Excel-DNA，源码在 src/
- 单元测试：在仓库根目录执行 `node test_suite_unit.cjs`，当前基线是 Pass = 233, Fail = 0
- 前端构建：`cd web && npx vite build`（产物输出到 ../bin/dist）

【铁律】
1. 只改我点名的文件/位置。任何越界文件，必须在交付说明里单独列出并给出理由。
2. 不改业务逻辑与契约：不改 bridge 的 action 名、不改字段名、不改状态值（'matched' / 'unconfirmed' /
   'failed' 这类枚举值），只改"显示层"。
3. 不新增第三方依赖。已有 lucide-svelte 图标库，图标只用它。
4. 不动 release/ 目录、不动 bin/ 下的编译产物、不重新打包 zip。
5. 不写"已完成""100%""全面覆盖""彻底解决"这类词。只报：改了哪些文件、每处改了什么、测试结果数字、
   哪些没做。做不到就直说做不到。
6. 文案是契约的一部分：test_suite_unit.cjs 里有对界面/服务文案的断言（例如 `身份核验失败`、
   `隔离工作副本`、`遇错即停`、`TASK-R5a`）。改任何界面文案前，先 grep 测试文件；如果断言命中，
   这属于"文案契约变更"，必须同步更新断言字符串并在说明里写明"改了哪条断言、为什么"，禁止直接删除断言。
7. 视觉规则（从本阶段起新写的样式必须遵守）：
   - 颜色只用 web/src/styles/office-fluent.css 里已有的 `--office-*` / `--excel-*` token，不写裸 #hex
   - 字号只用 12 / 13 / 16px；11px 仅限元信息（时间、计数）；禁止 10px 及以下
   - 圆角只用 4px / 6px / 999px
   - 不新增 emoji 字符做图标
8. 界面语言是中文。主界面文案（按钮、标题、字段名、状态、结果卡首行）一句话不超过 18 字，
   只写"用户能做的动作"和"用户能看到的结果"；机制、边界、参数解释放到展开区或 tooltip 里。

【每条指令的交付格式】
1) 改动清单：文件路径 + 每处改了什么（一行一条）
2) 越界说明（如果有）
3) 验收结果：把 `node test_suite_unit.cjs` 的结尾汇总行原样贴出来
4) 没做的事 / 有疑问的地方
```

---

## 【P0 · 功能区外科手术 + 三处露馅清理】

```text
本阶段目标：修掉 Excel 功能区第一眼就能看出来的问题，并清掉界面里 3 处"内部痕迹"。
本阶段不改任何交互逻辑，不改窗格布局。

一、重写功能区定义（src/LeeExcelAddIn.cs，GetCustomUI 方法里的 XML，约 32-103 行）

现状问题：tab 名和组名都叫 'ExcelMind AI'；组内唯一的大按钮 btnAssistant 没有 label 属性，
在 Excel 里渲染成"图标 + 空白标签区"；除它之外的 7 处图标全用 imageMso 老图标。

目标结构（一个 tab，4 组，8 个控件全部平铺，不用任何下拉壳包住固定功能）：

<tab id='LeeExcelTab' label='AI 助手'>
  <group id='GroupAssistant' label='助手'>
    <button id='btnAssistant' size='large' getLabel='GetAssistantLabel' getImage='GetIcon'
            keytip='W' onAction='OnTogglePane'
            screentip='AI 助手' supertip='用中文描述，AI 直接改表' />
  </group>
  <group id='GroupMacro' label='宏'>
    <dynamicMenu id='menuFavoriteMacros' size='normal' label='收藏宏' getImage='GetIcon'
                 getContent='GetFavoriteMacrosContent' keytip='F'
                 screentip='收藏宏' supertip='一键再跑收藏的宏' />
    <button id='btnMacroLibrary' size='normal' label='宏库' getImage='GetIcon' keytip='L'
            onAction='OnOpenMacroLibrary' screentip='宏库' supertip='查看、编辑、导入导出宏' />
  </group>
  <group id='GroupTasks' label='工具与任务'>
    <button id='btnDataTools' size='normal' label='工具' getImage='GetIcon' keytip='T'
            onAction='OnOpenDataTools' screentip='数据工具' supertip='去重、对账、汇总、图表、导入' />
    <button id='btnBatchModal' size='normal' label='批量处理' getImage='GetIcon' keytip='B'
            onAction='OnOpenBatch' screentip='批量处理' supertip='多个工作簿跑同一个宏' />
    <button id='btnWorkflowModal' size='normal' label='流水线' getImage='GetIcon' keytip='P'
            onAction='OnOpenWorkflow' screentip='流水线' supertip='两步串起来自动跑' />
  </group>
  <group id='GroupSettings' label='设置'>
    <button id='btnApiSettings' size='normal' label='设置' getImage='GetIcon' keytip='S'
            onAction='OnOpenSettings' screentip='设置' supertip='模型服务与参数' />
  </group>
</tab>

要求：
1. 原来的 <menu id='menuTasks'> 整个删掉，里面的 btnBatchModal / btnWorkflowModal 提升为顶层按钮。
2. 删掉原来的 '诊断' 需求：本阶段不要新增 btnDiagnostics（C# 侧没有现成的诊断入口），保持 7 个控件即可。
3. 所有 supertip 一律改成上面给的短句（不超过 20 字），删掉现在那种"查看、搜索、运行、重命名、导出和管理
   所有已保存的 VBA 宏与自动化脚本"式的长句。
4. 每个按钮都要有 label，不允许出现只有 getImage 没有 label 的按钮。

二、动态 label：新增 GetAssistantLabel(IRibbonControl control)
返回 task pane 当前可见状态对应的文字：可见 → "收起工作台"，不可见 → "打开工作台"。
在 OnTogglePane 切换可见性之后调用已有的 InvalidateRibbon()（或对该控件单独 InvalidateControl）刷新。

三、图标：把 imageMso 全部换成自绘图标
1. 在 src/BrandIconHelper.cs 里新增：public static Bitmap GetSmallIcon(string controlId)，返回 16x16 位图。
   沿用该类已有的绘制风格（#107C41→#094B26 渐变圆角底 + 白色线性符号），按 controlId 分派不同符号：
   - btnMacroLibrary：代码块 `</>`
   - menuFavoriteMacros：五角星
   - btnDataTools：扳手
   - btnBatchModal：三张叠放的矩形
   - btnWorkflowModal：两个节点 + 一条连线
   - btnApiSettings：齿轮
   每个符号都用 Graphics 基本图元画（线、矩形、圆），不要引入外部图片资源，不要用 emoji 字体。
   做好缓存（静态字段，按 id 存），避免每次重绘。
2. 把 LeeExcelAddIn.cs 里的 GetButtonImage 改造为统一回调 GetIcon(IRibbonControl control)：
   control.Id == "btnAssistant" 时返回现有 32px 品牌图（GetExcelMindIcon(32)），其余返回 GetSmallIcon(control.Id)。
   保留 try/catch，异常时返回 null。
3. 原 GetFavoriteMacrosContent 里动态生成的菜单项也有 3 处 imageMso（Info / MacroPlay / VisualBasic），
   一并换成 GetSmallIcon 对应的自绘图标（可以新增 fav_empty / fav_run / fav_lib 三个 id）。

四、清掉界面里的内部痕迹（共 5 处）
1. web/src/components/BatchModal.svelte 约 498 行：标题 "批量宏处理面板 (TASK-R4b)" → "批量处理"
2. web/src/components/WorkflowModal.svelte 约 239 行："双步骤任务流水线 (TASK-R5a)" → "任务流水线"
3. web/src/components/DataToolModal.svelte 约 1770 行："快捷图表工具定位与安全规范 (TASK-R5b-01)" → 删掉括号里的编号
4. web/src/components/BatchModal.svelte 约 638 行：placeholder 里是字面 `\n`（HTML 属性不解析转义），
   用户看到反斜杠 n。改成真正的多行占位：用 Svelte 表达式传字符串，例如
   placeholder={"Sub ProcessCurrent()\n    ActiveSheet.Range(\"A1\").Value = \"Processed\"\nEnd Sub"}
5. 全仓 grep `TASK-R` 与 `TASK-PHASE`，只在 .svelte 的"用户可见文本"里清理；代码注释里的任务编号保留。

【禁止】
- 不改窗格布局、不删 ⋯ 菜单、不改分段导航（那是下一阶段）
- 不改任何 bridge action 名、不改 C# 里 OnOpenXxx 的行为
- 不新增 tab、不新增按钮（除了上面点名的）
- 不动 ribbon 之外的 C# 逻辑

【验收】
1. `grep -rn "imageMso" src/LeeExcelAddIn.cs` 结果只剩注释（应为 0 处实际使用）
2. `grep -rn "TASK-R" web/src/components/*.svelte` 结果为空
3. `grep -rn "btnAssistant" -A 6 src/LeeExcelAddIn.cs` 能看到 label 或 getLabel 属性
4. `node test_suite_unit.cjs` → Pass = 233, Fail = 0（注意：测试里有对 `TASK-R5a` 的断言，
   如果因为它红了，说明断言写的是界面标题——这种情况把它改成新标题 `任务流水线`，并在说明里写明）
5. `cd web && npx vite build` 成功
```

---

## 【P1 · 反馈层统一：干掉 10 处原生弹窗】

```text
本阶段目标：把界面里所有 window.alert / window.confirm 换成应用内组件。这是后面所有改动的地基，
所以单独一步做，不和布局混在一起。

一、新建两个通用组件（放在 web/src/components/，样式必须用 office-fluent.css 的 token）
1. Toast：`web/src/components/Toast.svelte` + `web/src/services/toast.ts`
   - toast.ts 导出 `toast.success(msg)` / `toast.error(msg)` / `toast.info(msg)`
   - 单例，最多同时显示 3 条，3.5 秒自动消失，右下角堆叠（在 440px 窗格里改成底部居中，宽度自适应）
   - 不用 emoji，用 lucide 的 CheckCircle2 / XCircle / Info，尺寸 14
2. ConfirmDialog：`web/src/components/ConfirmDialog.svelte`
   - props：title、message（支持多行）、confirmText、cancelText、danger（布尔，true 时确认按钮为红色）
   - 返回 Promise<boolean>；ESC = 取消，Enter = 确认；打开时锁焦点
   - 不用 window.confirm

二、替换这 10 处（全部保留原有文案语义，只改呈现方式）
- web/src/App.svelte 约 252 行：跨工作簿选区的 alert → toast.error（文案不变，去掉 ⚠️ emoji）
- web/src/App.svelte 约 259 行：未存盘工作簿的 window.confirm → ConfirmDialog，
  title「工作簿尚未保存」，message 用现有的两句核心意思（会自动改表；建议先 Ctrl+S），confirmText「仍要继续」
- web/src/components/ExecutionCard.svelte 约 189 行：保存失败 alert → toast.error
- web/src/components/ScriptDrawer.svelte 约 154、239、902、915、933 行：5 处 alert → toast.error / toast.info
- web/src/components/ScriptDrawer.svelte 约 897 行：删除宏的 confirm → ConfirmDialog（danger=true，
  confirmText「删除」）
- web/src/components/WorkflowModal.svelte 约 170 行：删除工作流的 confirm → ConfirmDialog（danger=true）

三、文案规则（只针对上面这些弹窗）
- 去掉所有 ⚠️ / ❌ / ℹ️ emoji 前缀
- 标题 ≤ 12 字，正文 ≤ 2 行，按钮 ≤ 4 字
- 不写"为避免跨工作簿误操作"这类解释句，直接说结论 + 下一步动作

【禁止】
- 不改业务判断逻辑（什么情况拦截、什么情况放行，一律不动）
- 不删任何校验分支，只换呈现方式
- 不引入新依赖

【验收】
1. `grep -rn "alert(\|confirm(" web/src --include=*.svelte --include=*.ts` 结果为 0（`ConfirmDialog` 内部除外）
2. 手动跑一次前端：`cd web && npx vite`，在浏览器里触发删除宏 → 出现应用内确认框而不是系统灰框
3. `node test_suite_unit.cjs` → 233/0
```

---

## 【P2 · 入口收敛：一个功能一个家】

```text
本阶段目标：把重复的导航层删掉。窗格里只留"助手 / 宏"两页，管理类操作全部交给功能区。

一、删掉窗格顶栏的 ⋯ 菜单
web/src/components/Header.svelte 约 85-128 行：整块删除 more-menu-container（按钮 + 下拉 + 菜单项），
连带删除 Header 的 props onOpenBatch / onOpenWorkflow，以及 App.svelte 里对应的传参。
App.svelte 里 showBatchModal / showWorkflowModal 仍保留（功能区还能触发），只是窗格里不再有入口。

二、分段导航 3 项 → 2 项
web/src/components/Header.svelte 约 132-157 行：只保留「助手」「宏」两项，删掉「数据工具」。
activeWorkspace 的类型从 'assistant' | 'scripts' | 'datatools' 改成 'assistant' | 'scripts'。
注意 App.svelte 里还有 activeWorkspace === 'datatools' 的分支和 ?ws=datatools 的 URL 参数，一并清理。

三、「宏」改成内联页，不再用抽屉覆盖
现状：点「宏库」标签 → ScriptDrawer 以覆盖层盖住整屏 → 关掉又跳回助手，标签态和内容对不上。
目标：
1. web/src/components/ScriptDrawer.svelte 增加一个 `inline` prop（默认 false 保持旧行为）。
   inline=true 时：不渲染 .drawer-backdrop 遮罩、不渲染自带的标题栏和关闭按钮（×），
   直接铺满容器，作为页面内容。
2. App.svelte 里当 activeWorkspace === 'scripts' 时，把 ScriptDrawer 以 inline=true 渲染在
   .workspace-view 里（和助手页平级），而不是覆盖层。
3. 关掉"关抽屉自动跳回 assistant"的逻辑（App.svelte 约 873-877 行）。
4. 从功能区「宏库」进来时（bridge action open_macro_library / open_scripts / open_my_macros），
   行为改为：切到 activeWorkspace='scripts'，不再 showScripts=true。

四、删掉数据工具的旧弹窗路径
App.svelte 约 882-890 行：`{#if showDataTools}` 整块删除，连带 showDataTools 状态、
window.__openDataTools 里的赋值（改成切 activeWorkspace，或直接删掉这个调试钩子）。
DataToolModal 的 isWorkspaceView=false 分支保留（P3 还要用），但主流程不再走它。

【禁止】
- 不改 ScriptDrawer 内部的列表/导入/导出逻辑，只改"怎么呈现"
- 不改 bridge 的 action 名（open_scripts / open_macro_library 等仍要能触发）
- 不动 Excel 功能区（P0 已完成）

【验收】
1. 窗格里不存在任何覆盖式抽屉：`grep -rn "drawer-backdrop" web/src/components/*.svelte` 只剩 inline=false 的分支
2. 切换「助手 / 宏」时页面无遮罩、无跳动，标签高亮与内容一致
3. `grep -n "更多功能" web/src/components/Header.svelte` 为空
4. `node test_suite_unit.cjs` → 233/0（若有断言依赖 activeWorkspace 的三值类型，同步更新并说明）
```

---

## 【P3 · 数据工具面板化 + 底部输入区压缩】

```text
本阶段目标：两件事，都在 440px 宽度下验收（任务窗格固定 Width=440，见 LeeExcelAddIn.cs:247）。

一、数据工具从"5 个横排 tab"改成"网格 → 详情"
现状问题（实测）：web/src/components/DataToolModal.svelte 的工具 tab 条 scrollWidth=482 > clientWidth=440，
第 5 个工具"外部数据"被挤出可视区；而兜底的下拉写的是 @media (max-width:380px) / @container (max-width:380px)，
窗格固定 440px，永远不触发——这段是死代码。

目标：
1. 新增一个工具选择视图（同一个组件内用 `view: 'grid' | 'detail'` 状态控制）：
   - grid：标题「数据工具」+ 一行说明「选一个工具，按当前选区处理」+ 2 列网格，5 个工具全部可见
     工具名简化为：去重 / 对账 / 汇总 / 图表 / 导入数据；每个卡片一行副标题（≤8 字）
   - detail：现有的表单，顶部加一行返回：「← 数据工具」+ 当前工具名
2. 删掉横排 tab 条（.datatool-tabs-bar）和 380px/360px 的媒体查询兜底，以及 .datatool-selector-mobile 下拉
3. 详情页里"工具名"只出现一次：现在 tab、标题、下拉三处重复，改成只在返回行 + 标题处各一次
4. 主按钮（如"开始只读去重分析"）改为表单底部通栏按钮，宽度 100%，高度 34px

二、底部输入区 163px → ≤90px
现状：web/src/components/ChatInput.svelte 里 .mode-header（26px）+ .input-box（114px），
第一行塞了 操作/对话 + 全簿透视 + 附加当前选区 四个胶囊，下面还有一行快捷键提示。

目标：
1. 删除「全簿透视」（约 314-322 行，含 🌐 emoji）、「附加当前选区」（约 324-336 行）两个胶囊，
   收进输入框左侧的一个「＋」按钮的下拉菜单
   （菜单项：附加当前选区 / 整簿上下文）。选择后不再显示胶囊，改成输入框下方一行状态文字。
2. 输入框下方保留一行状态条（12px）：`操作 · 未附加选区`，其中「操作」可点切换 操作/对话，
   「未附加选区」可点触发附加。删除「Enter 发送 / Shift+Enter 换行」这行提示，改放进 ⓘ 的 title。
3. 输入框最小高度从现在的值降到 38px（单行），输入超过 3 行才长高，最大 96px。
4. 主按钮文案「执行操作」→「执行」，保持绿色实心；空态时保持禁用但不要浅到看不见（用 --office-dim 描边，
   不要降低整体透明度到 50% 以下）。

【禁止】
- 不改数据工具的任何计算逻辑、bridge 调用、参数名
- 不改 5 个工具的表单字段（只改呈现顺序与文案）
- 不改输入框的发送逻辑、快捷键行为（Enter 仍然发送）

【验收】（在浏览器里用 440×900 视口验收，前端可加 ?vw=440&vh=900 参数）
1. 打开数据工具：5 个工具全部可见，`document.querySelector('.datatool-tabs-bar')` 为 null，
   页面无横向滚动（scrollWidth === clientWidth）
2. 点任意工具 → 进详情 → 点返回 → 回到网格
3. 助手页：`document.querySelector('.app-header').height + .input-container 高度 ≤ 130`
4. `node test_suite_unit.cjs` → 233/0
```

---

## 【P4 · 文案清洗（三档分层）】

```text
本阶段目标：把界面上的"验收报告语言"降级到详情层。不改布局、不改交互，只改文字。

一、先建一张替换表（按下面规则执行，不要自由发挥）

A 档（主界面：按钮、标题、字段名、状态、结果卡首行）——只允许"动作 + 结果"：
  快照保护边界：…            → 可一键撤销这次改动。涉及外部文件的操作除外（点开看）
  隔离工作副本 · 遇错即停      → 出错就停
  开始只读去重分析            → 查找重复项
  判断主键列（单列或复合键）    → 按哪几列判断重复
  已选 1 列复合比较            → 按第 1 列判断
  身份核验一致                → 已确认写入目标文件
  抽样匹配                    → 已抽查 3 处，结果一致（数字取实际值，取不到就写"已抽查"）
  写后核验                    → 核对结果
  审计日志                    → 运行日志
  保存于 2026-09-25 10:15     → 3 天前保存（相对时间，超过 30 天再显示日期）

B 档（展开区 / tooltip / 确认卡正文）——允许讲机制和边界，但要短：
  保留「快照」「回滚」「预检」等词，但每段 ≤ 2 行，且必须先说"能做什么"再说"什么情况例外"。

C 档（日志、诊断包）——保持现状，不用改。

二、禁止词（A 档里出现即违规）
快照边界 / 只读 / 核验 / 隔离 / 副本 / 复合键 / 预检 / 兜底 / 白名单 / 指纹 / 幂等 / 源码 / 脱敏 /
降级 / 拦截 / 入口
例外：设置页可以保留 API Key / 模型 / Token / 参数（用户是懂行的人）。

三、具体到文件
1. web/src/components/DataToolModal.svelte：5 个工具的 desc、字段 label、主按钮、错误提示（约 2480 字，重点）
2. web/src/components/ScriptDrawer.svelte：卡片描述、导入导出说明、告警文案（约 1600 字，重点）
3. web/src/components/ExecutionCard.svelte：徽章、字段 label、底部说明（约 1277 字）
4. web/src/components/SettingsModal.svelte：删掉 4 段 helper text 里讲实现的部分
   （`/chat/completions`、`LocalStorage`、`max_tokens 字段`、`thinking 字段`），改成：
   - 接口地址 → 提示「一般以 /v1 结尾」
   - API Key → 提示「只保存在本机」
   - 温度 / 最大输出 → 收进「高级参数」折叠区，折叠区默认收起
5. web/src/components/BatchModal.svelte / WorkflowModal.svelte：标题、提示条、空状态

四、统一数字与单位
- 耗时统一用 ms（<1000）或 s（≥1000），同一屏不许两种写法
- 时间统一相对时间
- 数量加千分位

【硬约束】
- 每条界面文案 ≤ 18 字，超了就拆成"标题 + 详情"
- 改完 A 档后，`grep` 上面那批禁用词，A 档所在文件的模板区应命中 0
- test_suite_unit.cjs 里有 `身份核验失败` / `隔离工作副本` / `抽样匹配` / `遇错即停` 的断言：
  * 如果它断言的是"服务层 note 文案"（web/src/services/llm.ts:436 这类），**不要改那个 note**，
    它属于 B 档详情，改了会破坏契约
  * 如果它断言的是"组件里的展示文字"，同步更新断言并在说明里写明
- 不改任何 verificationStatus / precheckStatus 的枚举值，只改展示映射

【验收】
1. 交付一份"改了哪些文案"的对照表（旧 → 新，按文件分组）
2. `grep -rn "只读\|复合键\|隔离工作副本\|快照边界" web/src/components/*.svelte` 只允许出现在注释里
3. 主界面可见中文总量：改前 7504 字，目标 ≤ 4500 字。请写个小脚本统计
   （取每个 .svelte 的 </script> 之后、<style> 之前的文本节点 + placeholder/title/label 属性里的汉字数）并报数
4. `node test_suite_unit.cjs` → 233/0
```

---

## 【P5 · 撤销闭环 + L2 确认卡（要动 C#，单独排期）】

```text
本阶段目标：把"每次操作都弹确认"改成"默认可撤销，只有救不回来的才确认"。这一步会动执行链，
必须最小改动、可回退。

一、先确认现状（只读，不改）
1. src/NativeBridge.cs 里 execute_vba 的执行流程：快照在什么时候创建、失败时怎么回滚
2. web/src/components/ExecutionCard.svelte 里"恢复到执行前"按钮现在调用 `bridge.send('restore_snapshot', …)`
   （约 224 行），确认它的入参与幂等性
3. 把结论写给我，等我确认后再动代码（这一步不要直接改）

二、分级判定（在 C# 宿主侧，不看模型自述）
在 execute_vba 执行前，对即将执行的 VBA 源码做一次静态扫描，判定风险等级：
- L2（不可逆/越界）命中任一即算：
  · Worksheets(...).Delete / Sheets(...).Delete
  · Workbooks.Open / SaveAs / Close 且目标不是当前工作簿
  · CreateObject("Scripting.FileSystemObject") / "MSXML2.*" / "WinHttp.*" / "InternetExplorer.Application"
  · VBProject / ThisWorkbook.VBProject
  · Kill / FileCopy / Name ... As ...（VBA 文件操作语句）
  · 写入路径不在 targetWorkbookFullName 所在目录
- L1：有写操作（.Value = / .Formula = / .Interior / .Borders / Rows.Insert / Columns.Delete 等）但不命中 L2
- L0：无写操作
判定不出来（正则没匹配到任何写操作但也不能确定只读）→ 按 L2 处理。

三、执行链改造（等确认后再做）
1. L0：不建快照，直接执行
2. L1：建快照 → 直接执行 → 返回值里带 snapshotId + canUndo=true
3. L2：**不执行**，返回 { needsConfirm: true, riskLevel: 'L2', reasons: [...] }，
   由前端弹确认卡，用户确认后带 confirmToken 二次调用才执行
4. confirmToken 必须一次性、与 requestId 绑定，防止前端伪造

四、前端
1. ExecutionCard：L1 结果卡首行给「撤销这次改动」按钮（复用现有"恢复到执行前"逻辑）
2. 新增 L2 确认卡（可以复用 MacroConfirmCard 的样式，但内容改成风险导向）：
   标题写具体对象，例如「将删除工作表《汇总看板》」，正文一行说明后果，
   按钮「取消」/「仍要删除」（danger）
3. 删除"AI 生成→直接执行"和"宏库运行→确认卡"两套不一致的摩擦：宏库手动运行也走同一套分级
   （用户点了运行不等于知情，风险才是变量）

【禁止】
- 不为了让确认卡上线而放宽任何现有校验
- 不改快照的创建时机与内容
- 不在前端做风险判定（前端只看宿主返回的 riskLevel）

【验收】
1. 三类样本各跑一次：只读查询 / 写格式 / 删工作表，确认分别走 L0 / L1 / L2
2. L2 未确认时，工作簿必须完全没有被修改（对比执行前后的 usedRange 与行数）
3. `node test_suite_unit.cjs` → 233/0，且新增至少 3 条针对风险判定的单元测试
```

---

## 【P6 · 视觉规则闸门（放最后做）】

```text
本阶段目标：加一个能失败的检查脚本，让前面清干净的东西不会回退。

一、新增 scripts/check-tokens.mjs（不引入依赖，用 node 内置 fs）
扫描 web/src 下所有 .svelte 的 <style> 块和 .css 文件，报告：
1. 裸颜色：#hex / rgb() / rgba()，白名单只有 web/src/styles/office-fluent.css 的 token 定义区
2. 禁用字号：font-size 里的 9px / 10px / 10.5px / 11.5px（11px 允许，但要单独统计数量）
3. 圆角：border-radius 里除 4px / 6px / 999px / var(--office-radius*) 之外的值
4. emoji：模板区的 emoji 码点（U+1F300–U+1FAFF、U+2600–U+27BF）
5. 原生弹窗：alert( / confirm(
输出：每条一行 `文件:行 规则 内容`，结尾汇总 `裸色 N / 小字 N / 圆角 N / emoji N / 原生弹窗 N`

二、接入
1. package.json 加 `"check:ui": "node scripts/check-tokens.mjs"`
2. 先只做 warn（退出码 0），把当前数字报给我；
   等我确认存量清完，再改成有违规就 exit 1

【验收】
1. `node scripts/check-tokens.mjs` 能跑通并输出汇总行
2. 脚本能真的失败：手动在任意 .svelte 里加一行 `color:#ff0000` 和一个 emoji，脚本必须报出来，然后撤销这个测试改动
3. `node test_suite_unit.cjs` → 233/0
```

---

## 附：发完之后的复查口径

每阶段完成后，你可以把它的交付说明发我，我按下面四条复查（这也是我上一轮审查的口径）：

1. **改动范围**是否等于我点名的文件集；多出来的文件要有理由。
2. **验收命令**是否真的跑过（贴的是输出，不是"应该通过"）。
3. **有没有用"已完成/100%"替代事实**。
4. **有没有把断言删掉当作"修好了测试"**。
