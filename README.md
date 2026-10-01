# ExcelMind AI

Windows 平台 Microsoft Excel 原生桌面级 AI 智能加载项。基于 **Excel-DNA** 与 **Microsoft Edge WebView2** 架构构建，将大语言模型（LLM）的自然语言意图理解与 Office COM 自动化接口无缝结合，在 Excel 侧边栏提供沉浸式数据处理、公式推导、统计分析与表格美化体验。

---

## 核心特性

- **双模式分流架构**
  - **💬 自由问答 (CHAT)**：普通概念咨询、Excel 函数用法探讨或数据建模思路答疑，仅返回文本解释与公式建议，不触碰本地表格数据，安全省心。
  - **⚡ 表格自动化 (AUTOMATION)**：自然语言下发操作指令，自动生成并执行合规 VBA 脚本，完成多表清洗、多级汇总、透视表建立、条件格式设置与商业图表绘制。
- **全方位数据安全防护**
  - **目标工作簿绑定 (Target Binding)**：自动识别并显式绑定当前活动工作簿，杜绝代码跨工作簿误写入或污染后台其他已打开表格。
  - **执行前全本物理快照 (Pre-run Snapshot)**：执行任何变更操作前，毫秒级备份当前工作簿的完整副本，并在侧边栏提供“快照时间轴”，支持随时一键无损回滚。
  - **严格代码防御拦截**：内建结构完整性核查，拦截截断代码、未闭合过程（缺少 `End Sub`）及可能危害系统的高危语句。
- **个人脚本库 (Scripts Library)**
  - 自动沉淀优质代码：执行效果满意的 VBA 代码，支持一键加入“我的脚本库”（保存在 `%AppData%\ExcelMindAI\Scripts\`），方便后续日常重复调用。
- **主流大模型与自建服务全面兼容**
  - 内置 DeepSeek、OpenAI、Claude、月之暗面 (Kimi)、智谱 GLM、通义千问等主流厂商配置预设。
  - 支持任何兼容 OpenAI `/v1/chat/completions` 标准协议的本地或私有化大模型服务（如 Ollama、vLLM、OneAPI）。
  - 支持 DeepSeek 深度思考（Thinking Mode）模式切换与思考预算（Token Budget）精细调节。
  - API Key 仅保存在本机浏览器 LocalStorage 中，直接与大模型接口通信，不经过任何第三方转发服务。
- **多版本 Excel 与系统全自适应**
  - 兼容 32 位与 64 位 Office 架构，支持 Office 2010 / 2013 / 2016 / 2019 / 2021 及 Microsoft 365。
  - 支持 Windows 7 SP1 / Windows 10 / Windows 11。

---

## 界面与功能区入口

安装或启动加载项后，Excel 顶部功能区将自动常驻 **【ExcelMind AI】** 选项卡：

| 功能区按钮 | 图标 | 功能说明 |
| :--- | :--- | :--- |
| **ExcelMind AI** | 专属绿色品牌徽标 | 唤醒或收起右侧 440px 原生 AI 任务窗格 |
| **自动化脚本** | VBA 代码图标 | 直接打开个人本地脚本库抽屉，管理和快速重放已保存脚本 |
| **API 配置** | 服务连接图标 | 唤出大模型 API 密钥、接口地址及模型参数配置弹窗 |

---

## 快速使用

### 方式一：一键安装（永久常驻，推荐日常使用）

1. 下载最新发布包（如 `ExcelMind_AI_Release_v1.0.zip`）并解压至任意常用文件夹（如 `D:\Tools\ExcelMind_AI\`）；
2. 双击运行 **`安装插件(开启常驻).bat`**；
3. 脚本会自动解除 Windows 下载安全锁定、检测系统运行环境，并将加载项写入当前用户的 Excel 启动项；
4. 启动任意 Excel 文件，顶部功能区即可看到 **【ExcelMind AI】** 选项卡。

### 方式二：免安装便携启动（开箱即用，适合临时演示）

1. 解压发布包后，双击运行 **`免安装启动.bat`**；
2. 脚本自动检测当前电脑安装的 Excel 架构位数（32 位或 64 位），唤醒 Excel 并即时挂载对应 XLL 运行；
3. 不向注册表写入任何启动项，关闭 Excel 即退出，干净无残留。

### 卸载插件

若需移除插件自启动：双击运行 **`卸载插件.bat`**，脚本将自动清理 Excel 注册项，且绝不会误伤您的任何 Excel 业务文件。

---

## 初始配置指南

首次使用前，只需花费 1 分钟完成大模型 API 接入配置：

1. 点击功能区【API 配置】按钮，或点击侧边栏右上角的 ⚙️ 图标；
2. **选择厂商预设**：例如点击“DeepSeek”；
3. **填写 API Key**：输入申请到的 API 密钥（如 `sk-xxxxxxxx`）；
4. **模型参数设置（可选）**：
   - 默认模型为 `deepseek-chat`，如需深度推理可填 `deepseek-reasoner`；
   - 深度思考模式可自由切换【自动】、【关闭思考链】或【指定思考预算】；
5. 点击**【保存配置】**即可。

---

## 典型使用场景示例

### 场景 1：复杂公式与嵌套推导
> **提问（问答模式）**：  
> “在 Sheet1 中，A 列是工号，B 列是部门，C 列是打卡时间。如何在 D 列找出每个员工当天最早的一次打卡时间？请给出函数公式和逻辑解释。”

### 场景 2：数据清洗与格式规整
> **指令（自动化模式）**：  
> “将当前表格首行设为深绿底白色粗体表头，固定冻结首行。自动清除 A 到 H 列的所有前导空格，金额列（E列）格式化为带有千分位符并保留两位小数，整表加浅灰色实线边框。”

### 场景 3：多维汇总与图表生成
> **指令（自动化模式）**：  
> “根据当前工作表的销售流水，按‘大区’和‘品类’汇总‘销售额’与‘利润’，并在旁边新建一个漂亮的簇状柱状图对比各区域利润。”

*提示：每次自动化执行前均会自动拍摄快照。如果生成的表格格式或数据处理不符合预期，直接点击卡片下方的【⏪ 撤回本次操作 (恢复快照)】即可恢复到执行前状态。*

---

## 项目工程结构

```text
ExcelMind-AI/
├── .artifacts/                 # 运行生成物与构建物隔离目录（受 git 忽略保护）
│   ├── release/                # 打包生成的独立 ZIP 发布包
│   └── tests/                  # 自动化回归测试产物输出
├── bin/                        # 插件编译二进制文件与前端分发包
│   ├── LeeExcel.dll            # C# Any CPU 核心逻辑程序集
│   ├── LeeExcel.xll            # 32 位 Excel-DNA 引导入口
│   ├── LeeExcel64.xll          # 64 位 Excel-DNA 引导入口
│   ├── runtimes/               # 双架构 (x86/x64) WebView2 原生加载器
│   └── dist/                   # Svelte 5 前端静态资源 (HTML/JS/CSS)
├── packages/                   # 离线 NuGet 依赖 (Excel-DNA 1.9.0, WebView2 1.0.4191.47)
├── scripts/                    # 运维与自动化流水线脚本
│   ├── core/                   # 安装、便携启动、卸载及批处理模板
│   ├── package_release.ps1     # 跨机器发布包全自动打包流水线
│   └── clean_artifacts.ps1     # 临时产物安全清理工具
├── src/                        # C# 后端源代码
│   ├── BrandIconHelper.cs      # 功能区品牌矢量徽标动态绘制
│   ├── LeeExcelAddIn.cs        # 插件入口点、COM Ribbon 定义与任务窗格生命周期
│   ├── NativeBridge.cs         # WebView2 WebMessage 与 C# 双向通信网关
│   ├── ScriptManager.cs        # 本地 BAS 脚本持久化管理
│   ├── SimpleJson.cs           # 零外部依赖原生轻量 JSON 解析器
│   ├── SnapshotManager.cs      # 工作簿全量物理快照与还原引擎
│   ├── TaskPaneControl.cs      # WinForms 任务窗格宿主与 WebView2 初始化控制
│   └── VbaRunner.cs            # Excel COM 自动化调度与 VBA 执行沙箱
├── web/                        # 前端单页应用 (Svelte 5 + Vite 6 + TypeScript)
│   ├── src/                    # 前端源码 (Office Fluent 主题交互设计)
│   └── package.json            # 前端依赖配置
├── LeeExcel.dna                # 32 位 Excel-DNA 配置文件
├── LeeExcel64.dna              # 64 位 Excel-DNA 配置文件
├── test_suite_unit.cjs         # 离线核心单元测试门禁套件 (27 项)
└── test_regex_counter_example.cjs # 安全正则反例核验测试
```

---

## 本地二次开发与构建

### 运行环境准备
- Windows 10 / Windows 11
- .NET Framework 4.6.2 或更高版本（Windows 系统自带）
- Node.js 18+ 与 pnpm (或 npm)
- PowerShell 5.1 或 PowerShell 7 (pwsh)

### 1. 构建前端
```bash
cd web
pnpm install
pnpm run build
```
前端产物将自动输出至根目录的 `bin/dist/` 文件夹。

### 2. 编译 C# 程序集与全自动打包
在项目根目录下通过 PowerShell 执行一键打包脚本：
```powershell
pwsh -File scripts/package_release.ps1 -Version "v1.0"
```
打包流水线将自动完成：
- 重新触发前端 `pnpm run build`；
- 调用系统自带 64 位 `csc.exe` 编译 `src/*.cs` 输出 `bin/LeeExcel.dll`；
- 整合 32 位与 64 位双架构 XLL 引导加载器与 WebView2 原生驱动；
- 输出成品分发包至 `.artifacts/release/ExcelMind_AI_Release_v1.0.zip`。

### 3. 运行本地质量门禁
提交代码或发布前，确保离线测试套件通过率保持 100%：
```bash
node test_suite_unit.cjs
node test_regex_counter_example.cjs
```

---

## 运行环境与故障排查 (FAQ)

1. **点击按钮提示缺少 WebView2 运行时？**  
   Windows 11 及较新版本的 Windows 10 系统已默认内置 WebView2 运行时。若在精简版或老旧系统中运行，运行 `安装插件(开启常驻).bat` 时按回车即可自动从微软官方下载独立安装包（约 2MB），或前往 [Microsoft Edge WebView2 官网](https://developer.microsoft.com/microsoft-edge/webview2/) 下载独立 Evergreen 运行时。

2. **Excel 提示宏被禁用？**  
   本插件的核心能力之一是通过本地 COM 自动化执行规范的 VBA 代码。如果 Excel 弹窗提示宏被安全阻止，请在 Excel【文件】$\to$【选项】$\to$【信任中心】$\to$【信任中心设置】$\to$【宏设置】中，选择“禁用所有宏，并发出通知”或“启用所有宏”。

3. **从网络下载后无法加载（Mark of the Web 锁定）？**  
   从网盘或即时通讯工具下载的压缩包常带有 Windows 系统的“安全锁定”属性。使用本发布包提供的 `安装插件(开启常驻).bat` 或 `免安装启动.bat` 时，脚本已内建调用 `Unblock-File` 自动解除所有文件锁定。

---

## 开源许可

本项目遵循 MIT 协议开源。
