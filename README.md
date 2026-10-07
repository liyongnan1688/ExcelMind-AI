# ExcelMind AI

> **智能表格助手 · 桌面原生 · 安全可逆**  
> 专为 Windows Microsoft Excel 打造的 AI 智能加载项。结合大语言模型意图理解与 Office COM 原生自动化，让数据清洗、复杂计算、多维汇总与 VBA 宏复用像日常对话一样自然。

> 📌 **项目当前真实状态入口**：参见 [docs/CURRENT_STATE.md](docs/CURRENT_STATE.md)  
> 唯一主规划路线图：[docs/product-roadmap.md](docs/product-roadmap.md) · 测试资产手册：[tests/README.md](tests/README.md)

---

## 📌 当前发布基线 (v1.3.0-rc1)

- **当前发布版本**：[v1.3.0-rc1](https://github.com/liyongnan1688/ExcelMind-AI/releases/tag/v1.3.0-rc1)
- **发布提交**：`b4402a9ead90d8e74c451ccb6683a22e20da3ae4`
- **官方发行包**：`ExcelMindAI-v1.3.0-rc1.zip`
- **用户下载权威 SHA-256 校验值**：  
  `445ef8b481491bc99b31dcd286b1ac109ae6b519fafc51fc456aaa22687ff25e`

---

## ✨ 核心特性

- 💬 **对话与自动化双模式**
  - **问答模式 (CHAT)**：咨询公式用法、函数语法与建模方案，纯文本解答，不碰表格数据。
  - **操作模式 (AUTOMATION)**：自然语言下发操作指令，自动编写并执行 VBA，完成清洗、汇总、格式规整与图表生成。
- 📦 **VBA 宏管理与安全复用**
  - **多源导入**：支持从本地 `.bas`（标准模块）、`.vba`、`.txt` 导入或直接粘贴多过程源码。
  - **100% 原文保真**：严格保留 `Option Explicit`、中文注释与模块属性，不调用大模型替用户改写或伪造代码。
  - **多入口识别**：自动解析所有 Sub/Function 入口，支持下拉指定入口，必填非工作簿参数严格拦截。
  - **免 Key 离线执行**：运行已入库宏无需配置 API Key，直接由本地 COM 极速运行。
- 🛡️ **数据安全防损快照**
  - **执行前物理全本备份**：每次自动化修改或运行宏前，毫秒级备份当前工作簿独立副本。
  - **一键无损撤销**：执行效果不满意时，点击【一键恢复】即可完全还原工作簿数据。
  - **边界说明**：快照仅恢复目标工作簿本身的数据与工作表结构，不能撤销外部文件或网络副作用。
- 🌐 **全大模型服务兼容**
  - 内置 DeepSeek、OpenAI、Claude、月之暗面 (Kimi)、智谱 GLM、通义千问等主流厂商预设。
  - 兼容 Ollama、vLLM、OneAPI 等所有支持 OpenAI 标准协议的本地或私有化模型。
  - API Key 保存在本地 LocalStorage 中（未采用 DPAPI），纯本地直连模型服务，无第三方中转；外部数据源凭据独立采用 Windows DPAPI 本地机密存储。
- 💻 **Office 环境与自适应架构**
  - 架构自适应：包内同时提供 32 位与 64 位原生加载项。
  - 已验证环境：Windows 11 64 位 + 64 位 Microsoft Excel 宿主。
  - 运行未验证：32 位 Office 宿主及真实商业大模型 API 线上调用目前为“运行未验证”。

---

## 🧭 功能区 (Ribbon) 导航

启动后 Excel 顶部功能区展示 **【ExcelMind AI】** 选项卡：

| 功能区分组 | 控件形式 | 名称 / 图标 | 说明 |
| :--- | :--- | :--- | :--- |
| **ExcelMind AI** | 大按钮 (32×32) | 品牌图标 (无按钮文本) | 唤醒或收起右侧 AI 侧边栏 (Tooltip: 打开或收起 ExcelMind AI 工作台) |
| **宏与数据** | 纵向堆叠按钮 | **宏库** / **导入** / **常用宏** / **数据工具** | 打开宏管理列表、源码导入、常用宏动态下拉菜单与清洗对账数据工具 |
| **任务与设置** | 菜单与按钮 | **批量处理** / **工作流** / **设置** | 多文件批量宏队列、双步骤任务流水线与 API 基础配置 |

---

## 🚀 快速开始

从 [Releases](https://github.com/liyongnan1688/ExcelMind-AI/releases/tag/v1.3.0-rc1) 页面下载 `ExcelMindAI-v1.3.0-rc1.zip` 并校验 SHA-256 哈希：

> ⚠️ **未签名运行说明**：本候选包尚未导入商业 CA 证书。运行前请核对哈希值并遵守所在组织 IT 安全政策，由用户自主决定运行，本向导不默认引导用户绕过系统安全防护。

### 方式一：免安装便携启动（推荐体验）
- 双击 **`免安装启动.bat`**
- 自动检测 Excel 位数并即时挂载运行，关闭 Excel 即退出，注册表零残留。

### 方式二：一键安装常驻（日常使用）
- 双击 **`安装插件(开启常驻).bat`**
- 自动解除安全锁定并写入启动项，每次打开 Excel 顶部自动展现功能区。

### 卸载插件
- 双击 **`卸载插件.bat`**，自动清理 Excel 启动项，不影响任何业务表格与用户宏库。

---

## ⚙️ 1 分钟配置指南

首次使用仅需配置一次 API Key：
1. 点击功能区【设置】（或侧边栏右上角 ⚙️ 图标）；
2. 选择模型厂商（如 **DeepSeek**）或填入自定义接口地址；
3. 输入您的 API Key（如 `sk-xxxxxxxx`）；
4. 点击【保存配置】即可开始使用。

> 💡 *注：运行【宏库】中已保存的 VBA 宏与内置快捷工具不需要配置 API Key，完全离线运行。*

---

## 🛠️ 本地编译与构建

```powershell
# 1. 构建前端
cd web && pnpm install && pnpm run build && cd ..

# 2. 自动化打包跨机器发布包
pwsh -File scripts/package_release.ps1 -Version "v1.3.0-rc1"
```

离线自动化质量门禁检验：
```bash
node test_suite_unit.cjs
node test_regex_counter_example.cjs
```

---

## 📄 开源许可

本项目基于 [MIT License](LICENSE.md) 协议开源。
