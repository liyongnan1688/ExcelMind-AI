# Lee-Excel AI 协同开发规约 (AGENTS.md)

欢迎协助维护 Lee-Excel 项目。所有在此项目工作的 AI 助手（包括但不限于 Antigravity、Claude Code、Cursor 等）必须严格遵守以下工程规范：

---

## 1. 目录存储与产物归档铁律
- **新增测试源码**：必须保存在 `tests/tools/`（C# 程序源码）或 `tests/diagnostics/`（诊断脚本）下，不得堆积在项目根目录或 `scratch/`。
- **测试与临时产物**：执行测试产生的 `.exe`、`.dll`、`.log`、`.xlsx` 工作簿以及调试转储，必须显式输出至 `.artifacts/tests/<run-id>/` 或 `.artifacts/tmp/<session-id>/`。
- **严禁污染 `bin/` 与根目录**：`bin/` 仅用于承载 Excel 运行时的最终插件 DLL 及前端编译产物。严禁在 `bin/` 或根目录下生成测试可执行程序、临时工作簿或运行日志。
- **禁止将生成物当成正式源码提交**：临时运行日志、dump 文本、编译出的 EXE/DLL 严禁 `git add` 提交到仓库。
- **受保护的 11 个活跃脚本**：`scratch/call_llm.cjs`、`scratch/test_compile_behavior.ps1` 及其 9 个强依赖驱动脚本处于活跃开发保护中，不得擅自改动或强行迁移。

---

## 2. 外部交互与执行限制
- **不自动运行真实 API 测试**：涉及大模型 API 调用（如 `probe_model.cjs`、`call_llm.cjs` 等）会产生网络流量和账户扣费，**严禁在自动化回归或无用户明确授权的前提下自动运行真实 API**。
- **不私自启动/关闭本地已打开的 Excel 进程**：涉及 COM 自动化的工具必须做好异常处理（try/finally 中释放 COM 对象），严禁在未确认情况下强杀用户的 Excel 进程。
- **路径解析规范**：PowerShell 脚本一律使用 `$PSScriptRoot` 解析项目内相对路径，不得假定用户终端当前所在的 Working Directory。
- **C# 编译规范**：编译测试工具时必须使用系统自带 64 位 `v4.0.30319/csc.exe` 并显式传入 `/out:".artifacts/tests/..."`，严禁直接输出在源文件同级。

---

## 3. 安全清理守则
- **不未经确认执行实际清理**：所有文件删除、跨目录迁移操作，必须先进行只读调查并输出精确清单，得到用户显式确认后方可执行。
- **清理入口**：日常生成物清理必须且仅能通过 `scripts/clean_artifacts.ps1` 执行。默认仅运行预览模式，只有带 `-Apply` 参数时才执行物理清理。
- **严禁使用 `git clean`**：绝不允许使用 `git clean -fd` 等不可逆粗暴命令抹除未跟踪文件。

---

## 4. 交付与验证标准
- 任何代码或配置修改后，必须运行现有的离线核心门禁（`node test_suite_unit.cjs` 和 `node test_regex_counter_example.cjs`），确保 100% PASS。
- 绝不轻言“全部功能正常”，未实际在真实 Excel 宿主运行验证的，必须明确说明待验证范围。
