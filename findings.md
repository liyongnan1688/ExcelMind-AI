# 技术调研与关键发现 (Findings)

## 1. 用户开发环境现状
- **Node.js**: `v24.21.0`
- **pnpm**: `12.6.0`
- **Python**: `3.14.6`
- **.NET SDK**: 暂未安装现代 .NET SDK，但系统包含 `C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe`，且系统支持 `winget`。
- **Excel 版本**: Office 16.0 (Excel 2016 / 2019 / 2021 / Office 365)
- **宏安全注册表状态**:
  - `HKCU:\Software\Microsoft\Office\16.0\Excel\Security\AccessVBOM = 1`（已允许对 VBA 工程对象模型的编程访问！这是最关键的绿灯）。

## 2. 参考项目 (hewliyang/office-agents) 审计发现
- **许可证**: MIT License。
- **UI & 架构**: Svelte 5 + TailwindCSS + Vite，属于 Office.js Web Add-in。
- **模型配置**: `packages/sdk/src/provider-config.ts` 定义了高度成熟的国内主流厂商与自定义 API 配置，完全可直接复用。
- **折叠交互**: `packages/core/src/chat/compact-execution-block.svelte` 原生支持默认折叠状态条、展开查看详情与日志的交互模式。
- **Excel 操作真实机制**:
  - 原项目完全依赖 Office.js 网页沙箱 API（`set_cell_range`, `eval_officejs` 等），根本没有真正执行 VBA 的能力。
  - 其所谓的 `vba-generator.ts` 仅是将单元格写操作反向拼装为静态字符串供人工复制，并明确注明“禁止假装执行 VBA”。
  - 其所谓的撤销依赖 IndexedDB 记录单元格差异，无法用于还原 VBA 的全局副作用。

## 3. 底层 VBA 动态执行与工作簿安全
- **VBIDE 注入机制**: 通过 `wb.VBProject.VBComponents.Add(vbext_ct_StdModule)` 注入字符串，调用 `excelApp.Run` 执行，随后调用 `VBComponents.Remove` 瞬时清理。
- **.xlsx 纯净策略**: 模块执行完立即销毁，目标 `.xlsx` 永远不保存 VB 项目，彻底杜绝另存为 `.xlsm` 的弹窗骚扰。
- **物理快照备份与恢复**: 执行前使用 `wb.SaveCopyAs(path)` 毫秒级生成二进制全文件快照；恢复时 `wb.Close(false)` 释放文件锁，覆盖后重新打开。
