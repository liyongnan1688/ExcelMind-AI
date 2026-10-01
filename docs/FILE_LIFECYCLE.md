# Lee-Excel 项目文件生命周期与目录职责规范 (FILE_LIFECYCLE.md)

本文档明确 Lee-Excel 项目中各类文件的职责边界、生命周期、存储规范、清理机制及实验晋升流程，供项目维护人员及后续 AI 助手共同遵守。

---

## 一、 文件分类定义与存储边界

| 类别 | 属性定义 | 允许存放位置 | 纳入版本控制 (Git) | 清理政策 |
| :--- | :--- | :--- | :--- | :--- |
| **正式生产源码** | 插件核心运行必需的 C# 后台逻辑、Web 前端界面、资源、清单与依赖 | `src/` (后台 C#)<br>`web/src/` (Svelte/TS 前端)<br>`web/` (构建配置) | 强制跟踪 | 严禁清理 |
| **正式发布产物** | 最终用于 Office 加载运行的发行产物 | `bin/LeeExcel.dll`<br>`bin/dist/` (编译后前端) | 仅跟踪必要发布基线 | 仅随发布流程覆盖，严禁生成测试EXE或临时日志 |
| **长期测试资产** | 可重复执行的回归套件、诊断工具、固定 Mock 数据、编译说明 | `tests/tools/` (测试程序源码)<br>`tests/diagnostics/` (诊断脚本)<br>`test_suite_unit.cjs` (单元测试)<br>`test_regex_counter_example.cjs` | 强制跟踪 | 严禁存放日常生成的临时 EXE、DLL、工作簿或运行日志 |
| **测试与运行时生成物** | 每次执行测试、诊断或调试时生成的二进制物、日志、临时工作簿、内存Dump | `.artifacts/tests/<run-id>/`<br>（run-id 采用时间戳加唯一标识，如 `20261001_083000_regression`） | 必须忽略 (`.gitignore`) | 默认保留 14 天，由 `clean_artifacts.ps1` 安全清理 |
| **临时调试与实验产物** | 一次性排查脚本副本、中间切片、临时文本 | `.artifacts/tmp/<session-id>/` | 必须忽略 (`.gitignore`) | 默认保留 14 天，由 `clean_artifacts.ps1` 安全清理 |
| **当前临时实验区** | 当前开发迭代中的活跃驱动脚本（需人工参与的高阶 E2E 脚本） | `scratch/` (目前限定为 11 个受保护活跃脚本) | 暂留原位，不继续作为测试输出目录 | 严禁自动清理；待解耦完成后晋升或归档 |
| **历史追溯证据** | 证明过重大缺陷修复、历史版本验收通过的关键证据文件 | `docs/history/evidence_*/`<br>`docs/history/` | 跟踪管理 | 永久保留关键证据，不收纳过程无价值垃圾 |
| **备份与基线记录** | 清理、重构或重大迁移操作前的原始文件镜像、manifest、回滚脚本与基线 | `.backups/`<br>`.backup_cleanup_batch1_20261001/`<br>`.backup_cleanup_batch2_20261001/` | 忽略 (`.gitignore`)，持久化留存物理磁盘 | 不受普通临时文件清理脚本管理，供紧急回滚 |

---

## 二、 目录职责与禁令

### 1. `tests/` 目录
- **职能**：存放长期维护的测试工具源码（C#）与诊断脚本（PowerShell/Node）。
- **禁令**：
  - 严禁将编译出的 `.exe`、`.dll` 直接保存在 `tests/` 根目录或源码同级目录；
  - 严禁测试执行时在 `tests/` 目录下生成临时 `.xlsx` 工作簿或 `.log` 日志；
  - 编译 C# 测试工具时，必须显式通过 `/out:` 参数输出至 `.artifacts/tests/<run-id>/`。

### 2. `bin/` 目录
- **职能**：仅作为 Excel 加载插件的宿主目录，存放 `LeeExcel.dll` 及前端静态资源 `dist/`。
- **禁令**：
  - 严禁将任何临时测试 EXE、诊断工具、临时测试工作簿生成到 `bin/`；
  - 严禁在 `bin/` 写入任何调试日志或抓包转储。

### 3. `.artifacts/` 长期生成物目录
- **结构规范**：
  - `.artifacts/tests/<run-id>/`：测试运行专用。每次运行使用独立时间戳目录，避免覆盖破坏现场。
  - `.artifacts/tmp/<session-id>/`：临时会话调试专用。
- **并发与锁机制**：
  - 正在运行的测试任务需在对应目录创建 `.running` 标记文件；运行结束（成功或异常）后移除该标记。
  - 清理脚本检测到 `.running` 标记时自动跳过，绝不强删活跃运行环境。

---

## 三、 清理规范与生命周期管理

### 1. 清理条件与安全门禁
- **可清理对象**：仅限 `.artifacts/tests/` 和 `.artifacts/tmp/` 下符合规范的运行目录。
- **默认保留期**：最近 14 天内的产物一律保留，防止破坏近期测试追溯链。
- **安全检查**：
  - 严禁支持任意外部绝对路径；
  - 严禁触碰 `src/`、`web/`、`tests/`、`docs/`、`scratch/`、`bin/` 以及任何备份目录；
  - 拒绝路径穿越（`..`）与符号链接/重解析点（ReparsePoint）；
  - 对无有效运行元数据（`meta.json`、`run.json` 或时间戳规范前缀）的未知目录，默认只报告、不自动删除；
  - 被系统进程锁定的文件自动跳过，严禁强杀进程或强解文件句柄。

### 2. 操作指令与 RunId 规则

#### 核心 RunId 匹配与安全规则
- **完全精确匹配（Exact Match）**：`-RunId` 必须完全匹配目标子目录名称，禁止使用通配符（`*`、`?`）或前缀匹配。
- **非法输入严格拒绝**：拒绝包含路径分隔符（`\`、`/`）、目录穿越（`..`）或非法字符的 RunId 输入，发现即刻中断退出。
- **跨作用域歧义保护**：如果 `tests/` 与 `tmp/` 作用域下同时存在同名的 RunId，脚本主动报错并中止，绝不同时删除两个目录；调用者必须显式传入 `-Scope Tests` 或 `-Scope Tmp` 消除歧义。
- **运行锁保护生效**：显式指定 `-RunId` 可免除保留天数（RetentionDays）限制，但仍受全部核心安全门禁约束（若存在 `.running` 标记依然强制跳过，防止误删活跃测试）。
- **默认只读预览**：默认仅输出扫描结果与待清理目录清单，绝不自动删除；必须显式附带 `-Apply` 参数方可执行物理清理。

#### 预览清理候选（默认安全模式）
```powershell
# 仅预览超过 14 天的可清理目录（不执行物理删除）
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/clean_artifacts.ps1

# 预览超过 7 天的可清理目录
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/clean_artifacts.ps1 -RetentionDays 7

# 预览指定具体 RunId（不执行物理删除）
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/clean_artifacts.ps1 -RunId "run_20261001_083000"
```

#### 显式执行清理
```powershell
# 显式传入 -Apply 执行物理清理
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/clean_artifacts.ps1 -Apply

# 按指定具体运行 ID 执行清理（必须精确匹配且附带 -Apply）
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/clean_artifacts.ps1 -RunId "run_20261001_083000" -Apply

# 发生歧义时，显式指定作用域执行清理
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/clean_artifacts.ps1 -RunId "run_20261001_083000" -Scope Tests -Apply
```

#### 恢复与回滚
- 若误删或需追溯历史：
  - 第一批清理恢复：执行 `.backup_cleanup_batch1_20261001/rollback.ps1`
  - 第二批清理恢复：执行 `.backup_cleanup_batch2_20261001/rollback.ps1`
  - 最近路径修复前备份：保存在 `.backups/pre_path_fix_20261001/`，包含全部获批修改文件的原始副本及 `manifest_pre_fix.json`。
  - 回滚脚本具备 SHA256 自检机制，若目标文件存在且被修改会主动报错中止，保证无二次损坏。

---

## 四、 临时实验代码晋升为正式测试的机制

任何在 `.artifacts/tmp/` 或 `scratch/` 中编写的原型脚本，如需长期保留并纳入 CI/自动化验证，必须经过以下标准化晋升流程：

1. **去硬编码与路径规范化**：
   - 将硬编码的本地路径改为从命令行参数（`args`）或配置文件读取；
   - 默认输出路径统一指向 `.artifacts/tests/<run-id>/`；
   - PowerShell 脚本必须使用 `$PSScriptRoot` 动态解析相对路径，严禁依赖当前工作目录。
2. **副作用声明与环境解耦**：
   - 检查工具是否涉及 Excel COM 实例启动、Windows 弹窗句柄拦截、内存扫描或外部模型网络 API；
   - 涉及网络调用的必须提供 Mock 模式；不能离线运行的必须明确标注副作用。
3. **晋升入库**：
   - C# 工具源码移动至 `tests/tools/<ToolName>.cs`；
   - 诊断与排查脚本移动至 `tests/diagnostics/<script_name>.ps1`；
   - 在 `tests/README.md` 中登记其用途、依赖、编译参数、运行命令与副作用。
4. **清理临时残留**：
   - 晋升完成后，清除 `.artifacts/tmp/` 中的草稿，保持工作区干净。
