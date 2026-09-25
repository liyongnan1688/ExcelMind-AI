# 任务规划 (Task Plan): Excel AI 助手 (Lee-Excel) 系统性质量修复

## 核心目标与纠偏原则
1. **解除未经证实的假设**：
   - 不把 `temperature: 0.1` 假定为数值矩阵的唯一因果；移除写死强制温度，支持前端可选配置与按模型能力传参（默认留空不强传）。
   - 严禁把“统一使用 ActiveWorkbook”作为目标绑定方案；任务发起时在前端与宿主锁定具体目标工作簿（FullName / Name），执行前后校验身份，生成的 VBA 不受窗口切换置顶影响。
2. **拆分聊天与自动化通道 (CHAT vs AUTOMATION)**：
   - 问答、闲聊、身份咨询走 CHAT（零 VBA、零快照、零注入）。
   - 明确操作指令走 AUTOMATION。
   - 意图不明确时发出简短澄清，绝不乱执行。
3. **修复生成协议与代码提取**：
   - 彻底删除 D1:L9、35行上限、严禁逐格等偏置词。
   - 严禁凭 `includes('Sub')` 猜代码；响应截断或缺少 End Sub 时立即安全拦截。
4. **区分“宏已运行”与“任务完成”**：
   - COM 执行成功仅代表宏已运行。
   - 引入写后读回 (Readback)，比对起始单元格、算式特征、数据完整性；主观美化如实标为“效果待确认”。
5. **脱敏审计日志与极简界面**：
   - 气泡保持一句话汇报，默认折叠；审计日志记录脱敏参数、规整步骤与读回数据。
6. **跨任务全面回归**：
   - 覆盖纯聊天、代码解释、阶梯算式表、纯数字矩阵、数据格式化、公式、筛选、图表、截断保护与多窗口切换。

---

## 阶段进展

### Phase 1: 基础环境与回退基线
- **Status:** complete
- **Tasks:**
  - [x] 在 `feat/system-quality-fix` 分支保留 `fdab9e1` 可回退基线。

### Phase 2: C# 宿主引擎目标绑定与写后读回 (Readback)
- **Status:** complete
- **Tasks:**
  - [x] `src/VbaRunner.cs`：实现 `WorkbookReadback` 状态读回（区域、首末坐标、行列数、边框、底色、公式、抽样）。
  - [x] `src/VbaRunner.cs`：实现过程入口标准化与目标工作簿显式绑定（防止 ActiveWorkbook/ThisWorkbook 切换窗口串改）。
  - [x] `src/NativeBridge.cs`：实现 `FindTargetWorkbook`，针对目标工作簿执行与快照备份。
  - [x] `src/SnapshotManager.cs`：强化未保存工作簿防崩溃保护。
  - [x] 通过系统 `csc.exe` 成功编译 `bin/LeeExcel.dll` 与组织 `bin/LeeExcel64.xll`。

### Phase 3: 前端通道路由、严格协议与写后核验
- **Status:** complete
- **Tasks:**
  - [x] `web/src/services/config.ts`：增加可选 `temperature` 配置，移除写死温度。
  - [x] `web/src/services/llm.ts`：实现 `detectIntent`（精确划分 CHAT、AUTOMATION、AMBIGUOUS）。
  - [x] `web/src/services/llm.ts`：实现双通道系统提示词（CHAT 严禁代码块，AUTOMATION 规范输出并清除 D1:L9 偏置）。
  - [x] `web/src/services/llm.ts`：实现严格代码提取 `extractVbaCode` 与截断拦截。
  - [x] `web/src/services/llm.ts`：实现 `verifyExecutionResult`（客观核验算式文本、坐标起点与主观美化）。
  - [x] `web/src/components/ExecutionCard.svelte`：增加“写后核验”与“审计日志”选项卡，默认折叠，不伪称满分完成。
  - [x] `web/src/App.svelte`：实现通道分流调度与简短汇报。
  - [x] `npm run build` 构建成功，输出到 `bin/dist/`。

### Phase 4: 多场景自动化回归测试
- **Status:** complete
- **Tasks:**
  - [x] 运行 `node test_suite_unit.cjs`：30 项单元测试（多措辞意图、截断拦截、核验逻辑）100% 通过。
  - [x] 运行 `test_system_suite.ps1`：9 项独立 Excel COM 集成测试（多工作簿切换隔离、阶梯算式表、纯数字矩阵、数据完整性、公式、筛选、图表）100% 通过。
  - [x] 运行 `test_core.ps1`：存量基础功能 100% 通过。
