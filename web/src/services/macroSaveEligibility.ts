import type { VbaExecutionData } from './bridge';

export interface MacroSaveEligibilityResult {
  isSaveEligible: boolean;
  reason: string;
}

/**
 * 常用宏保存资格强校验纯函数门禁
 * 
 * 准入规则：
 * 1. 代码非空且当前非执行中
 * 2. 宏执行完成且无运行期异常 (execution.error 为空)
 * 3. 预检未失败 (precheckStatus !== 'failed')
 * 4. 未发生非目标工作簿越权修改 (!otherWorkbooksAffected)
 * 5. 目标身份核验未失败 (verificationStatus !== 'failed')
 * 
 * 注：抽样匹配 (matched) 与待确认 (unconfirmed) 均属于执行成功，均允许保存入库；
 * 仅安全违规、目标错乱或运行报错时进行物理拦截。
 */
export function checkMacroSaveEligibility(
  execution: VbaExecutionData | null | undefined,
  isExecuting: boolean,
  code?: string
): MacroSaveEligibilityResult {
  const hasCode = Boolean(code && code.trim().length > 0);
  const isSaveEligible = Boolean(
    hasCode &&
    !isExecuting &&
    execution &&
    !execution.error &&
    execution.precheckStatus !== 'failed' &&
    !execution.readback?.otherWorkbooksAffected &&
    execution.verificationStatus !== 'failed'
  );

  let reason = '';
  if (isExecuting) {
    reason = '宏正在执行中，请等待完成';
  } else if (!hasCode) {
    reason = '无可保存的宏代码';
  } else if (!execution) {
    reason = '代码尚未执行，暂不可存为常用宏';
  } else if (execution.error) {
    reason = `代码执行遇到异常 (${execution.error})，不可存为常用宏`;
  } else if (execution.precheckStatus === 'failed') {
    reason = '代码编译预检失败，不可存为常用宏';
  } else if (execution.readback?.otherWorkbooksAffected) {
    reason = '检测到修改了非目标工作簿，存在越权风险，不可存为常用宏';
  } else if (execution.verificationStatus === 'failed') {
    reason = '目标工作簿身份核验失败，不可存为常用宏';
  }

  return { isSaveEligible, reason };
}
