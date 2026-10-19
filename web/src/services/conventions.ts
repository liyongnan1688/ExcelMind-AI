/**
 * 企业视觉与排版公约常量规范 (Corporate Conventions Specification)
 * 
 * 职责：
 * 1. 规范化标准商务排版公约（千分位、保留两位小数、负数红字、企业级调色盘）；
 * 2. 作为静态基线常量解耦于提示词引擎，保持系统静态 System Prompt 100% 字节恒定（Prompt Caching 保障）；
 * 3. 预留用户级/客户级扩展挂载能力。
 */

export interface CorporatePalette {
  deepNavy: string;     // 科技深蓝 (#182D7B)
  techBlue: string;     // 科技蓝 (#0068B7)
  iceBlue: string;      // 浅冰蓝 (#DBEFF9)
  softIceBlue: string;  // 浅冰蓝底色 (#F0F6FA)
}

export const CORPORATE_PALETTE: CorporatePalette = {
  deepNavy: '#182D7B',
  techBlue: '#0068B7',
  iceBlue: '#DBEFF9',
  softIceBlue: '#F0F6FA',
};

export const CORPORATE_CONVENTIONS_BLOCK = `【企业排版与视觉公约规范 (Corporate Conventions)】：
1. 数值与货币排版：包含金额的列建议采用标准千分位格式（"#,##0.00" 或 "¥#,##0.00"）；百分比保留 2 位小数（"0.00%"）；整数保持常规或千分位整数（"#,##0"）；
2. 负数排版：负数建议采用红色或括号样式（如 "#,##0.00;[Red]-#,##0.00" 或 "#,##0.00;[Red](#,##0.00)"）；
3. 商务表头与配色标准：
   - 标题表头填充推荐科技深蓝 RGB(24, 45, 123)（十六进制 #182D7B）或科技蓝 RGB(0, 104, 183)（十六进制 #0068B7），搭配纯白加粗文字；
   - 汇总行/合计行采用浅冰蓝底色 RGB(219, 239, 249)（十六进制 #DBEFF9）并加粗；
   - 边框采用细实线（xlContinuous）；
4. 对齐规范：文本靠左，数值靠右，日期、代码与状态居中，表头居中并设置 WrapText = True。`;
