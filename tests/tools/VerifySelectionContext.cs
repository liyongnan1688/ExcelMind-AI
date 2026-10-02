using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using LeeExcel;

namespace LeeExcelTests
{
    public class VerifySelectionContext
    {
        private static int _pass = 0;
        private static int _fail = 0;

        private static void Assert(bool cond, string name, string detail = "")
        {
            if (cond)
            {
                Console.WriteLine("[PASS] " + name + (string.IsNullOrEmpty(detail) ? "" : " -> " + detail));
                _pass++;
            }
            else
            {
                Console.WriteLine("[FAIL] " + name + (string.IsNullOrEmpty(detail) ? "" : " -> " + detail));
                _fail++;
            }
        }

        public static int Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;
            Console.WriteLine("================================================================================");
            Console.WriteLine("       ExcelMind AI: R1a SelectionContextService 单元与数据契约验证        ");
            Console.WriteLine("================================================================================");

            // 1. 空 app 防御性测试
            var nullRes = SelectionContextService.GetSelectionContext(null);
            Assert(!nullRes.ok, "空宿主 app 拦截", "errorType=" + nullRes.errorType);
            Assert(nullRes.errorType == "no_workbook", "错误类型为 no_workbook");

            // 2. 数据结构构建与 SimpleJson 序列化测试
            var testData = new SelectionContextData
            {
                workbookName = "测试表_2026.xlsx",
                workbookFullName = @"C:\Users\35651\Documents\测试表_2026.xlsx",
                sheetName = "Sheet1",
                address = "$D$5:$F$10",
                totalRows = 6,
                totalColumns = 3,
                startRow = 5,
                startColumn = 4,
                endRow = 10,
                endColumn = 6,
                isSingleArea = true,
                sampleRowCount = 2,
                sampleColumnCount = 3,
                sampleAddress = "$D$5:$F$6",
                candidateHeaders = new List<string> { "部门", "预算", "公式" },
                sampleRows = new List<List<CellSampleItem>>
                {
                    new List<CellSampleItem>
                    {
                        new CellSampleItem { row = 5, col = 4, address = "$D$5", value = "部门", displayText = "部门", valueType = "string" },
                        new CellSampleItem { row = 5, col = 5, address = "$E$5", value = "预算", displayText = "预算", valueType = "string" },
                        new CellSampleItem { row = 5, col = 6, address = "$F$5", value = "公式", displayText = "公式", valueType = "string" }
                    },
                    new List<CellSampleItem>
                    {
                        new CellSampleItem { row = 6, col = 4, address = "$D$6", value = "研发", displayText = "研发", valueType = "string" },
                        new CellSampleItem { row = 6, col = 5, address = "$E$6", value = 12000.0, displayText = "12000", valueType = "number" },
                        new CellSampleItem { row = 6, col = 6, address = "$F$6", value = 13800.0, displayText = "13800", formula = "=E6*1.15", valueType = "number" }
                    }
                },
                formulaStatus = "sample_mixed",
                mergeStatus = "no_merged",
                visibilityStatus = "sample_scanned_only",
                isRowTruncated = true,
                isColumnTruncated = false,
                maxTextLengthLimit = 100,
                unscannedNotes = "选区共 6 行 × 3 列。本次仅安全抽样前 2 行 × 前 3 列。"
            };

            string json = SimpleJson.Serialize(testData);
            Assert(!string.IsNullOrEmpty(json), "JSON 序列化成功", "长度=" + json.Length);
            Assert(json.Contains("\"workbookName\":\"测试表_2026.xlsx\""), "包含工作簿名称");
            Assert(json.Contains("\"address\":\"$D$5:$F$10\""), "包含完整区域地址");
            Assert(json.Contains("\"candidateHeaders\":[\"部门\",\"预算\",\"公式\"]"), "包含候选表头");
            Assert(json.Contains("\"formula\":\"=E6*1.15\""), "包含单元格公式");
            Assert(json.Contains("\"sampleRows\":[[{"), "包含二维嵌套样本数据");

            // 3. 采样上限常量校验
            Assert(SelectionContextService.DefaultMaxSampleRows == 5, "默认采样最大行数为 5");
            Assert(SelectionContextService.DefaultMaxSampleCols == 15, "默认采样最大列数为 15");
            Assert(SelectionContextService.MaxTextCharLimit == 100, "单单元格最大字符截断为 100");

            Console.WriteLine(string.Format("\n测试汇总: 通过 = {0}, 失败 = {1}", _pass, _fail));
            return _fail > 0 ? 1 : 0;
        }
    }
}
