using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using LeeExcel;

namespace LeeExcelTests
{
    public class VerifyScriptHistoryResolution
    {
        private static int passed = 0;
        private static int failed = 0;

        private static void Assert(string testName, bool condition, string detail)
        {
            if (condition)
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("[PASS] " + testName);
                Console.ResetColor();
                passed++;
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("[FAIL] " + testName + (!string.IsNullOrEmpty(detail) ? " -> " + detail : ""));
                Console.ResetColor();
                failed++;
            }
        }

        private static string ComputeSha256(string text)
        {
            using (var sha = SHA256.Create())
            {
                byte[] bytes = Encoding.UTF8.GetBytes(text);
                return BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-", "").ToLowerInvariant();
            }
        }

        public static int Main(string[] args)
        {
            AppDomain.CurrentDomain.AssemblyResolve += (sender, resolveArgs) =>
            {
                try
                {
                    string asmName = new System.Reflection.AssemblyName(resolveArgs.Name).Name + ".dll";
                    string binPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, asmName);
                    if (File.Exists(binPath)) return System.Reflection.Assembly.LoadFrom(binPath);
                    string directBin = Path.GetFullPath(Path.Combine("bin", asmName));
                    if (File.Exists(directBin)) return System.Reflection.Assembly.LoadFrom(directBin);
                }
                catch { }
                return null;
            };

            return RunVerification(args);
        }

        private static int RunVerification(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;
            Console.WriteLine("================================================================================");
            Console.WriteLine("  ExcelMind AI: 宏库隔离、ResolveScriptFiles 身份歧义与隔离生命周期定向验证");
            Console.WriteLine("================================================================================");

            string sandboxDir = Path.Combine(Path.GetFullPath(".artifacts\\tests\\history_resolution_20261003_1335"), "sandbox");
            if (Directory.Exists(sandboxDir))
            {
                try { Directory.Delete(sandboxDir, true); } catch { }
            }
            Directory.CreateDirectory(sandboxDir);

            // 1. 设置沙箱隔离目录并执行安全阻断断言
            ScriptManager.SetCustomScriptsDirForTesting(sandboxDir);

            string defaultUserDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ExcelMindAI", "Scripts");
            string activeDir = ScriptManager.ScriptsDir;

            bool isIsolated = !string.Equals(Path.GetFullPath(activeDir), Path.GetFullPath(defaultUserDir), StringComparison.OrdinalIgnoreCase);
            Assert("测试宏库隔离断言：ScriptManager.ScriptsDir 严格指向临时沙箱，严禁指向正式用户宏库", isIsolated, "activeDir=" + activeDir + ", userDir=" + defaultUserDir);

            if (!isIsolated)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("CRITICAL: 沙箱隔离未生效，测试立即终止以保护用户宏库！");
                Console.ResetColor();
                return 1;
            }

            try
            {
                Console.WriteLine("  * 隔离沙箱路径: " + sandboxDir);

                // ============================================================================
                // 用例 1：反例验证——宏 A 文件名等于宏 B 元数据 ID
                // ============================================================================
                // 构造：
                // 宏 A: 文件名 macro_A.bas / macro_A.meta.json, 元数据 id = "macro_A_unique_id", displayName = "宏 A"
                // 宏 B: 文件名 macro_B.bas / macro_B.meta.json, 元数据 id = "macro_A",           displayName = "宏 B"
                string macroABas = Path.Combine(sandboxDir, "macro_A.bas");
                string macroAMeta = Path.Combine(sandboxDir, "macro_A.meta.json");
                string codeA = "Sub ProcA()\r\n    ' Macro A Body\r\nEnd Sub";
                File.WriteAllText(macroABas, codeA, Encoding.UTF8);
                var dictA = new Dictionary<string, string>
                {
                    { "id", "macro_A_unique_id" },
                    { "displayName", "宏 A (物理文件名 macro_A)" },
                    { "name", "宏 A" },
                    { "originalCodeHash", ComputeSha256(codeA) },
                    { "runHistory", "[]" }
                };
                File.WriteAllText(macroAMeta, SimpleJson.Serialize(dictA), Encoding.UTF8);

                string macroBBas = Path.Combine(sandboxDir, "macro_B.bas");
                string macroBMeta = Path.Combine(sandboxDir, "macro_B.meta.json");
                string codeB = "Sub ProcB()\r\n    ' Macro B Body\r\nEnd Sub";
                File.WriteAllText(macroBBas, codeB, Encoding.UTF8);
                var dictB = new Dictionary<string, string>
                {
                    { "id", "macro_A" }, // 宏 B 的元数据 ID 恰巧等于宏 A 的文件名！
                    { "displayName", "宏 B (ID 是 macro_A)" },
                    { "name", "宏 B" },
                    { "originalCodeHash", ComputeSha256(codeB) },
                    { "runHistory", "[]" }
                };
                File.WriteAllText(macroBMeta, SimpleJson.Serialize(dictB), Encoding.UTF8);

                // 查询稳定 ID "macro_A"：必须优先命中宏 B（因为宏 B 的元数据 ID == "macro_A"），严禁因文件名而误中宏 A！
                string resBas, resMeta, resId, resErr;
                bool okQueryB = ScriptManager.ResolveScriptFiles("macro_A", out resBas, out resMeta, out resId, out resErr);
                bool hitsMacroB = okQueryB && 
                                  Path.GetFileName(resBas).Equals("macro_B.bas", StringComparison.OrdinalIgnoreCase) &&
                                  Path.GetFileName(resMeta).Equals("macro_B.meta.json", StringComparison.OrdinalIgnoreCase) &&
                                  resId == "macro_A";

                Assert(
                    "反例验证 1：稳定 ID 优先于文件名（'宏 A 文件名 == 宏 B 元数据 ID' 正确命中宏 B 而非宏 A）",
                    hitsMacroB,
                    "resBas=" + resBas + ", resId=" + resId + ", err=" + resErr
                );

                // 追加历史记录至稳定 ID "macro_A"：必须写入宏 B 的元数据，宏 A 严禁被污染
                var recB = new RunRecordDto
                {
                    id = "run_rec_macro_b_001",
                    executedAt = "2026-10-03 14:00:00",
                    status = "success",
                    summary = "宏 B 执行成功"
                };
                string errAppendB;
                bool okAppendB = ScriptManager.AppendRunRecord("macro_A", recB, out errAppendB);

                var metaBAfter = SimpleJson.ParseFlatObject(File.ReadAllText(macroBMeta, Encoding.UTF8));
                var metaAAfter = SimpleJson.ParseFlatObject(File.ReadAllText(macroAMeta, Encoding.UTF8));

                bool bHasHistory = metaBAfter.ContainsKey("runHistory") && metaBAfter["runHistory"].Contains("run_rec_macro_b_001");
                bool aNotPolluted = !metaAAfter.ContainsKey("runHistory") || !metaAAfter["runHistory"].Contains("run_rec_macro_b_001");

                Assert(
                    "反例验证 1 扩展：使用稳定 ID 追加运行记录精准落入宏 B，宏 A 零污染",
                    okAppendB && bHasHistory && aNotPolluted,
                    "okAppend=" + okAppendB + ", bHas=" + bHasHistory + ", aClean=" + aNotPolluted + ", err=" + errAppendB
                );

                // ============================================================================
                // 用例 2：反例验证——多个元数据包含相同 ID 产生歧义冲突
                // ============================================================================
                // 构造：宏 C 与 宏 D 的元数据均包含相同的 ID "conflict_dup_id"
                string macroCBas = Path.Combine(sandboxDir, "macro_C.bas");
                string macroCMeta = Path.Combine(sandboxDir, "macro_C.meta.json");
                string codeC = "Sub ProcC()\r\nEnd Sub";
                File.WriteAllText(macroCBas, codeC, Encoding.UTF8);
                var dictC = new Dictionary<string, string>
                {
                    { "id", "conflict_dup_id" },
                    { "displayName", "宏 C" },
                    { "name", "宏 C" },
                    { "originalCodeHash", ComputeSha256(codeC) }
                };
                File.WriteAllText(macroCMeta, SimpleJson.Serialize(dictC), Encoding.UTF8);

                string macroDBas = Path.Combine(sandboxDir, "macro_D.bas");
                string macroDMeta = Path.Combine(sandboxDir, "macro_D.meta.json");
                string codeD = "Sub ProcD()\r\nEnd Sub";
                File.WriteAllText(macroDBas, codeD, Encoding.UTF8);
                var dictD = new Dictionary<string, string>
                {
                    { "id", "conflict_dup_id" }, // 与宏 C 相同的 ID
                    { "displayName", "宏 D" },
                    { "name", "宏 D" },
                    { "originalCodeHash", ComputeSha256(codeD) }
                };
                File.WriteAllText(macroDMeta, SimpleJson.Serialize(dictD), Encoding.UTF8);

                string resConflictBas, resConflictMeta, resConflictId, resConflictErr;
                bool okConflict = ScriptManager.ResolveScriptFiles("conflict_dup_id", out resConflictBas, out resConflictMeta, out resConflictId, out resConflictErr);

                bool conflictBlocked = !okConflict && 
                                       !string.IsNullOrEmpty(resConflictErr) && 
                                       resConflictErr.Contains("宏标识存在重复冲突") && 
                                       resConflictBas == null;

                Assert(
                    "反例验证 2：多元数据同 ID 明确阻断（严禁静默拾取第一个，抛出清晰歧义错误）",
                    conflictBlocked,
                    "ok=" + okConflict + ", err=" + resConflictErr
                );

                // ============================================================================
                // 用例 3：隔离生命周期完整流程（导入 → 标签/收藏更新 → 重命名 → 删除）
                // ============================================================================
                string lcId = "macro_lifecycle_001";
                string lcName = "生命周期验证宏";
                string lcCode = "Sub LifecycleTestProc()\r\n    Debug.Print \"Lifecycle\"\r\nEnd Sub";
                string savedId, errSave;

                // 3.1 导入 (SaveScript)
                bool okSave = ScriptManager.SaveScript(
                    id: lcId,
                    displayName: lcName,
                    code: lcCode,
                    description: "测试完整生命周期",
                    category: "测试",
                    sourceType: "import",
                    originalFileName: "LifecycleTest.bas",
                    encoding: "UTF-8",
                    entryPoint: "LifecycleTestProc",
                    rawBytesBase64: "",
                    overwrite: true,
                    savedId: out savedId,
                    error: out errSave
                );

                string lcBas = Path.Combine(sandboxDir, lcId + ".bas");
                string lcMeta = Path.Combine(sandboxDir, lcId + ".meta.json");
                bool filesExistAfterSave = File.Exists(lcBas) && File.Exists(lcMeta);

                Assert(
                    "隔离生命周期 1 - 导入宏：文件正确创建，元数据初始完整",
                    okSave && filesExistAfterSave && savedId == lcId,
                    "okSave=" + okSave + ", errSave=" + errSave
                );

                // 3.2 标签与收藏更新 (UpdateScriptTags / UpdateScriptFavorite)
                string errTag, errFav;
                var testTags = new List<string> { "审计", "2026年度", "高优先级" };
                bool okTag = ScriptManager.UpdateScriptTags(lcId, testTags, out errTag);
                bool okFav = ScriptManager.UpdateScriptFavorite(lcId, true, out errFav);

                var metaAfterTagFav = SimpleJson.ParseFlatObject(File.ReadAllText(lcMeta, Encoding.UTF8));
                bool tagUpdated = metaAfterTagFav.ContainsKey("tags") && metaAfterTagFav["tags"].Contains("2026年度");
                bool favUpdated = metaAfterTagFav.ContainsKey("isFavorite") && metaAfterTagFav["isFavorite"] == "true";
                // 源码与哈希严格保持不变
                string codeAfterTagFav = File.ReadAllText(lcBas, Encoding.UTF8);
                bool codeUnchanged1 = codeAfterTagFav == lcCode;

                Assert(
                    "隔离生命周期 2 - 标签与收藏更新：元数据增量更新，源码及哈希保持不变",
                    okTag && okFav && tagUpdated && favUpdated && codeUnchanged1,
                    "okTag=" + okTag + ", okFav=" + okFav + ", errTag=" + errTag + ", errFav=" + errFav
                );

                // 3.3 追加执行历史 (AppendRunRecord)
                var lcRec = new RunRecordDto
                {
                    id = "run_lc_001",
                    executedAt = "2026-10-03 14:05:00",
                    status = "success",
                    phase = "execution",
                    summary = "生命周期执行成功"
                };
                string errRec;
                bool okRec = ScriptManager.AppendRunRecord(lcId, lcRec, out errRec);
                var metaAfterRec = SimpleJson.ParseFlatObject(File.ReadAllText(lcMeta, Encoding.UTF8));
                bool recPresent = metaAfterRec.ContainsKey("runHistory") && metaAfterRec["runHistory"].Contains("run_lc_001");

                Assert(
                    "隔离生命周期 3 - 运行历史追加：成功落库并与宏关联",
                    okRec && recPresent,
                    "okRec=" + okRec + ", errRec=" + errRec
                );

                // 3.4 重命名 (RenameScript)
                string newName = "生命周期验证宏_已更名";
                string errRename;
                bool okRename = ScriptManager.RenameScript(lcId, newName, out errRename);
                var metaAfterRename = SimpleJson.ParseFlatObject(File.ReadAllText(lcMeta, Encoding.UTF8));
                bool nameUpdated = metaAfterRename.ContainsKey("displayName") && metaAfterRename["displayName"] == newName;
                bool idPreserved = metaAfterRename.ContainsKey("id") && metaAfterRename["id"] == lcId;
                bool historyPreserved = metaAfterRename.ContainsKey("runHistory") && metaAfterRename["runHistory"].Contains("run_lc_001");
                string codeAfterRename = File.ReadAllText(lcBas, Encoding.UTF8);
                bool codeUnchanged2 = codeAfterRename == lcCode;

                Assert(
                    "隔离生命周期 4 - 重命名：显示名称更新，ID 不变，运行历史与源码保真",
                    okRename && nameUpdated && idPreserved && historyPreserved && codeUnchanged2,
                    "okRename=" + okRename + ", err=" + errRename
                );

                // 3.5 删除 (DeleteScript)
                string errDel;
                bool okDel = ScriptManager.DeleteScript(lcId, out errDel);
                bool filesDeleted = !File.Exists(lcBas) && !File.Exists(lcMeta);

                Assert(
                    "隔离生命周期 5 - 删除：源码与元数据完全移除",
                    okDel && filesDeleted,
                    "okDel=" + okDel + ", filesDeleted=" + filesDeleted + ", err=" + errDel
                );

                // 3.6 零污染验证：检查其他宏（Macro A, Macro B, Macro C, Macro D）未受影响
                var metaAEnd = SimpleJson.ParseFlatObject(File.ReadAllText(macroAMeta, Encoding.UTF8));
                var metaBEnd = SimpleJson.ParseFlatObject(File.ReadAllText(macroBMeta, Encoding.UTF8));
                string codeAEnd = File.ReadAllText(macroABas, Encoding.UTF8);
                string codeBEnd = File.ReadAllText(macroBBas, Encoding.UTF8);

                bool macroAIntact = File.Exists(macroABas) && File.Exists(macroAMeta) && codeAEnd == codeA && metaAEnd["displayName"] == "宏 A (物理文件名 macro_A)";
                bool macroBIntact = File.Exists(macroBBas) && File.Exists(macroBMeta) && codeBEnd == codeB && metaBEnd["runHistory"].Contains("run_rec_macro_b_001");
                bool macroCIntact = File.Exists(macroCBas) && File.Exists(macroCMeta);
                bool macroDIntact = File.Exists(macroDBas) && File.Exists(macroDMeta);

                Assert(
                    "隔离生命周期 6 - 零副作用验证：全生命周期执行后，其他宏及历史绝对未受任何波及",
                    macroAIntact && macroBIntact && macroCIntact && macroDIntact,
                    "A=" + macroAIntact + ", B=" + macroBIntact + ", C=" + macroCIntact + ", D=" + macroDIntact
                );

                // ============================================================================
                // 用例 4：正式用户宏库保护验证（真实文件哈希未变）
                // ============================================================================
                string userAuditBas = Path.Combine(defaultUserDir, "macro_20261003_090144_873_2a40.bas");
                string expectedHash = "45368e4c1cb2dd133bef04c562a382a977b5e851d1555b1911f31da84743e8c4";
                string currentHash = "";
                if (File.Exists(userAuditBas))
                {
                    using (var sha = SHA256.Create())
                    {
                        currentHash = BitConverter.ToString(sha.ComputeHash(File.ReadAllBytes(userAuditBas))).Replace("-", "").ToLowerInvariant();
                    }
                }
                Assert(
                    "正式用户宏库保护：真实用户财务审计宏 (macro_20261003_090144_873_2a40.bas) 源码哈希严格恒等未被修改",
                    currentHash.Equals(expectedHash, StringComparison.OrdinalIgnoreCase),
                    "current=" + currentHash + ", expected=" + expectedHash
                );
            }
            finally
            {
                // 恢复默认配置
                ScriptManager.ResetCustomScriptsDirForTesting();
                // 清理沙箱测试目录
                try
                {
                    if (Directory.Exists(sandboxDir))
                    {
                        Directory.Delete(sandboxDir, true);
                    }
                }
                catch { }
            }

            Console.WriteLine("================================================================================");
            Console.WriteLine("  验证结果统计: PASSED = " + passed + " / FAILED = " + failed);
            Console.WriteLine("================================================================================");

            return failed == 0 ? 0 : 1;
        }
    }
}
