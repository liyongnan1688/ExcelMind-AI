using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;

public class VerifyPayload
{
    public static void Main(string[] args)
    {
        Console.OutputEncoding = Encoding.UTF8;
        string inputJson = args != null && args.Length > 0 ? args[0] : "payload.json";
        string outDir = args != null && args.Length > 1 ? args[1] : Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "../../.artifacts/tests/payload_verification");
        outDir = Path.GetFullPath(outDir);
        if (!Directory.Exists(outDir)) Directory.CreateDirectory(outDir);

        if (!File.Exists(inputJson))
        {
            Console.WriteLine("Input payload file not found: " + inputJson);
            Console.WriteLine("Usage: VerifyPayload.exe <inputJsonPath> [outputDir]");
            return;
        }

        string json = File.ReadAllText(inputJson, Encoding.UTF8);

        // 如果外层带反斜杠转义，去除最外层一层转义
        if (json.StartsWith("{\\\""))
        {
            var sb = new StringBuilder();
            for (int i = 0; i < json.Length; i++)
            {
                if (json[i] == '\\' && i + 1 < json.Length)
                {
                    char next = json[i + 1];
                    if (next == '"') { sb.Append('"'); i++; continue; }
                    if (next == '\\') { sb.Append('\\'); i++; continue; }
                }
                sb.Append(json[i]);
            }
            json = sb.ToString();
        }

        var jss = new JavaScriptSerializer();
        jss.MaxJsonLength = 10000000;
        var dict = jss.Deserialize<Dictionary<string, object>>(json);
        Console.WriteLine("Parsed Keys: " + string.Join(", ", dict.Keys));

        object dataObj;
        Dictionary<string, object> data = dict;
        if (dict.TryGetValue("data", out dataObj) && dataObj is Dictionary<string, object>)
        {
            data = (Dictionary<string, object>)dataObj;
            Console.WriteLine("Data Keys: " + string.Join(", ", data.Keys));
        }

        string rawModel = data.ContainsKey("rawModelResponse") ? (string)data["rawModelResponse"] : "";
        string origCode = data.ContainsKey("originalVbaCode") ? (string)data["originalVbaCode"] : "";
        string execCode = data.ContainsKey("executedVbaCode") ? (string)data["executedVbaCode"] : "";
        string wrapper  = data.ContainsKey("wrapperCode") ? (string)data["wrapperCode"] : "";
        string origHash = data.ContainsKey("originalCodeHash") ? (string)data["originalCodeHash"] : "";
        string execHash = data.ContainsKey("executedCodeHash") ? (string)data["executedCodeHash"] : "";
        string error    = dict.ContainsKey("error") ? (string)dict["error"] : (data.ContainsKey("error") ? (string)data["error"] : "");
        string precheck = data.ContainsKey("precheckStatus") ? (string)data["precheckStatus"] : "";
        string phase    = data.ContainsKey("executionPhase") ? (string)data["executionPhase"] : "";

        File.WriteAllText(Path.Combine(outDir, "VERIFIED_STAGE1_RAW_MODEL.vba"), rawModel, Encoding.UTF8);
        File.WriteAllText(Path.Combine(outDir, "VERIFIED_STAGE2_ORIGINAL.vba"), origCode, Encoding.UTF8);
        File.WriteAllText(Path.Combine(outDir, "VERIFIED_STAGE3_EXECUTED.vba"), execCode, Encoding.UTF8);
        File.WriteAllText(Path.Combine(outDir, "VERIFIED_WRAPPER.vba"), wrapper, Encoding.UTF8);

        string calcOrigHash = ComputeSha256(origCode);
        string calcExecHash = ComputeSha256(execCode);

        Console.WriteLine("=== 原始响应与代码比对结果 ===");
        Console.WriteLine("recorded originalCodeHash:   " + origHash);
        Console.WriteLine("calculated originalCodeHash: " + calcOrigHash);
        Console.WriteLine("MATCH? " + (origHash == calcOrigHash));
        Console.WriteLine("");
        Console.WriteLine("recorded executedCodeHash:   " + execHash);
        Console.WriteLine("calculated executedCodeHash: " + calcExecHash);
        Console.WriteLine("MATCH? " + (execHash == calcExecHash));
        Console.WriteLine("");
        Console.WriteLine("precheckStatus: " + precheck);
        Console.WriteLine("executionPhase: " + phase);
        Console.WriteLine("error message : " + error);
        Console.WriteLine("rawModelResponse contains 'Select': " + rawModel.Contains(".Select"));
        Console.WriteLine("originalVbaCode  contains 'Select': " + origCode.Contains(".Select"));
    }

    static string ComputeSha256(string text)
    {
        if (string.IsNullOrEmpty(text)) return "";
        using (var sha = SHA256.Create())
        {
            byte[] bytes = Encoding.UTF8.GetBytes(text);
            byte[] hash = sha.ComputeHash(bytes);
            var sb = new StringBuilder();
            foreach (byte b in hash) sb.Append(b.ToString("x2"));
            return sb.ToString();
        }
    }
}
