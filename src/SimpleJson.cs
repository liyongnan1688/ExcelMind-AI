using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

namespace LeeExcel
{
    /// <summary>
    /// 轻量级 JSON 工具，无需第三方外部依赖
    /// </summary>
    public static class SimpleJson
    {
        public static string Escape(string s)
        {
            if (s == null) return "";
            var sb = new StringBuilder();
            foreach (char c in s)
            {
                switch (c)
                {
                    case '\\': sb.Append("\\\\"); break;
                    case '\"': sb.Append("\\\""); break;
                    case '\n': sb.Append("\\n"); break;
                    case '\r': sb.Append("\\r"); break;
                    case '\t': sb.Append("\\t"); break;
                    default:
                        if (c < 32)
                            sb.AppendFormat("\\u{0:x4}", (int)c);
                        else
                            sb.Append(c);
                        break;
                }
            }
            return sb.ToString();
        }

        public static string Unescape(string s)
        {
            if (s == null) return "";
            return Regex.Replace(s, @"\\(u[0-9a-fA-F]{4}|[""\\/bfnrt])", m =>
            {
                string v = m.Value;
                if (v.StartsWith("\\u"))
                {
                    return ((char)Convert.ToInt32(v.Substring(2), 16)).ToString();
                }
                switch (v)
                {
                    case "\\\"": return "\"";
                    case "\\\\": return "\\";
                    case "\\/": return "/";
                    case "\\b": return "\b";
                    case "\\f": return "\f";
                    case "\\n": return "\n";
                    case "\\r": return "\r";
                    case "\\t": return "\t";
                    default: return v;
                }
            });
        }

        public static string Serialize(object obj)
        {
            if (obj == null) return "null";
            if (obj is string) return "\"" + Escape(obj as string) + "\"";
            if (obj is bool) return (bool)obj ? "true" : "false";
            if (obj is int || obj is long || obj is double || obj is float) return obj.ToString();

            var dict = obj as System.Collections.IDictionary;
            if (dict != null)
            {
                var sbDict = new StringBuilder("{");
                bool first = true;
                foreach (System.Collections.DictionaryEntry de in dict)
                {
                    if (!first) sbDict.Append(",");
                    first = false;
                    string k = de.Key != null ? de.Key.ToString() : "";
                    sbDict.Append("\"").Append(Escape(k)).Append("\":");
                    sbDict.Append(Serialize(de.Value));
                }
                sbDict.Append("}");
                return sbDict.ToString();
            }

            if (obj is System.Collections.IEnumerable && !(obj is IDictionary<string, object>))
            {
                var sbArr = new StringBuilder("[");
                bool first = true;
                foreach (var item in (System.Collections.IEnumerable)obj)
                {
                    if (!first) sbArr.Append(",");
                    first = false;
                    sbArr.Append(Serialize(item));
                }
                sbArr.Append("]");
                return sbArr.ToString();
            }

            var sb = new StringBuilder("{");
            var props = obj.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance);
            bool isFirst = true;
            foreach (var p in props)
            {
                if (!p.CanRead) continue;
                object val = p.GetValue(obj, null);
                if (!isFirst) sb.Append(",");
                isFirst = false;
                sb.Append("\"").Append(p.Name).Append("\":");
                sb.Append(Serialize(val));
            }
            sb.Append("}");
            return sb.ToString();
        }

        public static Dictionary<string, string> ParseFlatObject(string json)
        {
            var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (string.IsNullOrEmpty(json)) return dict;
            json = json.Trim();
            if (json.StartsWith("{")) json = json.Substring(1);
            if (json.EndsWith("}")) json = json.Substring(0, json.Length - 1);

            int i = 0;
            int len = json.Length;

            while (i < len)
            {
                // 跳过空白和逗号
                while (i < len && (char.IsWhiteSpace(json[i]) || json[i] == ',')) i++;
                if (i >= len) break;

                // 期望 key 开始
                if (json[i] != '"') { i++; continue; }
                i++; // 跳过开头的 "
                int keyStart = i;
                while (i < len && json[i] != '"')
                {
                    if (json[i] == '\\' && i + 1 < len) i += 2;
                    else i++;
                }
                string key = Unescape(json.Substring(keyStart, i - keyStart));
                if (i < len && json[i] == '"') i++; // 跳过结束的 "

                // 查找冒号 :
                while (i < len && json[i] != ':') i++;
                if (i < len && json[i] == ':') i++; // 跳过 :

                // 跳过空白
                while (i < len && char.IsWhiteSpace(json[i])) i++;
                if (i >= len) break;

                string val = "";
                if (json[i] == '"')
                {
                    // 字符串类型：扫描直到未转义的闭合引号
                    i++; // 跳过开头的 "
                    var sb = new StringBuilder();
                    while (i < len)
                    {
                        if (json[i] == '\\' && i + 1 < len)
                        {
                            char next = json[i + 1];
                            switch (next)
                            {
                                case '"': sb.Append('"'); break;
                                case '\\': sb.Append('\\'); break;
                                case '/': sb.Append('/'); break;
                                case 'b': sb.Append('\b'); break;
                                case 'f': sb.Append('\f'); break;
                                case 'n': sb.Append('\n'); break;
                                case 'r': sb.Append('\r'); break;
                                case 't': sb.Append('\t'); break;
                                case 'u':
                                    if (i + 5 < len)
                                    {
                                        string hex = json.Substring(i + 2, 4);
                                        int code;
                                        if (int.TryParse(hex, System.Globalization.NumberStyles.HexNumber, null, out code))
                                        {
                                            sb.Append((char)code);
                                            i += 6;
                                            continue;
                                        }
                                    }
                                    sb.Append("\\u");
                                    break;
                                default:
                                    sb.Append(next);
                                    break;
                            }
                            i += 2;
                        }
                        else if (json[i] == '"')
                        {
                            i++; // 跳过闭合 "
                            break;
                        }
                        else
                        {
                            sb.Append(json[i]);
                            i++;
                        }
                    }
                    val = sb.ToString();
                }
                else if (json[i] == '{' || json[i] == '[')
                {
                    // 复合对象或数组：深度平衡扫描直到闭合
                    int valStart = i;
                    char openChar = json[i];
                    char closeChar = openChar == '{' ? '}' : ']';
                    int depth = 0;
                    bool inString = false;
                    while (i < len)
                    {
                        char c = json[i];
                        if (inString)
                        {
                            if (c == '\\' && i + 1 < len) i += 2;
                            else
                            {
                                if (c == '"') inString = false;
                                i++;
                            }
                        }
                        else
                        {
                            if (c == '"') inString = true;
                            else if (c == openChar) depth++;
                            else if (c == closeChar)
                            {
                                depth--;
                                if (depth == 0)
                                {
                                    i++;
                                    break;
                                }
                            }
                            i++;
                        }
                    }
                    val = json.Substring(valStart, i - valStart).Trim();
                }
                else
                {
                    // 非字符串基础类型 (布尔值、数字、null 等)
                    int valStart = i;
                    while (i < len && json[i] != ',' && json[i] != '}') i++;
                    val = json.Substring(valStart, i - valStart).Trim();
                }

                dict[key] = val;
            }

            return dict;
        }

        public static List<T> DeserializeList<T>(string json) where T : new()
        {
            var list = new List<T>();
            if (string.IsNullOrEmpty(json)) return list;
            json = json.Trim();
            if (!json.StartsWith("[") || !json.EndsWith("]")) return list;

            int i = 1;
            int len = json.Length - 1;

            while (i < len)
            {
                while (i < len && (char.IsWhiteSpace(json[i]) || json[i] == ',')) i++;
                if (i >= len || json[i] != '{') break;

                // 精准匹配一个完整的顶层对象 { ... }
                int start = i;
                bool inStr = false;
                int braceDepth = 0;
                while (i < len)
                {
                    char c = json[i];
                    if (inStr)
                    {
                        if (c == '\\' && i + 1 < len) i += 2;
                        else
                        {
                            if (c == '"') inStr = false;
                            i++;
                        }
                    }
                    else
                    {
                        if (c == '"') inStr = true;
                        else if (c == '{') braceDepth++;
                        else if (c == '}')
                        {
                            braceDepth--;
                            if (braceDepth == 0)
                            {
                                i++;
                                break;
                            }
                        }
                        i++;
                    }
                }

                string objJson = json.Substring(start, i - start);
                var dict = ParseFlatObject(objJson);
                T item = new T();
                foreach (var p in typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance))
                {
                    if (dict.ContainsKey(p.Name) && p.CanWrite)
                    {
                        if (p.PropertyType == typeof(string))
                            p.SetValue(item, dict[p.Name], null);
                        else if (p.PropertyType == typeof(bool))
                            p.SetValue(item, dict[p.Name].ToLower() == "true", null);
                        else if (p.PropertyType == typeof(int))
                        {
                            int iv;
                            if (int.TryParse(dict[p.Name], out iv)) p.SetValue(item, iv, null);
                        }
                    }
                }
                list.Add(item);
            }
            return list;
        }
    }
}
