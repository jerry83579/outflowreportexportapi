using Newtonsoft.Json.Linq;
using OpenDocumentLib.doc;
using OpenDocumentLib.sheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Web;
using System.Web.Http;
using unoidl.com.sun.star.ucb;

namespace OutFlowReportExportAPI.Helpers
{
    /// <summary>
    /// 文件操作功能
    /// </summary>
    public class DocumentHelper
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="tmp"></param>
        /// <param name="data"></param>
        /// <param name="filename"></param>
        /// <returns></returns>
        public static string GetRptStream(string tmp, JObject data, string filename)
        {
            var tDoc = OpenTemplate(tmp);
            AppendLog(new string[] { $"{tmp} has data? {data != null}, template opened?{tDoc != null}" });
            if (null == tDoc || null == data)
                return "";
            else
            {
                string path = Path.Combine(HttpRuntime.BinDirectory, "..", "App_Data", "temp", "appendix", (string)data["OutflowControlPlan"]["OFP_NO"], $"{filename}");
                if (!Directory.Exists(Path.GetDirectoryName(path)))
                    Directory.CreateDirectory(Path.GetDirectoryName(path));

                foreach (var prop in data.Children<JProperty>())
                {
                    var d = prop.Value;
                    if (null == d) continue;
                    Type dt = d.GetType();
                    var mems = d.Children<JProperty>();

                    foreach (var mem in mems)
                    {
                        ReplaceText(tDoc, $":{prop.Name}.{mem.Name}", mem.Value?.Type == JTokenType.Date ? ((DateTime)mem.Value).ToString("yyyy/MM/dd") : mem.Value?.ToString() ?? "");
                    }
                }
                var now = DateTime.Now;
                ReplaceText(tDoc, ":now", $"民國{(now.Year - 1911):000}年{now.Month:00}月{now.Day:00}日");

                try
                {
                    AppendLog(new string[] { $"Generate word => {path}" });
                    tDoc.SaveAs(path);
                    return path;
                }
                finally
                {
                    tDoc.Dispose();
                }
            }
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="tmp">讀取的檔案名稱</param>
        /// <param name="databaseData">讀取資料庫的資料</param>
        /// <param name="filename">欲新增的檔案名稱</param>
        /// <returns></returns>
        public static string GetRptDatabase(string tmp, Dictionary<string, dynamic> databaseData, string filename)
        {
            var tDoc = OpenTemplate(tmp);
            AppendLog(new string[] { $"{tmp} has data? {databaseData != null}, template opened?{tDoc != null}" });
            if (null == tDoc || null == databaseData)
                return "";
            else
            {
                string path = Path.Combine(HttpRuntime.BinDirectory, "..", "App_Data", "temp", "appendix", databaseData["data"][0].OFP_No, $"{filename}");
                if (!Directory.Exists(Path.GetDirectoryName(path)))
                    Directory.CreateDirectory(Path.GetDirectoryName(path));
       
                // 重複資料
                if (databaseData["duplicateData"] != null)
                {
                    int index = 0;
                    string trueBox = "\u25A0";
                    string falseBox = "\u25A1";
                    foreach (var data in databaseData["duplicateData"])
                    {
                        foreach (var item in data)
                        {
                            var value = item.Value;
                            switch (item.Key)
                            {
                                case "IsQualified":
                                    if (value == true)
                                    {
                                        ReplaceText(tDoc, $":True{index}", ConvertData(trueBox));
                                        ReplaceText(tDoc, $":False{index}", ConvertData(falseBox));
                                    }
                                    else if (value == false)
                                    {
                                        ReplaceText(tDoc, $":True{index}", ConvertData(falseBox));
                                        ReplaceText(tDoc, $":False{index}", ConvertData(trueBox));
                                    }
                                    else
                                    {
                                        ReplaceText(tDoc, $":True{index}", ConvertData(falseBox));
                                        ReplaceText(tDoc, $":False{index}", ConvertData(falseBox));
                                    }
                                    break;
                                case "IsMatchPlan":
                                    if (value == null)
                                    {
                                        value = "";
                                        break;
                                    }
                                    value = value ? "是" : "否";
                                    break;
                                case "IsConductChangePlan":
                                    if (value == null)
                                    {
                                        value = "";
                                        break;
                                    }
                                    value = value ? "是" : "否";
                                    break;
                                case "PA_Num":
                                    if (value == null)
                                    {
                                        value = "";
                                        break;
                                    }
                                    value = Decry(value);
                                    break;
                                default:
                                    value = item.Value ?? "";
                                    break;
                            }
                            if (item.Key != "IsQualified")
                            {
                                ReplaceText(tDoc, $":{item.Key}{index}", ConvertData(value));
                            }
                        }
                        index++;
                    }

                    //  寫進空資料解決動態資料
                    List<string> listItemKey = new List<string>();
                    foreach (var data in databaseData["duplicateData"])
                    {
                        foreach (var item in data)
                        {
                            listItemKey.Add(item.Key);
                        }
                    }

                    for (var i = 0; i < 10; i++)
                    {
                        foreach (var item in listItemKey)
                        {
                            ReplaceText(tDoc, $":{item}{i}", "");
                            ReplaceText(tDoc, $":True{i}", "");
                            ReplaceText(tDoc, $":False{i}", "");
                        }
                    }
                }
                // 一般資料
                foreach (var data in databaseData["data"])
                {
                    foreach (var item in data)
                    {
                        var value = item.Value;
                        switch (item.Key)
                        {
                            case "PA_Num":
                                if (value == null)
                                {
                                    value = "";
                                    break;
                                }
                                value = Decry(value);
                                break;
                            case "FC_ResultStatus":
                                if (value == null)
                                {
                                    value = "";
                                    break;
                                }
                                value = value ? "已達完工標準" : "未達完工標準";
                                break;
                            case "MM_SVCheck_Status":
                                if (value == null)
                                {
                                    value = "";
                                    break;
                                }
                                value = value ? "是" : "否";
                                break;
                            case "MM_Check_Result":
                            case "MM_SVCheck_Result":
                                if (value == null)
                                {
                                    value = "";
                                    break;
                                }
                                    value = value ? "合格" : "不合格";
                                break;
                            default:
                                value = value ?? "";
                                break;
                        }
                        ReplaceText(tDoc, $":{item.Key}", ConvertData(value));
                    }
                }
                try
                {
                    AppendLog(new string[] { $"Generate word => {path}" });
                    tDoc.SaveAs(path);
                    return path;
                }
                finally
                {
                    tDoc.Dispose();
                }
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="v">要轉換的資料</param>
        /// <param name="b">是否有合格項目框</param>
        /// <returns></returns>
        public static string ConvertData<T>(T v)
        {
            string value;
            switch (v.GetType().Name)
            {
                case "String":
                case "Double":
                    return v.ToString();
                //case "Boolean":
                //        return v.ToString() == "True" ? "是" : "否";
                case "DateTime":
                    value = v.ToString();
                    int endIndex = value.ToString().IndexOf(" ");
                    string fullText = value.Substring(0, endIndex);
                    string[] DateTime = fullText.Split('/');
                    return $"民國{(Int32.Parse(DateTime[0]) - 1911)}年{DateTime[1]}月{DateTime[2]}日";
                default:
                    return "unknowType";
            }
        }

        /// <summary>
        /// 解密
        /// </summary>
        /// <param name="encry">透過Encry加密過的字串</param>
        /// <returns>解密內容</returns>
        public static string Decry(string encry)
        {
            var bytes = Convert.FromBase64String(encry);
            using (var ms = new MemoryStream(bytes))
            {
                using (var br = new BinaryReader(ms))
                {
                    var ivLen = br.ReadInt32();
                    var iv = br.ReadBytes(ivLen);
                    var keyLen = br.ReadInt32();
                    var key = br.ReadBytes(keyLen);
                    var r = Rijndael.Create();
                    r.Key = key;
                    r.IV = iv;
                    var trans = r.CreateDecryptor();
                    using (var cs = new CryptoStream(ms, trans, CryptoStreamMode.Read))
                    {
                        using (var rmv = new MemoryStream())
                        {
                            cs.CopyTo(rmv);
                            return System.Text.Encoding.UTF8.GetString(rmv.ToArray());
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 取代字串
        /// </summary>
        /// <param name="app"></param>
        /// <param name="ostr"></param>
        /// <param name="nstr"></param>
        private static void ReplaceText(Document app, string ostr, string nstr)
        {
            app.Replace(ostr, nstr);
        }

        private static Document OpenTemplate(string tmp)
        {
            string path = Path.GetFullPath(Path.Combine(HttpRuntime.BinDirectory, "..", "App_Data", "template", "appendix", $"{tmp}"));
            AppendLog(new string[] { $"Open Template => {path}" });
            return OpenWord(path);
        }
        /// <summary>
        /// 開啟樣板
        /// </summary>
        /// <param name="doc">word路徑</param>
        /// <returns></returns>
        private static Document OpenWord(string doc)
        {
            if (File.Exists(doc))
            {
                return new Document(doc);
            }
            else
                return null;
        }

        public static void AppendLog(string[] lines)
        {
            for (int i = 0; i < lines.Length; i++)
            {
                lines[i] = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} -- {lines[i]}";
            }

            string path = Path.Combine(HttpRuntime.BinDirectory, "..", "App_Data", "report.log");

            var fi = new FileInfo(path);
            if (fi.Exists && fi.Length > 102400)
                fi.Delete();

            File.AppendAllLines(path, lines);
        }
    }
}