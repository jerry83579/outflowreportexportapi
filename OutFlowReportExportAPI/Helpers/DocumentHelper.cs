using Newtonsoft.Json.Linq;
using OpenDocumentLib.doc;
using OpenDocumentLib.sheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace OutFlowReportExportAPI.Helpers
{
    /// <summary>
    /// 文件操作功能
    /// </summary>
    public class DocumentHelper
    {
        /// <summary>
        /// 產生資料內容
        /// </summary>
        /// <param name="tmp">樣板路徑</param>
        /// <param name="data">資料</param>
        /// <param name="filename">要回傳的檔名</param>
        /// <returns>回傳檔案路徑</returns>
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
        /// 從資料庫取內容
        /// </summary>
        /// <param name="tmp"></param>
        /// <param name="data"></param>
        /// <param name="filename"></param>
        /// <returns></returns>
        public static string GetRptDatabase(string tmp, List<dynamic> databaseData, List<dynamic> repeatData,string filename)
        {
            var tDoc = OpenTemplate(tmp);
            AppendLog(new string[] { $"{tmp} has data? {databaseData != null}, template opened?{tDoc != null}" });
            if (null == tDoc || null == databaseData)
                return "";
            else
            {
                string path = Path.Combine(HttpRuntime.BinDirectory, "..", "App_Data", "temp", "appendix", databaseData.FirstOrDefault().OFP_No, $"{filename}");
                if (!Directory.Exists(Path.GetDirectoryName(path)))
                    Directory.CreateDirectory(Path.GetDirectoryName(path));

                // 寫入資料
                if (repeatData != null)
                {
                    var index = 0;
                    foreach (var data in repeatData)
                    {
                        foreach (var item in data)
                        {
                            var value = item.Value ?? "";
                            ReplaceText(tDoc, $":{item.Key}{index}", ConvertData(value));
                        }
                        index++;
                    }
                }
                foreach (var data in databaseData)
                {
                    foreach (var item in data)
                    {
                        var value = item.Value ?? "";
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
        /// 判斷資料型態輸出
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="value"></param>
        /// <returns></returns>
        public static string ConvertData<T>(T v)
        {
            string value;
            switch (v.GetType().Name)
            {
                case "String":
                        return v.ToString();
                case "Double":
                        return v.ToString();
                case "Boolean":
                        return v.ToString() == "True" ? "是" : "否";
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
                lines[i] = $"{ DateTime.Now:yyyy-MM-dd HH:mm:ss} -- {lines[i]}";
            }

            string path = Path.Combine(HttpRuntime.BinDirectory, "..", "App_Data", "report.log");

            var fi = new FileInfo(path);
            if (fi.Exists && fi.Length > 102400)
                fi.Delete();

            File.AppendAllLines(path, lines);
        }
    }
}