using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Web;

namespace OutFlowReportExportAPI.Helpers
{
    public class HttpHelper
    {
        private static readonly HttpClient _client = new HttpClient();
        private static readonly string _outFlowBaseUrl = System.Web.Configuration.WebConfigurationManager.AppSettings["outFlowBaseUrl"];

        /// <summary>
        /// 取得API回傳資訊，使用在不需要登入取得的資訊
        /// </summary>
        /// <param name="method"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        public static string GetResponse(string method, object data)
        {
            var timeStemp = GetTimeStamp();
            var encryptData = GetEncryptUrlString(timeStemp, data);

            var path = Path.Combine(_outFlowBaseUrl, method) + $"?_={timeStemp}&d={encryptData}";

            var response = _client.GetStreamAsync(path).Result;

            return GetUnGzipString(response);
        }

        /// <summary>
        /// 透過Token登入後並取得回傳資料，使用在需要登入取得的資訊
        /// </summary>
        /// <param name="method"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        public static string GetResponseWithToken(string method, object data)
        {
            Dictionary<string, string> cookies = new Dictionary<string, string>() { { "06-outflowtoken", "6087323f-9d71-42c0-9422-40881a8e0060" } };

            // 加密資料
            var timeStemp = GetTimeStamp();
            var encryptData = GetEncryptUrlString(timeStemp, data);
            // 
            var path = Path.Combine(_outFlowBaseUrl, method) + $"?_={timeStemp}&d={encryptData}";


            var cookieContainer = new CookieContainer();
            using (var handler = new HttpClientHandler() { CookieContainer = cookieContainer })
            using (var client = new HttpClient(handler))
            {
                foreach(var cookie in cookies)
                {
                    cookieContainer.Add(new Uri(path), new Cookie(cookie.Key, cookie.Value));
                }

                var response = client.GetStreamAsync(path).Result;
                return GetUnGzipString(response);
            }
        }

        /// <summary>
        /// 透過網址取得回傳資訊
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>
        public static string GetResponse(string url)
        {
            var response = _client.GetStreamAsync(url).Result;

            return GetUnGzipString(response);
        }

        /// <summary>
        /// 透過網址取得回傳資訊
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>
        public static string GetResponseWithToken(string url)
        {
            Dictionary<string, string> cookies = new Dictionary<string, string>() { { "06-outflowtoken", "6087323f-9d71-42c0-9422-40881a8e0060" } };

            var cookieContainer = new CookieContainer();
            using (var handler = new HttpClientHandler() { CookieContainer = cookieContainer })
            using (var client = new HttpClient(handler))
            {
                foreach (var cookie in cookies)
                {
                    cookieContainer.Add(new Uri(url), new Cookie(cookie.Key, cookie.Value));
                }

                var response = client.GetStreamAsync(url).Result;
                return GetUnGzipString(response);
            }
        }

        /// <summary>
        /// 透過Route傳輸資料，
        /// </summary>
        /// <param name="method"></param>
        /// <param name="datas"></param>
        /// <returns></returns>
        public static string GetResponseWithRouteData(string method, string[] datas)
        {
            var path = Path.Combine(_outFlowBaseUrl, method);

            foreach (var data in datas)
            {
                path = Path.Combine(path, data);
            }

            var response = _client.GetStreamAsync(path).Result;

            return GetUnGzipString(response);
        }

        /// <summary>
        /// 透過Route傳輸資料，使用在需要登入取得的資訊
        /// </summary>
        /// <param name="method"></param>
        /// <param name="datas"></param>
        /// <returns></returns>
        public static string GetResponseWithRouteDataAndToken(string method, string[] datas)
        {

            Dictionary<string, string> cookies = new Dictionary<string, string>() { { "06-outflowtoken", "6087323f-9d71-42c0-9422-40881a8e0060" } };

            var cookieContainer = new CookieContainer();
            using (var handler = new HttpClientHandler() { CookieContainer = cookieContainer })
            using (var client = new HttpClient(handler))
            {
                var path = Path.Combine(_outFlowBaseUrl, method);

                foreach (var data in datas)
                {
                    path = Path.Combine(path, data);
                }

                foreach (var cookie in cookies)
                {
                    cookieContainer.Add(new Uri(path), new Cookie(cookie.Key, cookie.Value));
                }

                var response = client.GetStreamAsync(path).Result;
                return GetUnGzipString(response);
            }
        }

        /// <summary>
        /// 回傳檔案資訊
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="mime"></param>
        /// <returns></returns>
        public static HttpResponseMessage FileResult(string filePath, string mime = "application/octet-stream")
        {
            var result = new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StreamContent(File.OpenRead(filePath))
            };
            result.Content.Headers.ContentDisposition = new System.Net.Http.Headers.ContentDispositionHeaderValue("attachment")
            {
                FileName = Path.GetFileName(filePath)
            };
            result.Content.Headers.ContentType = new MediaTypeHeaderValue(mime);

            return result;
        }

        public static HttpResponseMessage FailResult(string msg)
        {
            return new HttpResponseMessage(HttpStatusCode.BadRequest)
            {
                Content = new StringContent(msg)
            };
        }


        public static long GetTimeStamp()
        {
            return new DateTimeOffset(DateTime.Now).ToUnixTimeMilliseconds();
        }

        /// <summary>
        /// 取得加密字串
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public static string GetEncryptUrlString(long timeStamp, object data)
        {
            var hex = timeStamp.ToString("x");

            return hex + EncodeData(HttpUtility.UrlEncode(JsonConvert.SerializeObject(data))) + hex;
        }

        /// <summary>
        /// 壓縮資料加密
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public static string GetUnGzipString(Stream input)
        {
            var inputByte = ReadFully(input);
            var decompressByte = Decompress(inputByte);

            using (var ms = new MemoryStream(decompressByte))
            using (var streamReader = new StreamReader(ms))
            {
                return streamReader.ReadToEnd();
            }
        }


        public static byte[] ReadFully(Stream input)
        {
            byte[] buffer = new byte[16 * 1024];
            using (MemoryStream ms = new MemoryStream())
            {
                int read;
                while ((read = input.Read(buffer, 0, buffer.Length)) > 0)
                {
                    ms.Write(buffer, 0, read);
                }
                return ms.ToArray();
            }
        }

        public static byte[] Decompress(byte[] data)
        {
            using (var compressedStream = new MemoryStream(data))
            using (var zipStream = new GZipStream(compressedStream, CompressionMode.Decompress))
            using (var resultStream = new MemoryStream())
            {
                zipStream.CopyTo(resultStream);
                return resultStream.ToArray();
            }
        }

        private static string EncodeData(string data)
        {
            var sb = new StringBuilder();
            for (int i = 0; i < data.Length; i++)
            {
                var code = ((int)data[i]).ToString("x");
                sb.Append(code);
            }

            return sb.ToString();
        }
    }
}