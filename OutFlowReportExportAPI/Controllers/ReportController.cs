using Dapper;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OpenDocumentLib.sheet;
using OutFlowReportExportAPI.Dtos;
using OutFlowReportExportAPI.Helpers;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Web;
using System.Web.Http;

namespace OutFlowReportExportAPI.Controllers
{
    public class ReportController : ApiController
    {
        private static readonly string _filePath = HttpContext.Current.Server.MapPath($"~/Files/");
        private static readonly string _appDataPath = HttpContext.Current.Server.MapPath($"~/App_Data/");
        private static readonly string _reportBaseUrl = System.Web.Configuration.WebConfigurationManager.AppSettings["reportBaseUrl"];
        private static readonly string _outFlowBaseUrl = System.Web.Configuration.WebConfigurationManager.AppSettings["outFlowBaseUrl"];
        private static readonly string _templateFilePath = HttpContext.Current.Server.MapPath("~/App_Data/template/ods");

        public ReportController()
        {
            if (!Directory.Exists(_filePath))
                Directory.CreateDirectory(_filePath);
        }

        #region 取得統計資料總表
        /// <summary>
        /// 取得統計資料總表
        /// </summary>
        /// <param name="caseCountInfoDto"></param>
        /// <returns></returns>
        [HttpGet, ActionName(nameof(GetSummaryFile))]
        public HttpResponseMessage GetSummaryFile([FromUri] FilterParams filterParams)
        {
            string FileName = "出流管制案件總統計表";
            var filePath = Path.Combine(_filePath, "ReportTempFile", $"{FileName}.ods");
            try
            {
                var responseString = HttpHelper.GetResponseWithToken("case_summary.ods", new { filter = filterParams });
                CaseCountInfoDto caseCountInfoDto = JsonConvert.DeserializeObject<CaseCountInfoDto>(responseString);

                using (var calc = new Calc())
                {
                    var tb = calc.Tables.AddNew(FileName);

                    // 樣式
                    Font headerFont = new Font("微軟正黑體", 30, FontStyle.Bold),
                        colFont = new Font("微軟正黑體", 14, FontStyle.Bold),
                        rowFont = new Font("微軟正黑體", 12);
                    var line = new Line() { Color = Color.Black, OuterWidth = 20 };


                    //表格標題
                    tb[0, 0].Formula = FileName;
                    tb["A1:G1"].Merge = true;
                    tb["A1:G1"].HorizonAlign = Alignment.HorizonAlignment.CENTER;
                    tb["A1:G1"].VerticalAlign = Alignment.VerticalAlignment.CENTER;
                    tb.Rows[0].Height = 20;

                    #region 統計資訊
                    // 調整這邊就可以改變統計的上下
                    int totalRow = 2;
                    tb[totalRow, 0].Formula = "總計";
                    tb[0, totalRow, 0, totalRow + 1].Merge = true;
                    tb[0, totalRow, 0, totalRow + 1].HorizonAlign = Alignment.HorizonAlignment.CENTER;
                    tb[0, totalRow, 0, totalRow + 1].VerticalAlign = Alignment.VerticalAlignment.CENTER;
                    tb[0, 0].SetFont(headerFont);
                    tb[totalRow, 1].Formula = "總案件數";
                    tb[totalRow, 2].Formula = "已核定管制案";
                    tb[totalRow, 3].Formula = "實質審查階段";
                    tb[totalRow, 4].Formula = "義務人填報階段";
                    tb[totalRow, 5].Formula = "核定滯洪池座數";
                    tb[totalRow, 6].Formula = "核定滯洪池體積(萬m3)";

                    tb[totalRow + 1, 1].Value = caseCountInfoDto.caseTotal;
                    tb[totalRow + 1, 2].Value = caseCountInfoDto.checkTotal;
                    tb[totalRow + 1, 3].Value = caseCountInfoDto.nocheckTotal;
                    tb[totalRow + 1, 4].Value = caseCountInfoDto.nosendTotal;
                    tb[totalRow + 1, 5].Value = caseCountInfoDto.facilityNumTotal;
                    tb[totalRow + 1, 6].Value = caseCountInfoDto.facilityAreaTotal;
                    // 設定樣式
                    tb[0, totalRow, 6, totalRow + 1].SetFont(colFont);
                    #endregion


                    int startRow = 5; // 要調整明細位置只要改這個數值就好了

                    if (caseCountInfoDto.groupby == "OFP_Cou")
                    {
                        tb[startRow, 0].Formula = "縣市別";
                    }
                    else if (caseCountInfoDto.groupby == "OFP_Year")
                    {
                        tb[startRow, 0].Formula = "年度";
                    }

                    tb[startRow, 1].Formula = "總案件數";
                    tb[startRow, 2].Formula = "已核定管制案";
                    tb[startRow, 3].Formula = "實質審查階段";
                    tb[startRow, 4].Formula = "義務人填報階段";
                    tb[startRow, 5].Formula = "核定滯洪池座數";
                    tb[startRow, 6].Formula = "核定滯洪池體積(萬m3)";
                    tb[0, startRow, 6, startRow].SetFont(colFont);
                    foreach (var item in caseCountInfoDto.details)
                    {
                        startRow++;
                        tb[0, startRow, 6, startRow].SetFont(rowFont);

                        if (caseCountInfoDto.groupby == "OFP_Cou")
                        {
                            tb[startRow, 0].Formula = item["OFP_Cou"].ToString();
                        }
                        else if (caseCountInfoDto.groupby == "OFP_Year")
                        {
                            tb[startRow, 0].Formula = item["OFP_Year"].ToString();
                        }
                        tb[startRow, 1].Formula = item["CaseCount"].ToString();
                        tb[startRow, 2].Formula = item["IsCheckCount"].ToString();
                        tb[startRow, 3].Formula = item["NoCheckCount"].ToString();
                        tb[startRow, 4].Formula = item["NoSendCount"].ToString();
                        tb[startRow, 5].Formula = item["FacilityNum"].ToString();
                        tb[startRow, 6].Formula = $"=ROUND({Convert.ToDouble(item["FacilityArea"].ToString())};5)";//取到小數點第五位，因為可能用Excel或Calc開啟，所以用這種計算方式較準確
                    }

                    for (int i = 0; i < tb.ColumnCount; i++)
                        tb.Columns[i].AutoWidth = true;
                    tb.Columns[0].HorizonAlign = Alignment.HorizonAlignment.CENTER;
                    calc.Tables.Remove(calc.Tables[0]);
                    calc.SaveAs(filePath);
                    calc.Close();
                }
            }
            catch (Exception ex)
            {
                return HttpHelper.FailResult(ex.ToString());
            }

            return HttpHelper.FileResult(filePath, "application/vnd.oasis.opendocument.spreadsheet");
        }
        #endregion

        #region 取得出流案件明細總表
        /// <summary>
        /// 取得出流案件明細
        /// </summary>
        /// <param name="filterParams"></param>
        /// <returns></returns>
        [HttpGet, ActionName(nameof(gov_detail))]
        public HttpResponseMessage gov_detail([FromUri] FilterParams filterParams)
        {
            string FileName = "出流管制案件詳細資訊表";
            var filePath = Path.Combine(_filePath, "ReportTempFile", $"{FileName}.ods");

            #region 寫入ods資料表
            try
            {
                var responseString = HttpHelper.GetResponseWithToken("gov_detail.ods", new { filter = filterParams });
                JArray CaseDetailObj = JsonConvert.DeserializeObject<JArray>(responseString);
                using (var calc = new Calc())
                {
                    var tb = calc.Tables.AddNew(FileName);

                    // 樣式
                    Font headerFont = new Font("微軟正黑體", 28, FontStyle.Bold),
                        colFont = new Font("微軟正黑體", 14, FontStyle.Bold),
                        rowFont = new Font("微軟正黑體", 12);
                    var line = new Line() { Color = Color.Black, OuterWidth = 20 };

                    //表格標題
                    tb[0, 0].Formula = FileName;
                    tb["A1:G1"].Merge = true;
                    tb[0, 0].SetFont(headerFont);

                    int startRow = 1; // 要調整明細位置只要改這個數值就好了

                    tb[startRow, 0].Formula = "案件名稱";
                    tb[startRow, 1].Formula = "義務人資訊";
                    tb[startRow, 2].Formula = "承覽單位/審查單位資訊";
                    tb[startRow, 3].Formula = "開發基地面積";
                    tb[startRow, 4].Formula = "開發基地排水資訊";
                    tb[startRow, 5].Formula = "滯洪池資訊";
                    tb[0, startRow, 5, startRow].SetFont(colFont);

                    foreach (var item in CaseDetailObj)
                    {
                        startRow++;
                        tb[0, startRow, 5, startRow].SetFont(rowFont);
                        string OFZ_Name = item["OFZ_Name"].ToString().Length > 40 ? $"{item["OFZ_Name"].ToString().Substring(0, 40)}\r\n{item["OFZ_Name"].ToString().Substring(40)}" : item["OFZ_Name"].ToString();

                        string caseInfo = $"案件名稱：{OFZ_Name}\r\n年度：{item["OFP_Year"]}\r\n管理系統案號：{item["OFP_NO"]}\r\n計畫類別：{item["OFP_Type"]}";

                        tb[startRow, 0].Formula = caseInfo;
                        // 義務人資訊
                        tb[startRow, 1].Formula = $"義務人：{item["Payer"]}\r\n代表人：{item["Representative"]}";
                        //承辦資訊
                        string EngInfo = $"承辦單位：{item["PracticeUnits"]}\r\n簽證技師：{item["EngineerName"]}\r\n簽證技師(科別)：{item["EngineerType"]}\r\n審查機關/單位：{item["OFP_Gov"]}";
                        tb[startRow, 2].Formula = EngInfo;
                        string Area = string.IsNullOrEmpty(item["OFP_Area"].ToString()) ? "0" : item["OFP_Area"].ToString();
                        tb[startRow, 3].Formula = $"{Area} 公頃";
                        //開發基地排水資訊
                        string RiversInfo = $"排入河川或排水名稱：{item["OFP_DrainName"].ToString()}\r\n所屬水系：{item["OFP_RiverName"].ToString()}";
                        tb[startRow, 4].Formula = RiversInfo;
                        //滯洪池資訊
                        string FacilityInfo = $"數量(座)：{item["OFP_FacilityNum"].ToString()}\r\n總體積(萬m3)：{item["OFP_FacilityArea"].ToString()}";
                        tb[startRow, 5].Formula = FacilityInfo;
                    }

                    for (int i = 0; i < tb.ColumnCount; i++)
                    {
                        tb.Columns[i].AutoWidth = true;
                        tb.Columns[i].VerticalAlign = Alignment.VerticalAlignment.CENTER;
                    }

                    calc.Tables.Remove(calc.Tables[0]);
                    calc.SaveAs(filePath);
                    calc.Close();
                }
            }
            catch (Exception ex)
            {
                return HttpHelper.FailResult(ex.ToString());
            }


            return HttpHelper.FileResult(filePath, "application/vnd.oasis.opendocument.spreadsheet");
            #endregion
        }
        #endregion

        #region 取得出流案件審查總表
        /// <summary>
        /// 取得出流案件審查資訊
        /// </summary>
        /// <param name="MeetingInfoObj"></param>
        /// <returns></returns>
        [HttpGet, ActionName(nameof(meeting))]
        public HttpResponseMessage meeting([FromUri] FilterParams filterParams)
        {
            string FileName = "出流管制案件審查資訊表";
            var filePath = Path.Combine(_filePath, "ReportTempFile", $"{FileName}.ods");
            #region 寫入ods資料表
            try
            {
                var responseString = HttpHelper.GetResponseWithToken("meeting.ods", new { filter = filterParams });
                JArray MeetingInfoObj = JsonConvert.DeserializeObject<JArray>(responseString);
                using (var calc = new Calc())
                {
                    var tb = calc.Tables.AddNew(FileName);

                    // 樣式
                    Font headerFont = new Font("微軟正黑體", 28, FontStyle.Bold),
                        colFont = new Font("微軟正黑體", 14, FontStyle.Bold),
                        rowFont = new Font("微軟正黑體", 12);
                    var line = new Line() { Color = Color.Black, OuterWidth = 20 };

                    //表格標題
                    tb[0, 0].Formula = FileName;
                    tb["A1:H1"].Merge = true;
                    tb[0, 0].SetFont(headerFont);

                    int startRow = 1; // 要調整明細位置只要改這個數值就好了

                    tb[startRow, 0].Formula = "案件名稱";
                    tb[startRow, 1].Formula = "義務人資訊";
                    tb[startRow, 2].Formula = "承覽單位/審查單位資訊";
                    tb[startRow, 3].Formula = "審查機關確認繳費日期";
                    tb[startRow, 4].Formula = "第1次審查開會日期";
                    tb[startRow, 5].Formula = "第2次審查開會日期";
                    tb[startRow, 6].Formula = "第3次審查開會日期";
                    tb[startRow, 7].Formula = "合計審查日數";
                    tb[0, startRow, 7, startRow].SetFont(colFont);
                    foreach (var item in MeetingInfoObj as JArray)
                    {
                        startRow++;
                        tb[0, startRow, 7, startRow].SetFont(rowFont);
                        string OFP_Name = item["OFP_Name"].ToString().Length > 40 ? $"{item["OFP_Name"].ToString().Substring(0, 40)}\r\n{item["OFP_Name"].ToString().Substring(40)}" : item["OFP_Name"].ToString();
                        string caseInfo = $"案件名稱：{OFP_Name}\r\n年度：{item["OFP_Year"]}\r\n管理系統案號：{item["OFP_NO"]}\r\n計畫類別：{item["OFP_Type"]}";
                        tb[startRow, 0].Formula = caseInfo;

                        // 義務人資訊
                        tb[startRow, 1].Formula = $"義務人：{item["Payer"]}\r\n代表人：{item["Representative"]}";
                        //承辦資訊
                        string EngInfo = $"承辦單位：{item["PracticeUnits"]}\r\n簽證技師：{item["EngineerName"]}\r\n簽證技師(科別)：{item["EngineerType"]}\r\n審查機關/單位：{item["OFP_Gov"]}";
                        tb[startRow, 2].Formula = EngInfo;
                        tb[startRow, 3].Formula = string.IsNullOrEmpty(item["RVGov_AccPayDate"].ToString()) ? "" : Convert.ToDateTime(item["RVGov_AccPayDate"].ToString()).ToString("yyyy-MM-dd");

                        JObject meetingInfo = JObject.Parse(item["Meeting"].ToString());
                        JArray meetings = JArray.Parse(meetingInfo["Lists"].ToString());
                        // 目前只顯示三次會議
                        int i = 4;
                        foreach (var meet in meetings)
                        {
                            if (i > 6) break;// 大於第三次會議就離開
                            // 審查開會日期
                            var reviewData = string.IsNullOrEmpty(meet["Review_Date"].ToString()) ? "" : Convert.ToDateTime(meet["Review_Date"].ToString()).ToString("yyyy-MM-dd");
                            // 機關審查開會通知文號
                            var ReviewNotify_No = meet["ReviewNotify_No"] ?? "";
                            // 修正本收件日期
                            var Fix_Receive_Date = meet["Fix_Receive_Date"] == null || string.IsNullOrEmpty(meet["Fix_Receive_Date"].ToString()) ? "" : Convert.ToDateTime(meet["Fix_Receive_Date"].ToString()).ToString("yyyy-MM-dd");

                            tb[startRow, i].Formula = $"審查開會日期：{reviewData}\r\n機關審查開會通知文號：{ReviewNotify_No}\r\n修正本收件日期：{Fix_Receive_Date}\r\n審查天數：{item["ReviewDays"]}";
                            i++;
                        }

                        tb[startRow, 7].Formula = meetingInfo["TotalDays"].ToString();
                    }
                    tb.Columns[3].NumberFormat = "YYYY-MM-DD";
                    tb.Columns[4].NumberFormat = "YYYY-MM-DD";
                    tb.Columns[5].NumberFormat = "YYYY-MM-DD";
                    tb.Columns[6].NumberFormat = "YYYY-MM-DD";
                    for (int i = 0; i < tb.ColumnCount; i++)
                    {
                        tb.Columns[i].AutoWidth = true;
                        tb.Columns[i].VerticalAlign = Alignment.VerticalAlignment.CENTER;
                    }
                    calc.Tables.Remove(calc.Tables[0]);
                    calc.SaveAs(filePath);
                    calc.Close();
                }
            }
            catch (Exception ex)
            {
                return HttpHelper.FailResult(ex.ToString());
            }

            return HttpHelper.FileResult(filePath, "application/vnd.oasis.opendocument.spreadsheet");
            #endregion
        }
        #endregion

        #region 取得檔案上傳總表
        /// <summary>
        /// 取得檔案上傳總表
        /// </summary>
        /// <param name="FileDataObj"></param>
        /// <returns></returns>
        [HttpGet, ActionName(nameof(case_files))]
        public HttpResponseMessage case_files([FromUri] FilterParams filterParams)
        {
            string FileName = "出流管制案件檔案列表";
            var filePath = Path.Combine(_filePath, "ReportTempFile", $"{FileName}.ods");

            #region 寫入ods
            try
            {
                var responseString = HttpHelper.GetResponseWithToken("case_files.ods", new { filter = filterParams });
                JObject FileDataObj = JsonConvert.DeserializeObject<JObject>(responseString);
                using (var calc = new Calc())
                {
                    #region 表頭設定
                    var tb = calc.Tables.AddNew(FileName);

                    // 樣式
                    Font headerFont = new Font("微軟正黑體", 28, FontStyle.Bold),
                        colFont = new Font("微軟正黑體", 14, FontStyle.Bold),
                        rowFont = new Font("微軟正黑體", 12);
                    var line = new Line() { Color = Color.Black, OuterWidth = 20 };

                    //表格標題
                    tb[0, 0].Formula = FileName;

                    //案件名稱
                    tb[1, 0].Formula = "案件名稱";
                    tb[0, 1, 0, 2].Merge = true;
                    tb[0, 1, 0, 2].HorizonAlign = Alignment.HorizonAlignment.CENTER;
                    tb[0, 1, 0, 2].VerticalAlign = Alignment.VerticalAlignment.CENTER;
                    tb[1, 0].SetFont(colFont);
                    #endregion

                    int startRow = 1; // 要調整明細位置只要改這個數值就好了

                    #region 透過分類的檔案創建欄位
                    // 取得檔案的分類
                    var fileCategoryString = HttpHelper.GetResponse("FilesCategory.json", new { filter = new JObject() });
                    JArray fileCategories = JsonConvert.DeserializeObject<JArray>(fileCategoryString);

                    #region 審查類別直接新增在程式中
                    int firstReviewsCount = fileCategories.Children<JObject>().Min(x => Convert.ToInt32(x["Sort"].ToString()));
                    int lastReviewsCount = fileCategories.Children<JObject>().Max(x => Convert.ToInt32(x["Sort"].ToString()));
                    fileCategories.Add(JObject.FromObject(new
                    {
                        FileDescription = "第1次行政審查補正報告",
                        FileCategory = "reviewsFile",
                        SubCategory = "TLCorrection",
                        SubCategoryName = "行政審查補正報告",
                        FileType = "1-C1",
                        Sort = firstReviewsCount - 3,
                        CanUploadExt = ".pdf"
                    }));

                    fileCategories.Add(JObject.FromObject(new
                    {
                        FileDescription = "第2次行政審查補正報告",
                        FileCategory = "reviewsFile",
                        SubCategory = "TLCorrection",
                        SubCategoryName = "行政審查補正報告",
                        FileType = "2-C1",
                        Sort = firstReviewsCount - 2,
                        CanUploadExt = ".pdf"
                    }));

                    fileCategories.Add(JObject.FromObject(new
                    {
                        FileDescription = "第3次行政審查補正報告",
                        FileCategory = "reviewsFile",
                        SubCategory = "TLCorrection",
                        SubCategoryName = "行政審查補正報告",
                        FileType = "3-C1",
                        Sort = firstReviewsCount - 1,
                        CanUploadExt = ".pdf"
                    }));

                    fileCategories.Add(JObject.FromObject(new
                    {
                        FileDescription = "第1次審查開會通知函",
                        FileCategory = "reviewsFile",
                        SubCategory = "Reviews",
                        SubCategoryName = "審查資料",
                        FileType = "1-D11,1-D21,1-D31",
                        Sort = lastReviewsCount + 1,
                        CanUploadExt = ".pdf"
                    }));

                    fileCategories.Add(JObject.FromObject(new
                    {
                        FileDescription = "第1次審查意見通知函",
                        FileCategory = "reviewsFile",
                        SubCategory = "Reviews",
                        SubCategoryName = "審查資料",
                        FileType = "1-D12,1-D22,1-D32",
                        Sort = lastReviewsCount + 2,
                        CanUploadExt = ".pdf"
                    }));

                    fileCategories.Add(JObject.FromObject(new
                    {
                        FileDescription = "第1次修正報告書",
                        FileCategory = "reviewsFile",
                        SubCategory = "Reviews",
                        SubCategoryName = "審查資料",
                        FileType = "1-E1",
                        Sort = lastReviewsCount + 3,
                        CanUploadExt = ".pdf"
                    }));

                    fileCategories.Add(JObject.FromObject(new
                    {
                        FileDescription = "第2次審查開會通知函",
                        FileCategory = "reviewsFile",
                        SubCategory = "Reviews",
                        SubCategoryName = "審查資料",
                        FileType = "2-D11,2-D21,2-D31",
                        Sort = lastReviewsCount + 4,
                        CanUploadExt = ".pdf"
                    }));

                    fileCategories.Add(JObject.FromObject(new
                    {
                        FileDescription = "第2次審查意見通知函",
                        FileCategory = "reviewsFile",
                        SubCategory = "Reviews",
                        SubCategoryName = "審查資料",
                        FileType = "2-D12,2-D22,2-D32",
                        Sort = lastReviewsCount + 5,
                        CanUploadExt = ".pdf"
                    }));

                    fileCategories.Add(JObject.FromObject(new
                    {
                        FileDescription = "第2次修正報告書",
                        FileCategory = "reviewsFile",
                        SubCategory = "Reviews",
                        SubCategoryName = "審查資料",
                        FileType = "2-E1",
                        Sort = lastReviewsCount + 6,
                        CanUploadExt = ".pdf"
                    }));

                    fileCategories.Add(JObject.FromObject(new
                    {
                        FileDescription = "第3次審查開會通知函",
                        FileCategory = "reviewsFile",
                        SubCategory = "Reviews",
                        SubCategoryName = "審查資料",
                        FileType = "3-D11,3-D21,3-D31",
                        Sort = lastReviewsCount + 7,
                        CanUploadExt = ".pdf"
                    }));

                    fileCategories.Add(JObject.FromObject(new
                    {
                        FileDescription = "第3次審查意見通知函",
                        FileCategory = "reviewsFile",
                        SubCategory = "Reviews",
                        SubCategoryName = "審查資料",
                        FileType = "3-D12,3-D22,3-D32",
                        Sort = lastReviewsCount + 8,
                        CanUploadExt = ".pdf"
                    }));

                    fileCategories.Add(JObject.FromObject(new
                    {
                        FileDescription = "第3次修正報告書",
                        FileCategory = "reviewsFile",
                        SubCategory = "Reviews",
                        SubCategoryName = "審查資料",
                        FileType = "3-E1",
                        Sort = lastReviewsCount + 9,
                        CanUploadExt = ".pdf"
                    }));

                    fileCategories.Add(JObject.FromObject(new
                    {
                        FileDescription = "審查核定檔案",
                        FileCategory = "reviewsFile",
                        SubCategory = "Reviews",
                        SubCategoryName = "審查資料",
                        FileType = "F1,F2,F3",
                        Sort = lastReviewsCount + 10,
                        CanUploadExt = ".pdf"
                    }));
                    #endregion

                    // 欄位位置
                    int col = 1;
                    // 將檔案類別關聯欄位，放入字典，讓後續資料寫入能找到對應的位置
                    Dictionary<string, int> RelatedTypeCol = new Dictionary<string, int>();
                    #region 寫入標題
                    // 群組相同的分類
                    var categories = fileCategories.OrderBy(x => new List<string>() { "fillFile", "reviewsFile", "approvedFile" }.IndexOf(x["FileCategory"].ToString())).ThenBy(x => x["Sort"])
                    .GroupBy(x => x["SubCategoryName"]).Select(group => new
                    {
                        typeName = group.Key,
                        total = group.Count(),
                        fileList = group.ToList()
                    }).ToList();

                    // 將每一個欄位放入
                    foreach (var item in categories)
                    {
                        // 分類欄位
                        tb[startRow, col].Formula = item.typeName.ToString();
                        tb[col, startRow, col + item.total - 1, startRow].Merge = true;
                        tb[col, startRow, col + item.total - 1, startRow].HorizonAlign = Alignment.HorizonAlignment.CENTER;

                        // 子項目欄位
                        foreach (var fileInfo in item.fileList)
                        {
                            tb[startRow + 1, col].Formula = fileInfo["FileDescription"].ToString();
                            // 如果有設定逗號，就將每個資料分開寫入字典，會用逗號分隔的是各種不同審查，不同審查只會放在一個位置
                            if (fileInfo["FileType"].ToString().Contains(","))
                            {
                                var fileTypes = fileInfo["FileType"].ToString().Split(',');
                                foreach (var fileType in fileTypes)
                                {
                                    RelatedTypeCol.Add(fileType, col);
                                }
                            }
                            else
                            {
                                RelatedTypeCol.Add(fileInfo["FileType"].ToString(), col);
                            }

                            col++;
                        }
                    }
                    #endregion
                    #endregion

                    #region 設定標題欄位格式，根據欄位長度設定
                    // 設定粗體
                    tb[0, startRow, col - 1, startRow].SetFont(colFont);

                    // 設定原本標頭的跨欄置中
                    tb[0, 0, col - 1, 0].Merge = true;
                    tb[0, 0].SetFont(headerFont);

                    startRow++;
                    // 設定粗體
                    tb[0, startRow, col - 1, startRow].SetFont(colFont);
                    #endregion
                    #region  寫入檔案
                    var mainObjs = FileDataObj["mainObj"] as JArray;
                    foreach (var ofp in mainObjs)
                    {
                        //換行
                        startRow++;

                        // 基本資訊
                        string OFP_Name = ofp["OFP_Name"].ToString().Length > 40 ? $"{ofp["OFP_Name"].ToString().Substring(0, 40)}\r\n{ofp["OFP_Name"].ToString().Substring(40)}" : ofp["OFP_Name"].ToString();
                        tb[startRow, 0].Formula = OFP_Name;

                        #region 判斷是否有上傳處理
                        var filesCategoryObj = FileDataObj["filesCategoryObj"] as JArray;
                        var filesCategory = filesCategoryObj.Where(x => x["OFP_ID"].ToString() == ofp["OFP_ID"].ToString());
                        foreach (var item in filesCategory)
                        {
                            if (item["FileType"].ToString().Length == 5 && Convert.ToInt32(item["FileType"].ToString().Substring(0, 1)) >= 4) { }
                            else { tb[startRow, Convert.ToInt32(RelatedTypeCol[item["FileType"].ToString()])].Formula = "V"; }
                        }
                        #endregion
                    }
                    #endregion

                    //調整寬的大小
                    for (int i = 0; i < tb.ColumnCount; i++)
                        tb.Columns[i].AutoWidth = true;
                    //移除預設頁
                    calc.Tables.Remove(calc.Tables[0]);
                    //儲存並關閉
                    calc.SaveAs(filePath);
                    calc.Close();
                }
            }
            catch (Exception ex)
            {
                return HttpHelper.FailResult(ex.ToString());
            }
            #endregion

            return HttpHelper.FileResult(filePath, "application/vnd.oasis.opendocument.spreadsheet");
        }
        #endregion

        #region 取得審查資訊明細
        /// <summary>
        /// 取得審查資訊明細
        /// </summary>
        /// <param name="CaseObj"></param>
        /// <returns></returns>
        [HttpGet, ActionName(nameof(case_detail))]
        [Route("api/Report/case_detail/{OFPID}")]
        public HttpResponseMessage case_detail(int OFPID)
        {
            string FileName = "出流管制審查案件明細";
            var filePath = string.Empty;


            #region 寫入ods
            try
            {
                var responseString = HttpHelper.GetResponseWithRouteDataAndToken("case_detail.ods", new string[] { OFPID.ToString() });
                JObject CaseObj = JsonConvert.DeserializeObject<JObject>(responseString);
                // 設定匯出的檔案名稱
                filePath = Path.Combine(_filePath, "ReportTempFile", $"{CaseObj["OFP_NO"].ToString()}-{FileName}.ods");
                using (Calc calc = new Calc(Path.Combine(_templateFilePath, $"{FileName}.ods")))
                {
                    Table tb = calc.Tables[0];

                    #region 案件編號到承辦人Email的區域
                    int nowRow = 0;
                    bool IsPlan = false;
                    //判斷是否為規劃書
                    if (CaseObj["OFP_Type"].ToString().Contains("規劃書"))
                        IsPlan = true;

                    tb[nowRow++, 0].Formula = string.Format("出流管制審查管控表\r\n{0}出流管制規劃書 {1}出流管制計畫書", IsPlan ? "■" : "□", !IsPlan ? "■" : "□");//標頭
                    tb[nowRow++, 1].Formula = CaseObj["OFP_NO"].ToString();//管理系統案號
                    tb[nowRow++, 1].Formula = CaseObj["OFP_Name"].ToString();//案件名稱

                    //審查單位區塊
                    var govUser = CaseObj["govUser"];
                    tb[nowRow++, 1].Formula = govUser["Gov_Name"].ToString();
                    tb[nowRow++, 1].Formula = govUser["Gov_address"].ToString();
                    tb[nowRow++, 1].Formula = govUser["Gov_Tel"].ToString();
                    tb[nowRow++, 1].Formula = govUser["Gov_FAX"].ToString();
                    tb[nowRow, 1].Formula = govUser["Gov_UserName"].ToString();
                    tb[nowRow, 3].Formula = govUser["Gov_Mail"].ToString();
                    tb[nowRow++, 4].Formula = govUser["Gov_cellnumber"].ToString();

                    nowRow++;

                    //義務人區塊
                    tb[nowRow++, 1].Formula = CaseObj["Payer"].ToString(); //義務人
                    tb[nowRow++, 1].Formula = CaseObj["Representative"].ToString(); //代表人
                    tb[nowRow++, 1].Formula = CaseObj["PA_address"].ToString(); //義務人地址
                    tb[nowRow++, 1].Formula = CaseObj["PA_Tel"].ToString(); //義務人電話
                    tb[nowRow++, 1].Formula = CaseObj["PA_FAX"].ToString(); //義務人傳真

                    //承辦單位區塊
                    nowRow++;
                    tb[nowRow++, 1].Formula = CaseObj["PracticeUnits"].ToString(); //承辦單位
                    tb[nowRow++, 1].Formula = CaseObj["E_Address"].ToString(); //承辦單位地址
                    tb[nowRow++, 1].Formula = CaseObj["EngineerName"].ToString(); //承辦技師
                                                                                  // 這三個同列
                    tb[nowRow, 1].Formula = CaseObj["EngineerName"].ToString(); //聯絡人(目前都是承辦技師)
                    tb[nowRow, 3].Formula = CaseObj["E_Email"].ToString(); //信箱
                    tb[nowRow++, 4].Formula = CaseObj["E_Cellnumber"].ToString(); //承辦技師手機
                    tb[nowRow++, 1].Formula = CaseObj["E_Tel"].ToString(); //承辦單位電話

                    //承辦單位傳真、Eamil沒有

                    #endregion

                    #region 委員區域
                    nowRow = 24;
                    var manList = CaseObj["manList"].ToList();

                    for (int row = 0; row < 8; row++)// 最多8列
                    {
                        int manCount = manList.Count;
                        if (row + 1 > manCount)
                        {
                            break;
                        }
                        tb[nowRow, 0].Formula = manList[row]["PeopleType"].ToString() == "Man" ? "委員" : "技師"; //目前固定委員
                        tb[nowRow, 1].Formula = manList[row]["PeopleName"].ToString(); //委員/技師姓名
                        tb[nowRow, 2].Formula = manList[row]["Tel"].ToString(); //電話
                        tb[nowRow, 3].Formula = manList[row]["Address"].ToString(); //地址
                        tb[nowRow, 5].Formula = manList[row]["Email"].ToString(); //Email
                        nowRow++;
                    }
                    #endregion

                    #region 開會日期區域
                    nowRow = 34;
                    tb[nowRow, 1].Formula = DateHelper.ToShortDate(CaseObj["Payer_Send_Date"].ToString());//義務人送件日期(對應審查發文日期 的 發文日期)
                    tb[nowRow, 2].Formula = DateHelper.ToShortDate(CaseObj["RVGov_AccPayDate"].ToString());//審查機關確認繳費日期
                    tb[nowRow++, 3].Formula = CaseObj["Payer_Send_No"].ToString(); //義務人送件文號
                    tb[nowRow, 1].Formula = DateHelper.ToShortDate(CaseObj["Notice_Payers_PayDate"].ToString());//通知義務人繳費日期(審查機關收文日期)
                    tb[nowRow, 2].Formula = DateHelper.ToShortDate(CaseObj["RVGov_AccPayDate"].ToString());//審查機關確認繳費日期
                    tb[nowRow++, 3].Formula = CaseObj["Notice_Payers_PayNo"].ToString();//通知義務人繳費文號
                    var meetinglist = CaseObj["Meeting"]["Lists"].ToList();
                    int meetCount = 1;
                    foreach (var item in meetinglist)
                    {
                        if (meetCount > 3)
                            break;

                        //審查核定
                        if (Convert.ToInt32(item["Correction_Status"]) == 1)
                        {
                            nowRow = 49;
                            tb[nowRow, 1].Formula = DateHelper.ToShortDate(item["ReviewNotify_Date"].ToString());
                            tb[nowRow, 3].Formula = item["ReviewNotify_No"].ToString();
                            tb[nowRow, 4].Formula = DateHelper.ToShortDate(item["Review_Date"].ToString());

                            break;
                        }
                        else
                        {
                            tb[nowRow, 1].Formula = DateHelper.ToShortDate(item["ReviewNotify_Date"].ToString());
                            tb[nowRow, 3].Formula = item["ReviewNotify_No"].ToString();
                            tb[nowRow, 4].Formula = DateHelper.ToShortDate(item["Review_Date"].ToString());
                            tb[nowRow, 5].Formula = item["ReviewDays"].ToString();
                            nowRow++;
                            tb[nowRow, 1].Formula = DateHelper.ToShortDate(item["ReviewReply_Date"].ToString());
                            tb[nowRow, 3].Formula = item["ReviewReply_No"].ToString();
                            nowRow++;
                            tb[nowRow, 1].Formula = item["Correctionfix_Date"] == null ? "" : DateHelper.ToShortDate(item["Correctionfix_Date"].ToString());
                            tb[nowRow, 2].Formula = item["Fix_Receive_Date"] == null ? "" : DateHelper.ToShortDate(item["Fix_Receive_Date"].ToString());
                            tb[nowRow, 3].Formula = item["Correctionfix_No"] == null ? "" : item["Correctionfix_No"].ToString();
                            nowRow++;
                            meetCount++;
                        }
                    }

                    tb[50, 5].Formula = CaseObj["Meeting"]["TotalDays"].ToString();
                    #endregion

                    calc.SaveAs(filePath);
                    calc.Close();

                }
            }
            catch (Exception ex)
            {
                return HttpHelper.FailResult(ex.ToString());
            }

            #endregion

            return HttpHelper.FileResult(filePath, "application/vnd.oasis.opendocument.spreadsheet");
        }
        #endregion

        #region 取得檔案繳交明細
        /// <summary>
        /// 取得檔案繳交明細
        /// </summary>
        /// <param name="CaseObj"></param>
        /// <returns></returns>
        [HttpGet, ActionName(nameof(file_detail))]
        [Route("api/Report/file_detail/{OFPID}")]
        public HttpResponseMessage file_detail(int OFPID)
        {
            string FileName = "出流管制審查案件成果繳交盤點單";
            var filePath = string.Empty;

            #region 寫入ods
            try
            {
                // 取得已上傳檔案的資料
                var responseString = HttpHelper.GetResponseWithRouteDataAndToken("file_detail.ods", new string[] { OFPID.ToString() });
                JObject CaseObj = JsonConvert.DeserializeObject<JObject>(responseString);
                var caseinfo = CaseObj["CaseInfo"];
                // 設定匯出的檔案名稱
                filePath = Path.Combine(_filePath, "ReportTempFile", $"{caseinfo["OFP_NO"].ToString()}-{FileName}.ods");
                using (Calc calc = new Calc(Path.Combine(_templateFilePath, $"{FileName}.ods")))
                {
                    Table tb = calc.Tables[0];
                    #region 基本資訊區域
                    // 左邊
                    tb[1, 1, 4, 1].Wrap = true;
                    tb[1, 1].Formula = caseinfo["OFP_Name"].ToString();
                    tb[2, 1].Formula = caseinfo["OFP_Gov"].ToString();
                    tb[3, 1].Formula = caseinfo["Payer"].ToString();
                    tb[4, 1].Formula = caseinfo["PracticeUnits"].ToString();
                    // 右邊            
                    tb[1, 6].Formula = caseinfo["OFP_NO"].ToString();
                    // 受理日期
                    tb[2, 6].Formula = string.IsNullOrEmpty(caseinfo["RVGov_AccPayDate"].ToString()) ? "" : Convert.ToDateTime(caseinfo["RVGov_AccPayDate"]).ToString("yyyy/MM/dd");
                    tb[3, 6].Formula = caseinfo["Representative"].ToString();
                    tb[4, 6].Formula = caseinfo["EngineerName"].ToString();
                    #endregion

                    #region 檔案資訊區域
                    #region 欄位資料製作
                    // 代碼對應的欄位名稱
                    Dictionary<string, string> typeName = new Dictionary<string, string>();
                    typeName.Add("fillFile", "申請階段");
                    typeName.Add("reviewsFile", "審查階段");
                    typeName.Add("approvedFile", "核定階段");
                    // 查詢分類的資訊
                    var fileCategoryString = HttpHelper.GetResponse("FilesCategory.json", new { filter = new JObject() });
                    JArray fileCategories = JsonConvert.DeserializeObject<JArray>(fileCategoryString);
                    // 新增審查的欄位
                    fileCategories.Add(JObject.FromObject(new { FileDescription = "行政審查補正報告", FileCategory = "reviewsFile", SubCategory = "TLCorrection", SubCategoryName = "行政審查補正報告", FileType = "C1", CanUploadExt = ".pdf、.zip" }));
                    fileCategories.Add(JObject.FromObject(new { FileDescription = "審查與修正檔案", FileCategory = "reviewsFile", SubCategory = "Reviews", SubCategoryName = "審查資料", FileType = "D,E,F", CanUploadExt = ".pdf、.zip" }));

                    // 群組相同的分類
                    var categories = fileCategories.GroupBy(x => x["FileCategory"]).Select(group => new
                    {
                        typeName = group.Key,
                        total = group.Count(),
                        fileList = group.GroupBy(y => y["SubCategory"]).Select(subgroup => new
                        {
                            subTypeName = subgroup.Key.ToString(),
                            subList = subgroup.ToList()
                        }).ToList()
                    }).ToList();
                    #endregion

                    JArray FilesObj = CaseObj["FilesObj"] as JArray;

                    // 起始位置
                    int startRow = 7;
                    // 根據群組的資訊去做繪製報表
                    foreach (var category in categories)
                    {
                        #region 繪製欄位
                        // 階段
                        tb[startRow, 0].Formula = typeName[category.typeName.ToString()];
                        // 合併首欄
                        tb[0, startRow, 0, startRow + category.total - 1].Merge = true;
                        #endregion

                        foreach (var subcategory in category.fileList)
                        {
                            // 這邊針對每個子類別去做查詢
                            var subFiles = FilesObj.Children<JObject>().Where(x => x["typeName"].ToString() == subcategory.subTypeName).FirstOrDefault();

                            bool isFirst = true;
                            foreach (var item in subcategory.subList)
                            {
                                // 第一次要繪製類別資訊
                                if (isFirst)
                                {
                                    tb[startRow, 1].Formula = item["SubCategoryName"].ToString();
                                    tb[1, startRow, 1, startRow + subcategory.subList.Count - 1].Merge = true;
                                }
                                // 繪製應包含項目
                                tb[startRow, 2].Formula = item["FileDescription"].ToString();
                                // 繪製提交資料格式
                                tb[startRow, 3].Formula = item["CanUploadExt"].ToString();

                                #region 取出要繪製的資訊
                                if (subFiles != null)
                                {
                                    #region 透過檔案型別取得上傳的檔案
                                    // 取得該分類檔案的清單
                                    JArray subFilesDetail = subFiles["List"] as JArray;
                                    var datas = new List<JObject>();
                                    // 如果是行政審查補正報告只需要判斷是否包含字串就好，因為檔案類別命名為1-C1、2-C1....
                                    if (item["SubCategory"].ToString() == "TLCorrection")
                                    {
                                        datas = subFilesDetail.Children<JObject>().Where(x => x["File_Type"].ToString().Contains(item["FileType"].ToString())).ToList();
                                    }
                                    // 如果是審查資料要判斷是否為D、E、F等類別的關鍵字，因為不同審查的類別代碼不同
                                    else if (item["SubCategory"].ToString() == "Reviews")
                                    {
                                        var fileTypeKeyWords = item["FileType"].ToString().Split(',');
                                        foreach (var key in fileTypeKeyWords)
                                        {
                                            datas.AddRange(subFilesDetail.Children<JObject>().Where(x => x["File_Type"].ToString().Contains(key)).ToList());
                                        }
                                    }
                                    // 其餘的查出與類別相同的資料就好
                                    else
                                    {
                                        datas = subFilesDetail.Children<JObject>().Where(x => x["File_Type"].ToString() == item["FileType"].ToString()).ToList();
                                    }
                                    #endregion

                                    #region 繪製取出的資訊
                                    // 有多筆檔案會放在同一個欄位
                                    StringBuilder sbFiles = new StringBuilder();
                                    foreach (var data in datas)
                                    {
                                        if (!string.IsNullOrEmpty(data["info"].ToString()))
                                            sbFiles.Append($"{data["info"]["File_name"].ToString()}\r\n");
                                    }
                                    // 繪製檔名
                                    tb[startRow, 4].Formula = sbFiles.ToString();
                                    // 該上傳檔案數量
                                    tb[startRow, 5].Value = datas.Count;
                                    // 已上傳數量
                                    tb[startRow, 6].Value = datas.Where(x => !string.IsNullOrEmpty(x["info"].ToString())).Count();
                                    // 未上傳數量
                                    tb[startRow, 7].Value = datas.Where(x => string.IsNullOrEmpty(x["info"].ToString())).Count();
                                    #endregion
                                }
                                else
                                {
                                    // 該上傳檔案數量
                                    tb[startRow, 5].Value = 0;
                                    // 已上傳數量
                                    tb[startRow, 6].Value = 0;
                                    // 未上傳數量
                                    tb[startRow, 7].Value = 0;
                                }
                                #endregion
                                startRow++;
                            }
                        }
                    }


                    #endregion

                    //調整寬的大小
                    for (int i = 0; i < 5; i++)
                        tb.Columns[i].AutoWidth = true;
                    calc.SaveAs(filePath);
                    calc.Close();

                }
            }
            catch (Exception ex)
            {
                return HttpHelper.FailResult(ex.ToString());
            }

            #endregion

            return HttpHelper.FileResult(filePath, "application/vnd.oasis.opendocument.spreadsheet");
        }
        #endregion

        #region 取得odt檔案
        [HttpGet, ActionName(nameof(outflowctrl))]
        [Route("api/Report/outflowctrl/{OFPNO}/{fileName}")]
        public HttpResponseMessage outflowctrl(string OFPNO, string fileName)
        {
            var filePath = string.Empty;

            fileName = $"{fileName}.odt";// odt為附檔名

            try
            {
                var responseString = HttpHelper.GetResponseWithToken(Path.Combine(_reportBaseUrl, "outflowctrl_api", OFPNO, fileName));
                JObject FileDataInfo = JsonConvert.DeserializeObject<JObject>(responseString);

                filePath = DocumentHelper.GetRptStream(fileName, FileDataInfo, $"{OFPNO}.odt");
            }
            catch (Exception ex)
            {
                return HttpHelper.FailResult(ex.ToString());
            }

            return HttpHelper.FileResult(filePath, "application/vnd.oasis.opendocument.spreadsheet");
        }
        #endregion

        #region 取得階段性報表
        [HttpGet, ActionName(nameof(summary_status))]
        public HttpResponseMessage summary_status()
        {
            string FileName = "出流管制案件檔案階段列表";
            var filePath = Path.Combine(_filePath, "ReportTempFile", $"{FileName}.ods");

            try
            {
                var responseString = HttpHelper.GetResponseWithToken(Path.Combine(_outFlowBaseUrl, "summary_status.ods"));
                JArray summaryObj = JArray.Parse(responseString);

                Dictionary<string, string> types = new Dictionary<string, string>();
                types.Add("中央列管出流管制計畫書審查調查表", "中央審查");
                types.Add("地方列管出流管制計畫書審查調查表", "地方審查");

                #region 寫入ods
                using (var calc = new Calc())
                {

                    foreach (var key in types.Keys)
                    {
                        int startRow = 0;
                        #region 表頭設定
                        var tb = calc.Tables.AddNew(key);

                        // 樣式
                        Font headerFont = new Font("微軟正黑體", 28, FontStyle.Bold),
                            colFont = new Font("微軟正黑體", 14, FontStyle.Bold),
                            rowFont = new Font("微軟正黑體", 12);
                        var line = new Line() { Color = Color.Black, OuterWidth = 20 };
                        var headerColor = Color.FromArgb(255, 192, 0);


                        tb[startRow, 0].Formula = key == "中央列管出流管制計畫書審查調查表" ? "轄管河川局" : " 直轄市、縣(市)";
                        tb[startRow, 1].Formula = "案件編號";
                        tb[startRow, 2].Formula = "開發計畫名稱";
                        tb[startRow, 3].Formula = "申請義務人";
                        tb[startRow, 4].Formula = "技師簽證單位";
                        tb[startRow, 5].Formula = "案件審查階段";
                        tb[startRow, 6].Formula = "案件填報狀態";
                        tb[startRow, 7].Formula = "案件待確認資料";
                        tb[startRow, 8].Formula = "應填報對象";
                        tb[startRow, 9].Formula = "案件創建日期";
                        tb[startRow, 10].Formula = "案件異動日期";
                        tb[startRow, 11].Formula = "資料未異動\r\n(日曆天)";
                        tb[0, startRow, 11, startRow].SetFont(colFont);
                        tb[0, startRow, 11, startRow].SetBorder(line);
                        tb[0, startRow, 11, startRow].Background = headerColor;
                        #endregion

                        #region 內文設定
                        JArray outflowPlanStatus = null;
                        JObject ouflowCompareStatus = null;
                        // 讀取案件目前對應狀態名稱
                        using (StreamReader r = new StreamReader(Path.Combine(_appDataPath, "outflowPlanStatus.json")))
                        {
                            string jsonStr = r.ReadToEnd();
                            outflowPlanStatus = JArray.Parse(jsonStr);
                        }
                        // 讀取目前狀態對應的資料
                        using (StreamReader r = new StreamReader(Path.Combine(_appDataPath, "outflowCompareStatus.json")))
                        {
                            string jsonStr = r.ReadToEnd();
                            ouflowCompareStatus = JObject.Parse(jsonStr);
                        }

                        // 中央或地方
                        var filterSummary = summaryObj.Children<JObject>().Where(x => x["OFP_ReviewClass"].ToString() == types[key]);

                        startRow = 1; // 要調整明細位置只要改這個數值就好了
                        foreach (var data in filterSummary)
                        {
                            tb[startRow, 0].Formula = data["OFP_Gov"].ToString();
                            tb[startRow, 1].Formula = data["OFP_NO"].ToString();
                            tb[startRow, 2].Formula = data["OFP_Name"].ToString();
                            tb[startRow, 3].Formula = data["Payer"].ToString();
                            tb[startRow, 4].Formula = data["PracticeUnits"].ToString();

                            #region 判斷案件階段資訊
                            // 從後面開始判斷，有判斷到就是最新的資訊
                            var reversed = outflowPlanStatus.Reverse();
                            foreach (var item in reversed)
                            {
                                var statusCol = item["statusCol"].ToString();
                                if (!string.IsNullOrEmpty(data[statusCol].ToString()) && data[statusCol].ToString() != "0")
                                {
                                    tb[startRow, 5].Formula = item["cText"].ToString();
                                    var status = item["status"] as JArray;
                                    // 取得對應text的值
                                    tb[startRow, 6].Formula = status.Children<JObject>().FirstOrDefault(x => x["val"].ToString() == data[statusCol].ToString())["text"].ToString();
                                    break;
                                }

                                // 已經是最後一個狀態的話就直接寫入
                                if (item["statusCol"].ToString() == reversed.LastOrDefault()["statusCol"].ToString())
                                {
                                    if (string.IsNullOrEmpty(data[item["statusCol"].ToString()].ToString()) || data[item["statusCol"].ToString()].ToString() == "0")
                                    {
                                        tb[startRow, 5].Formula = item["cText"].ToString();
                                        var status = item["status"] as JArray;
                                        // 取得對應text的值
                                        tb[startRow, 6].Formula = status.Children<JObject>().FirstOrDefault(x => x["val"].ToString() == "0")["text"].ToString();
                                        break;
                                    }
                                }
                            }

                            var ComfirmStatus = ouflowCompareStatus["ComfirmStatus"] as JObject;
                            if (ComfirmStatus.ContainsKey(tb[startRow, 6].Formula))
                            {
                                tb[startRow, 7].Formula = ComfirmStatus[tb[startRow, 6].Formula].ToString();
                            }

                            var NeedFillUser = ouflowCompareStatus["NeedFillUser"] as JObject;
                            if (NeedFillUser.ContainsKey(tb[startRow, 6].Formula))
                            {
                                tb[startRow, 8].Formula = NeedFillUser[tb[startRow, 6].Formula].ToString();
                            }
                            #endregion

                            tb[startRow, 9].Formula = $"創建人員：{data["CreateBy"]}\r\n創建時間：{data["CreateDate"]}";
                            tb[startRow, 10].Formula = $"更新人員：{data["UpdateBy"]}\r\n更新時間：{data["UpdateDate"]}";

                            var days = DateTime.Now.Subtract(Convert.ToDateTime(data["UpdateDate"])).Days;
                            tb[startRow, 11].Formula = data["CaseStatus"].ToString() != "已核定" ? days.ToString() : "";


                            startRow++;
                        }
                        #endregion

                        //調整寬的大小
                        for (int i = 0; i < tb.ColumnCount; i++)
                            tb.Columns[i].AutoWidth = true;
                    }






                    //移除預設頁
                    calc.Tables.Remove(calc.Tables[0]);
                    //儲存並關閉
                    calc.SaveAs(filePath);
                    calc.Close();
                }
                #endregion
            }
            catch (Exception ex)
            {
                return HttpHelper.FailResult(ex.ToString());
            }

            return HttpHelper.FileResult(filePath, "application/vnd.oasis.opendocument.spreadsheet");
        }
        #endregion

        #region 計畫書案件總表
        [HttpGet, ActionName(nameof(summary_all))]
        [Route("api/Report/summary_all")]
        public HttpResponseMessage summary_all()
        {
            string FileName = "summary_all";
            var filePath = Path.Combine(_filePath, "ReportTempFile", $"{FileName}.ods");

            try
            {

                var calc = new Calc();
                // Summary_OutflowCtrlPlans_All_Calc(calc);
                //Detail_OutflowCtrlPlans_All_Calc(calc);
                Summary_OutflowCtrlPlans_All_Calc_Drain(calc);
                Detail_OutflowCtrlPlans_All_Calc_Drain_Central(calc);
                Detail_OutflowCtrlPlans_All_Calc_Drain_Local(calc);
                Summary_OutflowCtrlPlans_All_Calc_Outflow(calc);
                Detail_OutflowCtrlPlans_All_Calc_Outflow_Central(calc);
                Detail_OutflowCtrlPlans_All_Calc_Outflow_Local(calc);
                calc.Tables.Remove(calc.Tables[0]);
                calc.SaveAs(filePath);
                calc.Close();
            }
            catch (Exception ex)
            {
                return HttpHelper.FailResult(ex.ToString());
            }
            return HttpHelper.FileResult(filePath, "application/vnd.oasis.opendocument.spreadsheet");
        }

        /// <summary>
        /// 產生地方排水計畫書審查調查表至Excel的工作表中
        /// </summary>
        /// <param name="calc">LibreOffice Calc主體</param>
        public void Detail_OutflowCtrlPlans_All_Calc_Drain_Central(Calc calc)
        {
            var values = new DynamicParameters();
            values.Add("@OFP_ReviewClass", "中央審查");
            values.Add("@OFP_Type", "排水計畫書");

            var data = UtilDB.GetDataList<dynamic>(Generate_Summary_Detail_QueryString(), values);
            var tmpData = GetDictionaryFromDapper(data);
            Table tb = calc.Tables.AddNew("中央列管排水計畫書審查調查表");

            tb[0, 0].Formula = "中央列管排水計畫書審查調查表";
            tb[0, 0, 28, 0].Merge = true;
            tb[0, 0].HorizonAlign = Alignment.HorizonAlignment.CENTER;
            tb[0, 0].SetFont(new Font("微軟正黑體", 14, FontStyle.Bold));

            var fields = new string[] {"", "轄管河川局"," 直轄市、縣(市)","開發計畫名稱","開發基地面積(公頃)","開發基地排入之河川或排水名稱","開發基地排入之河川或排水所屬水系(請依流域綜合治理計畫核定水系填寫)"
                , "是否需先送排水規劃書(即分兩階段審查)(請填: 是或否)", "排水規劃書受理日期(即第一階段)","排水規劃書核定函日期及文號(未核定者免填)","排水規劃書是否已審查通過(請填: 是或否)(毋需提送者請填: 毋需提送)"
                , "排水計畫書(即第二階段)受理日期","排水計畫書是否已審查通過)(請填: 是或否)","排水計畫核定函日期及文號(未核定者免填)","排水計畫核定滯(蓄)洪池座數(座)(未核定者免填)","排水計畫核定滯洪總體積(萬m3)(未核定者免填)"
                ,"排水計畫書核定減洪設施工程是否已申報開工(請填: 是或否)","排水計畫書核定減洪設施工程開工日(格式: 108.5.14)","排水計畫書核定減洪設施工程是否已完成(請填: 是或否)", "排水計畫書核定減洪設施工程完工日(格式: 108.5.14)"
                ,"累計查核開發基地次數(排水計畫書核定後)","簽證技師科別", "簽證技師","備註", "檔案收集","系統清查","案件編號","義務人","主管機關"};
            var fieldMap = Generate_Summary_Detail_fieldMap();
            var units = new Dictionary<string, string[]>()
            {
                {"第一河川局", new string[]{"宜蘭縣"} },
                {"第十河川局", new string[]{"臺北市", "基隆市", "新北市"} },
                {"第二河川局", new string[]{"桃園市", "新竹縣", "新竹市", "苗栗縣"} },
                {"第三河川局", new string[]{"臺中市", "南投縣"} },
                {"第四河川局", new string[]{"彰化縣"} },
                {"第五河川局", new string[]{"雲林縣", "嘉義縣", "嘉義市"} },
                {"第六河川局", new string[]{"臺南市", "高雄市"} },
                {"第七河川局", new string[]{"屏東縣", "澎湖縣"} },
                {"第八河川局", new string[]{"臺東縣", "金門縣"} },
                {"第九河川局", new string[]{"花蓮縣"} }
            };
            for (int i = 0; i < fields.Length; i++)
                tb[1, i].Formula = fields[i];
            tb[0, 1, fields.Length - 1, 1].HorizonAlign = Alignment.HorizonAlignment.CENTER;
            tb[0, 1, fields.Length - 1, 1].SetFont(new Font("微軟正黑體", 12, FontStyle.Bold));

            var sumCells = new Dictionary<string, List<string>>() { { "A", new List<string>()}, { "E", new List<string>()}, { "H", new List<string>()}, { "K", new List<string>() },
                {"M", new List<string>() }, { "O", new List<string>()}, { "P", new List<string>()}, {"Q", new List<string>() }, { "S", new List<string>()} };
            var start = 2;
            foreach (var unit in units.Keys)
            {
                foreach (var coun in units[unit])
                {
                    var list = tmpData.Where(d => d["OFZ_COU"] == coun);
                    int sn = 0;
                    foreach (var l in list)
                    {
                        sn++;
                        tb[start, 0].Value = sn;
                        tb[start, 1].Formula = unit;
                        tb[start, 2].Formula = coun;
                        foreach (var fi in fieldMap.Keys)
                        {
                            var aa = fieldMap[fi].Split(',').Select(f => f.Length == 0 ? "" : (l[f] ?? ""));
                            tb[start, fi].Formula = string.Join("\r\n", fieldMap[fi].Split(',').Select(f => f.Length == 0 ? "" : (l[f] ?? "").GetType() == typeof(DateTime) ? $"{((DateTime)l[f]).Year - 1911}.{l[f]:MM.dd}" : l[f] ?? ""));
                        }
                        start++;
                    }
                    if (sn == 0)
                    {
                        tb[start, 1].Formula = unit;
                        tb[start, 2].Formula = coun;
                        start++;
                    }
                    tb[start, 0].Formula = $"=A{start}";
                    tb[start, 1].Formula = "小計";
                    foreach (var key in sumCells.Keys)
                    {
                        var c = Convert.ToInt32(key[0]) - Convert.ToInt32('A');
                        var cell = tb[start, c];
                        if (fieldMap.ContainsKey(c))
                        {
                            if (Array.IndexOf(new string[] { "OFP_FacilityNum", "OFP_FacilityArea", "OFP_Area" }, fieldMap[c]) >= 0)
                                cell.Formula = sn > 0 ? $"=sum({key}{start - sn + 1}:{key}{start})".Replace($"{key}{start}:{key}{start}", $"{key}{start}") : $"=sum({key}{start})";
                            else
                                cell.Formula = sn > 0 ? $"=COUNTIF({key}{start - sn + 1}:{key}{start};\"是\")".Replace($"{key}{start}:{key}{start}", $"{key}{start}") : $"=COUNTIF({key}{start};\"是\")";
                        }
                        else
                            cell.Formula = $"={key}{start}";
                        sumCells[key].Add($"{key}{start + 1}");
                    }
                    start++;
                }
            }
            tb[start, 1].Formula = "合計";
            foreach (var key in sumCells.Keys)
                tb[start, Convert.ToInt32(key[0]) - Convert.ToInt32('A')].Formula = $"={string.Join("+", sumCells[key])}";
            var line = new Line() { Color = Color.Black, OuterWidth = 30 };
            tb[0, 1, fieldMap.Keys.Max(), start].SetBorder(line, line, line, line, line, line);
            tb[0, 1, fieldMap.Keys.Max(), start].Wrap = true;
            tb[0, 1, fieldMap.Keys.Max(), start].VerticalAlign = Alignment.VerticalAlignment.CENTER;
            tb[0, 1, 2, start].HorizonAlign = Alignment.HorizonAlignment.CENTER;

            tb.Columns[4].NumberFormat = "#,##0.##";
            tb.Columns[15].NumberFormat = "#,##0.##";
            tb.Columns[7].HorizonAlign = Alignment.HorizonAlignment.CENTER;
            tb.Columns[10].HorizonAlign = Alignment.HorizonAlignment.CENTER;
            tb.Columns[12].HorizonAlign = Alignment.HorizonAlignment.CENTER;
            tb.Columns[16].HorizonAlign = Alignment.HorizonAlignment.CENTER;
            tb.Columns[18].HorizonAlign = Alignment.HorizonAlignment.CENTER;
        }

        private List<Dictionary<string, dynamic>> GetDictionaryFromDapper(List<dynamic> data)
        {
            List<Dictionary<string, dynamic>> collection = new List<Dictionary<string, dynamic>>();
            foreach (dynamic rec in data)
            {
                Dictionary<string, dynamic> dictionary = new Dictionary<string, dynamic>();
                var d = rec as IDictionary<string, object>;
                foreach (var a in d)
                {
                    dictionary.Add(a.Key, a.Value);
                }
                collection.Add(dictionary);
            }
            return collection;
        }

        /// <summary>
        /// 產生地方排水計畫書審查調查表至Excel的工作表中
        /// </summary>
        /// <param name="calc">LibreOffice Calc主體</param>
        public void Detail_OutflowCtrlPlans_All_Calc_Drain_Local(Calc calc)
        {
            var values = new DynamicParameters();
            values.Add("@OFP_ReviewClass", "地方審查");
            values.Add("@OFP_Type", "排水計畫書");
            var data = UtilDB.GetDataList<dynamic>(Generate_Summary_Detail_QueryString(), values);
            var tmpData = GetDictionaryFromDapper(data);
            Table tb = calc.Tables.AddNew("地方列管排水計畫書審查調查表");

            tb[0, 0].Formula = "地方列管排水計畫書審查調查表";
            tb[0, 0, 28, 0].Merge = true;
            tb[0, 0].HorizonAlign = Alignment.HorizonAlignment.CENTER;
            tb[0, 0].SetFont(new Font("微軟正黑體", 14, FontStyle.Bold));

            var fields = new string[] {"", "轄管河川局"," 直轄市、縣(市)","開發計畫名稱","開發基地面積(公頃)","開發基地排入之河川或排水名稱","開發基地排入之河川或排水所屬水系(請依流域綜合治理計畫核定水系填寫)"
                , "是否需先送排水規劃書(即分兩階段審查)(請填: 是或否)", "排水規劃書受理日期(即第一階段)","排水規劃書核定函日期及文號(未核定者免填)","排水規劃書是否已審查通過(請填: 是或否)(毋需提送者請填: 毋需提送)"
                , "排水計畫書(即第二階段)受理日期","排水計畫書是否已審查通過)(請填: 是或否)","排水計畫核定函日期及文號(未核定者免填)","排水計畫核定滯(蓄)洪池座數(座)(未核定者免填)","排水計畫核定滯洪總體積(萬m3)(未核定者免填)"
                ,"排水計畫書核定減洪設施工程是否已申報開工(請填: 是或否)","排水計畫書核定減洪設施工程開工日(格式: 108.5.14)","排水計畫書核定減洪設施工程是否已完成(請填: 是或否)", "排水計畫書核定減洪設施工程完工日(格式: 108.5.14)"
                ,"累計查核開發基地次數(排水計畫書核定後)","簽證技師科別", "簽證技師","備註", "檔案收集","系統清查","案件編號","義務人","主管機關"};
            var fieldMap = Generate_Summary_Detail_fieldMap();
            var units = new Dictionary<string, string[]>()
            {
                {"第一河川局", new string[]{"宜蘭縣"} },
                {"第十河川局", new string[]{"臺北市", "基隆市", "新北市"} },
                {"第二河川局", new string[]{"桃園市", "新竹縣", "新竹市", "苗栗縣"} },
                {"第三河川局", new string[]{"臺中市", "南投縣"} },
                {"第四河川局", new string[]{"彰化縣"} },
                {"第五河川局", new string[]{"雲林縣", "嘉義縣", "嘉義市"} },
                {"第六河川局", new string[]{"臺南市", "高雄市"} },
                {"第七河川局", new string[]{"屏東縣", "澎湖縣"} },
                {"第八河川局", new string[]{"臺東縣", "金門縣"} },
                {"第九河川局", new string[]{"花蓮縣"} }
            };
            for (int i = 0; i < fields.Length; i++)
                tb[1, i].Formula = fields[i];
            tb[0, 1, fields.Length - 1, 1].HorizonAlign = Alignment.HorizonAlignment.CENTER;
            tb[0, 1, fields.Length - 1, 1].SetFont(new Font("微軟正黑體", 12, FontStyle.Bold));

            var sumCells = new Dictionary<string, List<string>>() { { "A", new List<string>()}, { "E", new List<string>()}, { "H", new List<string>()}, { "K", new List<string>() },
                {"M", new List<string>() }, { "O", new List<string>()}, { "P", new List<string>()}, {"Q", new List<string>() }, { "S", new List<string>()} };
            var start = 2;
            foreach (var unit in units.Keys)
            {
                foreach (var coun in units[unit])
                {
                    var list = tmpData.Where(d => d["OFZ_COU"] == coun);
                    int sn = 0;
                    foreach (var l in list)
                    {
                        sn++;
                        tb[start, 0].Value = sn;
                        tb[start, 1].Formula = unit;
                        tb[start, 2].Formula = coun;
                        foreach (var fi in fieldMap.Keys)
                        {
                            tb[start, fi].Formula = string.Join("\r\n", fieldMap[fi].Split(',').Select(f => f.Length == 0 ? "" : (l[f] ?? "").GetType() == typeof(DateTime) ? $"{((DateTime)l[f]).Year - 1911}.{l[f]:MM.dd}" : l[f] ?? ""));
                        }
                        start++;
                    }
                    if (sn == 0)
                    {
                        tb[start, 1].Formula = unit;
                        tb[start, 2].Formula = coun;
                        start++;
                    }
                    tb[start, 0].Formula = $"=A{start}";
                    tb[start, 1].Formula = "小計";
                    foreach (var key in sumCells.Keys)
                    {
                        var c = Convert.ToInt32(key[0]) - Convert.ToInt32('A');
                        var cell = tb[start, c];
                        if (fieldMap.ContainsKey(c))
                        {
                            if (Array.IndexOf(new string[] { "OFP_FacilityNum", "OFP_FacilityArea", "OFP_Area" }, fieldMap[c]) >= 0)
                                cell.Formula = sn > 0 ? $"=sum({key}{start - sn + 1}:{key}{start})".Replace($"{key}{start}:{key}{start}", $"{key}{start}") : $"=sum({key}{start})";
                            else
                                cell.Formula = sn > 0 ? $"=COUNTIF({key}{start - sn + 1}:{key}{start};\"是\")".Replace($"{key}{start}:{key}{start}", $"{key}{start}") : $"=COUNTIF({key}{start};\"是\")";
                        }
                        else
                            cell.Formula = $"={key}{start}";
                        sumCells[key].Add($"{key}{start + 1}");
                    }
                    start++;
                }
            }
            tb[start, 1].Formula = "合計";
            foreach (var key in sumCells.Keys)
                tb[start, Convert.ToInt32(key[0]) - Convert.ToInt32('A')].Formula = $"={string.Join("+", sumCells[key])}";
            var line = new Line() { Color = Color.Black, OuterWidth = 30 };
            tb[0, 1, fieldMap.Keys.Max(), start].SetBorder(line, line, line, line, line, line);
            tb[0, 1, fieldMap.Keys.Max(), start].Wrap = true;
            tb[0, 1, fieldMap.Keys.Max(), start].VerticalAlign = Alignment.VerticalAlignment.CENTER;
            tb[0, 1, 2, start].HorizonAlign = Alignment.HorizonAlignment.CENTER;

            tb.Columns[4].NumberFormat = "#,##0.##";
            tb.Columns[15].NumberFormat = "#,##0.##";
            tb.Columns[7].HorizonAlign = Alignment.HorizonAlignment.CENTER;
            tb.Columns[10].HorizonAlign = Alignment.HorizonAlignment.CENTER;
            tb.Columns[12].HorizonAlign = Alignment.HorizonAlignment.CENTER;
            tb.Columns[16].HorizonAlign = Alignment.HorizonAlignment.CENTER;
            tb.Columns[18].HorizonAlign = Alignment.HorizonAlignment.CENTER;
        }
        /// <summary>
        /// 產生地方排水計畫書審查調查表至Excel的工作表中
        /// </summary>
        /// <param name="calc">LibreOffice Calc主體</param>
        public void Detail_OutflowCtrlPlans_All_Calc_Outflow_Central(Calc calc)
        {
            var values = new DynamicParameters();
            values.Add("@OFP_ReviewClass", "中央審查");
            values.Add("@OFP_Type", "出流計畫書");
            var data = UtilDB.GetDataList<dynamic>(Generate_Summary_Detail_QueryString(), values);
            var tmpData = GetDictionaryFromDapper(data);
            Table tb = calc.Tables.AddNew("中央列管出流管制計畫書審查調查表");

            tb[0, 0].Formula = "中央列管出流管制計畫書審查調查表";
            tb[0, 0, 28, 0].Merge = true;
            tb[0, 0].HorizonAlign = Alignment.HorizonAlignment.CENTER;
            tb[0, 0].SetFont(new Font("微軟正黑體", 14, FontStyle.Bold));

            var fields = new string[] {"", "轄管河川局"," 直轄市、縣(市)","開發計畫名稱","開發基地面積(公頃)","開發基地排入之河川或排水名稱","開發基地排入之河川或排水所屬水系(請依流域綜合治理計畫核定水系填寫)"
                , "是否需先送出流管制規劃書(即分兩階段審查)(請填: 是或否)", "出流管制規劃書受理日期(即第一階段)","出流管制規劃書核定函日期及文號(未核定者免填)","出流管制規劃書是否已審查通過(請填: 是或否)(毋需提送者請填: 毋需提送)"
                , "出流管制計畫書(即第二階段)受理日期","出流管制計畫書是否已審查通過)(請填: 是或否)","出流管制計畫核定函日期及文號(未核定者免填)","出流管制計畫核定滯(蓄)洪池座數(座)(未核定者免填)","出流管制計畫核定滯洪總體積(萬m3)(未核定者免填)"
                ,"出流管制計畫書核定減洪設施工程是否已申報開工(請填: 是或否)","出流管制計畫書核定減洪設施工程開工日(格式: 108.5.14)","出流管制計畫書核定減洪設施工程是否已完成(請填: 是或否)", "出流管制計畫書核定減洪設施工程完工日(格式: 108.5.14)"
                ,"累計查核開發基地次數(出流管制計畫書核定後)","簽證技師科別", "簽證技師","備註", "檔案收集","系統清查","案件編號","義務人","主管機關"};
            var fieldMap = Generate_Summary_Detail_fieldMap();
            var units = new Dictionary<string, string[]>()
            {
                {"第一河川局", new string[]{"宜蘭縣"} },
                {"第十河川局", new string[]{"臺北市", "基隆市", "新北市"} },
                {"第二河川局", new string[]{"桃園市", "新竹縣", "新竹市", "苗栗縣"} },
                {"第三河川局", new string[]{"臺中市", "南投縣"} },
                {"第四河川局", new string[]{"彰化縣"} },
                {"第五河川局", new string[]{"雲林縣", "嘉義縣", "嘉義市"} },
                {"第六河川局", new string[]{"臺南市", "高雄市"} },
                {"第七河川局", new string[]{"屏東縣", "澎湖縣"} },
                {"第八河川局", new string[]{"臺東縣", "金門縣"} },
                {"第九河川局", new string[]{"花蓮縣"} }
            };
            for (int i = 0; i < fields.Length; i++)
                tb[1, i].Formula = fields[i];
            tb[0, 1, fields.Length - 1, 1].HorizonAlign = Alignment.HorizonAlignment.CENTER;
            tb[0, 1, fields.Length - 1, 1].SetFont(new Font("微軟正黑體", 12, FontStyle.Bold));

            var sumCells = new Dictionary<string, List<string>>() { { "A", new List<string>()}, { "E", new List<string>()}, { "H", new List<string>()}, { "K", new List<string>() },
                {"M", new List<string>() }, { "O", new List<string>()}, { "P", new List<string>()}, {"Q", new List<string>() }, { "S", new List<string>()} };
            var start = 2;
            foreach (var unit in units.Keys)
            {
                foreach (var coun in units[unit])
                {
                    var list = tmpData.Where(d => d["OFZ_COU"] == coun);
                    int sn = 0;
                    foreach (var l in list)
                    {
                        sn++;
                        tb[start, 0].Value = sn;
                        tb[start, 1].Formula = unit;
                        tb[start, 2].Formula = coun;
                        foreach (var fi in fieldMap.Keys)
                        {
                            tb[start, fi].Formula = string.Join("\r\n", fieldMap[fi].Split(',').Select(f => f.Length == 0 ? "" : (l[f] ?? "").GetType() == typeof(DateTime) ? $"{((DateTime)l[f]).Year - 1911}.{l[f]:MM.dd}" : l[f] ?? ""));
                        }
                        start++;
                    }
                    if (sn == 0)
                    {
                        tb[start, 1].Formula = unit;
                        tb[start, 2].Formula = coun;
                        start++;
                    }
                    tb[start, 0].Formula = $"=A{start}";
                    tb[start, 1].Formula = "小計";
                    foreach (var key in sumCells.Keys)
                    {
                        var c = Convert.ToInt32(key[0]) - Convert.ToInt32('A');
                        var cell = tb[start, c];
                        if (fieldMap.ContainsKey(c))
                        {
                            if (Array.IndexOf(new string[] { "OFP_FacilityNum", "OFP_FacilityArea", "OFP_Area" }, fieldMap[c]) >= 0)
                                cell.Formula = sn > 0 ? $"=sum({key}{start - sn + 1}:{key}{start})".Replace($"{key}{start}:{key}{start}", $"{key}{start}") : $"=sum({key}{start})";
                            else
                                cell.Formula = sn > 0 ? $"=COUNTIF({key}{start - sn + 1}:{key}{start};\"是\")".Replace($"{key}{start}:{key}{start}", $"{key}{start}") : $"=COUNTIF({key}{start};\"是\")";
                        }
                        else
                            cell.Formula = $"={key}{start}";
                        sumCells[key].Add($"{key}{start + 1}");
                    }
                    start++;
                }
            }
            tb[start, 1].Formula = "合計";
            foreach (var key in sumCells.Keys)
                tb[start, Convert.ToInt32(key[0]) - Convert.ToInt32('A')].Formula = $"={string.Join("+", sumCells[key])}";
            var line = new Line() { Color = Color.Black, OuterWidth = 30 };
            tb[0, 1, fieldMap.Keys.Max(), start].SetBorder(line, line, line, line, line, line);
            tb[0, 1, fieldMap.Keys.Max(), start].Wrap = true;
            tb[0, 1, fieldMap.Keys.Max(), start].VerticalAlign = Alignment.VerticalAlignment.CENTER;
            tb[0, 1, 2, start].HorizonAlign = Alignment.HorizonAlignment.CENTER;

            tb.Columns[4].NumberFormat = "#,##0.##";
            tb.Columns[15].NumberFormat = "#,##0.##";
            tb.Columns[7].HorizonAlign = Alignment.HorizonAlignment.CENTER;
            tb.Columns[10].HorizonAlign = Alignment.HorizonAlignment.CENTER;
            tb.Columns[12].HorizonAlign = Alignment.HorizonAlignment.CENTER;
            tb.Columns[16].HorizonAlign = Alignment.HorizonAlignment.CENTER;
            tb.Columns[18].HorizonAlign = Alignment.HorizonAlignment.CENTER;
        }
        /// <summary>
        /// 產生地方排水計畫書審查調查表至Excel的工作表中
        /// </summary>
        /// <param name="calc">LibreOffice Calc主體</param>
        public void Detail_OutflowCtrlPlans_All_Calc_Outflow_Local(Calc calc)
        {
            var values = new DynamicParameters();
            values.Add("@OFP_ReviewClass", "地方審查");
            values.Add("@OFP_Type", "出流計畫書");
            var data = UtilDB.GetDataList<dynamic>(Generate_Summary_Detail_QueryString(), values);
            var tmpData = GetDictionaryFromDapper(data);
            Table tb = calc.Tables.AddNew("地方列管出流管制計畫書審查調查表");

            tb[0, 0].Formula = "地方列管出流管制計畫書審查調查表";
            tb[0, 0, 28, 0].Merge = true;
            tb[0, 0].HorizonAlign = Alignment.HorizonAlignment.CENTER;
            tb[0, 0].SetFont(new Font("微軟正黑體", 14, FontStyle.Bold));

            var fields = new string[] {"", "轄管河川局"," 直轄市、縣(市)","開發計畫名稱","開發基地面積(公頃)","開發基地排入之河川或排水名稱","開發基地排入之河川或排水所屬水系(請依流域綜合治理計畫核定水系填寫)"
                , "是否需先送出流管制規劃書(即分兩階段審查)(請填: 是或否)", "出流管制規劃書受理日期(即第一階段)","出流管制規劃書核定函日期及文號(未核定者免填)","出流管制規劃書是否已審查通過(請填: 是或否)(毋需提送者請填: 毋需提送)"
                , "出流管制計畫書(即第二階段)受理日期","出流管制計畫書是否已審查通過)(請填: 是或否)","出流管制計畫核定函日期及文號(未核定者免填)","出流管制計畫核定滯(蓄)洪池座數(座)(未核定者免填)","出流管制計畫核定滯洪總體積(萬m3)(未核定者免填)"
                ,"出流管制計畫書核定減洪設施工程是否已申報開工(請填: 是或否)","出流管制計畫書核定減洪設施工程開工日(格式: 108.5.14)","出流管制計畫書核定減洪設施工程是否已完成(請填: 是或否)", "出流管制計畫書核定減洪設施工程完工日(格式: 108.5.14)"
                ,"累計查核開發基地次數(出流管制計畫書核定後)","簽證技師科別", "簽證技師","備註", "檔案收集","系統清查","案件編號","義務人","主管機關"};
            var fieldMap = Generate_Summary_Detail_fieldMap();
            var units = new Dictionary<string, string[]>()
            {
                {"第一河川局", new string[]{"宜蘭縣"} },
                {"第十河川局", new string[]{"臺北市", "基隆市", "新北市"} },
                {"第二河川局", new string[]{"桃園市", "新竹縣", "新竹市", "苗栗縣"} },
                {"第三河川局", new string[]{"臺中市", "南投縣"} },
                {"第四河川局", new string[]{"彰化縣"} },
                {"第五河川局", new string[]{"雲林縣", "嘉義縣", "嘉義市"} },
                {"第六河川局", new string[]{"臺南市", "高雄市"} },
                {"第七河川局", new string[]{"屏東縣", "澎湖縣"} },
                {"第八河川局", new string[]{"臺東縣", "金門縣"} },
                {"第九河川局", new string[]{"花蓮縣"} }
            };
            for (int i = 0; i < fields.Length; i++)
                tb[1, i].Formula = fields[i];
            tb[0, 1, fields.Length - 1, 1].HorizonAlign = Alignment.HorizonAlignment.CENTER;
            tb[0, 1, fields.Length - 1, 1].SetFont(new Font("微軟正黑體", 12, FontStyle.Bold));

            var sumCells = new Dictionary<string, List<string>>() { { "A", new List<string>()}, { "E", new List<string>()}, { "H", new List<string>()}, { "K", new List<string>() },
                {"M", new List<string>() }, { "O", new List<string>()}, { "P", new List<string>()}, {"Q", new List<string>() }, { "S", new List<string>()} };
            var start = 2;
            foreach (var unit in units.Keys)
            {
                foreach (var coun in units[unit])
                {
                    var list = tmpData.Where(d => d["OFZ_COU"] == coun);
                    int sn = 0;
                    foreach (var l in list)
                    {
                        sn++;
                        tb[start, 0].Value = sn;
                        tb[start, 1].Formula = unit;
                        tb[start, 2].Formula = coun;
                        foreach (var fi in fieldMap.Keys)
                        {
                            tb[start, fi].Formula = string.Join("\r\n", fieldMap[fi].Split(',').Select(f => f.Length == 0 ? "" : (l[f] ?? "").GetType() == typeof(DateTime) ? $"{((DateTime)l[f]).Year - 1911}.{l[f]:MM.dd}" : l[f] ?? ""));
                        }
                        start++;
                    }
                    if (sn == 0)
                    {
                        tb[start, 1].Formula = unit;
                        tb[start, 2].Formula = coun;
                        start++;
                    }
                    tb[start, 0].Formula = $"=A{start}";
                    tb[start, 1].Formula = "小計";
                    foreach (var key in sumCells.Keys)
                    {
                        var c = Convert.ToInt32(key[0]) - Convert.ToInt32('A');
                        var cell = tb[start, c];
                        if (fieldMap.ContainsKey(c))
                        {
                            if (Array.IndexOf(new string[] { "OFP_FacilityNum", "OFP_FacilityArea", "OFP_Area" }, fieldMap[c]) >= 0)
                                cell.Formula = sn > 0 ? $"=sum({key}{start - sn + 1}:{key}{start})".Replace($"{key}{start}:{key}{start}", $"{key}{start}") : $"=sum({key}{start})";
                            else
                                cell.Formula = sn > 0 ? $"=COUNTIF({key}{start - sn + 1}:{key}{start};\"是\")".Replace($"{key}{start}:{key}{start}", $"{key}{start}") : $"=COUNTIF({key}{start};\"是\")";
                        }
                        else
                            cell.Formula = $"={key}{start}";
                        sumCells[key].Add($"{key}{start + 1}");
                    }
                    start++;
                }
            }
            tb[start, 1].Formula = "合計";
            foreach (var key in sumCells.Keys)
                tb[start, Convert.ToInt32(key[0]) - Convert.ToInt32('A')].Formula = $"={string.Join("+", sumCells[key])}";
            var line = new Line() { Color = Color.Black, OuterWidth = 30 };
            tb[0, 1, fieldMap.Keys.Max(), start].SetBorder(line, line, line, line, line, line);
            tb[0, 1, fieldMap.Keys.Max(), start].Wrap = true;
            tb[0, 1, fieldMap.Keys.Max(), start].VerticalAlign = Alignment.VerticalAlignment.CENTER;
            tb[0, 1, 2, start].HorizonAlign = Alignment.HorizonAlignment.CENTER;

            tb.Columns[4].NumberFormat = "#,##0.##";
            tb.Columns[15].NumberFormat = "#,##0.##";
            tb.Columns[7].HorizonAlign = Alignment.HorizonAlignment.CENTER;
            tb.Columns[10].HorizonAlign = Alignment.HorizonAlignment.CENTER;
            tb.Columns[12].HorizonAlign = Alignment.HorizonAlignment.CENTER;
            tb.Columns[16].HorizonAlign = Alignment.HorizonAlignment.CENTER;
            tb.Columns[18].HorizonAlign = Alignment.HorizonAlignment.CENTER;
        }


        public string GetOutflowLegalCase()
        {
            var sql = @"  SELECT ofp.OFP_ID
                          , ofp.OFP_Year
                          , ofp.OFP_NO
                          , ofp.OFP_Name
                          , ofp.OFP_Type
                          , ofp.OFP_Gov
                          , ofp.OFP_SiteWGS84X
                          , ofp.OFP_SiteWGS84Y
                          , '已核定' AS IsChecked
                          , (
                                CASE WHEN ofp.OFP_Type LIKE '%計畫書%' THEN
                                (
                                    CASE WHEN otl.EN_END_Status = 1 THEN '已完工' 
                                    ELSE
                                    (
                                        CASE WHEN EN_ST_Status2 = 1 THEN '已開工' ELSE '未開工' END
                                    ) END
                                ) 
                                ELSE '--' END 
                            ) AS EN_Status
                          , lo.COUNTY
                          , lo.TOWN
                          , lo.SEC_NA
                          , lo.LANDNO
                          , lo.AREA
                          , lo.OFP_AREA
                          FROM OutflowControlPlan ofp
                          INNER JOIN OFPZoneTimeline otl ON ofp.OFP_ID = (CASE WHEN ofp.OFP_Type LIKE '%規劃書%' THEN otl.OFP_ID WHEN ofp.OFP_Type LIKE '%計畫書%' THEN otl.OFP_ID2 END)
                          INNER JOIN LandOwner lo ON ofp.OFP_ID = lo.OFP_ID
                          WHERE ofp.IsShow = 1 AND((ofp.OFP_Type LIKE '%規劃書%' AND(otl.CA_Acc_Status >= 1 AND Approved_Status = 1 AND CA_Acc_Status != 4)) OR (ofp.OFP_Type LIKE '%計畫書%' AND(otl.CA_Acc_Status2 >= 1 AND otl.Approved_Status2 = 1)))
                          ORDER BY ofp.OFP_Year";

            // 查詢
            var datas = UtilDB.GetDataList<dynamic>(sql, null);
            StringBuilder sbResult = new StringBuilder();
            sbResult.Append($"計畫ID,年度,案號,計畫名稱,計畫類別,受理主管機關,聯外排水路匯入區域排水或河川位置(經度),聯外排水路匯入區域排水或河川位置(緯度),審理階段,施工階段,縣市,鄉鎮,地段,地號,謄本面積,申請面積{Environment.NewLine}");
            foreach (var item in datas)
            {
                sbResult.Append($"{item.OFP_ID.ToString()},{item.OFP_Year.ToString()},{item.OFP_NO.ToString()},{item.OFP_Name.ToString()},{item.OFP_Type.ToString()},{item.OFP_Gov.ToString()},{item.OFP_SiteWGS84X.ToString()},{item.OFP_SiteWGS84Y.ToString()},{item.IsChecked.ToString()},{item.EN_Status.ToString()},{item.COUNTY.ToString()},{item.TOWN.ToString()},{item.SEC_NA.ToString()},\t{item.LANDNO.ToString()},{item.AREA.ToString()},{item.OFP_AREA.ToString()},{Environment.NewLine}");
            }

            return sbResult.ToString();

        }

        /// <summary>
        /// 產生Summary明細表的SQL語法
        /// </summary>
        /// <returns>查詢字串</returns>
        private string Generate_Summary_Detail_QueryString()
        {
            return @"select b.OFP_Cou OFZ_COU, b.OFP_Name OFZ_Name, b.OFP_NO, b.OFP_Gov, payer.Payer, a.OFP_Area
                    , case a.CA_Acc_Status when 4 then '否' else '是' end CA_Acc_Status
                    , a.RVGov_Accepted_Date, a.Approved_Date, a.Approved_NO
                    , case when (a.Plan_StatusSet in ('是', '1')) then '是' when (a.Plan_StatusSet in ('3')) then '毋需提送' else '否' end Plan_StatusSet
                    , a.RVGov_Accepted_Date2
                    , case when (a.Plan_StatusSet2 in ('1', '是')) then '是' when (a.Plan_StatusSet2 in ('0', '否')) then '否' else a.Plan_StatusSet2 end Plan_StatusSet2
                    , a.Approved_Date2, a.Approved_NO2
                    , case when (a.EN_ST_Status = 1) then '是' when (a.EN_ST_Status = 0) then '否' else cast(a.EN_ST_Status as varchar) end EN_ST_Status
                    , a.EN_StartDate
                    , case when (a.EN_END_Status = 1) then '是' when(a.EN_END_Status = 0) then '否' else cast(a.EN_END_Status as varchar) end EN_END_Status
                    , a.EN_EndDate, a.EN_END_Time, b.OFP_DrainName, b.OFP_RiverName, b.OFP_FacilityNum, b.OFP_FacilityArea, c.EngineerName, c.EngineerType, c.PracticeUnits
                    , case when (f.FileCount > 0) then '有' else '無' end hasFile
                    , --case when (f.FileCount > 0) then '有' else '無' end + '紀錄' sysCheckｍ,
                    datelist sysCheck
                       from OFPZoneTimeline a join OutflowControlPlan b on a.OFP_ID2 = b.OFP_ID
                       LEFT JOIN payer ON payer.PA_ID = b.PA_ID
                       LEFT JOIN (
                            SELECT ED_ID,EngineerName,PracticeUnits
                            , ( SELECT  EngineerType + '、'  FROM [06-outflow].[dbo].[V_Enginners] WHERE ED_ID = Ve.ED_ID For Xml Path('')) as EngineerType FROM [06-outflow].[dbo].[V_Enginners] Ve GROUP BY ED_ID, EngineerName, PracticeUnits
                       ) c ON b.Engineer  + ',' like cast(c.ED_ID as varchar) + ',%' 
                       --left join v_Enginners c on b.Engineer + ',' like cast(c.ED_ID as varchar) + ',%' 
                       left join (select OFP_ID, count(*) FileCount from [File] group by OFP_ID) f on b.OFP_ID = f.OFP_ID
                                        LEFT JOIN (
                                            SELECT DISTINCT  [OFP_ID],
                                                datelist=
                                                (
                                                    SELECT cast([File_Type] + '：' + [path] + [File_name] AS NVARCHAR(max) ) + '； ' 
                                                    FROM [File]      
                                                    WHERE [OFP_ID]=t0.[OFP_ID]    --把name一樣的加起來
                                                    FOR XML PATH('')
                                                )
                                            FROM [File] t0) g ON b.OFP_ID = g.OFP_ID
                    where b.IsShow = 1 and  b.OFP_ReviewClass = @OFP_ReviewClass and b.OFP_Type = @OFP_Type
                                        ";
        }

        /// <summary>
        /// 產生資料欄位應資訊
        /// </summary>
        /// <returns>對應的欄位資訊</returns>
        private Dictionary<int, string> Generate_Summary_Detail_fieldMap()
        {
            return new Dictionary<int, string>()
            {
                {3, "OFZ_Name" }, {4, "OFP_Area"}, {5, "OFP_DrainName"}, {6, "OFP_RiverName"}, {7, "CA_Acc_Status"}, {8, "RVGov_Accepted_Date"}, {9, "Approved_Date,Approved_NO"}, {10, "Plan_StatusSet"},
                {11, "RVGov_Accepted_Date2" }, {12, "Plan_StatusSet2"}, {13, "Approved_Date2,Approved_NO2"},{14, "OFP_FacilityNum"}, {15, "OFP_FacilityArea"},
                {16, "EN_ST_Status" }, {17, "EN_StartDate"}, {18, "EN_END_Status"}, {19, "EN_EndDate" }, {20, "EN_END_Time"}, {21, "EngineerType"}, {22, "EngineerName"}, {23, "PracticeUnits"}, {24, "hasFile"}, {25, "sysCheck"},
                {26, "OFP_NO" }, {27, "Payer"}, {28, "OFP_Gov"}
            };
        }

        /// <summary>
        /// 排水計畫書案件總表_輸出Calc
        /// </summary>
        /// <param name="calc">LibreOffice Calc主體</param>
        public void Summary_OutflowCtrlPlans_All_Calc_Drain(Calc calc)
        {
            var tb = calc.Tables.AddNew("排水計畫書調查總表");// calc.Tables[0];
            //tb.Name = "排水計畫書調查總表";
            Font headerFont = new Font("微軟正黑體", 14, FontStyle.Bold),
                colFont = new Font("微軟正黑體", 12, FontStyle.Bold);
            //表格標題
            tb[0, 0].Formula = "區域排水排水計畫書審查統計表（統計時間自98年起）";
            tb["A1:E1"].Merge = true;
            tb["A1:E1"].HorizonAlign = Alignment.HorizonAlignment.CENTER;
            tb[0, 0].SetFont(headerFont);

            tb[1, 4].Formula = $"{DateTime.Now.Year - 1911:000}.{DateTime.Now:MM.dd}";
            tb[1, 4].HorizonAlign = Alignment.HorizonAlignment.RIGHT;

            //總計
            var line = new Line() { Color = Color.Black, OuterWidth = 20 };
            tb[2, 0].Formula = "中央+地方";
            tb[2, 1].Formula = "受理件數";
            tb[2, 2].Formula = "通過件數";
            tb[2, 3].Formula = "核定滯洪池座數";
            tb[2, 4].Formula = "核定滯洪池體積(萬m3)";
            tb[3, 0].Formula = "總計";
            tb[3, 1].Formula = "=B30+B41";
            tb[3, 2].Formula = "=C30+C41";
            tb[3, 3].Formula = "=D30+D41";
            tb[3, 4].Formula = "=E30+E41";
            tb["A3:E4"].SetBorder(line, line, line, line, line, line);
            tb[0, 2, 4, 2].SetFont(colFont);

            //縣市
            var local = Summary_OutflowCtrlPlans_Local_All_Drain();
            var counties = new string[] { "宜蘭縣", "臺北市", "基隆市", "新北市", "桃園市", "新竹縣", "新竹市", "苗栗縣", "臺中市", "南投縣", "彰化縣", "雲林縣", "嘉義縣", "嘉義市", "臺南市", "高雄市", "屏東縣", "澎湖縣", "臺東縣", "金門縣", "花蓮縣" };
            tb[5, 0].Formula = "地方列管區域排水排水計畫書審查統計表（統計時間自98年起）";
            tb[5, 0].SetFont(headerFont);
            tb["A6:E6"].Merge = true;
            tb["A6:E6"].HorizonAlign = Alignment.HorizonAlignment.CENTER;
            tb[7, 0].Formula = "縣（市）";
            tb[7, 1].Formula = "受理件數";
            tb[7, 2].Formula = "通過件數";
            tb[7, 3].Formula = "核定滯洪池座數";
            tb[7, 4].Formula = "核定滯洪池體積(萬m3)";
            tb[0, 7, 4, 7].SetFont(colFont);

            for (int i = 8; i < 8 + counties.Length; i++)
            {
                var l = local.FirstOrDefault(ll => ll["OFP_Gov"] == counties[i - 8]);
                tb[i, 0].Formula = counties[i - 8];
                tb[i, 1].Formula = $"{l?["Count"] ?? 0}";
                tb[i, 2].Formula = $"{l?["App_Count"] ?? 0}";
                tb[i, 3].Formula = $"{l?["App_FacilityNum"] ?? 0}";
                tb[i, 4].Formula = $"{l?["App_FacilityArea"] ?? 0}";
            }
            var start = counties.Length + 8;
            tb[start, 0].Formula = "合計";
            tb[start, 1].Formula = $"=sum(B9:B{start})";
            tb[start, 2].Formula = $"=sum(C9:C{start})";
            tb[start, 3].Formula = $"=sum(D9:D{start})";
            tb[start, 4].Formula = $"=sum(E9:E{start})";
            tb[$"A8:E{start + 1}"].SetBorder(line, line, line, line, line, line);
            tb[0, 7, 4, start].SetBorder(line, line, line, line, line, line);

            start = start + 2;
            tb[start, 0].Formula = "中央列管區域排水排水計畫書審查統計表（統計時間自98年起）";
            tb[start, 0].SetFont(headerFont);
            tb[$"A{start + 1}:E{start + 1}"].Merge = true;
            tb[$"A{start + 1}:E{start + 1}"].HorizonAlign = Alignment.HorizonAlignment.CENTER;
            var unit = new string[] { "二", "三", "五", "六", "七", "十" };
            var central = Summary_OutflowCtrlPlans_Centroid_All_Drain();
            start = start + 2;
            tb[start, 0].Formula = "河川局";
            tb[start, 1].Formula = "受理件數";
            tb[start, 2].Formula = "通過件數";
            tb[start, 3].Formula = "核定滯洪池座數";
            tb[start, 4].Formula = "核定滯洪池體積(萬m3)";
            tb[0, start, 4, start].SetFont(colFont);
            start++;
            for (int i = 0; i < unit.Length; i++)
            {
                var c = central.FirstOrDefault(cc => cc["OFP_Gov"].IndexOf(unit[i]) > 0);
                tb[start + i, 0].Formula = unit[i];
                tb[start + i, 1].Formula = $"{c?["Count"] ?? 0}";
                tb[start + i, 2].Formula = $"{c?["App_Count"] ?? 0}";
                tb[start + i, 3].Formula = $"{c?["App_FacilityNum"] ?? 0}";
                tb[start + i, 4].Formula = $"{c?["App_FacilityArea"] ?? 0}";
            }
            start = start + unit.Length;
            tb[start, 0].Formula = "合計";
            tb[start, 1].Formula = $"=Sum(B{start - unit.Length + 1}:B{start})";
            tb[start, 2].Formula = $"=Sum(C{start - unit.Length + 1}:C{start})";
            tb[start, 3].Formula = $"=Sum(D{start - unit.Length + 1}:D{start})";
            tb[start, 4].Formula = $"=Sum(E{start - unit.Length + 1}:E{start})";

            tb[$"A{start - unit.Length}:E{start + 1}"].SetBorder(line, line, line, line, line, line);

            for (int i = 1; i < tb.ColumnCount; i++)
                tb.Columns[i].AutoWidth = true;
            tb.Columns[0].HorizonAlign = Alignment.HorizonAlignment.CENTER;
            tb.Columns[4].NumberFormat = "#,##0.##";
        }
        /// <summary>
        /// 出流計畫書案件總表_輸出Calc
        /// </summary>
        /// <param name="calc">LibreOffice Calc主體</param>
        public void Summary_OutflowCtrlPlans_All_Calc_Outflow(Calc calc)
        {
            var tb = calc.Tables.AddNew("出流管制計畫書調查總表"); // calc.Tables[0];
            //tb.Name = "出流管制計畫書調查總表";
            Font headerFont = new Font("微軟正黑體", 14, FontStyle.Bold),
                colFont = new Font("微軟正黑體", 12, FontStyle.Bold);
            //表格標題
            tb[0, 0].Formula = "出流管制計畫書審查統計表（統計時間自98年起）";
            tb["A1:E1"].Merge = true;
            tb["A1:E1"].HorizonAlign = Alignment.HorizonAlignment.CENTER;
            tb[0, 0].SetFont(headerFont);

            tb[1, 4].Formula = $"{DateTime.Now.Year - 1911:000}.{DateTime.Now:MM.dd}";
            tb[1, 4].HorizonAlign = Alignment.HorizonAlignment.RIGHT;

            //總計
            var line = new Line() { Color = Color.Black, OuterWidth = 20 };
            tb[2, 0].Formula = "中央+地方";
            tb[2, 1].Formula = "受理件數";
            tb[2, 2].Formula = "通過件數";
            tb[2, 3].Formula = "核定滯洪池座數";
            tb[2, 4].Formula = "核定滯洪池體積(萬m3)";
            tb[3, 0].Formula = "總計";
            tb[3, 1].Formula = "=B30+B41";
            tb[3, 2].Formula = "=C30+C41";
            tb[3, 3].Formula = "=D30+D41";
            tb[3, 4].Formula = "=E30+E41";
            tb["A3:E4"].SetBorder(line, line, line, line, line, line);
            tb[0, 2, 4, 2].SetFont(colFont);

            //縣市
            var local = Summary_OutflowCtrlPlans_Local_All_Outflow().ToArray();
            var counties = new string[] { "宜蘭縣", "臺北市", "基隆市", "新北市", "桃園市", "新竹縣", "新竹市", "苗栗縣", "臺中市", "南投縣", "彰化縣", "雲林縣", "嘉義縣", "嘉義市", "臺南市", "高雄市", "屏東縣", "澎湖縣", "臺東縣", "金門縣", "花蓮縣" };
            tb[5, 0].Formula = "地方列管出流管制計畫書審查統計表（統計時間自98年起）";
            tb[5, 0].SetFont(headerFont);
            tb["A6:E6"].Merge = true;
            tb["A6:E6"].HorizonAlign = Alignment.HorizonAlignment.CENTER;
            tb[7, 0].Formula = "縣（市）";
            tb[7, 1].Formula = "受理件數";
            tb[7, 2].Formula = "通過件數";
            tb[7, 3].Formula = "核定滯洪池座數";
            tb[7, 4].Formula = "核定滯洪池體積(萬m3)";
            tb[0, 7, 4, 7].SetFont(colFont);

            for (int i = 8; i < 8 + counties.Length; i++)
            {
                var l = local.FirstOrDefault(ll => ll["OFP_Gov"] == counties[i - 8]);
                tb[i, 0].Formula = counties[i - 8];
                tb[i, 1].Formula = $"{l?["Count"] ?? 0}";
                tb[i, 2].Formula = $"{l?["App_Count"] ?? 0}";
                tb[i, 3].Formula = $"{l?["App_FacilityNum"] ?? 0}";
                tb[i, 4].Formula = $"{l?["App_FacilityArea"] ?? 0}";
            }
            var start = counties.Length + 8;
            tb[start, 0].Formula = "合計";
            tb[start, 1].Formula = $"=sum(B9:B{start})";
            tb[start, 2].Formula = $"=sum(C9:C{start})";
            tb[start, 3].Formula = $"=sum(D9:D{start})";
            tb[start, 4].Formula = $"=sum(E9:E{start})";
            tb[$"A8:E{start + 1}"].SetBorder(line, line, line, line, line, line);
            tb[0, 7, 4, start].SetBorder(line, line, line, line, line, line);

            start = start + 2;
            tb[start, 0].Formula = "中央列管出流管制計畫書審查統計表（統計時間自98年起）";
            tb[start, 0].SetFont(headerFont);
            tb[$"A{start + 1}:E{start + 1}"].Merge = true;
            tb[$"A{start + 1}:E{start + 1}"].HorizonAlign = Alignment.HorizonAlignment.CENTER;
            var unit = new string[] { "二", "三", "五", "六", "七", "十" };
            var central = Summary_OutflowCtrlPlans_Centroid_All_Outflow();
            start = start + 2;
            tb[start, 0].Formula = "河川局";
            tb[start, 1].Formula = "受理件數";
            tb[start, 2].Formula = "通過件數";
            tb[start, 3].Formula = "核定滯洪池座數";
            tb[start, 4].Formula = "核定滯洪池體積(萬m3)";
            tb[0, start, 4, start].SetFont(colFont);
            start++;
            for (int i = 0; i < unit.Length; i++)
            {
                var c = central.FirstOrDefault(cc => cc["OFP_Gov"].IndexOf(unit[i]) > 0);
                tb[start + i, 0].Formula = unit[i];
                tb[start + i, 1].Formula = $"{c?["Count"] ?? 0}";
                tb[start + i, 2].Formula = $"{c?["App_Count"] ?? 0}";
                tb[start + i, 3].Formula = $"{c?["App_FacilityNum"] ?? 0}";
                tb[start + i, 4].Formula = $"{c?["App_FacilityArea"] ?? 0}";
            }
            start = start + unit.Length;
            tb[start, 0].Formula = "合計";
            tb[start, 1].Formula = $"=Sum(B{start - unit.Length + 1}:B{start})";
            tb[start, 2].Formula = $"=Sum(C{start - unit.Length + 1}:C{start})";
            tb[start, 3].Formula = $"=Sum(D{start - unit.Length + 1}:D{start})";
            tb[start, 4].Formula = $"=Sum(E{start - unit.Length + 1}:E{start})";

            tb[$"A{start - unit.Length}:E{start + 1}"].SetBorder(line, line, line, line, line, line);

            for (int i = 1; i < tb.ColumnCount; i++)
                tb.Columns[i].AutoWidth = true;
            tb.Columns[0].HorizonAlign = Alignment.HorizonAlignment.CENTER;
            tb.Columns[4].NumberFormat = "#,##0.##";
        }
        /// <summary>
        /// 中央管已受理計畫書案件總表
        /// </summary>
        /// <returns>[單位, 筆數]</returns>
        public IEnumerable<dynamic> Summary_OutflowCtrlPlans_Centroid()
        {
            var sql = @"select ofp.OFP_Gov, Count(*) Count from OutflowControlPlan ofp
                join OFPZoneTimeline otl on ofp.OFP_ID = otl.OFP_ID2 where ofp.IsShow = 1 and  ofp.OFP_ReviewClass = '中央審查' and ofp.OFP_Type like '%計畫書%' and otl.CA_Acc_Status2 > 0 group by ofp.OFP_Gov";
            return UtilDB.GetDataList<dynamic>(sql, null);
        }
        /// <summary>
        /// 中央管排水已受理計畫書案件總表
        /// </summary>
        /// <returns></returns>
        public IEnumerable<dynamic> Summary_OutflowCtrlPlans_Centroid_Drain()
        {
            var sql = @"select ofp.OFP_Gov, Count(*) Count from OutflowControlPlan ofp
                join OFPZoneTimeline otl on ofp.OFP_ID = otl.OFP_ID2 where ofp.IsShow = 1 and  ofp.OFP_ReviewClass = '中央審查' and ofp.OFP_Type = '排水計畫書' and otl.CA_Acc_Status2 > 0 group by ofp.OFP_Gov";
            return UtilDB.GetDataList<dynamic>(sql, null);
        }
        /// <summary>
        /// 中央管出流管制已受理計畫書案件總表
        /// </summary>
        /// <returns></returns>
        public IEnumerable<dynamic> Summary_OutflowCtrlPlans_Centroid_Outflow()
        {
            var sql = @"select ofp.OFP_Gov, Count(*) Count from OutflowControlPlan ofp
                join OFPZoneTimeline otl on ofp.OFP_ID = otl.OFP_ID2 where ofp.IsShow = 1 and ofp.OFP_ReviewClass = '中央審查' and ofp.OFP_Type = '出流計畫書' and otl.CA_Acc_Status2 > 0 group by ofp.OFP_Gov";
            return UtilDB.GetDataList<dynamic>(sql, null);
        }
        /// <summary>
        /// 中央管計畫總案件總表
        /// </summary>
        /// <returns>合併受理案件及已核定案件</returns>
        public IEnumerable<Dictionary<string, dynamic>> Summary_OutflowCtrlPlans_Centroid_All()
        {
            List<Dictionary<string, dynamic>> collection = new List<Dictionary<string, dynamic>>();
            var centroid = Summary_OutflowCtrlPlans_Centroid();
            var approved = Summary_OutflowCtrlPlans_Centroid_Approved();
            foreach (var citem in centroid)
            {
                Dictionary<string, dynamic> dictionary = new Dictionary<string, dynamic>();
                dictionary.Add("OFP_Gov", citem.OFP_Gov);
                dictionary.Add("Count", citem.Count);
                var first = approved.FirstOrDefault(a => a.OFP_Gov == citem.OFP_Gov);
                dictionary.Add("App_Count", first?.Count ?? 0);
                dictionary.Add("App_FacilityNum", first?.FacilityNum ?? 0);
                dictionary.Add("App_FacilityArea", first?.FacilityArea ?? 0);
                collection.Add(dictionary);
            }
            return collection;
        }
        /// <summary>
        /// 中央管排水計畫總案件總表
        /// </summary>
        /// <returns>合併受理案件及已核定案件</returns>
        public IEnumerable<Dictionary<string, dynamic>> Summary_OutflowCtrlPlans_Centroid_All_Drain()
        {
            List<Dictionary<string, dynamic>> collection = new List<Dictionary<string, dynamic>>();
            var centroid = Summary_OutflowCtrlPlans_Centroid_Drain();
            var approved = Summary_OutflowCtrlPlans_Centroid_Approved_Drain();
            foreach (var citem in centroid)
            {
                Dictionary<string, dynamic> dictionary = new Dictionary<string, dynamic>();
                dictionary.Add("OFP_Gov", citem.OFP_Gov);
                dictionary.Add("Count", citem.Count);
                var first = approved.FirstOrDefault(a => a.OFP_Gov == citem.OFP_Gov);
                dictionary.Add("App_Count", first?.Count ?? 0);
                dictionary.Add("App_FacilityNum", first?.FacilityNum ?? 0);
                dictionary.Add("App_FacilityArea", first?.FacilityArea ?? 0);
                collection.Add(dictionary);
            }
            return collection;
        }
        /// <summary>
        /// 中央管出流計畫總案件總表
        /// </summary>
        /// <returns>合併受理案件及已核定案件</returns>
        public IEnumerable<Dictionary<string, dynamic>> Summary_OutflowCtrlPlans_Centroid_All_Outflow()
        {
            List<Dictionary<string, dynamic>> collection = new List<Dictionary<string, dynamic>>();
            var centroid = Summary_OutflowCtrlPlans_Centroid_Outflow();
            var approved = Summary_OutflowCtrlPlans_Centroid_Approved_Outflow();
            foreach (var citem in centroid)
            {
                Dictionary<string, dynamic> dictionary = new Dictionary<string, dynamic>();
                dictionary.Add("OFP_Gov", citem.OFP_Gov);
                dictionary.Add("Count", citem.Count);
                var first = approved.FirstOrDefault(a => a.OFP_Gov == citem.OFP_Gov);
                dictionary.Add("App_Count", first?.Count ?? 0);
                dictionary.Add("App_FacilityNum", first?.FacilityNum ?? 0);
                dictionary.Add("App_FacilityArea", first?.FacilityArea ?? 0);
                collection.Add(dictionary);
            }
            return collection;
        }
        /// <summary>
        /// 中央管已核定計畫書案件總表
        /// </summary>
        /// <returns>[單位,筆數,滯洪池座數,滯洪池體積]</returns>
        public IEnumerable<dynamic> Summary_OutflowCtrlPlans_Centroid_Approved()
        {
            var sql = @"select ofp.OFP_Gov, Count(*) Count, sum(Isnull(ofp.OFP_FacilityNum, 0)) FacilityNum, sum(isnull(ofp.OFP_FacilityArea, 0)) FacilityArea from OutflowControlPlan ofp
                join OFPZoneTimeline otl on ofp.OFP_ID = otl.OFP_ID2 where ofp.IsShow = 1 and  ofp.OFP_ReviewClass='中央審查' and otl.CA_Approved_Status2 = 1 group by ofp.OFP_Gov";
            return UtilDB.GetDataList<dynamic>(sql, null);
        }
        /// <summary>
        /// 中央管排水已核定計畫書案件總表
        /// </summary>
        /// <returns></returns>
        public IEnumerable<dynamic> Summary_OutflowCtrlPlans_Centroid_Approved_Drain()
        {
            var sql = @"select ofp.OFP_Gov, Count(*) Count, sum(Isnull(ofp.OFP_FacilityNum, 0)) FacilityNum, sum(isnull(ofp.OFP_FacilityArea, 0)) FacilityArea from OutflowControlPlan ofp
                join OFPZoneTimeline otl on ofp.OFP_ID = otl.OFP_ID2 where ofp.OFP_Type = '排水計畫書' and ofp.IsShow = 1 and ofp.OFP_ReviewClass='中央審查' and Approved_Status2=1 group by ofp.OFP_Gov";
            return UtilDB.GetDataList<dynamic>(sql, null);
        }
        /// <summary>
        /// 中央管出流管制已核定計畫書案件總表
        /// </summary>
        /// <returns></returns>
        public IEnumerable<dynamic> Summary_OutflowCtrlPlans_Centroid_Approved_Outflow()
        {
            var sql = @"select ofp.OFP_Gov, Count(*) Count, sum(Isnull(ofp.OFP_FacilityNum, 0)) FacilityNum, sum(isnull(ofp.OFP_FacilityArea, 0)) FacilityArea from OutflowControlPlan ofp
                join OFPZoneTimeline otl on ofp.OFP_ID = otl.OFP_ID2 where ofp.OFP_Type = '出流計畫書' and ofp.IsShow = 1 and ofp.OFP_ReviewClass='中央審查' and Approved_Status2 = 1 group by ofp.OFP_Gov";
            return UtilDB.GetDataList<dynamic>(sql, null);
        }
        /// <summary>
        /// 地方管已受理計畫書案件總表
        /// </summary>
        /// <returns>[{OFP_Gov, Count}]</returns>
        public IEnumerable<dynamic> Summary_OutflowCtrlPlans_Local()
        {
            var sql = @"select ofp.OFP_Gov, Count(*) Count from OutflowControlPlan ofp
                join OFPZoneTimeline otl on ofp.OFP_ID = otl.OFP_ID2 where ofp.IsShow = 1 and ofp.OFP_ReviewClass = '地方審查' and ofp.OFP_Type like '%計畫書%' and CA_Acc_Status2 > 0 group by ofp.OFP_Gov";
            return UtilDB.GetDataList<dynamic>(sql, null);
        }
        /// <summary>
        /// 地方管排水已受理計畫書案件總表
        /// </summary>
        /// <returns></returns>
        public IEnumerable<dynamic> Summary_OutflowCtrlPlans_Local_Drain()
        {
            var sql = @"select ofp.OFP_Gov, Count(*) Count from OutflowControlPlan ofp
                join OFPZoneTimeline otl on ofp.OFP_ID = otl.OFP_ID2 where ofp.IsShow = 1 and ofp.OFP_ReviewClass = '地方審查' and ofp.OFP_Type = '排水計畫書' and CA_Acc_Status2 > 0 group by ofp.OFP_Gov";
            return UtilDB.GetDataList<dynamic>(sql, null);
        }
        /// <summary>
        /// 地方管出流管制已受理計畫書案件總表
        /// </summary>
        /// <returns></returns>
        public IEnumerable<dynamic> Summary_OutflowCtrlPlans_Local_Outflow()
        {
            var sql = @"select ofp.OFP_Gov, Count(*) Count from OutflowControlPlan ofp
                join OFPZoneTimeline otl on ofp.OFP_ID = otl.OFP_ID2 where ofp.IsShow = 1 and ofp.OFP_ReviewClass = '地方審查' and ofp.OFP_Type = '出流計畫書' and CA_Acc_Status2 > 0 group by ofp.OFP_Gov";
            return UtilDB.GetDataList<dynamic>(sql, null);
        }
        /// <summary>
        /// 地方管計畫總案件總表
        /// </summary>
        /// <returns>[{App_Count, App_FacilityNum, App_FacilityArea}]</returns>
        public IEnumerable<Dictionary<string, dynamic>> Summary_OutflowCtrlPlans_Local_All()
        {
            List<Dictionary<string, dynamic>> collection = new List<Dictionary<string, dynamic>>();
            var local = Summary_OutflowCtrlPlans_Local();
            var approved = Summary_OutflowCtrlPlans_Local_Approved();
            foreach (var citem in local)
            {
                Dictionary<string, dynamic> dictionary = new Dictionary<string, dynamic>();
                dictionary.Add("OFP_Gov", citem.OFP_Gov);
                dictionary.Add("Count", citem.Count);
                var first = approved.FirstOrDefault(a => a.OFP_Gov == citem.OFP_Gov);
                dictionary.Add("App_Count", first?.Count ?? 0);
                dictionary.Add("App_FacilityNum", first?.FacilityNum ?? 0);
                dictionary.Add("App_FacilityArea", first?.FacilityArea ?? 0);
                collection.Add(dictionary);
            }
            return collection;
        }
        /// <summary>
        /// 地方管排水計畫總案件總表
        /// </summary>
        /// <returns>[{App_Count, App_FacilityNum, App_FacilityArea}]</returns>
        public IEnumerable<Dictionary<string, dynamic>> Summary_OutflowCtrlPlans_Local_All_Drain()
        {
            List<Dictionary<string, dynamic>> collection = new List<Dictionary<string, dynamic>>();
            var local = Summary_OutflowCtrlPlans_Local_Drain();
            var approved = Summary_OutflowCtrlPlans_Local_Approved_Drain();
            foreach (var citem in local)
            {
                Dictionary<string, dynamic> dictionary = new Dictionary<string, dynamic>();
                dictionary.Add("OFP_Gov", citem.OFP_Gov);
                dictionary.Add("Count", citem.Count);
                var first = approved.FirstOrDefault(a => a.OFP_Gov == citem.OFP_Gov);
                dictionary.Add("App_Count", first?.Count ?? 0);
                dictionary.Add("App_FacilityNum", first?.FacilityNum ?? 0);
                dictionary.Add("App_FacilityArea", first?.FacilityArea ?? 0);
                collection.Add(dictionary);
            }
            return collection;
        }
        /// <summary>
        /// 地方管出流計畫總案件總表
        /// </summary>
        /// <returns>[{App_Count, App_FacilityNum, App_FacilityArea}]</returns>
        public IEnumerable<Dictionary<string, dynamic>> Summary_OutflowCtrlPlans_Local_All_Outflow()
        {
            List<Dictionary<string, dynamic>> collection = new List<Dictionary<string, dynamic>>();
            var local = Summary_OutflowCtrlPlans_Local_Outflow();
            var approved = Summary_OutflowCtrlPlans_Local_Approved_Outflow();
            foreach (var citem in local)
            {
                Dictionary<string, dynamic> dictionary = new Dictionary<string, dynamic>();
                dictionary.Add("OFP_Gov", citem.OFP_Gov);
                dictionary.Add("Count", citem.Count);
                var first = approved.FirstOrDefault(a => a.OFP_Gov == citem.OFP_Gov);
                dictionary.Add("App_Count", first?.Count ?? 0);
                dictionary.Add("App_FacilityNum", first?.FacilityNum ?? 0);
                dictionary.Add("App_FacilityArea", first?.FacilityArea ?? 0);
                collection.Add(dictionary);
            }
            return collection;
        }
        /// <summary>
        /// 地方管已核定計畫書案件總表
        /// </summary>
        /// <returns>[{OFP_GOv, Count, FacilityNum, FacilityArea}]</returns>
        public IEnumerable<dynamic> Summary_OutflowCtrlPlans_Local_Approved()
        {
            var sql = @"select ofp.OFP_Gov, Count(*) Count, sum(Isnull(ofp.OFP_FacilityNum, 0)) FacilityNum, sum(isnull(ofp.OFP_FacilityArea, 0)) FacilityArea from OutflowControlPlan ofp
                join OFPZoneTimeline otl on ofp.OFP_ID = otl.OFP_ID2 where ofp.IsShow = 1 and ofp.OFP_ReviewClass='地方審查' and otl.CA_Approved_Status2 = 1 group by ofp.OFP_Gov";
            return UtilDB.GetDataList<dynamic>(sql, null);
        }
        /// <summary>
        /// 地方管排水已核定計畫書案件總表
        /// </summary>
        /// <returns>[{OFP_GOv, Count, FacilityNum, FacilityArea}]</returns>
        public IEnumerable<dynamic> Summary_OutflowCtrlPlans_Local_Approved_Drain()
        {
            var sql = @"select ofp.OFP_Gov, Count(*) Count, sum(Isnull(ofp.OFP_FacilityNum, 0)) FacilityNum, sum(isnull(ofp.OFP_FacilityArea, 0)) FacilityArea from OutflowControlPlan ofp
                join OFPZoneTimeline otl on ofp.OFP_ID = otl.OFP_ID2 where ofp.OFP_Type = '排水計畫書' and ofp.IsShow = 1 and ofp.OFP_ReviewClass='地方審查' and otl.CA_Approved_Status2 = 1 group by ofp.OFP_Gov";
            return UtilDB.GetDataList<dynamic>(sql, null);
        }
        /// <summary>
        /// 地方管出流管制已核定計畫書案件總表
        /// </summary>
        /// <returns></returns>
        public IEnumerable<dynamic> Summary_OutflowCtrlPlans_Local_Approved_Outflow()
        {
            var sql = @"select ofp.OFP_Gov, Count(*) Count, sum(Isnull(ofp.OFP_FacilityNum, 0)) FacilityNum, sum(isnull(ofp.OFP_FacilityArea, 0)) FacilityArea from OutflowControlPlan ofp
                join OFPZoneTimeline otl on ofp.OFP_ID = otl.OFP_ID2 where ofp.OFP_Type = '出流計畫書' and ofp.IsShow = 1 and ofp.OFP_ReviewClass='地方審查' and otl.CA_Approved_Status2 = 1 group by ofp.OFP_Gov";
            return UtilDB.GetDataList<dynamic>(sql, null);
        }
        #endregion

        #region 取得Odt檔案(附件10,11,13,14,15,16,17,18)
        /// <summary>
        /// (附件10-EC_ID,附件11-ECS_ID,附件13-SC_ID,附件14-AC_ID,附件15-FC_ID,附件16-CP_ID,附件17-MM_RE_ID)
        /// </summary>
        /// <param name="OfpId">ID</param>
        /// <param name="fileName"></param>
        /// <returns></returns>
        [HttpGet, ActionName(nameof(OutflowctrlAppendix))]
        [Route("api/Report/OutflowctrlAppendix/{OfpId}/{fileName}")]
        public HttpResponseMessage OutflowctrlAppendix(OfpId OfpId, string fileName)
        {
            string filePath;
            string Sql;
            List<dynamic> data;
            List<dynamic> duplicateData = null;
            List<dynamic> duplicateBox = null;
            Dictionary<string, dynamic> databaseData = null;
            try
            {
                switch (OfpId)
                    {
                    case OfpId.EC_ID1:
                    case OfpId.EC_ID2:
                        databaseData = new Dictionary<string, dynamic>()
                            {
                              {"duplicateBox", duplicateBox},
                              {"duplicateData", duplicateData},
                              {"data", data = UtilDB.GetDataList<dynamic>(Sql = Get_CheckList(), new { OfpId })
                            } };
                        break;
                    case OfpId.ECS_ID1:
                    case OfpId.ECS_ID2:
                        databaseData = new Dictionary<string, dynamic>()
                            {
                              {"duplicateBox", duplicateBox},
                              {"duplicateData", duplicateData = UtilDB.GetDataList<dynamic>(Sql = Get_SVCheck_Item(), new { OfpId } )},
                              {"data", data = UtilDB.GetDataList<dynamic>(Sql = Get_SVCheck(), new { OfpId })
                            } };
                        break;
                    case OfpId.SC_ID:  
                       databaseData = new Dictionary<string, dynamic>()
                            {
                              {"duplicateBox", duplicateBox},
                              {"duplicateData", duplicateData = UtilDB.GetDataList<dynamic>(Sql = Get_ENENDSelfCheck_Item(), new { OfpId } )},
                              {"data", data = UtilDB.GetDataList<dynamic>(Sql = Get_ENENDSelfCheck(), new { OfpId })
                            } };
                        break;
                    case OfpId.AC_ID1:
                    case OfpId.AC_ID2:
                        databaseData = new Dictionary<string, dynamic>()
                            {
                              {"duplicateBox", duplicateBox},
                              {"duplicateData", duplicateData},
                              {"data", data = UtilDB.GetDataList<dynamic>(Sql =  Get_ENENDApplicationCompleted(), new { OfpId })
                            } };
                        break;
                    case OfpId.FC_ID:
                        databaseData = new Dictionary<string, dynamic>()
                            {
                              {"duplicateBox", data = UtilDB.GetDataList<dynamic>(Sql =  Get_ENENDFinalCheck_Item_Box(), new { OfpId } )},
                              {"duplicateData", duplicateData = UtilDB.GetDataList<dynamic>(Sql =  Get_ENENDFinalCheck_Item(), new { OfpId } )},
                              {"data", data = UtilDB.GetDataList<dynamic>(Sql =  Get_ENENDFinalCheck(), new { OfpId })
                            } };
                        break;
                    case OfpId.CP_ID:
                        databaseData = new Dictionary<string, dynamic>()
                            {
                              {"duplicateBox", duplicateBox },
                              {"duplicateData", duplicateData = UtilDB.GetDataList<dynamic>(Sql =  Get_ENENDCompletion_Item(), new { OfpId } )},
                              {"data", data = UtilDB.GetDataList<dynamic>(Sql =  Get_ENENDCompletion(), new { OfpId })
                            } };
                        break;
                    case OfpId.MM_RE_ID:
                        databaseData = new Dictionary<string, dynamic>()
                            {
                              {"duplicateBox", duplicateBox = UtilDB.GetDataList<dynamic>(Sql =  Get_MM_Record_Box(), new { OfpId } )},
                              {"duplicateData", duplicateData = UtilDB.GetDataList<dynamic>(Sql =  Get_MM_Record_Item(), new { OfpId } )},
                              {"data", data = UtilDB.GetDataList<dynamic>(Sql =  Get_MM_Record(), new { OfpId })
                            } };
                        break;
                    default:
                        break;
                    }
                
                filePath = DocumentHelper.GetRptDatabase($"{fileName}.odt", databaseData, $"{databaseData["data"][0].OFP_No}.odt");
                return HttpHelper.FileResult(filePath, "application/vnd.oasis.opendocument.spreadsheet"); 
            }
            catch (Exception ex)
            {
                return HttpHelper.FailResult(ex.ToString());
            }
        }
        /// <summary>
        /// 附件10
        /// </summary>
        /// <returns></returns>
        public string Get_CheckList()
        {
            var sql = @"SELECT ofp.OFP_Name, ofp.OFP_Gov, ofp.OFP_No, cl.EN_Checklist, cl.CheckSuggest, cl.Improve,
                      cl.ProphaseImprove, cl.Check_StartDate, cl.Check_EndDate, cl.EN_PreSchedule, cl.EN_Schedule,
                      payer.Payer, el.EngineerName
                      FROM [OutflowControlPlan] ofp
                      INNER JOIN [EN_Checklist] cl on ofp.OFP_ID = cl.OFP_ID 
                      INNER JOIN [payer] on ofp.PA_ID = payer.PA_ID 
                      INNER JOIN [EngineerList] el on ofp.SupervisorEngineer = el.ED_ID
                      WHERE cl.EC_ID = @OfpId";
            return sql;
        }
        /// <summary>
        /// 附件11
        /// </summary>
        /// <returns></returns>
        public string Get_SVCheck()
        {
            var sql = @"SELECT ofp.OFP_Name, ofp.OFP_No, ofp.OFP_Location, payer.Payer, payer.PA_Num, payer.PA_address, el.EngineerName, pl.PracticeUnits, pl.PracticeLicense,
                      pl.GUI, pl.Tel, start.EN_ST_Date, Approved.Approved_No, Approved.Approved_Date
                      FROM [EN_SVCheck] svc
                      INNER JOIN [OutflowControlPlan] ofp on svc.OFP_ID = ofp.OFP_ID
                      INNER JOIN [payer] on ofp.PA_ID = payer.PA_ID
                      INNER JOIN [EngineerList] el on ofp.SupervisorEngineer = el.ED_ID
                      INNER JOIN [PracticeList] pl on el.PU_ID = pl.PU_ID
                      INNER JOIN [EN_Starting] start on svc.ES_ID = start.ES_ID
                      INNER JOIN [Approved] on svc.OFP_ID = Approved.OFP_ID
                      WHERE svc.ECS_ID = @OfpId";
            return sql;
        }
        
        public string Get_SVCheck_Item()
        {
            var sql = @"SELECT svcI.Implementation_Situation, svcI.Note
                      FROM [EN_SVCheck_Item] svcI
                      WHERE svcI.ECS_ID = @OfpId";
            return sql;
        }
        /// <summary>
        /// 附件13
        /// </summary>
        /// <returns></returns>
        public string Get_ENENDSelfCheck()
        {
            var sql = @"SELECT ofp.OFP_Name, ofp.OFP_No, ofp.OFP_Location, payer.Payer, payer.PA_Num, payer.PA_address,
                      el.EngineerName, pl.PracticeUnits, pl.PracticeLicense, pl.GUI, pl.Tel, start.EN_ST_Date,
                      Approved.Approved_No, Approved.Approved_Date, eac.EN_END_Date 
                      FROM [ENEND_Self_Check] esc 
                      INNER JOIN [OutflowControlPlan] ofp on esc.OFP_ID = ofp.OFP_ID
                      INNER JOIN [payer] on ofp.PA_ID = payer.PA_ID
                      INNER JOIN [EngineerList] el on ofp.SupervisorEngineer = el.ED_ID
                      INNER JOIN [PracticeList] pl on el.PU_ID = pl.PU_ID
                      INNER JOIN [EN_Starting] start on esc.ES_ID = start.ES_ID
                      INNER JOIN [Approved] on esc.OFP_ID = Approved.OFP_ID
                      INNER JOIN [ENEND_Application_Completed] eac on esc.ES_ID = eac.ES_ID 
                      WHERE esc.SC_ID = @OfpId";
            return sql;
        }

        public string Get_ENENDSelfCheck_Item()
        {
            var sql = @"SELECT escI.IsMatchPlan, escI.IsConductChangePlan, escI.DefDescription, escI.Note
                      FROM [ENEND_Self_Check_Item] escI
                      WHERE escI.SC_ID = @OfpId";
            return sql;
        }

        /// <summary>
        /// 附件14
        /// </summary>
        /// <returns></returns>
        public string Get_ENENDApplicationCompleted()
        {
            var sql = @"SELECT eac.EN_END_Date, eac.EN_END_AppDate, ofp.OFP_Name, ofp.OFP_No, ofp.OFP_Location, payer.Payer, payer.PA_Num, payer.PA_address,
                      el.EngineerName, pl.PracticeUnits, pl.Address, pl.PracticeLicense, pl.GUI, pl.Tel, start.EN_ST_Date, Approved.Approved_No
                      FROM [ENEND_Application_Completed] eac 
                      INNER JOIN [OutflowControlPlan] ofp on eac.OFP_ID = ofp.OFP_ID
                      INNER JOIN [payer] on ofp.PA_ID = payer.PA_ID
                      INNER JOIN [EngineerList] el on ofp.SupervisorEngineer = el.ED_ID
                      INNER JOIN [PracticeList] pl on el.PU_ID = pl.PU_ID
                      INNER JOIN [EN_Starting] start on eac.ES_ID = start.ES_ID
                      INNER JOIN [Approved] on eac.OFP_ID = Approved.OFP_ID
                      WHERE eac.AC_ID = @OfpId";
            return sql;
        }
        /// <summary>
        /// 附件15
        /// </summary>
        /// <returns></returns>
        public string Get_ENENDFinalCheck()
        {
            var sql = @"SELECT ofp.OFP_Name, ofp.OFP_No, ofp.OFP_Location, payer.Payer, payer.PA_Num,  payer.PA_address, 
                      el.EngineerName, pl.PracticeUnits, pl.PracticeLicense, pl.GUI, pl.Tel, start.EN_ST_Date,
                      Approved_No, Approved.Approved_Date, eac.EN_END_Date, efc.FC_Date
                      FROM [ENEND_Final_Check] efc 
                      INNER JOIN [OutflowControlPlan] ofp on efc.OFP_ID = ofp.OFP_ID
                      INNER JOIN [payer] on ofp.PA_ID = payer.PA_ID
                      INNER JOIN [EngineerList] el on ofp.SupervisorEngineer = el.ED_ID
                      INNER JOIN [PracticeList] pl on el.PU_ID = pl.PU_ID
                      INNER JOIN [EN_Starting] start on efc.ES_ID = start.ES_ID
                      INNER JOIN [Approved] on efc.OFP_ID = Approved.OFP_ID
                      INNER JOIN [ENEND_Application_Completed] eac on efc.ES_ID = eac.ES_ID 
                      WHERE efc.FC_ID = @OfpId";
            return sql;
        }

        public string Get_ENENDFinalCheck_Item()
        {
            var sql = @"SELECT FacilityName, Location, ApprovedNum, ActualNum, DefPercentNum, DefPercentNum, ApprovedSize, ActualSize, DefPercentSize
                      FROM [ENEND_Final_Check_Item] efcI
                      WHERE efcI.FC_ID = @OfpId";
            return sql;
        }

        public string Get_ENENDFinalCheck_Item_Box()
        {
            var sql = @"SELECT efcI.IsQualified
                      FROM [ENEND_Final_Check_Item] efcI
                      WHERE efcI.FC_ID = @OfpId";
            return sql;
        }
        /// <summary>
        /// 附件16
        /// </summary>
        /// <returns></returns>
        /// 
        public string Get_ENENDCompletion()
        {
            var sql = @"SELECT ofp.OFP_Name, ofp.OFP_No, ofp.OFP_Location, payer.Payer, payer.PA_Num, 
                      payer.PA_address, start.EN_ST_Date, ec.CP_Date, ec.CP_NO, Approved.Approved_No
                      FROM [ENEND_Completion] ec 
                      INNER JOIN [OutflowControlPlan] ofp on ec.OFP_ID = ofp.OFP_ID
                      INNER JOIN [payer] on ofp.PA_ID = payer.PA_ID
                      INNER JOIN [Approved] on ec.OFP_ID = Approved.OFP_ID
                      INNER JOIN [EN_Starting] start on ec.ES_ID = start.ES_ID
                      WHERE ec.CP_ID = @OfpId";
            return sql;
        }

        public string Get_ENENDCompletion_Item()
        {
            var sql = @"SELECT el.EngineerName, el.EngineerLicense,  
                      pl.PracticeUnits, pl.Address, pl.GUI, pl.Tel
                      FROM [ENEND_Completion] ec 
                      INNER JOIN [OutflowControlPlan] ofp on ec.OFP_ID = ofp.OFP_ID
                      INNER JOIN [EngineerList] el on ofp.SupervisorEngineer = el.ED_ID or ofp.Engineer = el.ED_ID
                      INNER JOIN [PracticeList] pl on el.PU_ID = pl.PU_ID
                      WHERE ec.CP_ID = @OfpId";
            return sql;
        }
        /// <summary>
        /// 附件17
        /// </summary>
        /// <returns></returns>
        /// 

        public string Get_MM_Record()
        {
            var sql = "";
            return sql;
        }

        public string Get_MM_Record_Item()
        {
            var sql = "";
            return sql;
        }

        public string Get_MM_Record_Box()
        {
            var sql = @"SELECT mrI.IsQualified, mr.MM_Check_Result, mr.MM_Check_Note_Bef
                      FROM [MM_Record] mr
                      INNER JOIN [MM_Record_item] mrI on mr.MM_RE_ID = mrI.MM_RE_ID
                      WHERE mrI.MM_RE_ID = @OfpId";
            //var sql = @"SELECT mrI.IsQualified, mr.MM_Check_Result, mr.MM_Check_Note_Bef
            //          FROM [MM_Record_Item] mrI
            //          INNER JOIN [MM_Record] mr on mrI.MM_RE_ID = mr.MM_RE_ID
            //          WHERE mrI.MM_RE_ID = @OfpId";
            return sql;
        }
        #endregion
    }
}
