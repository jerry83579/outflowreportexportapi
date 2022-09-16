using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace OutFlowReportExportAPI.Helpers
{
    public class FilterParams
    {
        /// <summary>
        /// 計畫類別
        /// </summary>
        public string groupby { get; set; } = "OFP_Cou";
        /// <summary>
        /// 計畫類別
        /// </summary>
        public string OFP_Type { get; set; }
        /// <summary>
        /// 年度
        /// </summary>
        public string OFP_Year { get; set; }
        /// <summary>
        /// 縣市
        /// </summary>
        public string OFP_Cou { get; set; }
        /// <summary>
        /// 主管機關
        /// </summary>
        public string OFP_Gov { get; set; }
        /// <summary>
        /// 開發類別
        /// </summary>
        public string Develop_Type { get; set; }
        /// <summary>
        /// 計畫狀態
        /// </summary>
        public string OFP_Status { get; set; }
        /// <summary>
        /// 關鍵字
        /// </summary>
        public string OFP_KeyWord { get; set; }
        /// <summary>
        /// 開發樣態
        /// </summary>
        public string LanduseID { get; set; }
        /// <summary>
        /// 是否為光電案
        /// </summary>
        public string IsPhotovoltaic { get; set; }
    }
}