using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace OutFlowReportExportAPI.Helpers
{
    public class DateHelper
    {
        #region 將日期轉換成簡短日期
        /// <summary>
        /// 將日期字串轉換成簡短日期的字串
        /// </summary>
        /// <param name="dateTimeString"></param>
        /// <returns>簡短日期的字串</returns>
        public static string ToShortDate(string dateTimeString)
        {
            try
            {
                if (!string.IsNullOrEmpty(dateTimeString))
                {
                    dateTimeString = Convert.ToDateTime(dateTimeString).ToString("yyyy/MM/dd");
                    return dateTimeString;
                }
            }
            catch (Exception)
            {
                return "";
            }

            return "";
        }
        #endregion
    }
}