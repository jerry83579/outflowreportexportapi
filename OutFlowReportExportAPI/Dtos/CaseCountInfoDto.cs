using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace OutFlowReportExportAPI.Dtos
{
    /// <summary>
 /// 統計總表資訊
 /// </summary>
    public class CaseCountInfoDto
    {
        /// <summary>
        /// 統計標的
        /// </summary>
        public string groupby { get; set; }
        /// <summary>
        /// 案件總統計
        /// </summary>
        public int caseTotal { get; set; }

        /// <summary>
        /// 核定案件加總(已核定管制案)
        /// </summary>
        public int checkTotal { get; set; }

        /// <summary>
        /// 未核定案件加總(實質審查階段)
        /// </summary>
        public int nocheckTotal { get; set; }

        /// <summary>
        /// 未送審案件加總(義務人填報階段)
        /// </summary>
        public int nosendTotal { get; set; }
        /// <summary>
        /// 滯洪池數量
        /// </summary>
        public int facilityNumTotal { get; set; }
        /// <summary>
        /// 滯洪池面積
        /// </summary>
        public double facilityAreaTotal { get; set; }

        /// <summary>
        /// 統計資訊
        /// </summary>
        public IEnumerable<Dictionary<string, dynamic>> details { get; set; }
    }
}