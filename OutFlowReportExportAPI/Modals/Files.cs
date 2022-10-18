using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace OutFlowReportExportAPI.Modals
{
    public enum Files
    {
        [EnumMember(Value = "no10_PlanSupervision.odt")]
        PlanSupervision,
        [EnumMember(Value = "no11_EnSVChecks.odt")]
        EnSVChecks,
        [EnumMember(Value = "no13_EnEndSelfChecks.odt")]
        EnEndSelfChecks,
        [EnumMember(Value = "no14_EnEndApplicationCompleted.odt")]
        EnEndApplicationCompleted,
        [EnumMember(Value = "no15_EnEndFinalChecks.odt")]
        EnEndFinalChecks,
        [EnumMember(Value = "no16_EnEndCompletion.odt")]
        EnEndCompletion,
        [EnumMember(Value = "no17_MMRecord.odt")]
        MMRecord,
        [EnumMember(Value = "no18_MMSVRecord.odt")]
        MMSVRecord,
        [EnumMember(Value = "no20_EN_Starting.odt")]
        EN_Starting
    }
}
