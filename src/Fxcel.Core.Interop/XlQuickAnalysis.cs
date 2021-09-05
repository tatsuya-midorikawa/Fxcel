using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;
using MicrosoftQuickAnalysis = Microsoft.Office.Interop.Excel.QuickAnalysis;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public class XlQuickAnalysis : XlComObject
    {
        public XlQuickAnalysis(MicrosoftQuickAnalysis analysis) : base(analysis) { }
        private MicrosoftQuickAnalysis raw => (MicrosoftQuickAnalysis)_raw;
    }
}
