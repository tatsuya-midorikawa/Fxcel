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
    public readonly ref struct XlQuickAnalysis
    {
        internal readonly MicrosoftQuickAnalysis raw;
        public XlQuickAnalysis(MicrosoftQuickAnalysis analysis) => raw = analysis;

        public int Release() => ComHelper.Release(raw);
    }
}
