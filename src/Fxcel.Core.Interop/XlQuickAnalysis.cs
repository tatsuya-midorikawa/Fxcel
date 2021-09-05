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
    public sealed class XlQuickAnalysis : XlComObject
    {
        internal XlQuickAnalysis(MicrosoftQuickAnalysis com) => raw = com;
        internal MicrosoftQuickAnalysis raw;

        public override int Release() => ComHelper.Release(raw);
        public override void ForceRelease() => ComHelper.FinalRelease(raw);
        protected override void DidDispose()
        {
            raw = default!;
            base.DidDispose();
        }
    }
}
