using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Versioning;

namespace Fxcel.Core.Interop
{
    using MicrosoftRanges = Microsoft.Office.Interop.Excel.Ranges;

    [SupportedOSPlatform("windows")]
    public sealed class XlRanges : XlComObject
    {
        internal XlRanges(MicrosoftRanges com) => raw = com;
        internal MicrosoftRanges raw;

        public override int Release() => ComHelper.Release(raw);
        public override void ForceRelease() => ComHelper.FinalRelease(raw);
        protected override void DidDispose()
        {
            raw = default!;
            base.DidDispose();
        }
    }
}
