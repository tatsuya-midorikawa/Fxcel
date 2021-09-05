using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Versioning;

namespace Fxcel.Core.Interop
{
    using MicrosoftWorksheet = Microsoft.Office.Interop.Excel.Worksheet;

    [SupportedOSPlatform("windows")]
    public sealed class XlWorksheet : XlComObject
    {
        internal XlWorksheet(MicrosoftWorksheet com) => raw = com;
        internal MicrosoftWorksheet raw;

        public override int Release() => ComHelper.Release(raw);
        public override void ForceRelease() => ComHelper.FinalRelease(raw);
        protected override void DidDispose()
        {
            raw = default!;
            base.DidDispose();
        }
    }
}
