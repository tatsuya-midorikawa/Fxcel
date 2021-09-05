using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Versioning;

namespace Fxcel.Core.Interop
{
    using MicrosoftWorksheets = Microsoft.Office.Interop.Excel.Worksheets;

    [SupportedOSPlatform("windows")]
    public sealed class XlWorksheets : XlComObject
    {
        internal XlWorksheets(MicrosoftWorksheets com) => raw = com;
        internal MicrosoftWorksheets raw;

        public override int Release() => ComHelper.Release(raw);
        public override void FinalRelease() => ComHelper.FinalRelease(raw);
        protected override void DidDispose()
        {
            raw = default!;
            base.DidDispose();
        }
    }
}
