using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;

namespace Fxcel.Core.Interop
{
    using MicrosoftWorksheetFunction = Microsoft.Office.Interop.Excel.WorksheetFunction;

    [SupportedOSPlatform("windows")]
    public sealed class XlWorksheetFunction : XlComObject
    {
        internal XlWorksheetFunction(MicrosoftWorksheetFunction com) => raw = com;
        internal MicrosoftWorksheetFunction raw;

        public override int Release() => ComHelper.Release(raw);
        public override void FinalRelease() => ComHelper.FinalRelease(raw);
        protected override void DidDispose()
        {
            raw = default!;
            base.DidDispose();
        }
    }
}
