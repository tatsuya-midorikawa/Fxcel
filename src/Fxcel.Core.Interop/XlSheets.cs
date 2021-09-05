using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Versioning;

namespace Fxcel.Core.Interop
{
    using MicrosoftSheets = Microsoft.Office.Interop.Excel.Sheets;

    [SupportedOSPlatform("windows")]
    public sealed class XlSheets : XlComObject
    {
        internal XlSheets(MicrosoftSheets com) => raw = com;
        internal MicrosoftSheets raw;

        public override int Release() => ComHelper.Release(raw);
        public override void ForceRelease() => ComHelper.FinalRelease(raw);
        protected override void DidDispose()
        {
            raw = default!;
            base.DidDispose();
        }
    }
}
