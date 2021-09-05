using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;

namespace Fxcel.Core.Interop
{
    using MicrosoftOleDbErrors = Microsoft.Office.Interop.Excel.OLEDBErrors;

    [SupportedOSPlatform("windows")]
    public sealed class XlOleDbErrors : XlComObject
    {
        internal XlOleDbErrors(MicrosoftOleDbErrors com) => raw = com;
        internal MicrosoftOleDbErrors raw;

        public override int Release() => ComHelper.Release(raw);
        public override void ForceRelease() => ComHelper.FinalRelease(raw);
        protected override void DidDispose()
        {
            raw = default!;
            base.DidDispose();
        }
    }
}
