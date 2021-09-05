using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;

namespace Fxcel.Core.Interop
{
    using MicrosoftRange = Microsoft.Office.Interop.Excel.Range;

    [SupportedOSPlatform("windows")]
    public sealed class XlRange : XlComObject
    {
        internal XlRange(MicrosoftRange com) => raw = com;
        internal MicrosoftRange raw;

        public override int Release() => ComHelper.Release(raw);
        public override void ForceRelease() => ComHelper.FinalRelease(raw);
        protected override void DidDispose()
        {
            raw = default!;
            base.DidDispose();
        }
    }
}
