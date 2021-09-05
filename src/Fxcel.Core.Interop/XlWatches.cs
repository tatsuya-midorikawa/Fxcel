using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;

namespace Fxcel.Core.Interop
{
    using MicrosoftWatches = Microsoft.Office.Interop.Excel.Watches;

    [SupportedOSPlatform("windows")]
    public sealed class XlWatches : XlComObject
    {
        internal XlWatches(MicrosoftWatches com) => raw = com;
        internal MicrosoftWatches raw;

        public override int Release() => ComHelper.Release(raw);
        public override void ForceRelease() => ComHelper.FinalRelease(raw);
        protected override void DidDispose()
        {
            raw = default!;
            base.DidDispose();
        }
    }
}
