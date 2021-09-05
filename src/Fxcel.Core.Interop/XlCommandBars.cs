using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;

namespace Fxcel.Core.Interop
{
    using MicrosoftCommandBars = Microsoft.Office.Core.CommandBars;

    [SupportedOSPlatform("windows")]
    public sealed class XlCommandBars : XlComObject
    {
        internal XlCommandBars(MicrosoftCommandBars com) => raw = com;
        internal MicrosoftCommandBars raw;

        public override int Release() => ComHelper.Release(raw);
        public override void FinalRelease() => ComHelper.FinalRelease(raw);
        protected override void DidDispose()
        {
            raw = default!;
            base.DidDispose();
        }
    }
}
