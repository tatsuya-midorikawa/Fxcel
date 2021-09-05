using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;

namespace Fxcel.Core.Interop
{
    using MicrosoftComAddIns = Microsoft.Office.Core.COMAddIns;

    [SupportedOSPlatform("windows")]
    public sealed class XlComAddIns : XlComObject
    {
        internal XlComAddIns(MicrosoftComAddIns com) => raw = com;
        internal MicrosoftComAddIns raw;

        public override int Release() => ComHelper.Release(raw);
        public override void FinalRelease() => ComHelper.FinalRelease(raw);
        protected override void DidDispose()
        {
            raw = default!;
            base.DidDispose();
        }
    }
}
