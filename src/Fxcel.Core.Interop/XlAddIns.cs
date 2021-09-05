using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Versioning;

namespace Fxcel.Core.Interop
{
    using MicrosoftAddIns = Microsoft.Office.Interop.Excel.AddIns;

    [SupportedOSPlatform("windows")]
    public sealed class XlAddIns : XlComObject
    {
        internal XlAddIns(MicrosoftAddIns com) => raw = com;
        internal MicrosoftAddIns raw;

        public override int Release() => ComHelper.Release(raw);
        public override void FinalRelease() => ComHelper.FinalRelease(raw);
        protected override void DidDispose()
        {
            raw = default!;
            base.DidDispose();
        }
    }
}
