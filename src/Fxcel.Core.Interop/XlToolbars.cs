using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;

namespace Fxcel.Core.Interop
{
    using MicrosoftToolbars = Microsoft.Office.Interop.Excel.Toolbars;

    [SupportedOSPlatform("windows")]
    public sealed class XlToolbars : XlComObject
    {
        internal XlToolbars(MicrosoftToolbars com) => raw = com;
        internal MicrosoftToolbars raw;

        public override int Release() => ComHelper.Release(raw);
        public override void FinalRelease() => ComHelper.FinalRelease(raw);
        protected override void DidDispose()
        {
            raw = default!;
            base.DidDispose();
        }
    }
}
