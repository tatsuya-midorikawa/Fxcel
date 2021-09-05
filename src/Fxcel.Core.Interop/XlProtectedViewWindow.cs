using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;

namespace Fxcel.Core.Interop
{
    using MicrosoftProtectedViewWindow = Microsoft.Office.Interop.Excel.ProtectedViewWindow;

    [SupportedOSPlatform("windows")]
    public sealed class XlProtectedViewWindow : XlComObject
    {
        internal XlProtectedViewWindow(MicrosoftProtectedViewWindow com) => raw = com;
        internal MicrosoftProtectedViewWindow raw;

        public override int Release() => ComHelper.Release(raw);
        public override void ForceRelease() => ComHelper.FinalRelease(raw);
        protected override void DidDispose()
        {
            raw = default!;
            base.DidDispose();
        }
    }
}
