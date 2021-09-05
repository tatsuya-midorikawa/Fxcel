using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;

namespace Fxcel.Core.Interop
{
    using MicrosoftProtectedViewWindows = Microsoft.Office.Interop.Excel.ProtectedViewWindows;

    [SupportedOSPlatform("windows")]
    public sealed class XlProtectedViewWindows : XlComObject
    {
        internal XlProtectedViewWindows(MicrosoftProtectedViewWindows com) => raw = com;
        internal MicrosoftProtectedViewWindows raw;

        public override int Release() => ComHelper.Release(raw);
        public override void FinalRelease() => ComHelper.FinalRelease(raw);
        protected override void DidDispose()
        {
            raw = default!;
            base.DidDispose();
        }
    }
}
