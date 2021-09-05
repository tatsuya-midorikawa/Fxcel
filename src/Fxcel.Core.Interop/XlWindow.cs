using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;

namespace Fxcel.Core.Interop
{
    using MicrosoftWindow = Microsoft.Office.Interop.Excel.Window;

    [SupportedOSPlatform("windows")]
    public sealed class XlWindow : XlComObject
    {
        internal XlWindow(MicrosoftWindow com) => raw = com;
        internal MicrosoftWindow raw;

        public override int Release() => ComHelper.Release(raw);
        public override void FinalRelease() => ComHelper.FinalRelease(raw);
        protected override void DidDispose()
        {
            raw = default!;
            base.DidDispose();
        }
    }
}
