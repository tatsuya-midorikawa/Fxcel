using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;

namespace Fxcel.Core.Interop
{
    using MicrosoftWindows = Microsoft.Office.Interop.Excel.Windows;

    [SupportedOSPlatform("windows")]
    public sealed class XlWindows : XlComObject
    {
        internal XlWindows(MicrosoftWindows com) => raw = com;
        internal MicrosoftWindows raw;

        public override int Release() => ComHelper.Release(raw);
        public override void ForceRelease() => ComHelper.FinalRelease(raw);
        protected override void DidDispose()
        {
            raw = default!;
            base.DidDispose();
        }
    }
}
