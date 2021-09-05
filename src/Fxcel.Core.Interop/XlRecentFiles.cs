using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;

namespace Fxcel.Core.Interop
{
    using MicrosoftRecentFiles = Microsoft.Office.Interop.Excel.RecentFiles;

    [SupportedOSPlatform("windows")]
    public sealed class XlRecentFiles : XlComObject
    {
        internal XlRecentFiles(MicrosoftRecentFiles com) => raw = com;
        internal MicrosoftRecentFiles raw;

        public override int Release() => ComHelper.Release(raw);
        public override void FinalRelease() => ComHelper.FinalRelease(raw);
        protected override void DidDispose()
        {
            raw = default!;
            base.DidDispose();
        }
    }
}
