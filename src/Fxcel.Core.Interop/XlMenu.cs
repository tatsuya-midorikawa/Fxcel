using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;

namespace Fxcel.Core.Interop
{
    using MicrosoftMenu = Microsoft.Office.Interop.Excel.Menu;

    [SupportedOSPlatform("windows")]
    public sealed class XlMenu : XlComObject
    {
        internal XlMenu(MicrosoftMenu com) => raw = com;
        internal MicrosoftMenu raw;

        public override int Release() => ComHelper.Release(raw);
        public override void FinalRelease() => ComHelper.FinalRelease(raw);
        protected override void DidDispose()
        {
            raw = default!;
            base.DidDispose();
        }
    }
}
