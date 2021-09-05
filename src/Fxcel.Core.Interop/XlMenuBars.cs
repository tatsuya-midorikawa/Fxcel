using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;

namespace Fxcel.Core.Interop
{
    using MicrosoftMenuBars = Microsoft.Office.Interop.Excel.MenuBars;

    [SupportedOSPlatform("windows")]
    public sealed class XlMenuBars : XlComObject
    {
        internal XlMenuBars(MicrosoftMenuBars com) => raw = com;
        internal MicrosoftMenuBars raw;

        public override int Release() => ComHelper.Release(raw);
        public override void FinalRelease() => ComHelper.FinalRelease(raw);
        protected override void DidDispose()
        {
            raw = default!;
            base.DidDispose();
        }
    }
}
