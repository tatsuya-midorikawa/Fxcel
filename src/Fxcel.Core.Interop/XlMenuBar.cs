using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;

namespace Fxcel.Core.Interop
{
    using MicrosoftMenuBar = Microsoft.Office.Interop.Excel.MenuBar;

    [SupportedOSPlatform("windows")]
    public sealed class XlMenuBar : XlComObject
    {
        internal XlMenuBar(MicrosoftMenuBar com) => raw = com;
        internal MicrosoftMenuBar raw;

        public override int Release() => ComHelper.Release(raw);
        public override void FinalRelease() => ComHelper.FinalRelease(raw);
        protected override void DidDispose()
        {
            raw = default!;
            base.DidDispose();
        }
    }
}
