using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;

namespace Fxcel.Core.Interop
{
    using MicrosoftDialogs = Microsoft.Office.Interop.Excel.Dialogs;

    [SupportedOSPlatform("windows")]
    public class XlDialogs : XlComObject
    {
        internal XlDialogs(MicrosoftDialogs com) => raw = com;
        internal MicrosoftDialogs raw;

        public override int Release() => ComHelper.Release(raw);
        public override void ForceRelease() => ComHelper.FinalRelease(raw);
        protected override void DidDispose()
        {
            raw = default!;
            base.DidDispose();
        }
    }
}
