using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;

namespace Fxcel.Core.Interop
{
    using MicrosoftDialogSheet = Microsoft.Office.Interop.Excel.DialogSheet;

    [SupportedOSPlatform("windows")]
    public sealed class XlDialogSheet : XlComObject
    {
        internal XlDialogSheet(MicrosoftDialogSheet com) => raw = com;
        internal MicrosoftDialogSheet raw;

        public override int Release() => ComHelper.Release(raw);
        public override void FinalRelease() => ComHelper.FinalRelease(raw);
        protected override void DidDispose()
        {
            raw = default!;
            base.DidDispose();
        }
    }
}
