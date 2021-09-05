using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;

namespace Fxcel.Core.Interop
{
    using MicrosoftCellFormat = Microsoft.Office.Interop.Excel.CellFormat;

    [SupportedOSPlatform("windows")]
    public sealed class XlCellFormat : XlComObject
    {
        internal XlCellFormat(MicrosoftCellFormat com) => raw = com;
        internal MicrosoftCellFormat raw;

        public override int Release() => ComHelper.Release(raw);
        public override void ForceRelease() => ComHelper.FinalRelease(raw);
        protected override void DidDispose()
        {
            raw = default!;
            base.DidDispose();
        }
    }
}
