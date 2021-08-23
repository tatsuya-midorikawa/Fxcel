using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;
using MicrosoftDialogSheet = Microsoft.Office.Interop.Excel.DialogSheet;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public readonly ref struct XlDialogSheet
    {
        internal readonly MicrosoftDialogSheet raw;
        public XlDialogSheet(MicrosoftDialogSheet dialogSheet) => raw = dialogSheet;

        public int Release() => ComHelper.Release(raw);
    }
}
