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
    public class XlDialogSheet : XlComObject
    {
        public XlDialogSheet(MicrosoftDialogSheet dialogSheet) : base(dialogSheet) { }
        private MicrosoftDialogSheet raw => (MicrosoftDialogSheet)_raw;
    }
}
