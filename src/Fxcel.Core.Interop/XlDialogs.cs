using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;
using MicrosoftDialogs = Microsoft.Office.Interop.Excel.Dialogs;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public class XlDialogs : XlComObject
    {
        public XlDialogs(MicrosoftDialogs dialogs) : base(dialogs) { }
        private MicrosoftDialogs raw => (MicrosoftDialogs)_raw;
    }
}
