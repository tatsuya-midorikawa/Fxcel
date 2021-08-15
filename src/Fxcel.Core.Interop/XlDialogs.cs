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
    public readonly ref struct XlDialogs
    {
        internal readonly MicrosoftDialogs raw;
        public XlDialogs(MicrosoftDialogs dialogs) => raw = dialogs;
    }
}
