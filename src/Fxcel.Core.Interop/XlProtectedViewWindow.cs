using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;
using MicrosoftProtectedViewWindow = Microsoft.Office.Interop.Excel.ProtectedViewWindow;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public readonly ref struct XlProtectedViewWindow
    {
        internal readonly MicrosoftProtectedViewWindow raw;
        public XlProtectedViewWindow(MicrosoftProtectedViewWindow window) => raw = window;
    }
}
