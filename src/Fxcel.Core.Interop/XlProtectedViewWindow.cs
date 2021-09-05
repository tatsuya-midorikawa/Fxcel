using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;

namespace Fxcel.Core.Interop
{
    using MicrosoftProtectedViewWindow = Microsoft.Office.Interop.Excel.ProtectedViewWindow;

    [SupportedOSPlatform("windows")]
    public class XlProtectedViewWindow : XlComObject
    {
        public XlProtectedViewWindow(MicrosoftProtectedViewWindow window) : base(window) { }
        private MicrosoftProtectedViewWindow raw => (MicrosoftProtectedViewWindow)_raw;
    }
}
