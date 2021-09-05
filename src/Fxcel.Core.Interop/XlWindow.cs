using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;

namespace Fxcel.Core.Interop
{
    using MicrosoftWindow = Microsoft.Office.Interop.Excel.Window;

    [SupportedOSPlatform("windows")]
    public class XlWindow : XlComObject
    {
        public XlWindow(MicrosoftWindow window) : base(window) { }
        private MicrosoftWindow raw => (MicrosoftWindow)_raw;
    }
}
