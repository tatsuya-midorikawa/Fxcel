using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;
using MicrosoftWindow = Microsoft.Office.Interop.Excel.Window;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public readonly ref struct XlWindow
    {
        internal readonly MicrosoftWindow raw;
        public XlWindow(MicrosoftWindow window) => raw = window;
    }
}
