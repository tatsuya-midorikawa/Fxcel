using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;

namespace Fxcel.Core.Interop
{
    using MicrosoftWindows = Microsoft.Office.Interop.Excel.Windows;

    [SupportedOSPlatform("windows")]
    public class XlWindows : XlComObject
    {
        public XlWindows(MicrosoftWindows window) : base(window) { }
        private MicrosoftWindows raw => (MicrosoftWindows)_raw;
    }
}
