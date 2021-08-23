using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;
using MicrosoftWindows = Microsoft.Office.Interop.Excel.Windows;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public readonly ref struct XlWindows
    {
        internal readonly MicrosoftWindows raw;
        public XlWindows(MicrosoftWindows window) => raw = window;

        public int Release() => ComHelper.Release(raw);
    }
}
