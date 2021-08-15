using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;
using MicrosoftMenuBar = Microsoft.Office.Interop.Excel.MenuBar;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public readonly ref struct XlMenuBar
    {
        internal readonly MicrosoftMenuBar raw;
        public XlMenuBar(MicrosoftMenuBar menubar) => raw = menubar;
    }
}
