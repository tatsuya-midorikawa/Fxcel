using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;
using MicrosoftToolbars = Microsoft.Office.Interop.Excel.Toolbars;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public readonly ref struct XlToolbars
    {
        internal readonly MicrosoftToolbars raw;
        public XlToolbars(MicrosoftToolbars toolbars) => raw = toolbars;
    }
}
