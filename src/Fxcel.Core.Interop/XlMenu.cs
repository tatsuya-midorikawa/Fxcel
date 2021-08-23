using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;
using MicrosoftMenu = Microsoft.Office.Interop.Excel.Menu;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public readonly ref struct XlMenu
    {
        internal readonly MicrosoftMenu raw;
        public XlMenu(MicrosoftMenu menu) => raw = menu;

        public int Release() => ComHelper.Release(raw);
    }
}
