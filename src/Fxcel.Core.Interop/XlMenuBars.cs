using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;
using MicrosoftMenuBars = Microsoft.Office.Interop.Excel.MenuBars;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public readonly ref struct XlMenuBars
    {
        internal readonly MicrosoftMenuBars raw;
        public XlMenuBars(MicrosoftMenuBars menubars) => raw = menubars;

        public int Release() => ComHelper.Release(raw);
    }
}
