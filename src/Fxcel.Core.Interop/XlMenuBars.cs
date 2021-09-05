using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;

namespace Fxcel.Core.Interop
{
    using MicrosoftMenuBars = Microsoft.Office.Interop.Excel.MenuBars;

    [SupportedOSPlatform("windows")]
    public class XlMenuBars : XlComObject
    {
        public XlMenuBars(MicrosoftMenuBars menubars) : base(menubars) { }
        private MicrosoftMenuBars raw => (MicrosoftMenuBars)_raw;
    }
}
