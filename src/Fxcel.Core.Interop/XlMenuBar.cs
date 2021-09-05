using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;

namespace Fxcel.Core.Interop
{
    using MicrosoftMenuBar = Microsoft.Office.Interop.Excel.MenuBar;

    [SupportedOSPlatform("windows")]
    public class XlMenuBar : XlComObject
    {
        public XlMenuBar(MicrosoftMenuBar menubar) : base(menubar) { }
        private MicrosoftMenuBar raw => (MicrosoftMenuBar)_raw;
    }
}
