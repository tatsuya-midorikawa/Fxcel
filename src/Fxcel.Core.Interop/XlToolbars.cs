using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;

namespace Fxcel.Core.Interop
{
    using MicrosoftToolbars = Microsoft.Office.Interop.Excel.Toolbars;

    [SupportedOSPlatform("windows")]
    public class XlToolbars : XlComObject
    {
        public XlToolbars(MicrosoftToolbars toolbars) : base(toolbars) { }
        private MicrosoftToolbars raw => (MicrosoftToolbars)_raw;
    }
}
