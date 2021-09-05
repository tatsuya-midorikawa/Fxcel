using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;

namespace Fxcel.Core.Interop
{
    using MicrosoftOleDbErrors = Microsoft.Office.Interop.Excel.OLEDBErrors;

    [SupportedOSPlatform("windows")]
    public class XlOleDbErrors : XlComObject
    {
        public XlOleDbErrors(MicrosoftOleDbErrors oleDbErrors) : base(oleDbErrors) { }
        private MicrosoftOleDbErrors raw => (MicrosoftOleDbErrors)_raw;
    }
}
