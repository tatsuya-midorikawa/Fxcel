using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Versioning;

namespace Fxcel.Core.Interop
{
    using MicrosoftWorksheet = Microsoft.Office.Interop.Excel.Worksheet;

    [SupportedOSPlatform("windows")]
    public class XlWorksheet : XlComObject
    {
        public XlWorksheet(MicrosoftWorksheet worksheet) : base(worksheet) { }
        private MicrosoftWorksheet raw => (MicrosoftWorksheet)_raw;
    }
}
