using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;

namespace Fxcel.Core.Interop
{
    using MicrosoftWorksheetFunction = Microsoft.Office.Interop.Excel.WorksheetFunction;

    [SupportedOSPlatform("windows")]
    public class XlWorksheetFunction : XlComObject
    {
        public XlWorksheetFunction(MicrosoftWorksheetFunction function) : base(function) { }
        private MicrosoftWorksheetFunction raw => (MicrosoftWorksheetFunction)_raw;
    }
}
