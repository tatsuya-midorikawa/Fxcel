using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;
using MicrosoftWorksheetFunction = Microsoft.Office.Interop.Excel.WorksheetFunction;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public readonly ref struct XlWorksheetFunction
    {
        internal readonly MicrosoftWorksheetFunction raw;
        public XlWorksheetFunction(MicrosoftWorksheetFunction function) => raw = function;
    }
}
