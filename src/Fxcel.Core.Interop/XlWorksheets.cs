using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MicrosoftWorksheets = Microsoft.Office.Interop.Excel.Worksheets;
using System.Runtime.Versioning;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public readonly struct XlWorksheets
    {
        internal readonly MicrosoftWorksheets raw;
        public XlWorksheets(MicrosoftWorksheets worksheets) => raw = worksheets;
    }
}
