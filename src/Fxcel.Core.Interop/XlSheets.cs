using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Versioning;

namespace Fxcel.Core.Interop
{
    using MicrosoftSheets = Microsoft.Office.Interop.Excel.Sheets;

    [SupportedOSPlatform("windows")]
    public class XlSheets : XlComObject
    {
        public XlSheets(MicrosoftSheets worksheets) : base(worksheets) { }
        private MicrosoftSheets raw => (MicrosoftSheets)_raw;
    }
}
