using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MicrosoftSheets = Microsoft.Office.Interop.Excel.Sheets;
using System.Runtime.Versioning;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public readonly ref struct XlSheets
    {
        internal readonly MicrosoftSheets raw;
        public XlSheets(MicrosoftSheets worksheets) => raw = worksheets;

        public int Release() => ComHelper.Release(raw);
    }
}
