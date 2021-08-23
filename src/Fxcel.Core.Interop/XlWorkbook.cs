using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MicrosoftWorkbook = Microsoft.Office.Interop.Excel.Workbook;
using System.Runtime.Versioning;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public readonly ref struct XlWorkbook
    {
        internal readonly MicrosoftWorkbook raw;
        public XlWorkbook(MicrosoftWorkbook workbook) => raw = workbook;

        public int Release() => ComHelper.Release(raw);
    }
}
