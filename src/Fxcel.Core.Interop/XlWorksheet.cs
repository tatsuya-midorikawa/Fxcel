using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MicrosoftWorksheet = Microsoft.Office.Interop.Excel.Worksheet;
using System.Runtime.Versioning;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public readonly ref struct XlWorksheet
    {
        internal readonly MicrosoftWorksheet raw;
        public XlWorksheet(MicrosoftWorksheet worksheet) => raw = worksheet;

        public int Release() => ComHelper.Release(raw);
    }
}
