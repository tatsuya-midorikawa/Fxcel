using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;
using MicrosoftCellFormat = Microsoft.Office.Interop.Excel.CellFormat;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public readonly ref struct XlCellFormat
    {
        internal readonly MicrosoftCellFormat raw;
        public XlCellFormat(MicrosoftCellFormat format) => raw = format;
    }
}
