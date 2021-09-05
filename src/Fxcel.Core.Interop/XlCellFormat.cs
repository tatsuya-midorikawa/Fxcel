using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;

namespace Fxcel.Core.Interop
{
    using MicrosoftCellFormat = Microsoft.Office.Interop.Excel.CellFormat;

    [SupportedOSPlatform("windows")]
    public class XlCellFormat : XlComObject
    {
        public XlCellFormat(MicrosoftCellFormat format) : base(format) { }
        internal MicrosoftCellFormat raw => (MicrosoftCellFormat)_raw;
    }
}
