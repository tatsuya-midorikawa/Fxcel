using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MicrosoftRanges = Microsoft.Office.Interop.Excel.Ranges;
using System.Runtime.Versioning;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public readonly ref struct XlRanges
    {
        internal readonly MicrosoftRanges raw;
        public XlRanges(MicrosoftRanges ranges) => raw = ranges;
    }
}
