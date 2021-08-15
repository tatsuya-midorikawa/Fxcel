using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;
using MicrosoftRange = Microsoft.Office.Interop.Excel.Range;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public readonly ref struct XlRange
    {
        internal readonly MicrosoftRange raw;
        public XlRange(MicrosoftRange range) => raw = range;
    }
}
