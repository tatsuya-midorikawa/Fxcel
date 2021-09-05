using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;

namespace Fxcel.Core.Interop
{
    using MicrosoftRange = Microsoft.Office.Interop.Excel.Range;

    [SupportedOSPlatform("windows")]
    public class XlRange : XlComObject
    {
        public XlRange(MicrosoftRange range) : base(range) { }
        internal MicrosoftRange raw => (MicrosoftRange)_raw;
    }
}
