using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;

namespace Fxcel.Core.Interop
{
    using MicrosoftWatches = Microsoft.Office.Interop.Excel.Watches;

    [SupportedOSPlatform("windows")]
    public class XlWatches : XlComObject
    {
        public XlWatches(MicrosoftWatches watches) : base(watches) { }
        private MicrosoftWatches raw => (MicrosoftWatches)_raw;
    }
}
