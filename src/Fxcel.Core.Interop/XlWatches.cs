using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;
using MicrosoftWatches = Microsoft.Office.Interop.Excel.Watches;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public readonly ref struct XlWatches
    {
        internal readonly MicrosoftWatches raw;
        public XlWatches(MicrosoftWatches watches) => raw = watches;

        public int Release() => ComHelper.Release(raw);
    }
}
