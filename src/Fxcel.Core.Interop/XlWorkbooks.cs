using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MicrosoftWorkbooks = Microsoft.Office.Interop.Excel.Workbooks;
using System.Runtime.Versioning;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public readonly ref struct XlWorkbooks
    {
        internal readonly MicrosoftWorkbooks raw;
        public XlWorkbooks(MicrosoftWorkbooks workbooks) => raw = workbooks;

        public int Release() => ComHelper.Release(raw);
    }
}
