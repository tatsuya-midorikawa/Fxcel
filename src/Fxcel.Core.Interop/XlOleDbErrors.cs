using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;
using MicrosoftOleDbErrors = Microsoft.Office.Interop.Excel.OLEDBErrors;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public readonly ref struct XlOleDbErrors
    {
        internal readonly MicrosoftOleDbErrors raw;
        public XlOleDbErrors(MicrosoftOleDbErrors oleDbErrors) => raw = oleDbErrors;

        public int Release() => ComHelper.Release(raw);
    }
}
