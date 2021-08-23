using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;
using MicrosoftOdbcErrors = Microsoft.Office.Interop.Excel.ODBCErrors;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public readonly ref struct XlOdbcErrors
    {
        internal readonly MicrosoftOdbcErrors raw;
        public XlOdbcErrors(MicrosoftOdbcErrors odbcErrors) => raw = odbcErrors;

        public int Release() => ComHelper.Release(raw);
    }
}
