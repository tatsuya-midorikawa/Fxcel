using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;

namespace Fxcel.Core.Interop
{
    using MicrosoftOdbcErrors = Microsoft.Office.Interop.Excel.ODBCErrors;

    [SupportedOSPlatform("windows")]
    public class XlOdbcErrors : XlComObject
    {
        public XlOdbcErrors(MicrosoftOdbcErrors odbcErrors) : base(odbcErrors) { }
        private MicrosoftOdbcErrors raw => (MicrosoftOdbcErrors)_raw;
    }
}
