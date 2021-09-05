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
    public sealed class XlOdbcErrors : XlComObject
    {
        internal XlOdbcErrors(MicrosoftOdbcErrors com) => raw = com;
        internal MicrosoftOdbcErrors raw;

        public override int Release() => ComHelper.Release(raw);
        public override void ForceRelease() => ComHelper.FinalRelease(raw);
        protected override void DidDispose()
        {
            raw = default!;
            base.DidDispose();
        }
    }
}
