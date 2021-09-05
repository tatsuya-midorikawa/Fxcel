using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;

namespace Fxcel.Core.Interop
{
    using MicrosoftNames = Microsoft.Office.Interop.Excel.Names;

    [SupportedOSPlatform("windows")]
    public sealed class XlNames : XlComObject
    {
        internal XlNames(MicrosoftNames com) => raw = com;
        internal MicrosoftNames raw;

        public override int Release() => ComHelper.Release(raw);
        public override void ForceRelease() => ComHelper.FinalRelease(raw);
        protected override void DidDispose()
        {
            raw = default!;
            base.DidDispose();
        }
    }
}
