using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;

namespace Fxcel.Core.Interop
{
    using MicrosoftModules = Microsoft.Office.Interop.Excel.Modules;

    [SupportedOSPlatform("windows")]
    public sealed class XlModules : XlComObject
    {
        internal XlModules(MicrosoftModules com) => raw = com;
        internal MicrosoftModules raw;

        public override int Release() => ComHelper.Release(raw);
        public override void ForceRelease() => ComHelper.FinalRelease(raw);
        protected override void DidDispose()
        {
            raw = default!;
            base.DidDispose();
        }
    }
}
