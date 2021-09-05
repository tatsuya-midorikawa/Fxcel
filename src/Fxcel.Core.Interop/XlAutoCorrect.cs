using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;

namespace Fxcel.Core.Interop
{
    using MicrosoftAutoCorrect = Microsoft.Office.Interop.Excel.AutoCorrect;

    [SupportedOSPlatform("windows")]
    public sealed class XlAutoCorrect : XlComObject
    {
        internal XlAutoCorrect(MicrosoftAutoCorrect com) => raw = com;
        internal MicrosoftAutoCorrect raw;

        public override int Release() => ComHelper.Release(raw);
        public override void ForceRelease() => ComHelper.FinalRelease(raw);
        protected override void DidDispose()
        {
            raw = default!;
            base.DidDispose();
        }
    }
}
