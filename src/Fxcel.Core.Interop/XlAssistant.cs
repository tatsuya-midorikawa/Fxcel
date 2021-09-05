using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;

namespace Fxcel.Core.Interop
{
    using MicrosoftAssistant = Microsoft.Office.Core.Assistant;

    [SupportedOSPlatform("windows")]
    public sealed class XlAssistant : XlComObject
    {
        internal XlAssistant(MicrosoftAssistant com) => raw = com;
        internal MicrosoftAssistant raw;

        public override int Release() => ComHelper.Release(raw);
        public override void ForceRelease() => ComHelper.FinalRelease(raw);
        protected override void DidDispose()
        {
            raw = default!;
            base.DidDispose();
        }
    }
}
