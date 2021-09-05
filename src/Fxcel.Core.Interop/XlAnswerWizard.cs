using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;

namespace Fxcel.Core.Interop
{
    using MicrosoftAnswerWizard = Microsoft.Office.Core.AnswerWizard;

    [SupportedOSPlatform("windows")]
    public sealed class XlAnswerWizard : XlComObject
    {
        internal XlAnswerWizard(MicrosoftAnswerWizard com) => raw = com;
        internal MicrosoftAnswerWizard raw;

        public override int Release() => ComHelper.Release(raw);
        public override void ForceRelease() => ComHelper.FinalRelease(raw);
        protected override void DidDispose()
        {
            raw = default!;
            base.DidDispose();
        }
    }
}
