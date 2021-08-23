using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;
using MicrosoftAnswerWizard = Microsoft.Office.Core.AnswerWizard;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public readonly ref struct XlAnswerWizard
    {
        internal readonly MicrosoftAnswerWizard raw;
        public XlAnswerWizard(MicrosoftAnswerWizard wizard) => raw = wizard;

        public int Release() => ComHelper.Release(raw);
    }
}
