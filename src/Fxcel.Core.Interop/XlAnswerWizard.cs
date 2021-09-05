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
    public class XlAnswerWizard : XlComObject
    {
        public XlAnswerWizard(MicrosoftAnswerWizard wizard) : base(wizard) { }
        private MicrosoftAnswerWizard raw => (MicrosoftAnswerWizard)_raw;
    }
}
