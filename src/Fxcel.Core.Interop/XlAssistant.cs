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
    public class XlAssistant : XlComObject
    {
        public XlAssistant(MicrosoftAssistant assistant) : base(assistant) { }
        private MicrosoftAssistant raw => (MicrosoftAssistant)_raw;
    }
}
