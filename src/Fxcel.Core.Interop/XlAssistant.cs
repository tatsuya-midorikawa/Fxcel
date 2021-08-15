using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;
using MicrosoftAssistant = Microsoft.Office.Core.Assistant;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public readonly ref struct XlAssistant
    {
        internal readonly MicrosoftAssistant raw;
        public XlAssistant(MicrosoftAssistant assistant) => raw = assistant;
    }
}
