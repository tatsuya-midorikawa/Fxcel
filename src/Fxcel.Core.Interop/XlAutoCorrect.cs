using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;
using MicrosoftAutoCorrect = Microsoft.Office.Interop.Excel.AutoCorrect;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public readonly ref struct XlAutoCorrect
    {
        internal readonly MicrosoftAutoCorrect raw;
        public XlAutoCorrect(MicrosoftAutoCorrect autocorrect) => raw = autocorrect;
    }
}
