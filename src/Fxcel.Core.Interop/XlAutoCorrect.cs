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
    public class XlAutoCorrect : XlComObject
    {
        public XlAutoCorrect(MicrosoftAutoCorrect autocorrect) : base(autocorrect) { }
        private MicrosoftAutoCorrect raw => (MicrosoftAutoCorrect)_raw;
    }
}
