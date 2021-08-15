using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;
using MicrosoftComAddIns = Microsoft.Office.Core.COMAddIns;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public readonly ref struct XlComAddIns
    {
        internal readonly MicrosoftComAddIns raw;
        public XlComAddIns(MicrosoftComAddIns comAddIns) => raw = comAddIns;
    }
}
