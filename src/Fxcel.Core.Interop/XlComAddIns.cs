using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;

namespace Fxcel.Core.Interop
{
    using MicrosoftComAddIns = Microsoft.Office.Core.COMAddIns;

    [SupportedOSPlatform("windows")]
    public class XlComAddIns : XlComObject
    {
        public XlComAddIns(MicrosoftComAddIns comAddIns) : base(comAddIns) { }
        private MicrosoftComAddIns raw => (MicrosoftComAddIns)_raw;
    }
}
