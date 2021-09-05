using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Versioning;

namespace Fxcel.Core.Interop
{
    using MicrosoftAddIns = Microsoft.Office.Interop.Excel.AddIns;

    [SupportedOSPlatform("windows")]
    public class XlAddIns : XlComObject
    {
        public XlAddIns(MicrosoftAddIns addins) : base(addins) { }
        private MicrosoftAddIns raw => (MicrosoftAddIns)_raw;
    }
}
