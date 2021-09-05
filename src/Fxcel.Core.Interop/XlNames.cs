using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;

namespace Fxcel.Core.Interop
{
    using MicrosoftNames = Microsoft.Office.Interop.Excel.Names;

    [SupportedOSPlatform("windows")]
    public class XlNames : XlComObject
    {
        public XlNames(MicrosoftNames names) : base(names) { }
        private MicrosoftNames raw => (MicrosoftNames)_raw;
    }
}
