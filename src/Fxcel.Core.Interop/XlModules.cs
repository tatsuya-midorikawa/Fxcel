using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;

namespace Fxcel.Core.Interop
{
    using MicrosoftModules = Microsoft.Office.Interop.Excel.Modules;

    [SupportedOSPlatform("windows")]
    public class XlModules : XlComObject
    {
        public XlModules(MicrosoftModules modules) : base(modules) { }
        private MicrosoftModules raw => (MicrosoftModules)_raw;
    }
}
