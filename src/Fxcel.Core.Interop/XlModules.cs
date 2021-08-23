using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;
using MicrosoftModules = Microsoft.Office.Interop.Excel.Modules;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public readonly ref struct XlModules
    {
        internal readonly MicrosoftModules raw;
        public XlModules(MicrosoftModules modules) => raw = modules;

        public int Release() => ComHelper.Release(raw);
    }
}
