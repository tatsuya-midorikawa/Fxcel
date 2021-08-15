using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;
using MicrosoftCommandBars = Microsoft.Office.Core.CommandBars;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public readonly ref struct XlCommandBars
    {
        internal readonly MicrosoftCommandBars raw;
        public XlCommandBars(MicrosoftCommandBars commandBars) => raw = commandBars;
    }
}
