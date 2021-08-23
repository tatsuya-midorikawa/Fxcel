using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;
using MicrosoftDefaultWebOptions = Microsoft.Office.Interop.Excel.DefaultWebOptions;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public readonly ref struct XlDefaultWebOptions
    {
        internal readonly MicrosoftDefaultWebOptions raw;
        public XlDefaultWebOptions(MicrosoftDefaultWebOptions options) => raw = options;

        public int Release() => ComHelper.Release(raw);
    }
}
