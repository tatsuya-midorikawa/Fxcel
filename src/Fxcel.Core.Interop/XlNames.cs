using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;
using MicrosoftNames = Microsoft.Office.Interop.Excel.Names;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public readonly ref struct XlNames
    {
        internal readonly MicrosoftNames raw;
        public XlNames(MicrosoftNames names) => raw = names;

        public int Release() => ComHelper.Release(raw);
    }
}
