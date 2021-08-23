using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MicrosoftAddIns = Microsoft.Office.Interop.Excel.AddIns;
using System.Runtime.Versioning;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public readonly ref struct XlAddIns
    {
        internal readonly MicrosoftAddIns raw;
        public XlAddIns(MicrosoftAddIns addins) => raw = addins;

        public int Release() => ComHelper.Release(raw);
    }
}
