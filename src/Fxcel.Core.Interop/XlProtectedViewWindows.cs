using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;
using MicrosoftProtectedViewWindows = Microsoft.Office.Interop.Excel.ProtectedViewWindows;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public readonly ref struct XlProtectedViewWindows
    {
        internal readonly MicrosoftProtectedViewWindows raw;
        public XlProtectedViewWindows(MicrosoftProtectedViewWindows windows) => raw = windows;
    }
}
