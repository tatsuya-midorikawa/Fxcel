using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;

namespace Fxcel.Core.Interop
{
    using MicrosoftProtectedViewWindows = Microsoft.Office.Interop.Excel.ProtectedViewWindows;

    [SupportedOSPlatform("windows")]
    public class XlProtectedViewWindows : XlComObject
    {
        public XlProtectedViewWindows(MicrosoftProtectedViewWindows windows) : base(windows) { }
        private MicrosoftProtectedViewWindows raw => (MicrosoftProtectedViewWindows)_raw;
    }
}
