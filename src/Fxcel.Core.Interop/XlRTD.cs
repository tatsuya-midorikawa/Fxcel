using System.Runtime.Versioning;
using MicrosoftRTD = Microsoft.Office.Interop.Excel.RTD;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public readonly ref struct XlRTD
    {
        internal readonly MicrosoftRTD raw;
        public XlRTD(MicrosoftRTD rtd) => raw = rtd;
    }
}
