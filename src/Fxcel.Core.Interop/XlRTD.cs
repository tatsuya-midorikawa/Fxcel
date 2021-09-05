using System.Runtime.Versioning;

namespace Fxcel.Core.Interop
{
    using MicrosoftRTD = Microsoft.Office.Interop.Excel.RTD;

    [SupportedOSPlatform("windows")]
    public class XlRTD : XlComObject
    {
        public XlRTD(MicrosoftRTD rtd) : base(rtd) { }
        private MicrosoftRTD raw => (MicrosoftRTD)_raw;
    }
}
