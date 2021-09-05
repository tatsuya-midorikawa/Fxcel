using System.Runtime.Versioning;

namespace Fxcel.Core.Interop
{
    using MicrosoftErrorCheckingOptions = Microsoft.Office.Interop.Excel.ErrorCheckingOptions;

    [SupportedOSPlatform("windows")]
    public class XlErrorCheckingOptions : XlComObject
    {
        public XlErrorCheckingOptions(MicrosoftErrorCheckingOptions options) : base(options) { }
        private MicrosoftErrorCheckingOptions raw => (MicrosoftErrorCheckingOptions)_raw;
    }
}
