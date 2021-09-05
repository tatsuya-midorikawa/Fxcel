using System.Runtime.Versioning;

namespace Fxcel.Core.Interop
{
    using MicrosoftSmartTagRecognizers = Microsoft.Office.Interop.Excel.SmartTagRecognizers;

    [SupportedOSPlatform("windows")]
    public class XlSmartTagRecognizers : XlComObject
    {
        public XlSmartTagRecognizers(MicrosoftSmartTagRecognizers recognizers) : base(recognizers) { }
        private MicrosoftSmartTagRecognizers raw => (MicrosoftSmartTagRecognizers)_raw;
    }
}
