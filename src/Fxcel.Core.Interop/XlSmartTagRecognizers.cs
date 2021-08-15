using System.Runtime.Versioning;
using MicrosoftSmartTagRecognizers = Microsoft.Office.Interop.Excel.SmartTagRecognizers;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public readonly ref struct XlSmartTagRecognizers
    {
        internal readonly MicrosoftSmartTagRecognizers raw;
        public XlSmartTagRecognizers(MicrosoftSmartTagRecognizers recognizers) => raw = recognizers;
    }
}
