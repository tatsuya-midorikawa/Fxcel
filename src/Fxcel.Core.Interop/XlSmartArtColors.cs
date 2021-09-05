using System.Runtime.Versioning;

namespace Fxcel.Core.Interop
{
    using MicrosoftSmartArtColors = Microsoft.Office.Core.SmartArtColors;

    [SupportedOSPlatform("windows")]
    public class XlSmartArtColors : XlComObject
    {
        public XlSmartArtColors(MicrosoftSmartArtColors colors) : base(colors) { }
        private MicrosoftSmartArtColors raw => (MicrosoftSmartArtColors)_raw;
    }
}
