using System.Runtime.Versioning;

namespace Fxcel.Core.Interop
{
    using MicrosoftSmartArtLayouts = Microsoft.Office.Core.SmartArtLayouts;

    [SupportedOSPlatform("windows")]
    public class XlSmartArtLayouts : XlComObject
    {
        public XlSmartArtLayouts(MicrosoftSmartArtLayouts layouts) : base(layouts) { }
        private MicrosoftSmartArtLayouts raw => (MicrosoftSmartArtLayouts)_raw;
    }
}
