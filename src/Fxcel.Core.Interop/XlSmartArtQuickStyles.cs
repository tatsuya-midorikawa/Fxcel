using System.Runtime.Versioning;

namespace Fxcel.Core.Interop
{
    using MicrosoftSmartArtQuickStyles = Microsoft.Office.Core.SmartArtQuickStyles;

    [SupportedOSPlatform("windows")]
    public class XlSmartArtQuickStyles : XlComObject
    {
        public XlSmartArtQuickStyles(MicrosoftSmartArtQuickStyles styles) : base(styles) { }
        private MicrosoftSmartArtQuickStyles raw => (MicrosoftSmartArtQuickStyles)_raw;
    }
}
