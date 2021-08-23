using System.Runtime.Versioning;
using MicrosoftSmartArtQuickStyles = Microsoft.Office.Core.SmartArtQuickStyles;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public readonly ref struct XlSmartArtQuickStyles
    {
        internal readonly MicrosoftSmartArtQuickStyles raw;
        public XlSmartArtQuickStyles(MicrosoftSmartArtQuickStyles styles) => raw = styles;

        public int Release() => ComHelper.Release(raw);
    }
}
