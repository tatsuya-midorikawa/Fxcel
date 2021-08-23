using System.Runtime.Versioning;
using MicrosoftSmartArtColors = Microsoft.Office.Core.SmartArtColors;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public readonly ref struct XlSmartArtColors
    {
        internal readonly MicrosoftSmartArtColors raw;
        public XlSmartArtColors(MicrosoftSmartArtColors colors) => raw = colors;

        public int Release() => ComHelper.Release(raw);
    }
}
