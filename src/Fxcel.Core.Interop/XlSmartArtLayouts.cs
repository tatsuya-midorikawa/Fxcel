using System.Runtime.Versioning;
using MicrosoftSmartArtLayouts = Microsoft.Office.Core.SmartArtLayouts;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public readonly ref struct XlSmartArtLayouts
    {
        internal readonly MicrosoftSmartArtLayouts raw;
        public XlSmartArtLayouts(MicrosoftSmartArtLayouts layouts) => raw = layouts;

        public int Release() => ComHelper.Release(raw);
    }
}
