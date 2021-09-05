using System.Runtime.Versioning;

namespace Fxcel.Core.Interop
{
    using MicrosoftSmartArtColors = Microsoft.Office.Core.SmartArtColors;

    [SupportedOSPlatform("windows")]
    public sealed class XlSmartArtColors : XlComObject
    {
        internal XlSmartArtColors(MicrosoftSmartArtColors com) => raw = com;
        internal MicrosoftSmartArtColors raw;

        public override int Release() => ComHelper.Release(raw);
        public override void FinalRelease() => ComHelper.FinalRelease(raw);
        protected override void DidDispose()
        {
            raw = default!;
            base.DidDispose();
        }
    }
}
