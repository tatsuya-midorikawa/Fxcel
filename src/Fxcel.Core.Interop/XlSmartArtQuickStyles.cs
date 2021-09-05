using System.Runtime.Versioning;

namespace Fxcel.Core.Interop
{
    using MicrosoftSmartArtQuickStyles = Microsoft.Office.Core.SmartArtQuickStyles;

    [SupportedOSPlatform("windows")]
    public sealed class XlSmartArtQuickStyles : XlComObject
    {
        internal XlSmartArtQuickStyles(MicrosoftSmartArtQuickStyles com) => raw = com;
        internal MicrosoftSmartArtQuickStyles raw;

        public override int Release() => ComHelper.Release(raw);
        public override void ForceRelease() => ComHelper.FinalRelease(raw);
        protected override void DidDispose()
        {
            raw = default!;
            base.DidDispose();
        }
    }
}
