using System.Runtime.Versioning;

namespace Fxcel.Core.Interop
{
    using MicrosoftSmartArtLayouts = Microsoft.Office.Core.SmartArtLayouts;

    [SupportedOSPlatform("windows")]
    public sealed class XlSmartArtLayouts : XlComObject
    {
        internal XlSmartArtLayouts(MicrosoftSmartArtLayouts com) => raw = com;
        internal MicrosoftSmartArtLayouts raw;

        public override int Release() => ComHelper.Release(raw);
        public override void ForceRelease() => ComHelper.FinalRelease(raw);
        protected override void DidDispose()
        {
            raw = default!;
            base.DidDispose();
        }
    }
}
