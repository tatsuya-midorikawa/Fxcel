using System.Runtime.Versioning;

namespace Fxcel.Core.Interop
{
    using MicrosoftSmartTagRecognizers = Microsoft.Office.Interop.Excel.SmartTagRecognizers;

    [SupportedOSPlatform("windows")]
    public sealed class XlSmartTagRecognizers : XlComObject
    {
        internal XlSmartTagRecognizers(MicrosoftSmartTagRecognizers com) => raw = com;
        internal MicrosoftSmartTagRecognizers raw;

        public override int Release() => ComHelper.Release(raw);
        public override void FinalRelease() => ComHelper.FinalRelease(raw);
        protected override void DidDispose()
        {
            raw = default!;
            base.DidDispose();
        }
    }
}
