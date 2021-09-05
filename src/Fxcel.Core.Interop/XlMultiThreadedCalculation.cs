using System.Runtime.Versioning;

namespace Fxcel.Core.Interop
{
    using MicrosoftMultiThreadedCalculation = Microsoft.Office.Interop.Excel.MultiThreadedCalculation;

    [SupportedOSPlatform("windows")]
    public sealed class XlMultiThreadedCalculation : XlComObject
    {
        internal XlMultiThreadedCalculation(MicrosoftMultiThreadedCalculation com) => raw = com;
        internal MicrosoftMultiThreadedCalculation raw;

        public override int Release() => ComHelper.Release(raw);
        public override void FinalRelease() => ComHelper.FinalRelease(raw);
        protected override void DidDispose()
        {
            raw = default!;
            base.DidDispose();
        }
    }
}
