using System.Runtime.Versioning;

namespace Fxcel.Core.Interop
{
    using MicrosoftRTD = Microsoft.Office.Interop.Excel.RTD;

    [SupportedOSPlatform("windows")]
    public sealed class XlRTD : XlComObject
    {
        internal XlRTD(MicrosoftRTD com) => raw = com;
        internal MicrosoftRTD raw;

        public override int Release() => ComHelper.Release(raw);
        public override void ForceRelease() => ComHelper.FinalRelease(raw);
        protected override void DidDispose()
        {
            raw = default!;
            base.DidDispose();
        }
    }
}
