using System.Runtime.Versioning;

namespace Fxcel.Core.Interop
{
    using MicrosoftErrorCheckingOptions = Microsoft.Office.Interop.Excel.ErrorCheckingOptions;

    [SupportedOSPlatform("windows")]
    public sealed class XlErrorCheckingOptions : XlComObject
    {
        internal XlErrorCheckingOptions(MicrosoftErrorCheckingOptions com) => raw = com;
        internal MicrosoftErrorCheckingOptions raw;

        public override int Release() => ComHelper.Release(raw);
        public override void ForceRelease() => ComHelper.FinalRelease(raw);
        protected override void DidDispose()
        {
            raw = default!;
            base.DidDispose();
        }
    }
}
