using System.Runtime.Versioning;

namespace Fxcel.Core.Interop
{
    using MicrosoftAutoRecover = Microsoft.Office.Interop.Excel.AutoRecover;

    [SupportedOSPlatform("windows")]
    public sealed class XlAutoRecover : XlComObject
    {
        internal XlAutoRecover(MicrosoftAutoRecover com) => raw = com;
        internal MicrosoftAutoRecover raw;

        public override int Release() => ComHelper.Release(raw);
        public override void FinalRelease() => ComHelper.FinalRelease(raw);
        protected override void DidDispose()
        {
            raw = default!;
            base.DidDispose();
        }
    }
}
