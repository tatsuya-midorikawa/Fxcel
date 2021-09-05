using System.Runtime.Versioning;

namespace Fxcel.Core.Interop
{
    using MicrosoftAddIns2 = Microsoft.Office.Interop.Excel.AddIns2;

    [SupportedOSPlatform("windows")]
    public sealed class XlAddIns2 : XlComObject
    {
        internal XlAddIns2(MicrosoftAddIns2 com) => raw = com;
        internal MicrosoftAddIns2 raw;

        public override int Release() => ComHelper.Release(raw);
        public override void FinalRelease() => ComHelper.FinalRelease(raw);
        protected override void DidDispose()
        {
            raw = default!;
            base.DidDispose();
        }
    }
}
