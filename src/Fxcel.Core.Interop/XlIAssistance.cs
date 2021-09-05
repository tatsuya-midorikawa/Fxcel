using System.Runtime.Versioning;

namespace Fxcel.Core.Interop
{
    using MicrosoftIAssistance = Microsoft.Office.Core.IAssistance;

    [SupportedOSPlatform("windows")]
    public sealed class XlIAssistance : XlComObject
    {
        internal XlIAssistance(MicrosoftIAssistance com) => raw = com;
        internal MicrosoftIAssistance raw;

        public override int Release() => ComHelper.Release(raw);
        public override void ForceRelease() => ComHelper.FinalRelease(raw);
        protected override void DidDispose()
        {
            raw = default!;
            base.DidDispose();
        }
    }
}
