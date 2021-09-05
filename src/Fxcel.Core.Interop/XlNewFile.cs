using System.Runtime.Versioning;

namespace Fxcel.Core.Interop
{
    using MicrosoftNewFile = Microsoft.Office.Core.NewFile;

    [SupportedOSPlatform("windows")]
    public sealed class XlNewFile : XlComObject
    {
        internal XlNewFile(MicrosoftNewFile com) => raw = com;
        internal MicrosoftNewFile raw;

        public override int Release() => ComHelper.Release(raw);
        public override void ForceRelease() => ComHelper.FinalRelease(raw);
        protected override void DidDispose()
        {
            raw = default!;
            base.DidDispose();
        }
    }
}
