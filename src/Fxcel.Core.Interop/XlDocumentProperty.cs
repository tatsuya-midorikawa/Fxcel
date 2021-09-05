using System.Runtime.Versioning;

namespace Fxcel.Core.Interop
{
    using MicrosoftDocumentProperty = Microsoft.Office.Core.DocumentProperty;

    [SupportedOSPlatform("windows")]
    public sealed class XlDocumentProperty : XlComObject
    {
        internal XlDocumentProperty(MicrosoftDocumentProperty com) => raw = com;
        internal MicrosoftDocumentProperty raw;

        public override int Release() => ComHelper.Release(raw);
        public override void FinalRelease() => ComHelper.FinalRelease(raw);
        protected override void DidDispose()
        {
            raw = default!;
            base.DidDispose();
        }
    }
}
