using System.Runtime.Versioning;

namespace Fxcel.Core.Interop
{
    using MicrosoftSpeech = Microsoft.Office.Interop.Excel.Speech;

    [SupportedOSPlatform("windows")]
    public sealed class XlSpeech : XlComObject
    {
        internal XlSpeech(MicrosoftSpeech com) => raw = com;
        internal MicrosoftSpeech raw;

        public override int Release() => ComHelper.Release(raw);
        public override void FinalRelease() => ComHelper.FinalRelease(raw);
        protected override void DidDispose()
        {
            raw = default!;
            base.DidDispose();
        }
    }
}
