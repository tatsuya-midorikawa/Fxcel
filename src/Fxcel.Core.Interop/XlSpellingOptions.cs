using System.Runtime.Versioning;

namespace Fxcel.Core.Interop
{
    using MicrosoftSpellingOptions = Microsoft.Office.Interop.Excel.SpellingOptions;

    [SupportedOSPlatform("windows")]
    public sealed class XlSpellingOptions : XlComObject
    {
        internal XlSpellingOptions(MicrosoftSpellingOptions com) => raw = com;
        internal MicrosoftSpellingOptions raw;

        public override int Release() => ComHelper.Release(raw);
        public override void FinalRelease() => ComHelper.FinalRelease(raw);
        protected override void DidDispose()
        {
            raw = default!;
            base.DidDispose();
        }
    }
}
