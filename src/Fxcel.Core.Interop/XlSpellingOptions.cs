using System.Runtime.Versioning;

namespace Fxcel.Core.Interop
{
    using MicrosoftSpellingOptions = Microsoft.Office.Interop.Excel.SpellingOptions;

    [SupportedOSPlatform("windows")]
    public class XlSpellingOptions : XlComObject
    {
        public XlSpellingOptions(MicrosoftSpellingOptions options) : base(options) { }
        private MicrosoftSpellingOptions raw => (MicrosoftSpellingOptions)_raw;
    }
}
