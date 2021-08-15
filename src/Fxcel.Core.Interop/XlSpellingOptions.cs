using System.Runtime.Versioning;
using MicrosoftSpellingOptions = Microsoft.Office.Interop.Excel.SpellingOptions;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public readonly ref struct XlSpellingOptions
    {
        internal readonly MicrosoftSpellingOptions raw;
        public XlSpellingOptions(MicrosoftSpellingOptions options) => raw = options;
    }
}
