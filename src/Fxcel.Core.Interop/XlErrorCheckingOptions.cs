using System.Runtime.Versioning;
using MicrosoftErrorCheckingOptions = Microsoft.Office.Interop.Excel.ErrorCheckingOptions;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public readonly ref struct XlErrorCheckingOptions
    {
        internal readonly MicrosoftErrorCheckingOptions raw;
        public XlErrorCheckingOptions(MicrosoftErrorCheckingOptions options) => raw = options;

        public int Release() => ComHelper.Release(raw);
    }
}
