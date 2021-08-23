using System.Runtime.Versioning;
using MicrosoftFileExportConverters = Microsoft.Office.Interop.Excel.FileExportConverters;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public readonly ref struct XlFileExportConverters
    {
        internal readonly MicrosoftFileExportConverters raw;
        public XlFileExportConverters(MicrosoftFileExportConverters converters) => raw = converters;

        public int Release() => ComHelper.Release(raw);
    }
}
