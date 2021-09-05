using System.Runtime.Versioning;

namespace Fxcel.Core.Interop
{
    using MicrosoftFileExportConverters = Microsoft.Office.Interop.Excel.FileExportConverters;

    [SupportedOSPlatform("windows")]
    public class XlFileExportConverters : XlComObject
    {
        public XlFileExportConverters(MicrosoftFileExportConverters converters) : base(converters) { }
        private MicrosoftFileExportConverters raw => (MicrosoftFileExportConverters)_raw;
    }
}
