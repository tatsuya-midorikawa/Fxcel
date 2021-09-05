using System.Runtime.Versioning;

namespace Fxcel.Core.Interop
{
    using MicrosoftFileExportConverters = Microsoft.Office.Interop.Excel.FileExportConverters;

    [SupportedOSPlatform("windows")]
    public sealed class XlFileExportConverters : XlComObject
    {
        internal XlFileExportConverters(MicrosoftFileExportConverters com) => raw = com;
        internal MicrosoftFileExportConverters raw;

        public override int Release() => ComHelper.Release(raw);
        public override void ForceRelease() => ComHelper.FinalRelease(raw);
        protected override void DidDispose()
        {
            raw = default!;
            base.DidDispose();
        }
    }
}
