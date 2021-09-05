using System.Runtime.Versioning;

namespace Fxcel.Core.Interop
{
    using MicrosoftAutoRecover = Microsoft.Office.Interop.Excel.AutoRecover;

    [SupportedOSPlatform("windows")]
    public class XlAutoRecover : XlComObject
    {
        public XlAutoRecover(MicrosoftAutoRecover recover) : base(recover) { }
        private MicrosoftAutoRecover raw => (MicrosoftAutoRecover)_raw;
    }
}
