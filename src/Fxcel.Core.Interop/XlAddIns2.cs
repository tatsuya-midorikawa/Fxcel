using System.Runtime.Versioning;

namespace Fxcel.Core.Interop
{
    using MicrosoftAddIns2 = Microsoft.Office.Interop.Excel.AddIns2;

    [SupportedOSPlatform("windows")]
    public class XlAddIns2 : XlComObject
    {
        public XlAddIns2(MicrosoftAddIns2 addins) : base(addins) { }
        private MicrosoftAddIns2 raw => (MicrosoftAddIns2)_raw;
    }
}
