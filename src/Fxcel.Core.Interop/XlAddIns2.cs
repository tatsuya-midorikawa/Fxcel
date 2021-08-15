using MicrosoftAddIns2 = Microsoft.Office.Interop.Excel.AddIns2;
using System.Runtime.Versioning;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public readonly ref struct XlAddIns2
    {
        internal readonly MicrosoftAddIns2 raw;
        public XlAddIns2(MicrosoftAddIns2 addins) => raw = addins;
    }
}
