using System.Runtime.Versioning;
using MicrosoftAutoRecover = Microsoft.Office.Interop.Excel.AutoRecover;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public readonly ref struct XlAutoRecover
    {
        internal readonly MicrosoftAutoRecover raw;
        public XlAutoRecover(MicrosoftAutoRecover recover) => raw = recover;
    }
}
