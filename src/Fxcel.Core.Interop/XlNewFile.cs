using System.Runtime.Versioning;
using MicrosoftNewFile = Microsoft.Office.Core.NewFile;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public readonly ref struct XlNewFile
    {
        internal readonly MicrosoftNewFile raw;
        public XlNewFile(MicrosoftNewFile file) => raw = file;
    }
}
