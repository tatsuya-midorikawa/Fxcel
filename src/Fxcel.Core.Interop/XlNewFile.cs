using System.Runtime.Versioning;

namespace Fxcel.Core.Interop
{
    using MicrosoftNewFile = Microsoft.Office.Core.NewFile;

    [SupportedOSPlatform("windows")]
    public class XlNewFile : XlComObject
    {
        public XlNewFile(MicrosoftNewFile file) : base(file) { }
        private MicrosoftNewFile raw => (MicrosoftNewFile)_raw;
    }
}
