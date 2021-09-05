using System.Runtime.Versioning;
using MicrosoftDocumentProperty = Microsoft.Office.Core.DocumentProperty;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public readonly struct XlDocumentProperty
    {
        internal readonly MicrosoftDocumentProperty raw;
        public XlDocumentProperty(MicrosoftDocumentProperty speach) => raw = speach;

        public int Release() => ComHelper.Release(raw);
    }
}
