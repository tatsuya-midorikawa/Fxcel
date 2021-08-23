using System.Runtime.Versioning;
using MicrosoftIAssistance = Microsoft.Office.Core.IAssistance;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public readonly ref struct XlIAssistance
    {
        internal readonly MicrosoftIAssistance raw;
        public XlIAssistance(MicrosoftIAssistance assistance) => raw = assistance;

        public int Release() => ComHelper.Release(raw);
    }
}
