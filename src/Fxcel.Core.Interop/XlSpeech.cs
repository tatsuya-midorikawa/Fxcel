using System.Runtime.Versioning;
using MicrosoftSpeech = Microsoft.Office.Interop.Excel.Speech;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public readonly ref struct XlSpeech
    {
        internal readonly MicrosoftSpeech raw;
        public XlSpeech(MicrosoftSpeech speach) => raw = speach;

        public int Release() => ComHelper.Release(raw);
    }
}
