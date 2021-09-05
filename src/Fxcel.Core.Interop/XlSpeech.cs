using System.Runtime.Versioning;

namespace Fxcel.Core.Interop
{
    using MicrosoftSpeech = Microsoft.Office.Interop.Excel.Speech;

    [SupportedOSPlatform("windows")]
    public class XlSpeech : XlComObject
    {
        public XlSpeech(MicrosoftSpeech speach) : base(speach) { }
        private MicrosoftSpeech raw => (MicrosoftSpeech)_raw;
    }
}
