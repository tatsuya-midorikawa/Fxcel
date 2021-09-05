using System.Runtime.Versioning;

namespace Fxcel.Core.Interop
{
    using MicrosoftIAssistance = Microsoft.Office.Core.IAssistance;

    [SupportedOSPlatform("windows")]
    public class XlIAssistance : XlComObject
    {
        public XlIAssistance(MicrosoftIAssistance assistance) : base(assistance) { }
        private MicrosoftIAssistance raw => (MicrosoftIAssistance)_raw;
    }
}
