using System.Runtime.Versioning;

namespace Fxcel.Core.Interop
{
    using MicrosoftMultiThreadedCalculation = Microsoft.Office.Interop.Excel.MultiThreadedCalculation;

    [SupportedOSPlatform("windows")]
    public class XlMultiThreadedCalculation : XlComObject
    {
        public XlMultiThreadedCalculation(MicrosoftMultiThreadedCalculation calculation) : base(calculation) { }
        private MicrosoftMultiThreadedCalculation raw => (MicrosoftMultiThreadedCalculation)_raw;
    }
}
