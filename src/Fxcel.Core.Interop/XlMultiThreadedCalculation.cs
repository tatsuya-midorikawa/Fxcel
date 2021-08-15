using System.Runtime.Versioning;
using MicrosoftMultiThreadedCalculation = Microsoft.Office.Interop.Excel.MultiThreadedCalculation;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public readonly ref struct XlMultiThreadedCalculation
    {
        internal readonly MicrosoftMultiThreadedCalculation raw;
        public XlMultiThreadedCalculation(MicrosoftMultiThreadedCalculation calculation) => raw = calculation;
    }
}
