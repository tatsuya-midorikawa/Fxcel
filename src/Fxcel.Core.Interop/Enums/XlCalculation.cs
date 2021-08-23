using Microsoft.Office.Interop.Excel;
namespace Fxcel.Core.Interop
{
    [System.Runtime.Versioning.SupportedOSPlatform("windows")]
    public enum XlCalculation
    {
        Automatic = Constants.xlAutomatic,
        Manual = Constants.xlManual,
        Semiautomatic = Constants.xlSemiautomatic
    }
}
