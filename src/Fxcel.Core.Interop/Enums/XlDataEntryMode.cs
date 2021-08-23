using Microsoft.Office.Interop.Excel;
namespace Fxcel.Core.Interop
{
    [System.Runtime.Versioning.SupportedOSPlatform("windows")]
    public enum XlDataEntryMode
    {
        On = Constants.xlOn,
        Off = Constants.xlOff,
        Strict = Constants.xlStrict
    }
}
