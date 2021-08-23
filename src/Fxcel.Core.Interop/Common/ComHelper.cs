using System.Runtime.InteropServices;
using System.Runtime.Versioning;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    internal static class ComHelper
    {
        internal static int Release(object com) => Marshal.IsComObject(com) ? Marshal.ReleaseComObject(com) : 0;
        internal static int FinalRelease(object com) => Marshal.IsComObject(com) ? Marshal.FinalReleaseComObject(com) : 0;
    }
}
