using System.Runtime.InteropServices;
using System.Runtime.Versioning;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public static class ComHelper
    {
        public static int Release(object com) => Marshal.IsComObject(com) ? Marshal.ReleaseComObject(com) : 0;
        public static int FinalRelease(object com) => Marshal.IsComObject(com) ? Marshal.FinalReleaseComObject(com) : 0;
    }
}
