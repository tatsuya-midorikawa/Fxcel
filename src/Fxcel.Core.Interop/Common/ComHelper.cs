using System.Runtime.InteropServices;
using System.Runtime.Versioning;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public static class ComHelper
    {
        public static int Release(object com) => Marshal.ReleaseComObject(com);
        public static int FinalRelease(object com) => Marshal.FinalReleaseComObject(com);
    }
}
