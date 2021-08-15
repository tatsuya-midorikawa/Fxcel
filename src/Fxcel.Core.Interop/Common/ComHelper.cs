using System.Runtime.InteropServices;
using System.Runtime.Versioning;

namespace Fxcel.Core.Interop
{
    public interface IComObject
    {
        int ComRelease();
    }

    [SupportedOSPlatform("windows")]
    internal static class ComHelper
    {
        internal static int Release(object com) => Marshal.ReleaseComObject(com);
    }
}
