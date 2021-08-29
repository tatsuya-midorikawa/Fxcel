namespace Fxcel.Core.Interop
{
    [System.Runtime.Versioning.SupportedOSPlatform("windows")]
    [System.Flags]
    public enum XlInputType
    {
        Formula     = 0,        // 0
        Number      = 1 << 0,   // 1
        String      = 1 << 1,   // 2
        Boolean     = 1 << 2,   // 4
        RangeObject = 1 << 3,   // 8
        Error       = 1 << 4,   // 16
        Array       = 1 << 6,   // 64
    }
}
