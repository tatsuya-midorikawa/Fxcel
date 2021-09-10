namespace Fxcel.Core.Interop
{
    [System.Runtime.Versioning.SupportedOSPlatform("windows")]
    public enum XlSaveConflictResolution
    {
        LocalSessionChanges = 2,
        OtherSessionChanges = 3,
        UserResolution = 1
    }
}
