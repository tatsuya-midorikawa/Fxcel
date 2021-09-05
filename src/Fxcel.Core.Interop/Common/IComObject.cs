namespace Fxcel.Core.Interop.Common
{
    public interface IComObject
    {
        int Release();
        void FinalRelease();
    }
}
