using System;

namespace Fxcel.Core.Interop.Common
{
    public interface IComObject : IDisposable
    {
        int Release();
        void ForceRelease();
    }
}
