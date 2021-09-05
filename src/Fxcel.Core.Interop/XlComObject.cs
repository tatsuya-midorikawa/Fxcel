using System;
using System.Collections.Generic;
using System.Runtime.Versioning;
using Fxcel.Core.Interop.Common;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public abstract class XlComObject : IDisposable, IComObject
    {
        private readonly List<IComObject> _garbage = new();
        private bool _disposed= false;
        protected object _raw;

        protected XlComObject(object com) => _raw = com;

        private void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    // release managed objects
                    WillDispose();

                    for (var i = 0; i < _garbage.Count; i++)
                    {
                        try { _garbage[i]?.FinalRelease(); } 
                        catch { /* ignore */ }
                        finally { _garbage[i] = default!; }
                    }

                    OnDisposing();
                    
                    FinalRelease();
                    _raw = default!;

                    DidDispose();
                }

                // release unmanaged objects

                // update status
                _disposed = true;
            }
        }

        public void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }

        internal V ManageCom<V>(V target) where V : IComObject
        {
            _garbage.Add(target);
            return target;
        }
        protected virtual void WillDispose() { }
        protected virtual void OnDisposing() { }
        protected virtual void DidDispose() { }
        public int Release() => ComHelper.Release(_raw);
        public void FinalRelease() => ComHelper.FinalRelease(_raw);
    }
}
