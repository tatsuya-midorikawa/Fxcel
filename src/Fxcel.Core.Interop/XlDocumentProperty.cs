using System;
using System.Runtime.CompilerServices;
using System.Runtime.Versioning;
using Fxcel.Core.Interop.Common;

namespace Fxcel.Core.Interop
{
    using MicrosoftDocumentProperty = Microsoft.Office.Core.DocumentProperty;

    [SupportedOSPlatform("windows")]
    public readonly struct XlDocumentProperty : IComObject
    {
        internal readonly MicrosoftDocumentProperty raw;
        private readonly ComCollector collector;
        private readonly bool disposed;

        internal XlDocumentProperty(MicrosoftDocumentProperty com)
        {
            raw = com;
            collector = new();
            disposed = false;
        }

        public readonly void Dispose()
        {
            if (!disposed)
            {
                // release managed objects
                collector.Collect();
                ForceRelease();

                // update status
                Unsafe.AsRef(disposed) = true;
            }
            GC.SuppressFinalize(this);
        }

        public readonly int Release() => ComHelper.Release(raw);
        public readonly void ForceRelease() => ComHelper.FinalRelease(raw);
    }
}
