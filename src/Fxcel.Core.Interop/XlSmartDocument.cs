using System;
using System.Runtime.CompilerServices;
using System.Runtime.Versioning;
using Fxcel.Core.Interop.Common;

namespace Fxcel.Core.Interop
{
    using MicrosoftSmartDocument = Microsoft.Office.Core.SmartDocument;

    [SupportedOSPlatform("windows")]
    public readonly struct XlSmartDocument : IComObject
    {
        internal readonly MicrosoftSmartDocument raw;
        private readonly ComCollector collector;
        private readonly bool disposed;

        internal XlSmartDocument(MicrosoftSmartDocument com)
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
        }

        public readonly int Release() => ComHelper.Release(raw);
        public readonly void ForceRelease() => ComHelper.FinalRelease(raw);
    }
}
