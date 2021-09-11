using System;
using System.Runtime.CompilerServices;
using System.Runtime.Versioning;
using Fxcel.Core.Interop.Common;

namespace Fxcel.Core.Interop
{
    using MicrosoftSmartArtLayouts = Microsoft.Office.Core.SmartArtLayouts;

    [SupportedOSPlatform("windows")]
    public readonly struct XlSmartArtLayouts : IComObject
    {
        internal readonly MicrosoftSmartArtLayouts raw;
        private readonly ComCollector collector;
        private readonly bool disposed;

        internal XlSmartArtLayouts(MicrosoftSmartArtLayouts com)
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
