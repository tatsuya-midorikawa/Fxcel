using System;
using System.Runtime.CompilerServices;
using System.Runtime.Versioning;
using Fxcel.Core.Interop.Common;

namespace Fxcel.Core.Interop
{
    using MicrosoftIAssistance = Microsoft.Office.Core.IAssistance;

    [SupportedOSPlatform("windows")]
    public readonly struct XlIAssistance : IComObject
    {
        internal readonly MicrosoftIAssistance raw;
        private readonly ComCollector collector;
        private readonly bool disposed;

        internal XlIAssistance(MicrosoftIAssistance com)
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
