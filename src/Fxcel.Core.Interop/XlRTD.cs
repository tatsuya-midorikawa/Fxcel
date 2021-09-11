using System;
using System.Runtime.CompilerServices;
using System.Runtime.Versioning;
using Fxcel.Core.Interop.Common;

namespace Fxcel.Core.Interop
{
    using MicrosoftRTD = Microsoft.Office.Interop.Excel.RTD;

    [SupportedOSPlatform("windows")]
    public readonly struct XlRTD : IComObject
    {
        internal readonly MicrosoftRTD raw;
        private readonly ComCollector collector;
        private readonly bool disposed;

        internal XlRTD(MicrosoftRTD com)
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
