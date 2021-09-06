using System;
using System.Runtime.CompilerServices;
using System.Runtime.Versioning;
using Fxcel.Core.Interop.Common;

namespace Fxcel.Core.Interop
{
    using MicrosoftAddIns2 = Microsoft.Office.Interop.Excel.AddIns2;

    [SupportedOSPlatform("windows")]
    public readonly struct XlAddIns2 : IComObject
    {
        internal readonly MicrosoftAddIns2 raw;
        private readonly ComCollector collector;
        private readonly bool disposed;

        internal XlAddIns2(MicrosoftAddIns2 com)
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
