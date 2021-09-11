using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;
using Fxcel.Core.Interop.Common;

namespace Fxcel.Core.Interop
{
    using MicrosoftProtectedViewWindows = Microsoft.Office.Interop.Excel.ProtectedViewWindows;

    [SupportedOSPlatform("windows")]
    public readonly struct XlProtectedViewWindows : IComObject
    {
        internal readonly MicrosoftProtectedViewWindows raw;
        private readonly ComCollector collector;
        private readonly bool disposed;

        internal XlProtectedViewWindows(MicrosoftProtectedViewWindows com)
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
