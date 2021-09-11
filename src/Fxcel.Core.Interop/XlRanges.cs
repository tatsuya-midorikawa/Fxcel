using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Versioning;
using Fxcel.Core.Interop.Common;
using System.Runtime.CompilerServices;

namespace Fxcel.Core.Interop
{
    using MicrosoftRanges = Microsoft.Office.Interop.Excel.Ranges;

    [SupportedOSPlatform("windows")]
    public readonly struct XlRanges : IComObject
    {
        internal readonly MicrosoftRanges raw;
        private readonly ComCollector collector;
        private readonly bool disposed;

        internal XlRanges(MicrosoftRanges com)
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
