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
    using MicrosoftQueries = Microsoft.Office.Interop.Excel.Queries;

    [SupportedOSPlatform("windows")]
    public readonly struct XlQueries : IComObject
    {
        internal readonly MicrosoftQueries raw;
        private readonly ComCollector collector;
        private readonly bool disposed;

        internal XlQueries(MicrosoftQueries com)
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
