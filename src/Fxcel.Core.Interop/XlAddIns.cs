using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Versioning;
using System.Runtime.CompilerServices;

using Fxcel.Core.Interop.Common;

namespace Fxcel.Core.Interop
{
    using MicrosoftAddIns = Microsoft.Office.Interop.Excel.AddIns;

    [SupportedOSPlatform("windows")]
    public readonly struct XlAddIns : IComObject
    {
        internal readonly MicrosoftAddIns raw;
        private readonly ComCollector collector;
        private readonly bool disposed;

        internal XlAddIns(MicrosoftAddIns com)
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
