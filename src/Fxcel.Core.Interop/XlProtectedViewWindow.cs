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
    using MicrosoftProtectedViewWindow = Microsoft.Office.Interop.Excel.ProtectedViewWindow;

    [SupportedOSPlatform("windows")]
    public readonly struct XlProtectedViewWindow : IComObject
    {
        internal readonly MicrosoftProtectedViewWindow raw;
        private readonly ComCollector collector;
        private readonly bool disposed;

        internal XlProtectedViewWindow(MicrosoftProtectedViewWindow com)
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
