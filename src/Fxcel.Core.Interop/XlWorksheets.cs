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
    using MicrosoftWorksheets = Microsoft.Office.Interop.Excel.Worksheets;

    [SupportedOSPlatform("windows")]
    public readonly struct XlWorksheets : IComObject
    {
        internal readonly MicrosoftWorksheets raw;
        private readonly ComCollector collector;
        private readonly bool disposed;

        internal XlWorksheets(MicrosoftWorksheets com)
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
