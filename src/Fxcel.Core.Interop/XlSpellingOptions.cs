using System;
using System.Runtime.CompilerServices;
using System.Runtime.Versioning;
using Fxcel.Core.Interop.Common;

namespace Fxcel.Core.Interop
{
    using MicrosoftSpellingOptions = Microsoft.Office.Interop.Excel.SpellingOptions;

    [SupportedOSPlatform("windows")]
    public readonly struct XlSpellingOptions : IComObject
    {
        internal readonly MicrosoftSpellingOptions raw;
        private readonly ComCollector collector;
        private readonly bool disposed;

        internal XlSpellingOptions(MicrosoftSpellingOptions com)
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
