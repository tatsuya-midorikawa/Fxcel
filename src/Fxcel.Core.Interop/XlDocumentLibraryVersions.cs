using System;
using System.Runtime.CompilerServices;
using System.Runtime.Versioning;
using Fxcel.Core.Interop.Common;

namespace Fxcel.Core.Interop
{
    using MicrosoftDocumentLibraryVersions = Microsoft.Office.Core.DocumentLibraryVersions;

    [SupportedOSPlatform("windows")]
    public readonly struct XlDocumentLibraryVersions : IComObject
    {
        internal readonly MicrosoftDocumentLibraryVersions raw;
        private readonly ComCollector collector;
        private readonly bool disposed;

        internal XlDocumentLibraryVersions(MicrosoftDocumentLibraryVersions com)
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
