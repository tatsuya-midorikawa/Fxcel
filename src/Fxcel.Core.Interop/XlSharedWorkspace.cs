﻿using System;
using System.Runtime.CompilerServices;
using System.Runtime.Versioning;
using Fxcel.Core.Interop.Common;

namespace Fxcel.Core.Interop
{
    using MicrosoftSharedWorkspace = Microsoft.Office.Core.SharedWorkspace;

    [SupportedOSPlatform("windows")]
    public readonly struct XlSharedWorkspace : IComObject
    {
        internal readonly MicrosoftSharedWorkspace raw;
        private readonly ComCollector collector;
        private readonly bool disposed;

        internal XlSharedWorkspace(MicrosoftSharedWorkspace com)
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
