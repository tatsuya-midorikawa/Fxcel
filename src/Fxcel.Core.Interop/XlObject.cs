using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Runtime.Versioning;
using Fxcel.Core.Interop.Common;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public readonly struct XlObject : IComObject
    {
        internal readonly object raw;
        private readonly ComCollector collector;
        private readonly bool disposed;

        internal XlObject(object com)
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

        public bool TryGetValue<T>(out T value)
        {
            var isValid = typeof(T) == raw?.GetType();
#pragma warning disable CS8600, CS8601
            value = isValid ? (T)raw : default!;
#pragma warning restore CS8600, CS8601
            return isValid;
        }
        public T GetValue<T>() => (T)raw;
        public new Type GetType() => raw is null ? typeof(XlObject) : raw.GetType();
    }
}
