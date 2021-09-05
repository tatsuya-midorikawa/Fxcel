using System;
using System.Collections.Generic;
using System.Runtime.Versioning;
using Fxcel.Core.Interop.Common;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public sealed class XlObject : XlComObject
    {
        internal XlObject(object com) => raw = com;
        internal object raw;
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

        public override int Release() => ComHelper.Release(raw);
        public override void ForceRelease() => ComHelper.FinalRelease(raw);
        protected override void DidDispose()
        {
            raw = default!;
            base.DidDispose();
        }

    }
}
