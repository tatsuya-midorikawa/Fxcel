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
            var isValid = typeof(T) == raw.GetType();
            value = isValid ? (T)raw : default!;
            return isValid;
        }
        public T GetValue<T>() => (T)raw;
        public new Type GetType() => raw.GetType();

        public override int Release() => ComHelper.Release(raw);
        public override void ForceRelease() => ComHelper.FinalRelease(raw);
        protected override void DidDispose()
        {
            raw = default!;
            base.DidDispose();
        }

    }
}
