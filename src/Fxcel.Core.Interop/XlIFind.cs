using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;

namespace Fxcel.Core.Interop
{
    using MicrosoftIFind = Microsoft.Office.Core.IFind;

    [SupportedOSPlatform("windows")]
    public sealed class XlIFind : XlComObject
    {
        internal XlIFind(MicrosoftIFind com) => raw = com;
        internal MicrosoftIFind raw;

        public override int Release() => ComHelper.Release(raw);
        public override void ForceRelease() => ComHelper.FinalRelease(raw);
        protected override void DidDispose()
        {
            raw = default!;
            base.DidDispose();
        }
    }
}
