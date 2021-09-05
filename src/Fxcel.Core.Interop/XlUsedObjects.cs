using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;

namespace Fxcel.Core.Interop
{
    using MicrosoftUsedObjects = Microsoft.Office.Interop.Excel.UsedObjects;

    [SupportedOSPlatform("windows")]
    public sealed class XlUsedObjects : XlComObject
    {
        internal XlUsedObjects(MicrosoftUsedObjects com) => raw = com;
        internal MicrosoftUsedObjects raw;

        public override int Release() => ComHelper.Release(raw);
        public override void FinalRelease() => ComHelper.FinalRelease(raw);
        protected override void DidDispose()
        {
            raw = default!;
            base.DidDispose();
        }
    }
}
