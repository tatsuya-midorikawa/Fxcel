using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;

namespace Fxcel.Core.Interop
{
    using MicrosoftDefaultWebOptions = Microsoft.Office.Interop.Excel.DefaultWebOptions;

    [SupportedOSPlatform("windows")]
    public sealed class XlDefaultWebOptions : XlComObject
    {
        internal XlDefaultWebOptions(MicrosoftDefaultWebOptions com) => raw = com;
        internal MicrosoftDefaultWebOptions raw;

        public override int Release() => ComHelper.Release(raw);
        public override void FinalRelease() => ComHelper.FinalRelease(raw);
        protected override void DidDispose()
        {
            raw = default!;
            base.DidDispose();
        }
    }
}
