using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;

namespace Fxcel.Core.Interop
{
    using MicrosoftFileSearch = Microsoft.Office.Core.FileSearch;

    [SupportedOSPlatform("windows")]
    public sealed class XlFileSearch : XlComObject
    {
        internal XlFileSearch(MicrosoftFileSearch com) => raw = com;
        internal MicrosoftFileSearch raw;

        public override int Release() => ComHelper.Release(raw);
        public override void ForceRelease() => ComHelper.FinalRelease(raw);
        protected override void DidDispose()
        {
            raw = default!;
            base.DidDispose();
        }
    }
}
