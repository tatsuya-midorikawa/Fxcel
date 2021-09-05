using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;

namespace Fxcel.Core.Interop
{
    using MicrosoftFileDialog = Microsoft.Office.Core.FileDialog;

    [SupportedOSPlatform("windows")]
    public sealed class XlFileDialog : XlComObject
    {
        internal XlFileDialog(MicrosoftFileDialog com) => raw = com;
        internal MicrosoftFileDialog raw;

        public override int Release() => ComHelper.Release(raw);
        public override void FinalRelease() => ComHelper.FinalRelease(raw);
        protected override void DidDispose()
        {
            raw = default!;
            base.DidDispose();
        }
    }
}
