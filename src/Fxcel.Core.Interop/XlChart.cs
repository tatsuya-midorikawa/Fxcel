using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;

namespace Fxcel.Core.Interop
{
    using MicrosoftChart = Microsoft.Office.Interop.Excel.Chart;

    [SupportedOSPlatform("windows")]
    public sealed class XlChart : XlComObject
    {
        internal XlChart(MicrosoftChart com) => raw = com;
        internal MicrosoftChart raw;

        public override int Release() => ComHelper.Release(raw);
        public override void FinalRelease() => ComHelper.FinalRelease(raw);
        protected override void DidDispose()
        {
            raw = default!;
            base.DidDispose();
        }
    }
}
