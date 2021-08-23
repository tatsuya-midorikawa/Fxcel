using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;
using MicrosoftChart = Microsoft.Office.Interop.Excel.Chart;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public readonly ref struct XlChart
    {
        internal readonly MicrosoftChart raw;
        public XlChart(MicrosoftChart chart) => raw = chart;

        public int Release() => ComHelper.Release(raw);
    }
}
