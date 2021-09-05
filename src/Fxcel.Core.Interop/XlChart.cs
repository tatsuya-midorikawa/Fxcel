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
    public class XlChart : XlComObject
    {
        public XlChart(MicrosoftChart chart) : base(chart) { }
        private MicrosoftChart raw => (MicrosoftChart)_raw;
    }
}
