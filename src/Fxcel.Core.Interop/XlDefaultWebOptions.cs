using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;
using MicrosoftDefaultWebOptions = Microsoft.Office.Interop.Excel.DefaultWebOptions;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public class XlDefaultWebOptions : XlComObject
    {
        public XlDefaultWebOptions(MicrosoftDefaultWebOptions options) : base(options) { }
        private MicrosoftDefaultWebOptions raw => (MicrosoftDefaultWebOptions)_raw;
    }
}
