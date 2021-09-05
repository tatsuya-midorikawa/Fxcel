using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;
using MicrosoftFileSearch = Microsoft.Office.Core.FileSearch;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public class XlFileSearch : XlComObject
    {
        public XlFileSearch(MicrosoftFileSearch fileSearch) : base(fileSearch) { }
        private MicrosoftFileSearch raw => (MicrosoftFileSearch)_raw;
    }
}
