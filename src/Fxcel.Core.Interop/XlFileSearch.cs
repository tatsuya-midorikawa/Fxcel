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
    public readonly ref struct XlFileSearch
    {
        internal readonly MicrosoftFileSearch raw;
        public XlFileSearch(MicrosoftFileSearch fileSearch) => raw = fileSearch;
    }
}
