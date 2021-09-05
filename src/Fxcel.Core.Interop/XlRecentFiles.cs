using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;

namespace Fxcel.Core.Interop
{
    using MicrosoftRecentFiles = Microsoft.Office.Interop.Excel.RecentFiles;

    [SupportedOSPlatform("windows")]
    public class XlRecentFiles : XlComObject
    {
        public XlRecentFiles(MicrosoftRecentFiles recentFiles) : base(recentFiles) { }
        private MicrosoftRecentFiles raw => (MicrosoftRecentFiles)_raw;
    }
}
