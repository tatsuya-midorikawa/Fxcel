using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;
using MicrosoftRecentFiles = Microsoft.Office.Interop.Excel.RecentFiles;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public readonly ref struct XlRecentFiles
    {
        internal readonly MicrosoftRecentFiles raw;
        public XlRecentFiles(MicrosoftRecentFiles recentFiles) => raw = recentFiles;

        public int Release() => ComHelper.Release(raw);
    }
}
