using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;
using MicrosoftFileDialog = Microsoft.Office.Core.FileDialog;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public readonly ref struct XlFileDialog
    {
        internal readonly MicrosoftFileDialog raw;
        public XlFileDialog(MicrosoftFileDialog dialog) => raw = dialog;

        public int Release() => ComHelper.Release(raw);
    }
}
