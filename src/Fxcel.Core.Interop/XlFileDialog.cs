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
    public class XlFileDialog : XlComObject
    {
        public XlFileDialog(MicrosoftFileDialog dialog) : base(dialog) { }
        private MicrosoftFileDialog raw => (MicrosoftFileDialog)_raw;
    }
}
