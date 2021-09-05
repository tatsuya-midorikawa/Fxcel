using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;

namespace Fxcel.Core.Interop
{
    using MicrosoftIFind = Microsoft.Office.Core.IFind;

    [SupportedOSPlatform("windows")]
    public class XlIFind : XlComObject
    {
        public XlIFind(MicrosoftIFind ifind) : base(ifind) { }
        private MicrosoftIFind raw => (MicrosoftIFind)_raw;
    }
}
