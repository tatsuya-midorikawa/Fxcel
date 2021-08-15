using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;
using MicrosoftIFind = Microsoft.Office.Core.IFind;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public readonly ref struct XlIFind
    {
        internal readonly MicrosoftIFind raw;
        public XlIFind(MicrosoftIFind ifind) => raw = ifind;
    }
}
