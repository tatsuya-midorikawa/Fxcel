using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;
using MicrosoftUsedObjects = Microsoft.Office.Interop.Excel.UsedObjects;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public readonly ref struct XlUsedObjects
    {
        internal readonly MicrosoftUsedObjects raw;
        public XlUsedObjects(MicrosoftUsedObjects obj) => raw = obj;

        public int Release() => ComHelper.Release(raw);
    }
}
