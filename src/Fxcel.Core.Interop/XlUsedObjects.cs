using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;

namespace Fxcel.Core.Interop
{
    using MicrosoftUsedObjects = Microsoft.Office.Interop.Excel.UsedObjects;

    [SupportedOSPlatform("windows")]
    public class XlUsedObjects : XlComObject
    {
        public XlUsedObjects(MicrosoftUsedObjects obj) : base(obj) { }
        private MicrosoftUsedObjects raw => (MicrosoftUsedObjects)_raw;
    }
}
