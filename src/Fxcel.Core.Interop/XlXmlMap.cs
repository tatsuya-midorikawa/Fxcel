using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MicrosoftXmlMap = Microsoft.Office.Interop.Excel.XmlMap;
using System.Runtime.Versioning;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public readonly ref struct XlXmlMap
    {
        internal readonly MicrosoftXmlMap raw;
        public XlXmlMap(MicrosoftXmlMap xmlmap) => raw = xmlmap;

        public int Release() => ComHelper.Release(raw);
    }
}
