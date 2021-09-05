using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Versioning;

namespace Fxcel.Core.Interop
{
    using MicrosoftXmlMap = Microsoft.Office.Interop.Excel.XmlMap;

    [SupportedOSPlatform("windows")]
    public sealed class XlXmlMap : XlComObject
    {
        internal XlXmlMap(MicrosoftXmlMap com) => raw = com;
        internal MicrosoftXmlMap raw;

        public override int Release() => ComHelper.Release(raw);
        public override void FinalRelease() => ComHelper.FinalRelease(raw);
        protected override void DidDispose()
        {
            raw = default!;
            base.DidDispose();
        }
    }
}
