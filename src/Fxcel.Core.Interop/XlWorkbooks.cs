using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Versioning;
using System.Runtime.InteropServices;
using Fxcel.Core.Interop.Common;
using System.Collections;

namespace Fxcel.Core.Interop
{
    using MicrosoftWorkbook = Microsoft.Office.Interop.Excel.Workbook;
    using MicrosoftWorkbooks = Microsoft.Office.Interop.Excel.Workbooks;

    [SupportedOSPlatform("windows")]
    public sealed class XlWorkbooks : XlComObject, IEnumerable<XlWorkbook>
    {
        internal XlWorkbooks(MicrosoftWorkbooks com) => raw = com;
        internal MicrosoftWorkbooks raw;

        public override int Release() => ComHelper.Release(raw);
        public override void ForceRelease() => ComHelper.FinalRelease(raw);
        protected override void DidDispose()
        {
            raw = default!;
            base.DidDispose();
        }

        public IEnumerator<XlWorkbook> GetEnumerator() =>
            raw.OfType<MicrosoftWorkbook>().Select(wb => ManageCom(new XlWorkbook(wb))).GetEnumerator();

        IEnumerator IEnumerable.GetEnumerator() =>
            raw.OfType<MicrosoftWorkbook>().Select(wb => ManageCom(new XlWorkbook(wb))).GetEnumerator();

        public XlWorkbook this[int index] => new(raw[index]);
        public XlWorkbook this[string name] => new(raw[name]);

        public XlApplication Application => ManageCom(new XlApplication(raw.Application));

        public XlWorkbook Add([Optional][In][MarshalAs(UnmanagedType.Struct)] string template) => 
            new(string.IsNullOrEmpty(template) ? raw.Add() : raw.Add(template));

        public void Close() => raw.Close();

    }
}
