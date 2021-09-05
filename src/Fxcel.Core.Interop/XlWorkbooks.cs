using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Versioning;
using System.Runtime.InteropServices;
using Fxcel.Core.Interop.Common;

namespace Fxcel.Core.Interop
{
    using MicrosoftWorkbooks = Microsoft.Office.Interop.Excel.Workbooks;

    [SupportedOSPlatform("windows")]
    public class XlWorkbooks : XlComObject
    {
        public XlWorkbooks(MicrosoftWorkbooks workbooks) : base(workbooks) { }
        private MicrosoftWorkbooks raw => (MicrosoftWorkbooks)_raw;

        public XlWorkbook this[int index] => new(raw[index]);
        public XlWorkbook this[string name] => new(raw[name]);
        public XlWorkbook Add([Optional][In][MarshalAs(UnmanagedType.Struct)] string template) => 
            new(string.IsNullOrEmpty(template) ? raw.Add() : raw.Add(template));
    }
}
