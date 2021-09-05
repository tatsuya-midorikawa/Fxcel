using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Versioning;
using System.Runtime.InteropServices;
using Fxcel.Core.Interop.Common;
using System.Collections;
using System.Runtime.CompilerServices;

namespace Fxcel.Core.Interop
{
    using MicrosoftWorkbook = Microsoft.Office.Interop.Excel.Workbook;
    using MicrosoftWorkbooks = Microsoft.Office.Interop.Excel.Workbooks;

    [SupportedOSPlatform("windows")]
    public readonly struct XlWorkbooks : IComObject, IEnumerable<XlWorkbook>
    {
        internal readonly MicrosoftWorkbooks raw;
        private readonly bool disposed;
        private readonly ComCollector collector;

        internal XlWorkbooks(MicrosoftWorkbooks com)
        {
            raw = com;
            disposed = false;
            collector = new();
        }

        public readonly void Dispose()
        {
            if (!disposed)
            {
                // release managed objects
                collector.Collect();
                ForceRelease();

                // update status
                Unsafe.AsRef(disposed) = true;
            }
            GC.SuppressFinalize(this);
        }

        public readonly int Release() => ComHelper.Release(raw);
        public readonly void ForceRelease() => ComHelper.FinalRelease(raw);


        public readonly IEnumerator<XlWorkbook> GetEnumerator()
        {
            var c = collector;
            return raw.OfType<MicrosoftWorkbook>().Select(wb => c.Mark(new XlWorkbook(wb))).GetEnumerator();
        }

        readonly IEnumerator IEnumerable.GetEnumerator()
        {
            var c = collector;
            return raw.OfType<MicrosoftWorkbook>().Select(wb => c.Mark(new XlWorkbook(wb))).GetEnumerator();
        }

        public readonly XlWorkbook this[int index] => new(raw[index]);
        public readonly XlWorkbook this[string name] => new(raw[name]);

        public readonly XlApplication Application => collector.Mark(new XlApplication(raw.Application));
        public readonly XlCreator Creator => (XlCreator)raw.Creator;
        public readonly XlApplication Parent => collector.Mark(new XlApplication(raw.Parent));
        public readonly int Count => raw.Count;

        public readonly XlWorkbook Add([Optional][In][MarshalAs(UnmanagedType.Struct)] string template) => 
            new(string.IsNullOrEmpty(template) ? raw.Add() : raw.Add(template));

        public readonly void Close() => raw.Close();

    }
}
