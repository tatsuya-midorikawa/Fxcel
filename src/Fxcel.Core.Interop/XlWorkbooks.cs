﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Versioning;
using System.Runtime.InteropServices;
using Fxcel.Core.Interop.Common;
using System.Collections;
using System.Runtime.CompilerServices;
using System.Diagnostics;

namespace Fxcel.Core.Interop
{
    using MicrosoftWorkbook = Microsoft.Office.Interop.Excel.Workbook;
    using MicrosoftWorkbooks = Microsoft.Office.Interop.Excel.Workbooks;

    [SupportedOSPlatform("windows")]
    public readonly struct XlWorkbooks : IComObject, IEnumerable<XlWorkbook>
    {
        internal readonly MicrosoftWorkbooks raw;
        private readonly ComCollector collector;
        private readonly bool disposed;

        internal XlWorkbooks(MicrosoftWorkbooks com)
        {
            raw = com;
            collector = new();
            disposed = false;
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
        }

        public readonly int Release() => ComHelper.Release(raw);
        public readonly void ForceRelease() => ComHelper.FinalRelease(raw);

        public readonly IEnumerator<XlWorkbook> GetEnumerator()
        {
            var collector = this.collector;
            return raw.OfType<MicrosoftWorkbook>().Select(wb => collector.Mark(new XlWorkbook(wb))).GetEnumerator();
        }

        readonly IEnumerator IEnumerable.GetEnumerator()
        {
            var collector = this.collector;
            return raw.OfType<MicrosoftWorkbook>().Select(wb => collector.Mark(new XlWorkbook(wb))).GetEnumerator();
        }

        public readonly XlWorkbook this[int index] => collector.Mark(new XlWorkbook(raw[index]));
        public readonly XlWorkbook this[string name] => collector.Mark(new XlWorkbook(raw[name]));

        public readonly XlApplication Application => collector.Mark(new XlApplication(raw.Application));
        public readonly XlCreator Creator => (XlCreator)raw.Creator;
        public readonly XlApplication Parent => collector.Mark(new XlApplication(raw.Parent));
        public readonly int Count => raw.Count;

        public readonly XlWorkbook Add() => collector.Mark(new XlWorkbook(raw.Add()));
        public readonly XlWorkbook Add(string template) => collector.Mark(new XlWorkbook(raw.Add(template)));

        public readonly void Close() => raw.Close();
        public readonly XlWorkbook Open(string filename) => collector.Mark(new XlWorkbook(raw.Open(Filename: filename)));
        public readonly XlWorkbook Open(string filename, string password) => collector.Mark(new XlWorkbook(raw.Open(Filename: filename, Password: password)));
        public readonly XlWorkbook Open(string filename, bool @readonly) => collector.Mark(new XlWorkbook(raw.Open(Filename: filename, ReadOnly: @readonly)));
        public readonly XlWorkbook Open(string filename, string password, bool @readonly) => collector.Mark(new XlWorkbook(raw.Open(Filename: filename, Password: password, ReadOnly: @readonly)));

    }
}
