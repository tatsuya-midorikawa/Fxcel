using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Versioning;
using Fxcel.Core.Interop.Common;
using System.Runtime.CompilerServices;

namespace Fxcel.Core.Interop
{
    using MicrosoftWorkbook = Microsoft.Office.Interop.Excel.Workbook;
    using MicrosoftDocumentProperty = Microsoft.Office.Core.DocumentProperty;

    [SupportedOSPlatform("windows")]
    public readonly struct XlWorkbook : IComObject
    {
        internal readonly MicrosoftWorkbook raw;
        private readonly ComCollector collector;
        private readonly bool disposed;

        internal XlWorkbook(MicrosoftWorkbook com)
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
            GC.SuppressFinalize(this);
        }

        public readonly int Release() => ComHelper.Release(raw);
        public readonly void ForceRelease() => ComHelper.FinalRelease(raw);

        public readonly XlApplication Application => collector.Mark(new XlApplication(raw.Application));
        public readonly XlCreator Creator => (XlCreator)raw.Creator;
        public readonly XlApplication Parent => collector.Mark(new XlApplication(raw.Parent));
        public readonly bool AcceptLabelsInFormulas
        {
            get => raw.AcceptLabelsInFormulas;
            set => raw.AcceptLabelsInFormulas = value;
        }
        public readonly XlChart ActiveChart => collector.Mark(new XlChart(raw.ActiveChart));
        public readonly XlWorksheet ActiveSheet => collector.Mark(new XlWorksheet(raw.ActiveSheet));
        public readonly string Author
        {
            get => raw.Author;
            set => raw.Author = value;
        }
        public int AutoUpdateFrequency
        {
            get => raw.AutoUpdateFrequency;
            set => raw.AutoUpdateFrequency = value;
        }
        public bool AutoUpdateSaveChanges
        {
            get => raw.AutoUpdateSaveChanges;
            set => raw.AutoUpdateSaveChanges = value;
        }
        public int ChangeHistoryDuration
        {
            get => raw.ChangeHistoryDuration;
            set => raw.ChangeHistoryDuration = value;
        }
        public IEnumerable<XlDocumentProperty> BuiltinDocumentProperties
        {
            get
            {
                var c = collector;
                return ((IEnumerable<MicrosoftDocumentProperty>)raw.BuiltinDocumentProperties).Select(p => c.Mark(new XlDocumentProperty(p)));
            }
        }
        public XlSheets Charts => new(raw.Charts);
        public string CodeName => raw.CodeName;
        // TODO:
        /// <see href="https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel._workbook.colors?view=excel-pia" />
        public object Colors => raw.Colors;
    }
}
