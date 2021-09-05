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
    public readonly struct XlWorkbook : IDisposable, IComObject
    {
        internal readonly MicrosoftWorkbook raw;
        private readonly bool disposed;
        private readonly ComCollector collector;

        internal XlWorkbook(MicrosoftWorkbook com)
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

        public int Release() => ComHelper.Release(raw);
        public void ForceRelease() => ComHelper.FinalRelease(raw);

        public XlApplication Application => new(raw.Application);
        public XlCreator Creator => (XlCreator)raw.Creator;
        public XlApplication Parent => new(raw.Parent);
        public bool AcceptLabelsInFormulas
        {
            get => raw.AcceptLabelsInFormulas;
            set => raw.AcceptLabelsInFormulas = value;
        }
        public XlChart ActiveChart => new(raw.ActiveChart);
        public XlWorksheet ActiveSheet => new(raw.ActiveSheet);
        public string Author
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
        public IEnumerable<XlDocumentProperty> BuiltinDocumentProperties => ((IEnumerable<MicrosoftDocumentProperty>)raw.BuiltinDocumentProperties).Select(p => new XlDocumentProperty(p));
        public XlSheets Charts => new(raw.Charts);
        public string CodeName => raw.CodeName;
        // TODO:
        /// <see href="https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel._workbook.colors?view=excel-pia" />
        public object Colors => raw.Colors;
    }
}
