using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Versioning;
using Fxcel.Core.Interop.Common;

namespace Fxcel.Core.Interop
{
    using MicrosoftWorkbook = Microsoft.Office.Interop.Excel.Workbook;
    using MicrosoftDocumentProperty = Microsoft.Office.Core.DocumentProperty;

    [SupportedOSPlatform("windows")]
    public class XlWorkbook : XlComObject
    {
        public XlWorkbook(MicrosoftWorkbook workbook) : base(workbook) { }
        private MicrosoftWorkbook raw => (MicrosoftWorkbook)_raw;

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
