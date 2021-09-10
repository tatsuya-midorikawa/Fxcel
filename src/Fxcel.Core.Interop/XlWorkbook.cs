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
    using MicrosoftDisplayDrawingObjects = Microsoft.Office.Interop.Excel.XlDisplayDrawingObjects;

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

        public XlCommandBars CommandBars => collector.Mark(new XlCommandBars(raw.CommandBars));
        public string Comments => raw.Comments;
        public XlSaveConflictResolution ConflictResolution => (XlSaveConflictResolution)raw.ConflictResolution;

        // TODO:
        /// <see href="https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel._workbook.container?view=excel-pia" />
        public XlObject Container => collector.Mark(new XlObject(raw.Container));
        public bool CreateBackup => raw.CreateBackup;
        // TODO:
        /// <see href="https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel._workbook.customdocumentproperties?view=excel-pia" />
        public XlObject CustomDocumentProperties => collector.Mark(new XlObject(raw.CustomDocumentProperties));
        public bool Date1904 { get => raw.Date1904; set => raw.Date1904 = value; }
        public XlSheets DialogSheets => collector.Mark(new XlSheets(raw.DialogSheets));
        public XlDisplayDrawingObjects DisplayDrawingObjects
        {
            get => (XlDisplayDrawingObjects)raw.DisplayDrawingObjects;
            set => raw.DisplayDrawingObjects = (MicrosoftDisplayDrawingObjects)value;
        }
        public XlFileFormat FileFormat => (XlFileFormat)raw.FileFormat;
        public string FullName => raw.FullName;
        public bool HasMailer { get => raw.HasMailer; set => raw.HasMailer = value; }
        public bool HasPassword => raw.HasPassword;
        public bool HasRoutingSlip { get => raw.HasRoutingSlip; set => raw.HasRoutingSlip = value; }
        public bool IsAddin { get => raw.IsAddin; set => raw.IsAddin = value; }
        public string Keywords { get => raw.Keywords; set => raw.Keywords = value; }
        public XlMailer Mailer => collector.Mark(new XlMailer(raw.Mailer));
        public XlSheets Modules => collector.Mark(new XlSheets(raw.Modules));
        public bool MultiUserEditing => raw.MultiUserEditing;
        public string Name => raw.Name;
        public XlNames Names => collector.Mark(new XlNames(raw.Names));
        public string OnSave { get => raw.OnSave; set => raw.OnSave = value; }
        public string OnSheetActivate { get => raw.OnSheetActivate; set => raw.OnSheetActivate = value; }
        public string OnSheetDeactivate { get => raw.OnSheetDeactivate; set => raw.OnSheetDeactivate = value; }
        public string Path => raw.Path;
        public bool PersonalViewListSettings { get => raw.PersonalViewListSettings; set => raw.PersonalViewListSettings = value; }
        public bool PersonalViewPrintSettings { get => raw.PersonalViewPrintSettings; set => raw.PersonalViewPrintSettings = value; }
        public bool PrecisionAsDisplayed { get => raw.PrecisionAsDisplayed; set => raw.PrecisionAsDisplayed = value; }
        public bool ProtectStructure => raw.ProtectStructure;
        public bool ProtectWindows => raw.ProtectWindows;
        public bool ReadOnly => raw.ReadOnly;
        public int RevisionNumber => raw.RevisionNumber;
        public bool Routed => raw.Routed;
        public XlRoutingSlip RoutingSlip => collector.Mark(new XlRoutingSlip(raw.RoutingSlip));
        public bool Saved { get => raw.Saved; set => raw.Saved = value; }
        public bool SaveLinkValues { get => raw.SaveLinkValues; set => raw.SaveLinkValues = value; }
        public XlSheets Sheets => collector.Mark(new XlSheets(raw.Sheets));
        public bool ShowConflictHistory { get => raw.ShowConflictHistory; set => raw.ShowConflictHistory = value; }
        public XlStyles Styles => collector.Mark(new XlStyles(raw.Styles));
        public string Subject { get => raw.Subject; set => raw.Subject = value; }
        public string Title { get => raw.Title; set => raw.Title = value; }
        public bool UpdateRemoteReferences { get => raw.UpdateRemoteReferences; set => raw.UpdateRemoteReferences = value; }
        public bool UserControl { get => raw.UserControl; set => raw.UserControl = value; }
        // TODO:
        /// <see href="https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel._workbook.userstatus?view=excel-pia" />
        public XlObject UserStatus => collector.Mark(new XlObject(raw.UserStatus));
        public XlCustomViews CustomViews => collector.Mark(new XlCustomViews(raw.CustomViews));
        public XlWindows Windows => collector.Mark(new XlWindows(raw.Windows));
    }
}
