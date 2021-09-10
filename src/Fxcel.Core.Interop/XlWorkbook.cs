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
    using MicrosoftXlUpdateLinks = Microsoft.Office.Interop.Excel.XlUpdateLinks;

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
        public XlSheets Worksheets => collector.Mark(new XlSheets(raw.Worksheets));
        public bool WriteReserved => raw.WriteReserved;
        public string WriteReservedBy => raw.WriteReservedBy;
        public XlSheets Excel4IntlMacroSheets => collector.Mark(new XlSheets(raw.Excel4IntlMacroSheets));
        public XlSheets Excel4MacroSheets => collector.Mark(new XlSheets(raw.Excel4MacroSheets));
        public bool TemplateRemoveExtData { get => raw.TemplateRemoveExtData; set => raw.TemplateRemoveExtData = value; }
        public bool HighlightChangesOnScreen { get => raw.HighlightChangesOnScreen; set => raw.HighlightChangesOnScreen = value; }
        public bool KeepChangeHistory { get => raw.KeepChangeHistory; set => raw.KeepChangeHistory = value; }
        public bool ListChangesOnNewSheet { get => raw.ListChangesOnNewSheet; set => raw.ListChangesOnNewSheet = value; }
        public bool IsInplace => raw.IsInplace;
        public XlPublishObjects PublishObjects => collector.Mark(new XlPublishObjects(raw.PublishObjects));
        public XlWebOptions WebOptions => collector.Mark(new XlWebOptions(raw.WebOptions));
        public XlHtmlProject HtmlProject => collector.Mark(new XlHtmlProject(raw.HTMLProject));
        public bool EnvelopeVisible { get => raw.EnvelopeVisible; set => raw.EnvelopeVisible = value; }
        public int CalculationVersion => raw.CalculationVersion;
        public bool VbaSigned => raw.VBASigned;
        public bool ShowPivotTableFieldList { get => raw.ShowPivotTableFieldList; set => raw.ShowPivotTableFieldList = value; }
        public XlUpdateLinks UpdateLinks { get => (XlUpdateLinks)raw.UpdateLinks; set => raw.UpdateLinks = (MicrosoftXlUpdateLinks)value; }
        public bool EnableAutoRecover { get => raw.EnableAutoRecover; set => raw.EnableAutoRecover = value; }
        public bool RemovePersonalInformation { get => raw.RemovePersonalInformation; set => raw.RemovePersonalInformation = value; }
        public string FullNameURLEncoded => raw.FullNameURLEncoded;
        public string Password { get => raw.Password; set => raw.Password = value; }
        public string WritePassword { get => raw.WritePassword; set => raw.WritePassword = value; }
        public string PasswordEncryptionProvider => raw.PasswordEncryptionProvider;
        public string PasswordEncryptionAlgorithm => raw.PasswordEncryptionAlgorithm;
        public int PasswordEncryptionKeyLength => raw.PasswordEncryptionKeyLength;
        public bool PasswordEncryptionFileProperties => raw.PasswordEncryptionFileProperties;
        public bool ReadOnlyRecommended { get => raw.ReadOnlyRecommended; set => raw.ReadOnlyRecommended = value; }
        public XlSmartTagOptions SmartTagOptions => collector.Mark(new XlSmartTagOptions(raw.SmartTagOptions));
        public XlPermission Permission => collector.Mark(new XlPermission(raw.Permission));
        public XlSharedWorkspace SharedWorkspace => collector.Mark(new XlSharedWorkspace(raw.SharedWorkspace));
        public XlSync Sync => collector.Mark(new XlSync(((Microsoft.Office.Interop.Excel._Workbook)raw).Sync));
        public XlXmlNamespaces XmlNamespaces => collector.Mark(new XlXmlNamespaces(raw.XmlNamespaces));
        public XlXmlMaps XmlMaps => collector.Mark(new XlXmlMaps(raw.XmlMaps));
        public XlSmartDocument SmartDocument => collector.Mark(new XlSmartDocument(raw.SmartDocument));
        public XlDocumentLibraryVersions DocumentLibraryVersions => collector.Mark(new XlDocumentLibraryVersions(raw.DocumentLibraryVersions));
        public bool InactiveListBorderVisible { get => raw.InactiveListBorderVisible; set => raw.InactiveListBorderVisible = value; }
        public bool DisplayInkComments { get => raw.DisplayInkComments; set => raw.DisplayInkComments = value; }
        public XlMetaProperties ContentTypeProperties => collector.Mark(new XlMetaProperties(raw.ContentTypeProperties));
        public XlConnections Connections => collector.Mark(new XlConnections(raw.Connections));
        public XlSignatureSet SignatureSet => collector.Mark(new XlSignatureSet(raw.Signatures));
        public XlServerPolicy ServerPolicy => collector.Mark(new XlServerPolicy(raw.ServerPolicy));
        public XlDocumentInspectors DocumentInspectors => collector.Mark(new XlDocumentInspectors(raw.DocumentInspectors));
        public XlServerViewableItems ServerViewableItems => collector.Mark(new XlServerViewableItems(raw.ServerViewableItems));
        public XlTableStyles TableStyles => collector.Mark(new XlTableStyles(raw.TableStyles));

        // TODO:
        /// <see href="https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel._workbook.defaulttablestyle?view=excel-pia" />
        public XlObject DefaultTableStyle => collector.Mark(new XlObject(raw.DefaultTableStyle));
        // TODO:
        /// <see href="https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel._workbook.defaultpivottablestyle?view=excel-pia" />
        public XlObject DefaultPivotTableStyle => collector.Mark(new XlObject(raw.DefaultPivotTableStyle));

        public bool CheckCompatibility { get => raw.CheckCompatibility; set => raw.CheckCompatibility = value; }
        public bool HasVBProject => raw.HasVBProject;

        public XlCustomXmlParts CustomXmlParts => collector.Mark(new XlCustomXmlParts(raw.CustomXMLParts));
        public bool Final { get => raw.Final; set => raw.Final = value; }
        public XlResearch Research => collector.Mark(new XlResearch(raw.Research));
        public XlOfficeTheme Theme => collector.Mark(new XlOfficeTheme(raw.Theme));
        public bool Excel8CompatibilityMode => raw.Excel8CompatibilityMode;
        public bool ConnectionsDisabled => raw.ConnectionsDisabled;
        public bool ShowPivotChartActiveFields { get => raw.ShowPivotChartActiveFields; set => raw.ShowPivotChartActiveFields = value; }
        public XlIconSets ThIconSetseme => collector.Mark(new XlIconSets(raw.IconSets));
        public string EncryptionProvider { get => raw.EncryptionProvider; set => raw.EncryptionProvider = value; }
        public bool DoNotPromptForConvert { get => raw.DoNotPromptForConvert; set => raw.DoNotPromptForConvert = value; }
        public bool ForceFullCalculation { get => raw.ForceFullCalculation; set => raw.ForceFullCalculation = value; }
        public XlSlicerCaches SlicerCaches => collector.Mark(new XlSlicerCaches(raw.SlicerCaches));
        public XlSlicer ActiveSlicer => collector.Mark(new XlSlicer(raw.ActiveSlicer));

        // TODO:
        /// <see href="https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel._workbook.defaultslicerstyle?view=excel-pia" />
        public XlObject DefaultSlicerStyle => collector.Mark(new XlObject(raw.DefaultSlicerStyle));

        public int AccuracyVersion { get => raw.AccuracyVersion; set => raw.AccuracyVersion = value; }
        public bool CaseSensitive => raw.CaseSensitive;
        public bool UseWholeCellCriteria => raw.UseWholeCellCriteria;
        public bool UseWildcards => raw.UseWildcards;

        // TODO:
        /// <see href="https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel._workbook.pivottables?view=excel-pia" />
        public XlObject PivotTables => collector.Mark(new XlObject(raw.PivotTables));

        public XlModel Model => collector.Mark(new XlModel(raw.Model));
        public bool ChartDataPointTrack { get => raw.ChartDataPointTrack; set => raw.ChartDataPointTrack = value; }

        // TODO:
        /// <see href="https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel._workbook.defaulttimelinestyle?view=excel-pia" />
        public XlObject DefaultTimelineStyle => collector.Mark(new XlObject(raw.DefaultTimelineStyle));

        public XlQueries Queries => collector.Mark(new XlQueries(raw.Queries));
        public string WorkIdentity { get => raw.WorkIdentity; set => raw.WorkIdentity = value; }
        public bool AutoSaveOn { get => raw.AutoSaveOn; set => raw.AutoSaveOn = value; }
        public XlSensitivityLabel SensitivityLabel => collector.Mark(new XlSensitivityLabel(raw.SensitivityLabel));
    }
}
