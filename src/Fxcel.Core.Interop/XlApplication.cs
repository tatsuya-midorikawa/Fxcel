using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;

namespace Fxcel.Core.Interop
{
    using MicrosoftApplication = Microsoft.Office.Interop.Excel.Application;
    using MicrosoftWorksheet = Microsoft.Office.Interop.Excel.Worksheet;
    using MicrosoftRange = Microsoft.Office.Interop.Excel.Range;
    using MicrosoftXlFileFormat = Microsoft.Office.Interop.Excel.XlFileFormat;
    using MicrosoftXlReferenceStyle = Microsoft.Office.Interop.Excel.XlReferenceStyle;
    using MicrosoftXlDirection = Microsoft.Office.Interop.Excel.XlDirection;
    using MicrosoftXlEnableCancelKey = Microsoft.Office.Interop.Excel.XlEnableCancelKey;
    using MicrosoftXlCommentDisplayMode = Microsoft.Office.Interop.Excel.XlCommentDisplayMode;
    using MicrosoftXlCutCopyMode = Microsoft.Office.Interop.Excel.XlCutCopyMode;
    using MicrosoftXlMousePointer = Microsoft.Office.Interop.Excel.XlMousePointer;
    using MicrosoftXlCalculation = Microsoft.Office.Interop.Excel.XlCalculation;
    using MicrosoftXlWindowState = Microsoft.Office.Interop.Excel.XlWindowState;
    using MicrosoftMsoFeatureInstall = Microsoft.Office.Core.MsoFeatureInstall;
    using MicrosoftMsoAutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity;
    using MicrosoftMsoFileDialogType = Microsoft.Office.Core.MsoFileDialogType;
    using MicrosoftXlCalculationInterruptKey = Microsoft.Office.Interop.Excel.XlCalculationInterruptKey;

    [SupportedOSPlatform("windows")]
    public readonly struct XlApplication : IComObject
    {
        internal static readonly List<XlApplication> apps = new();
        internal readonly MicrosoftApplication raw;
        internal XlApplication(MicrosoftApplication excel) => raw = excel;

        public int ComRelease() => ComHelper.Release(raw);

        public XlApplication Application => new(raw.Application);
        public XlCreator Creator => (XlCreator)raw.Creator;
        public XlApplication Parent => new(raw.Parent);
        public XlRange ActiveCell => new(raw.ActiveCell);
        public XlChart ActiveChart => new(raw.ActiveChart);
        public XlDialogSheet ActiveDialog => new(raw.ActiveDialog);
        public XlMenuBar ActiveMenuBar => new(raw.ActiveMenuBar);
        public string ActivePrinter { get => raw.ActivePrinter; set => raw.ActivePrinter = value; }
        public XlWorksheet ActiveSheet => new((MicrosoftWorksheet)raw.ActiveSheet);
        public XlWindow ActiveWindow => new(raw.ActiveWindow);
        public XlWorkbook ActiveWorkbook => new(raw.ActiveWorkbook);
        public XlAddIns AddIns => new(raw.AddIns);
        public XlAssistant Assistant => new(raw.Assistant);
        public XlRange Cells => new(raw.Cells);
        public XlSheets Charts => new(raw.Charts);
        public XlRange Columns => new(raw.Columns);
        public int DDEAppReturnCode => raw.DDEAppReturnCode;
        public XlSheets DialogSheets => new(raw.DialogSheets);
        public XlMenuBars MenuBars => new(raw.MenuBars);
        public XlModules Modules => new(raw.Modules);
        public XlNames Names => new(raw.Names);
        public XlRange Rows => new(raw.Rows);
        public object Selection => raw.Selection;
        public XlSheets Sheets => new(raw.Sheets);
        public XlMenu ShortcutMenus(int index) => new(raw.ShortcutMenus[index]);
        public XlWorkbook ThisWorkbook => new(raw.ThisWorkbook);
        public XlToolbars Toolbars => new(raw.Toolbars);
        public XlWindows Windows => new(raw.Windows);
        public XlWorkbooks Workbooks => new(raw.Workbooks);
        public XlWorksheetFunction WorksheetFunction => new(raw.WorksheetFunction);
        public XlSheets Worksheets => new(raw.Worksheets);
        public XlSheets Excel4IntlMacroSheets => new(raw.Excel4IntlMacroSheets);
        public XlSheets Excel4MacroSheets => new(raw.Excel4MacroSheets);
        public bool AlertBeforeOverwriting { get => raw.AlertBeforeOverwriting; set => raw.AlertBeforeOverwriting = value; }
        public string AltStartupPath { get => raw.AltStartupPath; set => raw.AltStartupPath = value; }
        public bool AskToUpdateLinks { get => raw.AskToUpdateLinks; set => raw.AskToUpdateLinks = value; }
        public bool EnableAnimations { get => raw.EnableAnimations; set => raw.EnableAnimations = value; }
        public XlAutoCorrect AutoCorrect => new(raw.AutoCorrect);
        public int Build => raw.Build;
        public bool CalculateBeforeSave { get => raw.CalculateBeforeSave; set => raw.CalculateBeforeSave = value; }
        public XlCalculation Calculation { get => (XlCalculation)raw.Calculation; set => raw.Calculation = (MicrosoftXlCalculation)value; }
        public object Caller => raw.Caller;
        public bool CanPlaySounds => raw.CanPlaySounds;
        public bool CanRecordSounds => raw.CanRecordSounds;
        public string Caption { get => raw.Caption; set => raw.Caption = value; }
        public bool CellDragAndDrop { get => raw.CellDragAndDrop; set => raw.CellDragAndDrop = value; }
        public XlClipboardFormat[] ClipboardFormats => ((object[])raw.ClipboardFormats).Select(f => (XlClipboardFormat)f).ToArray();
        public bool DisplayClipboardWindow { get => raw.DisplayClipboardWindow; set => raw.DisplayClipboardWindow = value; }
        public bool ColorButtons { get => raw.ColorButtons; set => raw.ColorButtons = value; }
        public XlCommandUnderlines CommandUnderlines => (XlCommandUnderlines)raw.CommandUnderlines;
        public bool ConstrainNumeric { get => raw.ConstrainNumeric; set => raw.ConstrainNumeric = value; }
        public bool CopyObjectsWithCells { get => raw.CopyObjectsWithCells; set => raw.CopyObjectsWithCells = value; }
        public XlMousePointer Cursor { get => (XlMousePointer)raw.Cursor; set => raw.Cursor = (MicrosoftXlMousePointer)value; }
        public int CustomListCount => raw.CustomListCount;
        public XlCutCopyMode CutCopyMode { get => (XlCutCopyMode)raw.CutCopyMode; set => raw.CutCopyMode = (MicrosoftXlCutCopyMode)value; }
        public XlDataEntryMode DataEntryMode { get => (XlDataEntryMode)raw.DataEntryMode; set => raw.DataEntryMode = (int)value; }
        public string _Default => raw._Default;
        public string DefaultFilePath { get => raw.DefaultFilePath; set => raw.DefaultFilePath = value; }
        public XlDialogs Dialogs => new(raw.Dialogs);
        public bool DisplayAlerts { get => raw.DisplayAlerts; set => raw.DisplayAlerts = value; }
        public bool DisplayFormulaBar { get => raw.DisplayFormulaBar; set => raw.DisplayFormulaBar = value; }
        public bool DisplayFullScreen { get => raw.DisplayFullScreen; set => raw.DisplayFullScreen = value; }
        public bool DisplayNoteIndicator { get => raw.DisplayNoteIndicator; set => raw.DisplayNoteIndicator = value; }
        public XlCommentDisplayMode DisplayCommentIndicator { get => (XlCommentDisplayMode)raw.DisplayCommentIndicator; set => raw.DisplayCommentIndicator = (MicrosoftXlCommentDisplayMode)value; }
        public bool DisplayExcel4Menus { get => raw.DisplayExcel4Menus; set => raw.DisplayExcel4Menus = value; }
        public bool DisplayRecentFiles { get => raw.DisplayRecentFiles; set => raw.DisplayRecentFiles = value; }
        public bool DisplayScrollBars { get => raw.DisplayScrollBars; set => raw.DisplayScrollBars = value; }
        public bool DisplayStatusBar { get => raw.DisplayStatusBar; set => raw.DisplayStatusBar = value; }
        public bool EditDirectlyInCell { get => raw.EditDirectlyInCell; set => raw.EditDirectlyInCell = value; }
        public bool EnableAutoComplete { get => raw.EnableAutoComplete; set => raw.EnableAutoComplete = value; }
        public XlEnableCancelKey EnableCancelKey { get => (XlEnableCancelKey)raw.EnableCancelKey; set => raw.EnableCancelKey = (MicrosoftXlEnableCancelKey)value; }
        public bool EnableSound { get => raw.EnableSound; set => raw.EnableSound = value; }
        public bool EnableTipWizard { get => raw.EnableTipWizard; set => raw.EnableTipWizard = value; }
        public object FileConverters => raw.FileConverters;
        public XlFileSearch FileSearch => new(raw.FileSearch);
        public XlIFind FileFind => new(raw.FileFind);
        public bool FixedDecimal { get => raw.FixedDecimal; set => raw.FixedDecimal = value; }
        public int FixedDecimalPlaces { get => raw.FixedDecimalPlaces; set => raw.FixedDecimalPlaces = value; }
        public double Height { get => raw.Height; set => raw.Height = value; }
        public bool IgnoreRemoteRequests { get => raw.IgnoreRemoteRequests; set => raw.IgnoreRemoteRequests = value; }
        public bool Interactive { get => raw.Interactive; set => raw.Interactive = value; }
        public object International(XlApplicationInternational index) => raw.International[index];
        public bool Iteration { get => raw.Iteration; set => raw.Iteration = value; }
        public bool LargeButtons { get => raw.LargeButtons; set => raw.LargeButtons = value; }
        public double Left { get => raw.Left; set => raw.Left = value; }
        public string LibraryPath => raw.LibraryPath;
        public string MailSession => (string)raw.MailSession;
        public XlMailSystem MailSystem => (XlMailSystem)raw.MailSystem;
        public bool MathCoprocessorAvailable => raw.MathCoprocessorAvailable;
        public double MaxChange { get => raw.MaxChange; set => raw.MaxChange = value; }
        public int MaxIterations { get => raw.MaxIterations; set => raw.MaxIterations = value; }
        public int MemoryFree => raw.MemoryFree;
        public int MemoryTotal => raw.MemoryTotal;
        public int MemoryUsed => raw.MemoryUsed;
        public bool MouseAvailable => raw.MouseAvailable;
        public bool MoveAfterReturn { get => raw.MoveAfterReturn; set => raw.MoveAfterReturn = value; }
        public XlDirection MoveAfterReturnDirection { get => (XlDirection)raw.MoveAfterReturnDirection; set => raw.MoveAfterReturnDirection = (MicrosoftXlDirection)value; }
        public XlRecentFiles RecentFiles => new(raw.RecentFiles);
        public string Name => raw.Name;
        public string NetworkTemplatesPath => raw.NetworkTemplatesPath;
        public XlOdbcErrors OdbcErrors => new(raw.ODBCErrors);
        public int OdbcTimeout { get => raw.ODBCTimeout; set => raw.ODBCTimeout = value; }
        public string OnCalculate { get => raw.OnCalculate; set => raw.OnCalculate = value; }
        public string OnData { get => raw.OnData; set => raw.OnData = value; }
        public string OnDoubleClick { get => raw.OnDoubleClick; set => raw.OnDoubleClick = value; }
        public string OnEntry { get => raw.OnEntry; set => raw.OnEntry = value; }
        public string OnSheetActivate { get => raw.OnSheetActivate; set => raw.OnSheetActivate = value; }
        public string OnSheetDeactivate { get => raw.OnSheetDeactivate; set => raw.OnSheetDeactivate = value; }
        public string OnWindow { get => raw.OnWindow; set => raw.OnWindow = value; }
        public string OperatingSystem => raw.OperatingSystem;
        public string OrganizationName => raw.OrganizationName;
        public string Path => raw.Path;
        public string PathSeparator => raw.PathSeparator;
        public XlRange PreviousSelections(int index) => new((MicrosoftRange)raw.PreviousSelections[index]);
        public bool PivotTableSelection { get => raw.PivotTableSelection; set => raw.PivotTableSelection = value; }
        public bool PromptForSummaryInfo { get => raw.PromptForSummaryInfo; set => raw.PromptForSummaryInfo = value; }
        public bool RecordRelative => raw.RecordRelative;
        public XlReferenceStyle ReferenceStyle { get => (XlReferenceStyle)raw.ReferenceStyle; set => raw.ReferenceStyle = (MicrosoftXlReferenceStyle)value; }
        public object RegisteredFunctions => raw.RegisteredFunctions;
        public bool RollZoom { get => raw.RollZoom; set => raw.RollZoom = value; }
        public bool ScreenUpdating { get => raw.ScreenUpdating; set => raw.ScreenUpdating = value; }
        public int SheetsInNewWorkbook { get => raw.SheetsInNewWorkbook; set => raw.SheetsInNewWorkbook = value; }
        public bool ShowChartTipNames { get => raw.ShowChartTipNames; set => raw.ShowChartTipNames = value; }
        public bool ShowChartTipValues { get => raw.ShowChartTipValues; set => raw.ShowChartTipValues = value; }
        public string StandardFont { get => raw.StandardFont; set => raw.StandardFont = value; }
        public double StandardFontSize { get => raw.StandardFontSize; set => raw.StandardFontSize = value; }
        public string StartupPath => raw.StartupPath;
        public bool StatusBar { get => (bool)raw.StatusBar; set => raw.StatusBar = value; }
        public string TemplatesPath => raw.TemplatesPath;
        public bool ShowToolTips { get => raw.ShowToolTips; set => raw.ShowToolTips = value; }
        public double Top { get => raw.Top; set => raw.Top = value; }
        public XlFileFormat DefaultSaveFormat { get => (XlFileFormat)raw.DefaultSaveFormat; set => raw.DefaultSaveFormat = (MicrosoftXlFileFormat)value; }
        public string TransitionMenuKey { get => raw.TransitionMenuKey; set => raw.TransitionMenuKey = value; }
        public int TransitionMenuKeyAction { get => raw.TransitionMenuKeyAction; set => raw.TransitionMenuKeyAction = value; }
        public bool TransitionNavigKeys { get => raw.TransitionNavigKeys; set => raw.TransitionNavigKeys = value; }
        public double UsableHeight => raw.UsableHeight;
        public double UsableWidth => raw.UsableWidth;
        public bool UserControl { get => raw.UserControl; set => raw.UserControl = value; }
        public string UserName { get => raw.UserName; set => raw.UserName = value; }
        public string Value => raw.Value;
        public string Version => raw.Version;
        public bool Visible { get => raw.Visible; set => raw.Visible = value; }
        public double Width { get => raw.Width; set => raw.Width = value; }
        public bool WindowsForPens => raw.WindowsForPens;
        public XlWindowState WindowState { get => (XlWindowState)raw.WindowState; set => raw.WindowState = (MicrosoftXlWindowState)value; }
        public int UILanguage { get => raw.UILanguage; set => raw.UILanguage = value; }
        public int DefaultSheetDirection { get => raw.DefaultSheetDirection; set => raw.DefaultSheetDirection = value; }
        public int CursorMovement { get => raw.CursorMovement; set => raw.CursorMovement = value; }
        public bool ControlCharacters { get => raw.ControlCharacters; set => raw.ControlCharacters = value; }
        public bool EnableEvents { get => raw.EnableEvents; set => raw.EnableEvents = value; }
        public bool DisplayInfoWindow { get => raw.DisplayInfoWindow; set => raw.DisplayInfoWindow = value; }
        public bool ExtendList { get => raw.ExtendList; set => raw.ExtendList = value; }
        public XlOleDbErrors OleDbErrors => new(raw.OLEDBErrors);
        public XlComAddIns ComAddIns => new(raw.COMAddIns);
        public XlDefaultWebOptions DefaultWebOptions => new(raw.DefaultWebOptions);
        public string ProductCode => raw.ProductCode;
        public string UserLibraryPath => raw.UserLibraryPath;
        public bool AutoPercentEntry { get => raw.AutoPercentEntry; set => raw.AutoPercentEntry = value; }
        public XlLanguageSettings LanguageSettings => new(raw.LanguageSettings);
        public XlAnswerWizard AnswerWizard => new(raw.AnswerWizard);
        public int CalculationVersion => raw.CalculationVersion;
        public bool ShowWindowsInTaskbar { get => raw.ShowWindowsInTaskbar; set => raw.ShowWindowsInTaskbar = value; }
        public XlMsoFeatureInstall FeatureInstall { get => (XlMsoFeatureInstall)raw.FeatureInstall; set => raw.FeatureInstall = (MicrosoftMsoFeatureInstall)value; }
        public bool Ready => raw.Ready;
        public XlCellFormat FindFormat { get => new(raw.FindFormat); set => raw.FindFormat = value.raw; }
        public XlCellFormat ReplaceFormat { get => new(raw.ReplaceFormat); set => raw.ReplaceFormat = value.raw; }
        public XlUsedObjects UsedObjects => new(raw.UsedObjects);
        public XlCalculationState CalculationState => (XlCalculationState)raw.CalculationState;
        public XlCalculationInterruptKey CalculationInterruptKey { get => (XlCalculationInterruptKey)raw.CalculationInterruptKey; set => raw.CalculationInterruptKey = (MicrosoftXlCalculationInterruptKey)value; }
        public XlWatches Watches => new(raw.Watches);
        public bool DisplayFunctionToolTips { get => raw.DisplayFunctionToolTips; set => raw.DisplayFunctionToolTips = value; }
        public XlMsoAutomationSecurity AutomationSecurity { get => (XlMsoAutomationSecurity)raw.AutomationSecurity; set => raw.AutomationSecurity = (MicrosoftMsoAutomationSecurity)value; }
        public XlFileDialog FileDialog(XlMsoFileDialogType type) => new(raw.FileDialog[(MicrosoftMsoFileDialogType)type]);
        public bool DisplayPasteOptions { get => raw.DisplayPasteOptions; set => raw.DisplayPasteOptions = value; }
        public bool DisplayInsertOptions { get => raw.DisplayInsertOptions; set => raw.DisplayInsertOptions = value; }
        public bool GenerateGetPivotData { get => raw.GenerateGetPivotData; set => raw.GenerateGetPivotData = value; }
        public XlAutoRecover AutoRecover => new(raw.AutoRecover);
        public int Hwnd => raw.Hwnd;
        public int Hinstance => raw.Hinstance;
        public XlErrorCheckingOptions ErrorCheckingOptions => new(raw.ErrorCheckingOptions);
        public bool AutoFormatAsYouTypeReplaceHyperlinks { get => raw.AutoFormatAsYouTypeReplaceHyperlinks; set => raw.AutoFormatAsYouTypeReplaceHyperlinks = value; }
        public XlSmartTagRecognizers SmartTagRecognizers => new(raw.SmartTagRecognizers);
        public XlNewFile NewWorkbook => new(((Microsoft.Office.Interop.Excel._Application)raw).NewWorkbook);
        public XlSpellingOptions SpellingOptions => new(raw.SpellingOptions);
        public XlSpeech Speech => new(raw.Speech);
        public bool MapPaperSize { get => raw.MapPaperSize; set => raw.MapPaperSize = value; }
        public bool ShowStartupDialog { get => raw.ShowStartupDialog; set => raw.ShowStartupDialog = value; }
        public string DecimalSeparator { get => raw.DecimalSeparator; set => raw.DecimalSeparator = value; }
        public string ThousandsSeparator { get => raw.ThousandsSeparator; set => raw.ThousandsSeparator = value; }
        public bool UseSystemSeparators { get => raw.UseSystemSeparators; set => raw.UseSystemSeparators = value; }
        public XlRange ThisCell => new(raw.ThisCell);
        public XlRTD RTD => new(raw.RTD);
        public bool DisplayDocumentActionTaskPane { get => raw.DisplayDocumentActionTaskPane; set => raw.DisplayDocumentActionTaskPane = value; }
        public bool ArbitraryXMLSupportAvailable => raw.ArbitraryXMLSupportAvailable;
        public int MeasurementUnit { get => raw.MeasurementUnit; set => raw.MeasurementUnit = value; }
        public bool ShowSelectionFloaties { get => raw.ShowSelectionFloaties; set => raw.ShowSelectionFloaties = value; }
        public bool ShowMenuFloaties { get => raw.ShowMenuFloaties; set => raw.ShowMenuFloaties = value; }
        public bool ShowDevTools { get => raw.ShowDevTools; set => raw.ShowDevTools = value; }
        public bool EnableLivePreview { get => raw.EnableLivePreview; set => raw.EnableLivePreview = value; }
        public bool DisplayDocumentInformationPanel { get => raw.DisplayDocumentInformationPanel; set => raw.DisplayDocumentInformationPanel = value; }
        public bool AlwaysUseClearType { get => raw.AlwaysUseClearType; set => raw.AlwaysUseClearType = value; }
        public bool WarnOnFunctionNameConflict { get => raw.WarnOnFunctionNameConflict; set => raw.WarnOnFunctionNameConflict = value; }
        public int FormulaBarHeight { get => raw.FormulaBarHeight; set => raw.FormulaBarHeight = value; }
        public bool DisplayFormulaAutoComplete { get => raw.DisplayFormulaAutoComplete; set => raw.DisplayFormulaAutoComplete = value; }
    }
}
