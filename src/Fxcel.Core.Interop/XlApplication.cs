using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
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
    using MicrosoftMsoFileValidationMode = Microsoft.Office.Core.MsoFileValidationMode;
    using MicrosoftXlCalculationInterruptKey = Microsoft.Office.Interop.Excel.XlCalculationInterruptKey;
    using MicrosoftXlGenerateTableRefs = Microsoft.Office.Interop.Excel.XlGenerateTableRefs;
    using MicrosoftXlFileValidationPivotMode = Microsoft.Office.Interop.Excel.XlFileValidationPivotMode;

    [SupportedOSPlatform("windows")]
    public readonly ref struct XlApplication
    {
        internal readonly MicrosoftApplication raw;
        internal XlApplication(MicrosoftApplication excel) => raw = excel;

        public int Release() => ComHelper.Release(raw);

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
        public XlGenerateTableRefs GenerateTableRefs { get => (XlGenerateTableRefs)raw.GenerateTableRefs; set => raw.GenerateTableRefs = (MicrosoftXlGenerateTableRefs)value; }
        public XlIAssistance Assistance => new(raw.Assistance);
        public bool EnableLargeOperationAlert { get => raw.EnableLargeOperationAlert; set => raw.EnableLargeOperationAlert = value; }
        public int LargeOperationCellThousandCount { get => raw.LargeOperationCellThousandCount; set => raw.LargeOperationCellThousandCount = value; }
        public bool DeferAsyncQueries { get => raw.DeferAsyncQueries; set => raw.DeferAsyncQueries = value; }
        public XlMultiThreadedCalculation MultiThreadedCalculation => new(raw.MultiThreadedCalculation);
        public int ActiveEncryptionSession => raw.ActiveEncryptionSession;
        public bool HighQualityModeForGraphics { get => raw.HighQualityModeForGraphics; set => raw.HighQualityModeForGraphics = value; }
        public XlFileExportConverters FileExportConverters => new(raw.FileExportConverters);
        public XlSmartArtLayouts SmartArtLayouts => new(raw.SmartArtLayouts);
        public XlSmartArtQuickStyles SmartArtQuickStyles => new(raw.SmartArtQuickStyles);
        public XlSmartArtColors SmartArtColors => new(raw.SmartArtColors);
        public XlAddIns2 AddIns2 => new(raw.AddIns2);
        public bool PrintCommunication { get => raw.PrintCommunication; set => raw.PrintCommunication = value; }
        public bool UseClusterConnector { get => raw.UseClusterConnector; set => raw.UseClusterConnector = value; }
        public string ClusterConnector { get => raw.ClusterConnector; set => raw.ClusterConnector = value; }
        public bool Quitting => raw.Quitting;
        public XlProtectedViewWindows ProtectedViewWindows => new(raw.ProtectedViewWindows);
        public XlProtectedViewWindow ActiveProtectedViewWindow => new(raw.ActiveProtectedViewWindow);
        public bool IsSandboxed => raw.IsSandboxed;
        public bool SaveISO8601Dates { get => raw.SaveISO8601Dates; set => raw.SaveISO8601Dates = value; }
        public XlMsoFileValidationMode FileValidation { get => (XlMsoFileValidationMode)raw.FileValidation; set => raw.FileValidation = (MicrosoftMsoFileValidationMode)value; }
        public XlFileValidationPivotMode FileValidationPivot { get => (XlFileValidationPivotMode)raw.FileValidationPivot; set => raw.FileValidationPivot = (MicrosoftXlFileValidationPivotMode)value; }
        public bool ShowQuickAnalysis { get => raw.ShowQuickAnalysis; set => raw.ShowQuickAnalysis = value; }
        public XlQuickAnalysis QuickAnalysis => new(raw.QuickAnalysis);
        public bool FlashFill { get => raw.FlashFill; set => raw.FlashFill = value; }
        public bool EnableMacroAnimations { get => raw.EnableMacroAnimations; set => raw.EnableMacroAnimations = value; }
        public bool ChartDataPointTrack { get => raw.ChartDataPointTrack; set => raw.ChartDataPointTrack = value; }
        public bool FlashFillMode { get => raw.FlashFillMode; set => raw.FlashFillMode = value; }
        public bool MergeInstances { get => raw.MergeInstances; set => raw.MergeInstances = value; }
        public bool EnableCheckFileExtensions { get => raw.EnableCheckFileExtensions; set => raw.EnableCheckFileExtensions = value; }

        public void Calculate() => raw.Calculate();
        /// <summary>指定したDDEチャネルを介してコマンドの実行や別のアプリケーションでアクションの実行をする.</summary>
        /// <param name="channel">DdeInitiateの戻り値.</param>
        /// <param name="message">受信アプリケーションで定義されたメッセージ.</param>
        /// <see href="https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel._application.ddeexecute?view=excel-pia" />
        public void DdeExecute(int channel, string message) => raw.DDEExecute(channel, message);
        /// <summary>アプリケーションへのDDEチャネルを開く.</summary>
        /// <param name="app">アプリケーション名.</param>
        /// <param name="topic">チャネルを開いているアプリケーション内容についての説明.</param>
        /// <returns>チャネルID.</returns>
        /// <see href="https://docs.microsoft.com/ja-jp/dotnet/api/microsoft.office.interop.excel._application.ddeinitiate?view=excel-pia" />
        public int DdeInitiate(string app, string topic) => raw.DDEInitiate(app, topic);
        /// <summary>アプリケーションにデータを送信する.</summary>
        /// <param name="channel">DdeInitiateの戻り値.</param>
        /// <param name="item">データの送信先アイテム.</param>
        /// <param name="data">アプリケーションに送信されるデータ.</param>
        /// <see href="https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel._application.ddepoke?view=excel-pia" />
        public void DdePoke(int channel, object item, object data) => raw.DDEPoke(channel, item, data);
        /// <summary>指定したアプリケーションに情報を要求する.</summary>
        /// <param name="channel">DdeInitiateの戻り値.</param>
        /// <param name="item">リクエストするアイテム.</param>
        /// <returns>配列アイテム.</returns>
        /// <see href="https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel._application.dderequest?view=excel-pia" />
        public object DdeRequest(int channel, string item) => raw.DDERequest(channel, item);
        /// <summary>チャネルを閉じる.</summary>
        /// <param name="channel">DdeInitiateの戻り値.</param>
        /// <see href="https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel._application.ddeterminate?view=excel-pia" />
        public void DdeTerminate(int channel) => raw.DDETerminate(channel);
        /// <summary></summary>
        /// <param name="name"></param>
        /// <returns></returns>
        /// <see href="https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel._application.evaluate?view=excel-pia" />
        public object Evaluate(string name) => raw.Evaluate(name);
        /// <summary></summary>
        /// <param name="function"></param>
        /// <returns></returns>
        /// <see href="https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel._application.executeexcel4macro?view=excel-pia" />
        public object ExecuteExcel4Macro(string function) => raw.ExecuteExcel4Macro(function);

        /// <summary></summary>
        /// <param name="arg1"></param>
        /// <param name="arg2"></param>
        /// <returns></returns>
        /// <see href="https://docs.microsoft.com/ja-jp/dotnet/api/microsoft.office.interop.excel._application.intersect?view=excel-pia" />
        public XlRange Intersect(XlRange arg1, XlRange arg2) => new(raw.Intersect(arg1.raw, arg2.raw));
        public XlRange Intersect(XlRange arg1, XlRange arg2, XlRange arg3) => new(raw.Intersect(arg1.raw, arg2.raw, arg3.raw));
        public XlRange Intersect(XlRange arg1, XlRange arg2, XlRange arg3, XlRange arg4) => new(raw.Intersect(arg1.raw, arg2.raw, arg3.raw, arg4.raw));
        public XlRange Intersect(XlRange arg1, XlRange arg2, XlRange arg3, XlRange arg4, XlRange arg5) => new(raw.Intersect(arg1.raw, arg2.raw, arg3.raw, arg4.raw, arg5.raw));
        public XlRange Intersect(XlRange arg1, XlRange arg2, XlRange arg3, XlRange arg4, XlRange arg5, XlRange arg6) => new(raw.Intersect(arg1.raw, arg2.raw, arg3.raw, arg4.raw, arg5.raw, arg6.raw));
        public XlRange Intersect(XlRange arg1, XlRange arg2, XlRange arg3, XlRange arg4, XlRange arg5, XlRange arg6, XlRange arg7) => new(raw.Intersect(arg1.raw, arg2.raw, arg3.raw, arg4.raw, arg5.raw, arg6.raw, arg7.raw));
        public XlRange Intersect(XlRange arg1, XlRange arg2, XlRange arg3, XlRange arg4, XlRange arg5, XlRange arg6, XlRange arg7, XlRange arg8) => new(raw.Intersect(arg1.raw, arg2.raw, arg3.raw, arg4.raw, arg5.raw, arg6.raw, arg7.raw, arg8.raw));
        public XlRange Intersect(XlRange arg1, XlRange arg2, XlRange arg3, XlRange arg4, XlRange arg5, XlRange arg6, XlRange arg7, XlRange arg8, XlRange arg9) => new(raw.Intersect(arg1.raw, arg2.raw, arg3.raw, arg4.raw, arg5.raw, arg6.raw, arg7.raw, arg8.raw, arg9.raw));
        public XlRange Intersect(XlRange arg1, XlRange arg2, XlRange arg3, XlRange arg4, XlRange arg5, XlRange arg6, XlRange arg7, XlRange arg8, XlRange arg9, XlRange arg10) => new(raw.Intersect(arg1.raw, arg2.raw, arg3.raw, arg4.raw, arg5.raw, arg6.raw, arg7.raw, arg8.raw, arg9.raw, arg10.raw));
        public XlRange Intersect(XlRange arg1, XlRange arg2, XlRange arg3, XlRange arg4, XlRange arg5, XlRange arg6, XlRange arg7, XlRange arg8, XlRange arg9, XlRange arg10, XlRange arg11) => new(raw.Intersect(arg1.raw, arg2.raw, arg3.raw, arg4.raw, arg5.raw, arg6.raw, arg7.raw, arg8.raw, arg9.raw, arg10.raw, arg11.raw));
        public XlRange Intersect(XlRange arg1, XlRange arg2, XlRange arg3, XlRange arg4, XlRange arg5, XlRange arg6, XlRange arg7, XlRange arg8, XlRange arg9, XlRange arg10, XlRange arg11, XlRange arg12) => new(raw.Intersect(arg1.raw, arg2.raw, arg3.raw, arg4.raw, arg5.raw, arg6.raw, arg7.raw, arg8.raw, arg9.raw, arg10.raw, arg11.raw, arg12.raw));
        public XlRange Intersect(XlRange arg1, XlRange arg2, XlRange arg3, XlRange arg4, XlRange arg5, XlRange arg6, XlRange arg7, XlRange arg8, XlRange arg9, XlRange arg10, XlRange arg11, XlRange arg12, XlRange arg13) => new(raw.Intersect(arg1.raw, arg2.raw, arg3.raw, arg4.raw, arg5.raw, arg6.raw, arg7.raw, arg8.raw, arg9.raw, arg10.raw, arg11.raw, arg12.raw, arg13.raw));
        public XlRange Intersect(XlRange arg1, XlRange arg2, XlRange arg3, XlRange arg4, XlRange arg5, XlRange arg6, XlRange arg7, XlRange arg8, XlRange arg9, XlRange arg10, XlRange arg11, XlRange arg12, XlRange arg13, XlRange arg14) => new(raw.Intersect(arg1.raw, arg2.raw, arg3.raw, arg4.raw, arg5.raw, arg6.raw, arg7.raw, arg8.raw, arg9.raw, arg10.raw, arg11.raw, arg12.raw, arg13.raw, arg14.raw));
        public XlRange Intersect(XlRange arg1, XlRange arg2, XlRange arg3, XlRange arg4, XlRange arg5, XlRange arg6, XlRange arg7, XlRange arg8, XlRange arg9, XlRange arg10, XlRange arg11, XlRange arg12, XlRange arg13, XlRange arg14, XlRange arg15) => new(raw.Intersect(arg1.raw, arg2.raw, arg3.raw, arg4.raw, arg5.raw, arg6.raw, arg7.raw, arg8.raw, arg9.raw, arg10.raw, arg11.raw, arg12.raw, arg13.raw, arg14.raw, arg15.raw));
        public XlRange Intersect(XlRange arg1, XlRange arg2, XlRange arg3, XlRange arg4, XlRange arg5, XlRange arg6, XlRange arg7, XlRange arg8, XlRange arg9, XlRange arg10, XlRange arg11, XlRange arg12, XlRange arg13, XlRange arg14, XlRange arg15, XlRange arg16) => new(raw.Intersect(arg1.raw, arg2.raw, arg3.raw, arg4.raw, arg5.raw, arg6.raw, arg7.raw, arg8.raw, arg9.raw, arg10.raw, arg11.raw, arg12.raw, arg13.raw, arg14.raw, arg15.raw, arg16.raw));
        public XlRange Intersect(XlRange arg1, XlRange arg2, XlRange arg3, XlRange arg4, XlRange arg5, XlRange arg6, XlRange arg7, XlRange arg8, XlRange arg9, XlRange arg10, XlRange arg11, XlRange arg12, XlRange arg13, XlRange arg14, XlRange arg15, XlRange arg16, XlRange arg17) => new(raw.Intersect(arg1.raw, arg2.raw, arg3.raw, arg4.raw, arg5.raw, arg6.raw, arg7.raw, arg8.raw, arg9.raw, arg10.raw, arg11.raw, arg12.raw, arg13.raw, arg14.raw, arg15.raw, arg16.raw, arg17.raw));
        public XlRange Intersect(XlRange arg1, XlRange arg2, XlRange arg3, XlRange arg4, XlRange arg5, XlRange arg6, XlRange arg7, XlRange arg8, XlRange arg9, XlRange arg10, XlRange arg11, XlRange arg12, XlRange arg13, XlRange arg14, XlRange arg15, XlRange arg16, XlRange arg17, XlRange arg18) => new(raw.Intersect(arg1.raw, arg2.raw, arg3.raw, arg4.raw, arg5.raw, arg6.raw, arg7.raw, arg8.raw, arg9.raw, arg10.raw, arg11.raw, arg12.raw, arg13.raw, arg14.raw, arg15.raw, arg16.raw, arg17.raw, arg18.raw));
        public XlRange Intersect(XlRange arg1, XlRange arg2, XlRange arg3, XlRange arg4, XlRange arg5, XlRange arg6, XlRange arg7, XlRange arg8, XlRange arg9, XlRange arg10, XlRange arg11, XlRange arg12, XlRange arg13, XlRange arg14, XlRange arg15, XlRange arg16, XlRange arg17, XlRange arg18, XlRange arg19) => new(raw.Intersect(arg1.raw, arg2.raw, arg3.raw, arg4.raw, arg5.raw, arg6.raw, arg7.raw, arg8.raw, arg9.raw, arg10.raw, arg11.raw, arg12.raw, arg13.raw, arg14.raw, arg15.raw, arg16.raw, arg17.raw, arg18.raw, arg19.raw));
        public XlRange Intersect(XlRange arg1, XlRange arg2, XlRange arg3, XlRange arg4, XlRange arg5, XlRange arg6, XlRange arg7, XlRange arg8, XlRange arg9, XlRange arg10, XlRange arg11, XlRange arg12, XlRange arg13, XlRange arg14, XlRange arg15, XlRange arg16, XlRange arg17, XlRange arg18, XlRange arg19, XlRange arg20) => new(raw.Intersect(arg1.raw, arg2.raw, arg3.raw, arg4.raw, arg5.raw, arg6.raw, arg7.raw, arg8.raw, arg9.raw, arg10.raw, arg11.raw, arg12.raw, arg13.raw, arg14.raw, arg15.raw, arg16.raw, arg17.raw, arg18.raw, arg19.raw, arg20.raw));
        public XlRange Intersect(XlRange arg1, XlRange arg2, XlRange arg3, XlRange arg4, XlRange arg5, XlRange arg6, XlRange arg7, XlRange arg8, XlRange arg9, XlRange arg10, XlRange arg11, XlRange arg12, XlRange arg13, XlRange arg14, XlRange arg15, XlRange arg16, XlRange arg17, XlRange arg18, XlRange arg19, XlRange arg20, XlRange arg21) => new(raw.Intersect(arg1.raw, arg2.raw, arg3.raw, arg4.raw, arg5.raw, arg6.raw, arg7.raw, arg8.raw, arg9.raw, arg10.raw, arg11.raw, arg12.raw, arg13.raw, arg14.raw, arg15.raw, arg16.raw, arg17.raw, arg18.raw, arg19.raw, arg20.raw, arg21.raw));
        public XlRange Intersect(XlRange arg1, XlRange arg2, XlRange arg3, XlRange arg4, XlRange arg5, XlRange arg6, XlRange arg7, XlRange arg8, XlRange arg9, XlRange arg10, XlRange arg11, XlRange arg12, XlRange arg13, XlRange arg14, XlRange arg15, XlRange arg16, XlRange arg17, XlRange arg18, XlRange arg19, XlRange arg20, XlRange arg21, XlRange arg22) => new(raw.Intersect(arg1.raw, arg2.raw, arg3.raw, arg4.raw, arg5.raw, arg6.raw, arg7.raw, arg8.raw, arg9.raw, arg10.raw, arg11.raw, arg12.raw, arg13.raw, arg14.raw, arg15.raw, arg16.raw, arg17.raw, arg18.raw, arg19.raw, arg20.raw, arg21.raw, arg22.raw));
        public XlRange Intersect(XlRange arg1, XlRange arg2, XlRange arg3, XlRange arg4, XlRange arg5, XlRange arg6, XlRange arg7, XlRange arg8, XlRange arg9, XlRange arg10, XlRange arg11, XlRange arg12, XlRange arg13, XlRange arg14, XlRange arg15, XlRange arg16, XlRange arg17, XlRange arg18, XlRange arg19, XlRange arg20, XlRange arg21, XlRange arg22, XlRange arg23) => new(raw.Intersect(arg1.raw, arg2.raw, arg3.raw, arg4.raw, arg5.raw, arg6.raw, arg7.raw, arg8.raw, arg9.raw, arg10.raw, arg11.raw, arg12.raw, arg13.raw, arg14.raw, arg15.raw, arg16.raw, arg17.raw, arg18.raw, arg19.raw, arg20.raw, arg21.raw, arg22.raw, arg23.raw));
        public XlRange Intersect(XlRange arg1, XlRange arg2, XlRange arg3, XlRange arg4, XlRange arg5, XlRange arg6, XlRange arg7, XlRange arg8, XlRange arg9, XlRange arg10, XlRange arg11, XlRange arg12, XlRange arg13, XlRange arg14, XlRange arg15, XlRange arg16, XlRange arg17, XlRange arg18, XlRange arg19, XlRange arg20, XlRange arg21, XlRange arg22, XlRange arg23, XlRange arg24) => new(raw.Intersect(arg1.raw, arg2.raw, arg3.raw, arg4.raw, arg5.raw, arg6.raw, arg7.raw, arg8.raw, arg9.raw, arg10.raw, arg11.raw, arg12.raw, arg13.raw, arg14.raw, arg15.raw, arg16.raw, arg17.raw, arg18.raw, arg19.raw, arg20.raw, arg21.raw, arg22.raw, arg23.raw, arg24.raw));
        public XlRange Intersect(XlRange arg1, XlRange arg2, XlRange arg3, XlRange arg4, XlRange arg5, XlRange arg6, XlRange arg7, XlRange arg8, XlRange arg9, XlRange arg10, XlRange arg11, XlRange arg12, XlRange arg13, XlRange arg14, XlRange arg15, XlRange arg16, XlRange arg17, XlRange arg18, XlRange arg19, XlRange arg20, XlRange arg21, XlRange arg22, XlRange arg23, XlRange arg24, XlRange arg25) => new(raw.Intersect(arg1.raw, arg2.raw, arg3.raw, arg4.raw, arg5.raw, arg6.raw, arg7.raw, arg8.raw, arg9.raw, arg10.raw, arg11.raw, arg12.raw, arg13.raw, arg14.raw, arg15.raw, arg16.raw, arg17.raw, arg18.raw, arg19.raw, arg20.raw, arg21.raw, arg22.raw, arg23.raw, arg24.raw, arg25.raw));
        public XlRange Intersect(XlRange arg1, XlRange arg2, XlRange arg3, XlRange arg4, XlRange arg5, XlRange arg6, XlRange arg7, XlRange arg8, XlRange arg9, XlRange arg10, XlRange arg11, XlRange arg12, XlRange arg13, XlRange arg14, XlRange arg15, XlRange arg16, XlRange arg17, XlRange arg18, XlRange arg19, XlRange arg20, XlRange arg21, XlRange arg22, XlRange arg23, XlRange arg24, XlRange arg25, XlRange arg26) => new(raw.Intersect(arg1.raw, arg2.raw, arg3.raw, arg4.raw, arg5.raw, arg6.raw, arg7.raw, arg8.raw, arg9.raw, arg10.raw, arg11.raw, arg12.raw, arg13.raw, arg14.raw, arg15.raw, arg16.raw, arg17.raw, arg18.raw, arg19.raw, arg20.raw, arg21.raw, arg22.raw, arg23.raw, arg24.raw, arg25.raw, arg26.raw));
        public XlRange Intersect(XlRange arg1, XlRange arg2, XlRange arg3, XlRange arg4, XlRange arg5, XlRange arg6, XlRange arg7, XlRange arg8, XlRange arg9, XlRange arg10, XlRange arg11, XlRange arg12, XlRange arg13, XlRange arg14, XlRange arg15, XlRange arg16, XlRange arg17, XlRange arg18, XlRange arg19, XlRange arg20, XlRange arg21, XlRange arg22, XlRange arg23, XlRange arg24, XlRange arg25, XlRange arg26, XlRange arg27) => new(raw.Intersect(arg1.raw, arg2.raw, arg3.raw, arg4.raw, arg5.raw, arg6.raw, arg7.raw, arg8.raw, arg9.raw, arg10.raw, arg11.raw, arg12.raw, arg13.raw, arg14.raw, arg15.raw, arg16.raw, arg17.raw, arg18.raw, arg19.raw, arg20.raw, arg21.raw, arg22.raw, arg23.raw, arg24.raw, arg25.raw, arg26.raw, arg27.raw));
        public XlRange Intersect(XlRange arg1, XlRange arg2, XlRange arg3, XlRange arg4, XlRange arg5, XlRange arg6, XlRange arg7, XlRange arg8, XlRange arg9, XlRange arg10, XlRange arg11, XlRange arg12, XlRange arg13, XlRange arg14, XlRange arg15, XlRange arg16, XlRange arg17, XlRange arg18, XlRange arg19, XlRange arg20, XlRange arg21, XlRange arg22, XlRange arg23, XlRange arg24, XlRange arg25, XlRange arg26, XlRange arg27, XlRange arg28) => new(raw.Intersect(arg1.raw, arg2.raw, arg3.raw, arg4.raw, arg5.raw, arg6.raw, arg7.raw, arg8.raw, arg9.raw, arg10.raw, arg11.raw, arg12.raw, arg13.raw, arg14.raw, arg15.raw, arg16.raw, arg17.raw, arg18.raw, arg19.raw, arg20.raw, arg21.raw, arg22.raw, arg23.raw, arg24.raw, arg25.raw, arg26.raw, arg27.raw, arg28.raw));
        public XlRange Intersect(XlRange arg1, XlRange arg2, XlRange arg3, XlRange arg4, XlRange arg5, XlRange arg6, XlRange arg7, XlRange arg8, XlRange arg9, XlRange arg10, XlRange arg11, XlRange arg12, XlRange arg13, XlRange arg14, XlRange arg15, XlRange arg16, XlRange arg17, XlRange arg18, XlRange arg19, XlRange arg20, XlRange arg21, XlRange arg22, XlRange arg23, XlRange arg24, XlRange arg25, XlRange arg26, XlRange arg27, XlRange arg28, XlRange arg29) => new(raw.Intersect(arg1.raw, arg2.raw, arg3.raw, arg4.raw, arg5.raw, arg6.raw, arg7.raw, arg8.raw, arg9.raw, arg10.raw, arg11.raw, arg12.raw, arg13.raw, arg14.raw, arg15.raw, arg16.raw, arg17.raw, arg18.raw, arg19.raw, arg20.raw, arg21.raw, arg22.raw, arg23.raw, arg24.raw, arg25.raw, arg26.raw, arg27.raw, arg28.raw, arg29.raw));
        public XlRange Intersect(XlRange arg1, XlRange arg2, XlRange arg3, XlRange arg4, XlRange arg5, XlRange arg6, XlRange arg7, XlRange arg8, XlRange arg9, XlRange arg10, XlRange arg11, XlRange arg12, XlRange arg13, XlRange arg14, XlRange arg15, XlRange arg16, XlRange arg17, XlRange arg18, XlRange arg19, XlRange arg20, XlRange arg21, XlRange arg22, XlRange arg23, XlRange arg24, XlRange arg25, XlRange arg26, XlRange arg27, XlRange arg28, XlRange arg29, XlRange arg30) => new(raw.Intersect(arg1.raw, arg2.raw, arg3.raw, arg4.raw, arg5.raw, arg6.raw, arg7.raw, arg8.raw, arg9.raw, arg10.raw, arg11.raw, arg12.raw, arg13.raw, arg14.raw, arg15.raw, arg16.raw, arg17.raw, arg18.raw, arg19.raw, arg20.raw, arg21.raw, arg22.raw, arg23.raw, arg24.raw, arg25.raw, arg26.raw, arg27.raw, arg28.raw, arg29.raw, arg30.raw));

        // TODO: 
        /// <summary></summary>
        /// <param name="macro"></param>
        /// <returns></returns>
        /// <see href="https://docs.microsoft.com/ja-jp/dotnet/api/microsoft.office.interop.excel._application.run?view=excel-pia" />
        //public object Run(string macro) => raw.Run(macro);
        //public object Run(string macro, object arg1) => raw.Run(macro, arg1);
        //public object Run(string macro, object arg1, object arg2) => raw.Run(macro, arg1, arg2);
        //public object Run(string macro, object arg1, object arg2, object arg3) => raw.Run(macro, arg1, arg2, arg3);
        //public object Run(string macro, object arg1, object arg2, object arg3, object arg4) => raw.Run(macro, arg1, arg2, arg3, arg4);
        //public object Run(string macro, object arg1, object arg2, object arg3, object arg4, object arg5) => raw.Run(macro, arg1, arg2, arg3, arg4, arg5);
        //public object Run(string macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6) => raw.Run(macro, arg1, arg2, arg3, arg4, arg5, arg6);
        //public object Run(string macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7) => raw.Run(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7);
        //public object Run(string macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8) => raw.Run(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8);
        //public object Run(string macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9) => raw.Run(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9);
        //public object Run(string macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10) => raw.Run(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10);
        //public object Run(string macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11) => raw.Run(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11);
        //public object Run(string macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12) => raw.Run(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12);
        //public object Run(string macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13) => raw.Run(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13);
        //public object Run(string macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14) => raw.Run(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14);
        //public object Run(string macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15) => raw.Run(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15);
        //public object Run(string macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16) => raw.Run(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16);
        //public object Run(string macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17) => raw.Run(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17);
        //public object Run(string macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18) => raw.Run(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18);
        //public object Run(string macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19) => raw.Run(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19);
        //public object Run(string macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20) => raw.Run(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20);
        //public object Run(string macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21) => raw.Run(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21);
        //public object Run(string macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22) => raw.Run(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22);
        //public object Run(string macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23) => raw.Run(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23);
        //public object Run(string macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24) => raw.Run(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24);
        //public object Run(string macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25) => raw.Run(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25);
        //public object Run(string macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26) => raw.Run(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26);
        //public object Run(string macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27) => raw.Run(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27);
        //public object Run(string macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28) => raw.Run(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28);
        //public object Run(string macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28, object arg29) => raw.Run(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29);
        //public object Run(string macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28, object arg29, object arg30) => raw.Run(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30);
        public object Run(string macro, object? arg1 = null, object? arg2 = null, object? arg3 = null, object? arg4 = null, object? arg5 = null, object? arg6 = null, object? arg7 = null, object? arg8 = null, object? arg9 = null, object? arg10 = null, object? arg11 = null, object? arg12 = null, object? arg13 = null, object? arg14 = null, object? arg15 = null, object? arg16 = null, object? arg17 = null, object? arg18 = null, object? arg19 = null, object? arg20 = null, object? arg21 = null, object? arg22 = null, object? arg23 = null, object? arg24 = null, object? arg25 = null, object? arg26 = null, object? arg27 = null, object? arg28 = null, object? arg29 = null, object? arg30 = null) => raw.Run(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30);

        /// <summary>アクティブなアプリケーションにキーストロークを送信する.</summary>
        /// <param name="keys">送信するキーの組み合わせ.</param>
        /// <param name="wait">マクロに制御を戻す前に, キーが処理されるのを待機させる場合は true を, キーが処理されるのを待機せずにマクロの実行をさせる場合は false を指定する. (default: false)</param>
        /// <see href="https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel._application.sendkeys?view=excel-pia" />
        public void SendKeys(string keys, bool wait = false) => raw.SendKeys(keys, wait);

        /// <summary></summary>
        /// <param name="arg1"></param>
        /// <param name="arg2"></param>
        /// <returns></returns>
        /// <see href="https://docs.microsoft.com/ja-jp/dotnet/api/microsoft.office.interop.excel._application.union?view=excel-pia" />
        public XlRange Union(XlRange arg1, XlRange arg2) => new(raw.Union(arg1.raw, arg2.raw));
        public XlRange Union(XlRange arg1, XlRange arg2, XlRange arg3) => new(raw.Union(arg1.raw, arg2.raw, arg3.raw));
        public XlRange Union(XlRange arg1, XlRange arg2, XlRange arg3, XlRange arg4) => new(raw.Union(arg1.raw, arg2.raw, arg3.raw, arg4.raw));
        public XlRange Union(XlRange arg1, XlRange arg2, XlRange arg3, XlRange arg4, XlRange arg5) => new(raw.Union(arg1.raw, arg2.raw, arg3.raw, arg4.raw, arg5.raw));
        public XlRange Union(XlRange arg1, XlRange arg2, XlRange arg3, XlRange arg4, XlRange arg5, XlRange arg6) => new(raw.Union(arg1.raw, arg2.raw, arg3.raw, arg4.raw, arg5.raw, arg6.raw));
        public XlRange Union(XlRange arg1, XlRange arg2, XlRange arg3, XlRange arg4, XlRange arg5, XlRange arg6, XlRange arg7) => new(raw.Union(arg1.raw, arg2.raw, arg3.raw, arg4.raw, arg5.raw, arg6.raw, arg7.raw));
        public XlRange Union(XlRange arg1, XlRange arg2, XlRange arg3, XlRange arg4, XlRange arg5, XlRange arg6, XlRange arg7, XlRange arg8) => new(raw.Union(arg1.raw, arg2.raw, arg3.raw, arg4.raw, arg5.raw, arg6.raw, arg7.raw, arg8.raw));
        public XlRange Union(XlRange arg1, XlRange arg2, XlRange arg3, XlRange arg4, XlRange arg5, XlRange arg6, XlRange arg7, XlRange arg8, XlRange arg9) => new(raw.Union(arg1.raw, arg2.raw, arg3.raw, arg4.raw, arg5.raw, arg6.raw, arg7.raw, arg8.raw, arg9.raw));
        public XlRange Union(XlRange arg1, XlRange arg2, XlRange arg3, XlRange arg4, XlRange arg5, XlRange arg6, XlRange arg7, XlRange arg8, XlRange arg9, XlRange arg10) => new(raw.Union(arg1.raw, arg2.raw, arg3.raw, arg4.raw, arg5.raw, arg6.raw, arg7.raw, arg8.raw, arg9.raw, arg10.raw));
        public XlRange Union(XlRange arg1, XlRange arg2, XlRange arg3, XlRange arg4, XlRange arg5, XlRange arg6, XlRange arg7, XlRange arg8, XlRange arg9, XlRange arg10, XlRange arg11) => new(raw.Union(arg1.raw, arg2.raw, arg3.raw, arg4.raw, arg5.raw, arg6.raw, arg7.raw, arg8.raw, arg9.raw, arg10.raw, arg11.raw));
        public XlRange Union(XlRange arg1, XlRange arg2, XlRange arg3, XlRange arg4, XlRange arg5, XlRange arg6, XlRange arg7, XlRange arg8, XlRange arg9, XlRange arg10, XlRange arg11, XlRange arg12) => new(raw.Union(arg1.raw, arg2.raw, arg3.raw, arg4.raw, arg5.raw, arg6.raw, arg7.raw, arg8.raw, arg9.raw, arg10.raw, arg11.raw, arg12.raw));
        public XlRange Union(XlRange arg1, XlRange arg2, XlRange arg3, XlRange arg4, XlRange arg5, XlRange arg6, XlRange arg7, XlRange arg8, XlRange arg9, XlRange arg10, XlRange arg11, XlRange arg12, XlRange arg13) => new(raw.Union(arg1.raw, arg2.raw, arg3.raw, arg4.raw, arg5.raw, arg6.raw, arg7.raw, arg8.raw, arg9.raw, arg10.raw, arg11.raw, arg12.raw, arg13.raw));
        public XlRange Union(XlRange arg1, XlRange arg2, XlRange arg3, XlRange arg4, XlRange arg5, XlRange arg6, XlRange arg7, XlRange arg8, XlRange arg9, XlRange arg10, XlRange arg11, XlRange arg12, XlRange arg13, XlRange arg14) => new(raw.Union(arg1.raw, arg2.raw, arg3.raw, arg4.raw, arg5.raw, arg6.raw, arg7.raw, arg8.raw, arg9.raw, arg10.raw, arg11.raw, arg12.raw, arg13.raw, arg14.raw));
        public XlRange Union(XlRange arg1, XlRange arg2, XlRange arg3, XlRange arg4, XlRange arg5, XlRange arg6, XlRange arg7, XlRange arg8, XlRange arg9, XlRange arg10, XlRange arg11, XlRange arg12, XlRange arg13, XlRange arg14, XlRange arg15) => new(raw.Union(arg1.raw, arg2.raw, arg3.raw, arg4.raw, arg5.raw, arg6.raw, arg7.raw, arg8.raw, arg9.raw, arg10.raw, arg11.raw, arg12.raw, arg13.raw, arg14.raw, arg15.raw));
        public XlRange Union(XlRange arg1, XlRange arg2, XlRange arg3, XlRange arg4, XlRange arg5, XlRange arg6, XlRange arg7, XlRange arg8, XlRange arg9, XlRange arg10, XlRange arg11, XlRange arg12, XlRange arg13, XlRange arg14, XlRange arg15, XlRange arg16) => new(raw.Union(arg1.raw, arg2.raw, arg3.raw, arg4.raw, arg5.raw, arg6.raw, arg7.raw, arg8.raw, arg9.raw, arg10.raw, arg11.raw, arg12.raw, arg13.raw, arg14.raw, arg15.raw, arg16.raw));
        public XlRange Union(XlRange arg1, XlRange arg2, XlRange arg3, XlRange arg4, XlRange arg5, XlRange arg6, XlRange arg7, XlRange arg8, XlRange arg9, XlRange arg10, XlRange arg11, XlRange arg12, XlRange arg13, XlRange arg14, XlRange arg15, XlRange arg16, XlRange arg17) => new(raw.Union(arg1.raw, arg2.raw, arg3.raw, arg4.raw, arg5.raw, arg6.raw, arg7.raw, arg8.raw, arg9.raw, arg10.raw, arg11.raw, arg12.raw, arg13.raw, arg14.raw, arg15.raw, arg16.raw, arg17.raw));
        public XlRange Union(XlRange arg1, XlRange arg2, XlRange arg3, XlRange arg4, XlRange arg5, XlRange arg6, XlRange arg7, XlRange arg8, XlRange arg9, XlRange arg10, XlRange arg11, XlRange arg12, XlRange arg13, XlRange arg14, XlRange arg15, XlRange arg16, XlRange arg17, XlRange arg18) => new(raw.Union(arg1.raw, arg2.raw, arg3.raw, arg4.raw, arg5.raw, arg6.raw, arg7.raw, arg8.raw, arg9.raw, arg10.raw, arg11.raw, arg12.raw, arg13.raw, arg14.raw, arg15.raw, arg16.raw, arg17.raw, arg18.raw));
        public XlRange Union(XlRange arg1, XlRange arg2, XlRange arg3, XlRange arg4, XlRange arg5, XlRange arg6, XlRange arg7, XlRange arg8, XlRange arg9, XlRange arg10, XlRange arg11, XlRange arg12, XlRange arg13, XlRange arg14, XlRange arg15, XlRange arg16, XlRange arg17, XlRange arg18, XlRange arg19) => new(raw.Union(arg1.raw, arg2.raw, arg3.raw, arg4.raw, arg5.raw, arg6.raw, arg7.raw, arg8.raw, arg9.raw, arg10.raw, arg11.raw, arg12.raw, arg13.raw, arg14.raw, arg15.raw, arg16.raw, arg17.raw, arg18.raw, arg19.raw));
        public XlRange Union(XlRange arg1, XlRange arg2, XlRange arg3, XlRange arg4, XlRange arg5, XlRange arg6, XlRange arg7, XlRange arg8, XlRange arg9, XlRange arg10, XlRange arg11, XlRange arg12, XlRange arg13, XlRange arg14, XlRange arg15, XlRange arg16, XlRange arg17, XlRange arg18, XlRange arg19, XlRange arg20) => new(raw.Union(arg1.raw, arg2.raw, arg3.raw, arg4.raw, arg5.raw, arg6.raw, arg7.raw, arg8.raw, arg9.raw, arg10.raw, arg11.raw, arg12.raw, arg13.raw, arg14.raw, arg15.raw, arg16.raw, arg17.raw, arg18.raw, arg19.raw, arg20.raw));
        public XlRange Union(XlRange arg1, XlRange arg2, XlRange arg3, XlRange arg4, XlRange arg5, XlRange arg6, XlRange arg7, XlRange arg8, XlRange arg9, XlRange arg10, XlRange arg11, XlRange arg12, XlRange arg13, XlRange arg14, XlRange arg15, XlRange arg16, XlRange arg17, XlRange arg18, XlRange arg19, XlRange arg20, XlRange arg21) => new(raw.Union(arg1.raw, arg2.raw, arg3.raw, arg4.raw, arg5.raw, arg6.raw, arg7.raw, arg8.raw, arg9.raw, arg10.raw, arg11.raw, arg12.raw, arg13.raw, arg14.raw, arg15.raw, arg16.raw, arg17.raw, arg18.raw, arg19.raw, arg20.raw, arg21.raw));
        public XlRange Union(XlRange arg1, XlRange arg2, XlRange arg3, XlRange arg4, XlRange arg5, XlRange arg6, XlRange arg7, XlRange arg8, XlRange arg9, XlRange arg10, XlRange arg11, XlRange arg12, XlRange arg13, XlRange arg14, XlRange arg15, XlRange arg16, XlRange arg17, XlRange arg18, XlRange arg19, XlRange arg20, XlRange arg21, XlRange arg22) => new(raw.Union(arg1.raw, arg2.raw, arg3.raw, arg4.raw, arg5.raw, arg6.raw, arg7.raw, arg8.raw, arg9.raw, arg10.raw, arg11.raw, arg12.raw, arg13.raw, arg14.raw, arg15.raw, arg16.raw, arg17.raw, arg18.raw, arg19.raw, arg20.raw, arg21.raw, arg22.raw));
        public XlRange Union(XlRange arg1, XlRange arg2, XlRange arg3, XlRange arg4, XlRange arg5, XlRange arg6, XlRange arg7, XlRange arg8, XlRange arg9, XlRange arg10, XlRange arg11, XlRange arg12, XlRange arg13, XlRange arg14, XlRange arg15, XlRange arg16, XlRange arg17, XlRange arg18, XlRange arg19, XlRange arg20, XlRange arg21, XlRange arg22, XlRange arg23) => new(raw.Union(arg1.raw, arg2.raw, arg3.raw, arg4.raw, arg5.raw, arg6.raw, arg7.raw, arg8.raw, arg9.raw, arg10.raw, arg11.raw, arg12.raw, arg13.raw, arg14.raw, arg15.raw, arg16.raw, arg17.raw, arg18.raw, arg19.raw, arg20.raw, arg21.raw, arg22.raw, arg23.raw));
        public XlRange Union(XlRange arg1, XlRange arg2, XlRange arg3, XlRange arg4, XlRange arg5, XlRange arg6, XlRange arg7, XlRange arg8, XlRange arg9, XlRange arg10, XlRange arg11, XlRange arg12, XlRange arg13, XlRange arg14, XlRange arg15, XlRange arg16, XlRange arg17, XlRange arg18, XlRange arg19, XlRange arg20, XlRange arg21, XlRange arg22, XlRange arg23, XlRange arg24) => new(raw.Union(arg1.raw, arg2.raw, arg3.raw, arg4.raw, arg5.raw, arg6.raw, arg7.raw, arg8.raw, arg9.raw, arg10.raw, arg11.raw, arg12.raw, arg13.raw, arg14.raw, arg15.raw, arg16.raw, arg17.raw, arg18.raw, arg19.raw, arg20.raw, arg21.raw, arg22.raw, arg23.raw, arg24.raw));
        public XlRange Union(XlRange arg1, XlRange arg2, XlRange arg3, XlRange arg4, XlRange arg5, XlRange arg6, XlRange arg7, XlRange arg8, XlRange arg9, XlRange arg10, XlRange arg11, XlRange arg12, XlRange arg13, XlRange arg14, XlRange arg15, XlRange arg16, XlRange arg17, XlRange arg18, XlRange arg19, XlRange arg20, XlRange arg21, XlRange arg22, XlRange arg23, XlRange arg24, XlRange arg25) => new(raw.Union(arg1.raw, arg2.raw, arg3.raw, arg4.raw, arg5.raw, arg6.raw, arg7.raw, arg8.raw, arg9.raw, arg10.raw, arg11.raw, arg12.raw, arg13.raw, arg14.raw, arg15.raw, arg16.raw, arg17.raw, arg18.raw, arg19.raw, arg20.raw, arg21.raw, arg22.raw, arg23.raw, arg24.raw, arg25.raw));
        public XlRange Union(XlRange arg1, XlRange arg2, XlRange arg3, XlRange arg4, XlRange arg5, XlRange arg6, XlRange arg7, XlRange arg8, XlRange arg9, XlRange arg10, XlRange arg11, XlRange arg12, XlRange arg13, XlRange arg14, XlRange arg15, XlRange arg16, XlRange arg17, XlRange arg18, XlRange arg19, XlRange arg20, XlRange arg21, XlRange arg22, XlRange arg23, XlRange arg24, XlRange arg25, XlRange arg26) => new(raw.Union(arg1.raw, arg2.raw, arg3.raw, arg4.raw, arg5.raw, arg6.raw, arg7.raw, arg8.raw, arg9.raw, arg10.raw, arg11.raw, arg12.raw, arg13.raw, arg14.raw, arg15.raw, arg16.raw, arg17.raw, arg18.raw, arg19.raw, arg20.raw, arg21.raw, arg22.raw, arg23.raw, arg24.raw, arg25.raw, arg26.raw));
        public XlRange Union(XlRange arg1, XlRange arg2, XlRange arg3, XlRange arg4, XlRange arg5, XlRange arg6, XlRange arg7, XlRange arg8, XlRange arg9, XlRange arg10, XlRange arg11, XlRange arg12, XlRange arg13, XlRange arg14, XlRange arg15, XlRange arg16, XlRange arg17, XlRange arg18, XlRange arg19, XlRange arg20, XlRange arg21, XlRange arg22, XlRange arg23, XlRange arg24, XlRange arg25, XlRange arg26, XlRange arg27) => new(raw.Union(arg1.raw, arg2.raw, arg3.raw, arg4.raw, arg5.raw, arg6.raw, arg7.raw, arg8.raw, arg9.raw, arg10.raw, arg11.raw, arg12.raw, arg13.raw, arg14.raw, arg15.raw, arg16.raw, arg17.raw, arg18.raw, arg19.raw, arg20.raw, arg21.raw, arg22.raw, arg23.raw, arg24.raw, arg25.raw, arg26.raw, arg27.raw));
        public XlRange Union(XlRange arg1, XlRange arg2, XlRange arg3, XlRange arg4, XlRange arg5, XlRange arg6, XlRange arg7, XlRange arg8, XlRange arg9, XlRange arg10, XlRange arg11, XlRange arg12, XlRange arg13, XlRange arg14, XlRange arg15, XlRange arg16, XlRange arg17, XlRange arg18, XlRange arg19, XlRange arg20, XlRange arg21, XlRange arg22, XlRange arg23, XlRange arg24, XlRange arg25, XlRange arg26, XlRange arg27, XlRange arg28) => new(raw.Union(arg1.raw, arg2.raw, arg3.raw, arg4.raw, arg5.raw, arg6.raw, arg7.raw, arg8.raw, arg9.raw, arg10.raw, arg11.raw, arg12.raw, arg13.raw, arg14.raw, arg15.raw, arg16.raw, arg17.raw, arg18.raw, arg19.raw, arg20.raw, arg21.raw, arg22.raw, arg23.raw, arg24.raw, arg25.raw, arg26.raw, arg27.raw, arg28.raw));
        public XlRange Union(XlRange arg1, XlRange arg2, XlRange arg3, XlRange arg4, XlRange arg5, XlRange arg6, XlRange arg7, XlRange arg8, XlRange arg9, XlRange arg10, XlRange arg11, XlRange arg12, XlRange arg13, XlRange arg14, XlRange arg15, XlRange arg16, XlRange arg17, XlRange arg18, XlRange arg19, XlRange arg20, XlRange arg21, XlRange arg22, XlRange arg23, XlRange arg24, XlRange arg25, XlRange arg26, XlRange arg27, XlRange arg28, XlRange arg29) => new(raw.Union(arg1.raw, arg2.raw, arg3.raw, arg4.raw, arg5.raw, arg6.raw, arg7.raw, arg8.raw, arg9.raw, arg10.raw, arg11.raw, arg12.raw, arg13.raw, arg14.raw, arg15.raw, arg16.raw, arg17.raw, arg18.raw, arg19.raw, arg20.raw, arg21.raw, arg22.raw, arg23.raw, arg24.raw, arg25.raw, arg26.raw, arg27.raw, arg28.raw, arg29.raw));
        public XlRange Union(XlRange arg1, XlRange arg2, XlRange arg3, XlRange arg4, XlRange arg5, XlRange arg6, XlRange arg7, XlRange arg8, XlRange arg9, XlRange arg10, XlRange arg11, XlRange arg12, XlRange arg13, XlRange arg14, XlRange arg15, XlRange arg16, XlRange arg17, XlRange arg18, XlRange arg19, XlRange arg20, XlRange arg21, XlRange arg22, XlRange arg23, XlRange arg24, XlRange arg25, XlRange arg26, XlRange arg27, XlRange arg28, XlRange arg29, XlRange arg30) => new(raw.Union(arg1.raw, arg2.raw, arg3.raw, arg4.raw, arg5.raw, arg6.raw, arg7.raw, arg8.raw, arg9.raw, arg10.raw, arg11.raw, arg12.raw, arg13.raw, arg14.raw, arg15.raw, arg16.raw, arg17.raw, arg18.raw, arg19.raw, arg20.raw, arg21.raw, arg22.raw, arg23.raw, arg24.raw, arg25.raw, arg26.raw, arg27.raw, arg28.raw, arg29.raw, arg30.raw));

        /// <summary></summary>
        /// <param name="application"></param>
        /// <see href="https://docs.microsoft.com/ja-jp/dotnet/api/microsoft.office.interop.excel._application.activatemicrosoftapp?view=excel-pia" />
        public void ActivateMicrosoftApp(XlMsApplication application) => raw.ActivateMicrosoftApp((Microsoft.Office.Interop.Excel.XlMSApplication)application);

        /// <summary></summary>
        /// <param name="chart"></param>
        /// <param name="name"></param>
        /// <see href="https://docs.microsoft.com/ja-jp/dotnet/api/microsoft.office.interop.excel._application.addchartautoformat?view=excel-pia" />
        public void AddChartAutoFormat(object chart, string name) => raw.AddChartAutoFormat(chart, name);
        public void AddChartAutoFormat(object chart, string name, object description) => raw.AddChartAutoFormat(chart, name, description);

        /// <summary>ユーザー設定リストに追加する</summary>
        /// <param name="listArray">追加する文字列を配列で指定</param>
        /// <see href="https://docs.microsoft.com/ja-jp/dotnet/api/microsoft.office.interop.excel._application.addcustomlist?view=excel-pia" />
        public void AddCustomList(string[] listArray) => raw.AddCustomList(listArray);
        /// <summary>ユーザー設定リストに追加する</summary>
        /// <param name="listArray">追加する文字列を配列で指定</param>
        /// <param name="byRow">行単位の場合はtrue, 列単位の場合はfalseを指定</param>
        /// <see href="https://docs.microsoft.com/ja-jp/dotnet/api/microsoft.office.interop.excel._application.addcustomlist?view=excel-pia" />
        public void AddCustomList(string[] listArray, bool byRow) => raw.AddCustomList(listArray, byRow);
        /// <summary>ユーザー設定リストに追加する</summary>
        /// <param name="listArray">追加する文字列をセル範囲で指定</param>
        /// <see href="https://docs.microsoft.com/ja-jp/dotnet/api/microsoft.office.interop.excel._application.addcustomlist?view=excel-pia" />
        public void AddCustomList(XlRange listArray) => raw.AddCustomList(listArray.raw);
        /// <summary>ユーザー設定リストに追加する</summary>
        /// <param name="listArray">追加する文字列をセル範囲で指定</param>
        /// <param name="byRow">行単位の場合はtrue, 列単位の場合はfalseを指定</param>
        /// <see href="https://docs.microsoft.com/ja-jp/dotnet/api/microsoft.office.interop.excel._application.addcustomlist?view=excel-pia" />
        public void AddCustomList(XlRange listArray, bool byRow) => raw.AddCustomList(listArray.raw, byRow);

        /// <summary></summary>
        /// <param name="centimeters"></param>
        /// <see href="https://docs.microsoft.com/ja-jp/dotnet/api/microsoft.office.interop.excel._application.centimeterstopoints?view=excel-pia" />
        public void CentimetersToPoints(double centimeters) => raw.CentimetersToPoints(centimeters);

        /// <summary></summary>
        /// <param name="word"></param>
        /// <returns></returns>
        /// <see href="https://docs.microsoft.com/ja-jp/dotnet/api/microsoft.office.interop.excel._application.checkspelling?view=excel-pia" />
        public bool CheckSpelling(string word) => raw.CheckSpelling(Word: word);
        public bool CheckSpelling(string word, string customDirectoryPath) => raw.CheckSpelling(Word: word, CustomDictionary: customDirectoryPath);
        public bool CheckSpelling(string word, bool ignoreUppercase) => raw.CheckSpelling(Word: word, IgnoreUppercase: ignoreUppercase);
        public bool CheckSpelling(string word, string customDirectoryPath, bool ignoreUppercase) => raw.CheckSpelling(Word: word, CustomDictionary: customDirectoryPath, IgnoreUppercase: ignoreUppercase);
        
        // TODO: 戻り値の型を調査する.
        /// <summary></summary>
        /// <returns></returns>
        /// <see href="https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel._application.convertformula?view=excel-pia" />
        public object ConvertFormula(string formula, XlReferenceStyle fromReferenceStyle) => raw.ConvertFormula(Formula: formula, FromReferenceStyle: (MicrosoftXlReferenceStyle)fromReferenceStyle);
        public object ConvertFormulaRef(string formula, XlReferenceStyle fromReferenceStyle, XlReferenceStyle toReferenceStyle) => raw.ConvertFormula(Formula: formula, FromReferenceStyle: (MicrosoftXlReferenceStyle)fromReferenceStyle, ToReferenceStyle: (MicrosoftXlReferenceStyle)toReferenceStyle);
        public object ConvertFormulaAbs(string formula, XlReferenceStyle fromReferenceStyle, XlReferenceStyle toAbsolute) => raw.ConvertFormula(Formula: formula, FromReferenceStyle: (MicrosoftXlReferenceStyle)fromReferenceStyle, ToAbsolute: (MicrosoftXlReferenceStyle)toAbsolute);
        public object ConvertFormula(string formula, XlReferenceStyle fromReferenceStyle, XlReferenceStyle toReferenceStyle, XlReferenceStyle toAbsolute) => raw.ConvertFormula(Formula: formula, FromReferenceStyle: (MicrosoftXlReferenceStyle)fromReferenceStyle, ToReferenceStyle: (MicrosoftXlReferenceStyle)toReferenceStyle, ToAbsolute: (MicrosoftXlReferenceStyle)toAbsolute);
        public object ConvertFormula(string formula, XlReferenceStyle fromReferenceStyle, XlRange relativeTo) => raw.ConvertFormula(Formula: formula, FromReferenceStyle: (MicrosoftXlReferenceStyle)fromReferenceStyle, relativeTo.raw);
        public object ConvertFormulaRef(string formula, XlReferenceStyle fromReferenceStyle, XlReferenceStyle toReferenceStyle, XlRange relativeTo) => raw.ConvertFormula(Formula: formula, FromReferenceStyle: (MicrosoftXlReferenceStyle)fromReferenceStyle, ToReferenceStyle: (MicrosoftXlReferenceStyle)toReferenceStyle, RelativeTo: relativeTo.raw);
        public object ConvertFormulaAbs(string formula, XlReferenceStyle fromReferenceStyle, XlReferenceStyle toAbsolute, XlRange relativeTo) => raw.ConvertFormula(Formula: formula, FromReferenceStyle: (MicrosoftXlReferenceStyle)fromReferenceStyle, ToAbsolute: (MicrosoftXlReferenceStyle)toAbsolute, RelativeTo: relativeTo.raw);
        public object ConvertFormula(string formula, XlReferenceStyle fromReferenceStyle, XlReferenceStyle toReferenceStyle, XlReferenceStyle toAbsolute, XlRange relativeTo) => raw.ConvertFormula(Formula: formula, FromReferenceStyle: (MicrosoftXlReferenceStyle)fromReferenceStyle, ToReferenceStyle: (MicrosoftXlReferenceStyle)toReferenceStyle, ToAbsolute: (MicrosoftXlReferenceStyle)toAbsolute, RelativeTo: relativeTo.raw);

        /// <summary></summary>
        /// <param name="name"></param>
        /// <see href="https://docs.microsoft.com/ja-jp/dotnet/api/microsoft.office.interop.excel._application.deletechartautoformat?view=excel-pia" />
        public void DeleteChartAutoFormat(string name) => raw.DeleteChartAutoFormat(name);
        /// <summary></summary>
        /// <param name="listNumber"></param>
        /// <see href="https://docs.microsoft.com/ja-jp/dotnet/api/microsoft.office.interop.excel._application.deletecustomlist?view=excel-pia" />
        public void DeleteCustomList(int listNumber) => raw.DeleteCustomList(listNumber);
        /// <summary></summary>
        /// <see href="https://docs.microsoft.com/ja-jp/dotnet/api/microsoft.office.interop.excel._application.doubleclick?view=excel-pia" />
        public void DoubleClick() => raw.DoubleClick();
        /// <summary></summary>
        /// <param name="listNumber"></param>
        /// <returns></returns>
        /// <see href="https://docs.microsoft.com/ja-jp/dotnet/api/microsoft.office.interop.excel._application.getcustomlistcontents?view=excel-pia#Microsoft_Office_Interop_Excel__Application_GetCustomListContents_System_Int32_" />
        public string[] GetCustomListContents(int listNumber) => raw.GetCustomListContents(listNumber);
        /// <summary></summary>
        /// <param name="list"></param>
        /// <returns></returns>
        /// <see href="https://docs.microsoft.com/ja-jp/dotnet/api/microsoft.office.interop.excel._application.getcustomlistnum?view=excel-pia" />
        public int GetCustomListNum(string[] list) => raw.GetCustomListNum(list);

        // TODO: 
        /// <summary></summary>
        /// <param name="fileFilter"></param>
        /// <param name="filterIndex"></param>
        /// <param name="title"></param>
        /// <param name="buttonText"></param>
        /// <returns></returns>
        /// <see href="https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel._application.getopenfilename?view=excel-pia" />
        public string GetOpenFilename(string? fileFilter = null, int filterIndex = 1, string? title = null) => (string)raw.GetOpenFilename(FileFilter: fileFilter, FilterIndex: filterIndex, Title: title, MultiSelect: false);
        //public string GetOpenFilename(string fileFilter = "All Files (.),.", int filterIndex = 1, string title = "Open") => (string)raw.GetOpenFilename(FileFilter: fileFilter, FilterIndex: filterIndex, Title: title, MultiSelect: false);

        /// <summary></summary>
        /// <param name="fileFilter"></param>
        /// <param name="filterIndex"></param>
        /// <param name="title"></param>
        /// <param name="buttonText"></param>
        /// <returns></returns>
        /// <see href="https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel._application.getopenfilename?view=excel-pia" />
        public string[] GetOpenMultiFilename(string? fileFilter = null, int filterIndex = 1, string? title = null) => (string[])raw.GetOpenFilename(FileFilter: fileFilter, FilterIndex: filterIndex, Title: title, MultiSelect: true);
        //public string[] GetOpenMultiFilename(string fileFilter = "All Files (.),.", int filterIndex = 1, string title = "Open") => (string[])raw.GetOpenFilename(FileFilter: fileFilter, FilterIndex: filterIndex, Title: title, MultiSelect: true);

        // TODO: 
        /// <summary></summary>
        /// <param name="initialFilename"></param>
        /// <param name="fileFilter"></param>
        /// <param name="filterIndex"></param>
        /// <param name="title"></param>
        /// <returns></returns>
        /// <see href="https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel._application.getsaveasfilename?view=excel-pia" />
        public string GetSaveAsFilename(string? initialFilename = null, string? fileFilter = null, int filterIndex = 1, string? title = null) => (string)raw.GetSaveAsFilename(InitialFilename: initialFilename, FileFilter: fileFilter, FilterIndex: filterIndex, Title: title);
        //public object GetSaveAsFilename(string? initialFilename = null, string fileFilter = "All Files (.),.", int filterIndex = 1, string title = "Save As") => raw.GetSaveAsFilename(FileFilter: fileFilter, FilterIndex: filterIndex, Title: title);

        /// <summary></summary>
        /// <see href="https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel._application.goto?view=excel-pia" />
        public void Goto() => raw.Goto();
        public void Goto(XlRange reference, bool scroll = false) => raw.Goto(Reference: reference.raw, Scroll: scroll);
        public void Goto(string reference, bool scroll = false) => raw.Goto(Reference: reference, Scroll: scroll);

        /// <summary></summary>
        /// <see href="https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel._application.help?view=excel-pia" />
        public void Help() => raw.Help();
        public void Help(string helpFile) => raw.Help(HelpFile: helpFile);
        public void Help(string helpFile, int helpContextID) => raw.Help(HelpFile: helpFile, HelpContextID: helpContextID);

        /// <summary></summary>
        /// <param name="Inches"></param>
        /// <returns></returns>
        /// <see href="https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel._application.inchestopoints?view=excel-pia" />
        public double InchesToPoints(double Inches) => raw.InchesToPoints(Inches);

        // TODO:
        /// <summary></summary>
        /// <param name="prompt"></param>
        /// <param name="title"></param>
        /// <param name="defaultValue"></param>
        /// <param name="left"></param>
        /// <param name="top"></param>
        /// <param name="helpFile"></param>
        /// <param name="helpContextID"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        /// <see href="https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel._application.inputbox?view=excel-pia" />
        /// <see href="https://docs.microsoft.com/en-us/office/vba/api/excel.application.inputbox" />
        public object InputBox(string prompt, string? title = null, string? defaultValue = null, double? left = null, double? top = null, string? helpFile = null, int? helpContextID = null, XlInputType type = XlInputType.String) =>
            raw.InputBox(Prompt: prompt, Title: title, Default: defaultValue, Left: left, Top: top, HelpFile: helpFile, HelpContextID: helpContextID, Type: type);

        // TODO:
        /// <summary></summary>
        /// <param name="macro"></param>
        /// <param name="description"></param>
        /// <param name="hasShortcutKey"></param>
        /// <param name="shortcutKey"></param>
        /// <param name="category"></param>
        /// <param name="statusBar"></param>
        /// <param name="helpFile"></param>
        /// <param name="helpContextID"></param>
        /// <see href="https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel._application.macrooptions?view=excel-pia" />
        public void MacroOptions(string? macro = null, string? description = null, bool hasShortcutKey = false, string? shortcutKey = null, XlMacroOptionsCategory? category = null, string? statusBar = null, string? helpFile = null, int? helpContextID = null) => raw.MacroOptions(Macro: macro, Description: description, HasShortcutKey: hasShortcutKey, ShortcutKey: shortcutKey, Category: category, StatusBar: statusBar, HelpContextID: helpContextID, HelpFile: helpFile);

        // TODO:
        /// <summary></summary>
        /// <param name="macro"></param>
        /// <param name="description"></param>
        /// <param name="hasShortcutKey"></param>
        /// <param name="shortcutKey"></param>
        /// <param name="category"></param>
        /// <param name="statusBar"></param>
        /// <param name="helpFile"></param>
        /// <param name="helpContextID"></param>
        /// <param name="argumentDescriptions"></param>
        /// <see href="https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel._application.macrooptions2?view=excel-pia" />
        public void MacroOptions2(string? macro = null, string? description = null, bool hasShortcutKey = false, string? shortcutKey = null, XlMacroOptionsCategory? category = null, string? statusBar = null, string? helpFile = null, int? helpContextID = null, string? argumentDescriptions = null) => raw.MacroOptions2(Macro: macro, Description: description, HasShortcutKey: hasShortcutKey, ShortcutKey: shortcutKey, Category: category, StatusBar: statusBar, HelpContextID: helpContextID, HelpFile: helpFile, ArgumentDescriptions: argumentDescriptions);

        /// <summary></summary>
        /// <see href="https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel._application.maillogoff?view=excel-pia" />
        public void MailLogoff() => raw.MailLogoff();

        // TODO:
        /// <summary></summary>
        /// <param name="name"></param>
        /// <param name="password"></param>
        /// <param name="downloadNewMail"></param>
        /// <see href="https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel._application.maillogon?view=excel-pia" />
        public void MailLogon(string? name = null, string? password = null, bool? downloadNewMail = null) => raw.MailLogon(name, password, downloadNewMail);

        /// <summary></summary>
        /// <returns></returns>
        /// <see href="https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel._application.nextletter?view=excel-pia" />
        public XlWorkbook NextLetter() => new XlWorkbook(raw.NextLetter());

        // TODO:
        /// <summary></summary>
        /// <param name="key"></param>
        /// <param name="procedure"></param>
        /// <see href="https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel._application.onkey?view=excel-pia" />
        public void OnKey(string key, string? procedure = null) => raw.OnKey(key, procedure);

        /// <summary></summary>
        /// <param name="text"></param>
        /// <param name="procedure"></param>
        /// <see href="https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel._application.onrepeat?view=excel-pia" />
        public void OnRepeat(
            [In][MarshalAs(UnmanagedType.BStr)] string text, 
            [In][MarshalAs(UnmanagedType.BStr)] string procedure
        ) =>
            raw.OnRepeat(text, procedure);

        /// <summary></summary>
        /// <param name="earliestTime"></param>
        /// <param name="procedure"></param>
        /// <param name="latestTime"></param>
        /// <param name="schedule"></param>
        /// <see href="https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel._application.ontime?view=excel-pia" />
        public void OnTime(
            [In][MarshalAs(UnmanagedType.Struct)] DateTime earliestTime,
            [In][MarshalAs(UnmanagedType.BStr)] string procedure,
            [Optional][In][MarshalAs(UnmanagedType.Struct)] DateTime latestTime, 
            [Optional][In][MarshalAs(UnmanagedType.Struct)] bool schedule
        ) => 
            raw.OnTime(earliestTime, procedure, latestTime, schedule);

        /// <summary></summary>
        /// <param name="text"></param>
        /// <param name="procedure"></param>
        /// <see href="https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel._application.onundo?view=excel-pia" />
        public void OnUndo(
            [In][MarshalAs(UnmanagedType.BStr)] string text, 
            [In][MarshalAs(UnmanagedType.BStr)] string procedure
        ) => 
            raw.OnUndo(text, procedure);

        /// <summary></summary>
        /// <see href="https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel._application.quit?view=excel-pia" />
        public void Quit() => raw.Quit();

        /// <summary></summary>
        /// <param name="basicCode"></param>
        /// <param name="xlmCode"></param>
        /// <see href="https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel._application.recordmacro?view=excel-pia" />
        public void RecordMacro(
            [Optional][In][MarshalAs(UnmanagedType.Struct)] string basicCode, 
            [Optional][In][MarshalAs(UnmanagedType.Struct)] string xlmCode
        ) => 
            raw.RecordMacro(basicCode, xlmCode);

        /// <summary></summary>
        /// <param name="filename"></param>
        /// <returns></returns>
        /// <see href="https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel._application.registerxll?view=excel-pia" />
        public bool RegisterXLL([In][MarshalAs(UnmanagedType.BStr)] string filename) => raw.RegisterXLL(filename);

        /// <summary></summary>
        /// <see href="https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel._application.repeat?view=excel-pia" />
        public void Repeat() => raw.Repeat();

        /// <summary></summary>
        /// /// <see href="https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel._application.repeat?view=excel-pia" />
        public void ResetTipWizard() => raw.ResetTipWizard();

        /// <summary></summary>
        /// <param name="filename"></param>
        /// <see href="https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel._application.save?view=excel-pia" />
        public void Save([Optional][In][MarshalAs(UnmanagedType.Struct)] string filename) => raw.Save(filename);

        /// <summary></summary>
        /// <param name="filename"></param>
        /// <see href="https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel._application.saveworkspace?view=excel-pia" />
        public void SaveWorkspace([Optional][In][MarshalAs(UnmanagedType.Struct)] string filename) => raw.SaveWorkspace(filename);

        /// <summary></summary>
        /// <param name="formatName"></param>
        /// <param name="gallery"></param>
        /// <see href="https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel._application.setdefaultchart?view=excel-pia" />
        public void SetDefaultChart(
            [Optional][In][MarshalAs(UnmanagedType.Struct)] string formatName, 
            [Optional][In][MarshalAs(UnmanagedType.Struct)] string gallery
        ) => 
            raw.SetDefaultChart(formatName, gallery);

        /// <summary></summary>
        /// <see href="https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel._application.undo?view=excel-pia" />
        public void Undo() => raw.Undo();

        /// <summary></summary>
        /// <param name="isVolatile"></param>
        /// <see href="https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel._application.volatile?view=excel-pia#Microsoft_Office_Interop_Excel__Application_Volatile_System_Object_" />
        public void Volatile([Optional][In][MarshalAs(UnmanagedType.Struct)] bool isVolatile) => raw.Volatile(isVolatile);

        /// <summary></summary>
        /// <param name="time"></param>
        /// <returns></returns>
        /// <see href="https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel._application.wait?view=excel-pia" />
        public bool Wait([In][MarshalAs(UnmanagedType.Struct)] DateTime time) => raw.Wait(time);

        /// <summary></summary>
        /// <param name="text"></param>
        /// <returns></returns>
        /// <see href="https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel._application.getphonetic?view=excel-pia#Microsoft_Office_Interop_Excel__Application_GetPhonetic_System_Object_" />
        public string GetPhonetic([Optional][In][MarshalAs(UnmanagedType.Struct)] string text) => raw.GetPhonetic(text);

        /// <summary></summary>
        /// <see href="https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel._application.calculatefull?view=excel-pia#Microsoft_Office_Interop_Excel__Application_CalculateFull" />
        public void CalculateFull() => raw.CalculateFull();

        /// <summary></summary>
        /// <see href="https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel._application.calculatefullrebuild?view=excel-pia" />
        public void CalculateFullRebuild() => raw.CalculateFullRebuild();

        /// <summary></summary>
        /// <returns></returns>
        /// <see href="https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel._application.findfile?view=excel-pia" />
        public bool FindFile() => raw.FindFile();

        /// <summary></summary>
        /// <param name="keepAbort"></param>
        /// <see href="https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel._application.checkabort?view=excel-pia" />
        public void CheckAbort([Optional][In][MarshalAs(UnmanagedType.Struct)] bool keepAbort) => raw.CheckAbort(keepAbort);

        /// <summary></summary>
        /// <param name="xmlMap"></param>
        public void DisplayXMLSourcePane([Optional][In] XlXmlMap xmlMap) => raw.DisplayXMLSourcePane(xmlMap.raw);

    }
}
