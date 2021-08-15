using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using MicrosoftApplication = Microsoft.Office.Interop.Excel.Application;
using MicrosoftWorksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public readonly struct XlApplication : IComObject
    {
        internal static readonly List<XlApplication> apps = new();
        internal readonly MicrosoftApplication raw;
        internal XlApplication(MicrosoftApplication excel) => raw = excel;

        public int ComRelease() => ComHelper.Release(raw);

        internal XlApplication Application => new XlApplication(raw.Application);
        internal XlCreator Creator => (XlCreator)raw.Creator;
        internal XlApplication Parent => new XlApplication(raw.Parent);
        internal XlRange ActiveCell => new XlRange(raw.ActiveCell);
        internal XlChart ActiveChart => new XlChart(raw.ActiveChart);
        internal XlDialogSheet ActiveDialog => new XlDialogSheet(raw.ActiveDialog);
        internal XlMenuBar ActiveMenuBar => new XlMenuBar(raw.ActiveMenuBar);
        internal string ActivePrinter => raw.ActivePrinter;
        internal XlWorksheet ActiveSheet => new XlWorksheet((MicrosoftWorksheet)raw.ActiveSheet);
        internal XlWindow ActiveWindow => new XlWindow(raw.ActiveWindow);
        internal XlWorkbook ActiveWorkbook => new XlWorkbook(raw.ActiveWorkbook);
        internal XlAddIns AddIns => new XlAddIns(raw.AddIns);
        internal XlAssistant Assistant => new XlAssistant(raw.Assistant);
        internal XlRange Cells => new XlRange(raw.Cells);
        internal XlSheets Charts => new XlSheets(raw.Charts);
        internal XlRange Columns => new XlRange(raw.Columns);
        internal int DDEAppReturnCode => raw.DDEAppReturnCode;
        internal XlSheets DialogSheets => new XlSheets(raw.DialogSheets);
        internal XlMenuBars MenuBars => new XlMenuBars(raw.MenuBars);
        internal XlModules Modules => new XlModules(raw.Modules);
        internal XlNames Names => new XlNames(raw.Names);
        internal XlRange Rows => new XlRange(raw.Rows);
        internal object Selection => raw.Selection;
        internal XlSheets Sheets => new XlSheets(raw.Sheets);
        internal XlMenu ShortcutMenus(int index) => new XlMenu(raw.ShortcutMenus[index]);
        internal XlWorkbook ThisWorkbook => new XlWorkbook(raw.ThisWorkbook);
        internal XlToolbars Toolbars => new XlToolbars(raw.Toolbars);
        internal XlWindows Windows => new XlWindows(raw.Windows);
        internal XlWorkbooks Workbooks => new XlWorkbooks(raw.Workbooks);
        internal XlWorksheetFunction WorksheetFunction => new XlWorksheetFunction(raw.WorksheetFunction);
        internal XlSheets Worksheets => new XlSheets(raw.Worksheets);
        internal XlSheets Excel4IntlMacroSheets => new XlSheets(raw.Excel4IntlMacroSheets);
        internal XlSheets Excel4MacroSheets => new XlSheets(raw.Excel4MacroSheets);
        internal bool AlertBeforeOverwriting => raw.AlertBeforeOverwriting;
        internal string AltStartupPath => raw.AltStartupPath;
        internal bool AskToUpdateLinks => raw.AskToUpdateLinks;
        internal bool EnableAnimations => raw.EnableAnimations;
        internal XlAutoCorrect AutoCorrect => new XlAutoCorrect(raw.AutoCorrect);
        internal int Build => raw.Build;
        internal bool CalculateBeforeSave => raw.CalculateBeforeSave;
        internal XlCalculation Calculation => (XlCalculation)raw.Calculation;
        internal object Caller => raw.Caller;
        internal bool CanPlaySounds => raw.CanPlaySounds;
        internal bool CanRecordSounds => raw.CanRecordSounds;
        internal string Caption => raw.Caption;
        internal bool CellDragAndDrop => raw.CellDragAndDrop;
        internal XlClipboardFormat[] ClipboardFormats => ((object[])raw.ClipboardFormats).Select(f => (XlClipboardFormat)f).ToArray();
    }
}
