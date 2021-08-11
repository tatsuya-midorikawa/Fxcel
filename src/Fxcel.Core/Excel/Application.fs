namespace Fxcel.Core.Excel

open System
open System.IO
open System.Runtime.CompilerServices
open Fxcel.Core
open Fxcel.Core.Common
open Fxcel.Core.Excel.Constant

module Application =
  type internal MicrosoftCalculation = Microsoft.Office.Interop.Excel.XlCalculation

/// <summary>Excel Application</summary>
[<IsReadOnly;Struct;>]
type Application internal (excel: MicrosoftExcel, status: DisposeStatus, workbooks: ResizeArray<Workbook>) =
  interface IDisposable with
    member __.Dispose() = __.dispose()
    
  /// <summary></summary>
  member __.window_handle with get () : int<handle> = excel.Hwnd |> to_handle
  /// <summary></summary>
  member __.ignore_remote_requests with get () : bool = excel.IgnoreRemoteRequests
  member __.set_ignore_remote_requests (ignore_remote_requests: bool) = excel.IgnoreRemoteRequests <- ignore_remote_requests
  /// <summary></summary>
  member __.display_alerts with get () : bool = excel.DisplayAlerts
  member __.set_display_alerts (display_alerts: bool) = excel.DisplayAlerts <- display_alerts
  /// <summary></summary>
  member __.visible with get () : bool = excel.Visible
  member __.set_visible (visible: bool) = excel.Visible <- visible
  /// <summary></summary>
  member __.calculation with get () : Calculation = excel.Calculation |> (int >> to_enum<Calculation>)
  member __.set_calculations (calculation: Calculation) = excel.Calculation  <- calculation |> (int >> to_enum<Application.MicrosoftCalculation>)
    /// <summary></summary>
  member __.active_workbook 
    with get () : Workbook = 
      let book = new Workbook(excel.ActiveWorkbook, { Disposed = false })
      workbooks.Add(book)
      book

  /// <summary></summary>
  member __.blank_workbook () =
    let book = new Workbook (excel.Workbooks.Add (), { Disposed = false })
    workbooks.Add(book)
    book
  /// <summary></summary>
  member __.open_file (file: string) =
    let book = new Workbook (Path.GetFullPath(file) |> excel.Workbooks.Open, { Disposed = false })
    workbooks.Add(book)
    book
  /// <summary></summary>
  member __.create_from (template: string) =
    let book = new Workbook (excel.Workbooks.Add (Path.GetFullPath(template)), { Disposed = false })
    workbooks.Add(book)
    book

  /// <summary></summary>
  member __.quit () = excel.Quit()
  /// <summary></summary>
  member __.undo () = excel.Undo()
  /// <summary>Run excel macro.</summary>
  member __.run (macro: string, ?arg1: obj, ?arg2: obj, ?arg3: obj, ?arg4: obj, ?arg5: obj, ?arg6: obj, ?arg7: obj, ?arg8: obj, ?arg9: obj, ?arg10: obj, ?arg11: obj, ?arg12: obj, ?arg13: obj, ?arg14: obj, ?arg15: obj, ?arg16: obj, ?arg17: obj, ?arg18: obj, ?arg19: obj, ?arg20: obj, ?arg21: obj, ?arg22: obj, ?arg23: obj, ?arg24: obj, ?arg25: obj, ?arg26: obj, ?arg27: obj, ?arg28: obj, ?arg29: obj, ?arg30: obj) = 
    excel.Run(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30)
  // TODO
  /// <summary>Run excel macro.</summary>
  member __.input_box() = excel.InputBox("Prompt")

  /// <summary></summary>
  member __.dispose () =
    if not status.Disposed then
      // 子要素を解放
      workbooks |> Seq.iter (fun wb -> wb.dispose())
      // 自分自身を解放
      if not __.visible then
        __.quit ()
        Process.kill __.window_handle
      Com.release' excel
      // 後始末
      status.Disposed <- true
      GC.Collect()
