namespace Fxcel.Core.Excel

open System
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

  member __.window_handle with get() : int<handle> = excel.Hwnd |> to_handle
  member __.ignore_remote_requests with get() : bool = excel.IgnoreRemoteRequests and set(v) = excel.IgnoreRemoteRequests <- v
  member __.display_alerts with get() : bool = excel.DisplayAlerts and set(v) = excel.DisplayAlerts <- v
  member __.visible with get() : bool = excel.Visible and set (v) = excel.Visible <- v
  member __.calculation 
    with get() : Calculation = excel.Calculation |> (int >> to_enum<Calculation>)
    and set(v: Calculation) = excel.Calculation  <- v |> (int >> to_enum<Application.MicrosoftCalculation>)

  member __.blank_workbook () =
    let book = new Workbook (excel.Workbooks.Add(), { Disposed = false })
    workbooks.Add(book)
    book

  member __.quit () = excel.Quit()

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

module Excel =        
  /// <summary></summary>
  let create () =
    let excel = Com.new'<MicrosoftExcel> Interop.excel'id
    new Application (excel, { Disposed= false }, ResizeArray<Workbook>())
