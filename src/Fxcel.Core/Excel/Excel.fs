namespace Fxcel.Core.Excel

open System
open System.Runtime.CompilerServices
open Fxcel.Core
open Fxcel.Core.Common

/// <summary>Excel Application</summary>
[<IsReadOnly;Struct;>]
type Application internal (excel: MicrosoftExcel, status: DisposeStatus, workbooks: ResizeArray<Workbook>) =
  interface IDisposable with
    member __.Dispose() = __.dispose()

  member __.Hwnd with get () : int<handle> = excel.Hwnd |> to_handle

  member __.blank_workbook () =
    let book = new Workbook(excel.Workbooks.Add(), { Disposed = false })
    workbooks.Add(book)
    book

  member __.dispose () =
    if not status.Disposed then
      // 子要素を解放
      workbooks |> Seq.iter (fun wb -> wb.dispose())
      // 自分自身を解放
      excel.IgnoreRemoteRequests <- false
      excel.DisplayAlerts <- true
      Com.release' excel
      // 後始末
      status.Disposed <- true
      GC.Collect()

module Excel =        
  /// <summary></summary>
  let create () =
    let excel = Com.new'<MicrosoftExcel> Interop.excel'id
    //excel.IgnoreRemoteRequests <- true
    //excel.DisplayAlerts <- false
    //excel.Visible <- false
    new Application (excel, { Disposed= false }, ResizeArray<Workbook>())
