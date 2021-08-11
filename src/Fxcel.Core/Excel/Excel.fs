namespace Fxcel.Core.Excel

open System
open System.Runtime.CompilerServices
open Fxcel.Core
open Fxcel.Core.Common

/// <summary>Excel Application</summary>
[<IsReadOnly;Struct;>]
type Application internal (excel: MicrosoftExcel, status: DisposeStatus) =
  interface IDisposable with
    member __.Dispose() = __.dispose()

  member __.Hwnd with get() : int<handle> = excel.Hwnd |> to_handle

  member __.dispose() =
    if not status.Disposed then
      excel.IgnoreRemoteRequests <- false
      excel.DisplayAlerts <- true
      Com.release' excel
      status.Disposed <- true
      GC.Collect()

module Excel =        
  /// <summary></summary>
  let create () =
    let excel = Com.new'<MicrosoftExcel> Interop.excel'id
    excel.IgnoreRemoteRequests <- true
    excel.DisplayAlerts <- false
    excel.Visible <- false
    new Application (excel, { Disposed= false })
