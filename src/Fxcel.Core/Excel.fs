namespace Fxcel.Core

open System
open System.Runtime.CompilerServices
open Fxcel.Core.Common

module Excel =
  type internal MicrosoftExcel = Microsoft.Office.Interop.Excel.Application
  type internal MicrosoftWorkbook = Microsoft.Office.Interop.Excel.Workbook
  type internal MicrosoftWorksheet = Microsoft.Office.Interop.Excel.Worksheet
  type internal DisposeStatus = { mutable Disposed: bool }
  
  /// <summary>Excel Worksheet</summary>
  [<IsReadOnly;Struct;>]
  type Worksheet internal (worksheet: MicrosoftWorksheet) =
    member __.Name with get() : string = worksheet.Name and set(name) = worksheet.Name <- name

  /// <summary>Excel Workbook</summary>
  [<IsReadOnly;Struct;>]
  type Workbook internal (workbook: MicrosoftWorkbook) =
    member __.Name with get() = workbook.Name

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

  let create () =
    let excel = Com.new'<MicrosoftExcel> Interop.excel'id
    excel.IgnoreRemoteRequests <- true
    excel.DisplayAlerts <- false
    excel.Visible <- false
    new Application (excel, { Disposed= false })
