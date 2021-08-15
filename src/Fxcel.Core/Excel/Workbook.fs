namespace Fxcel.Core.Excel

open System
open System.Collections
open System.Collections.Generic
open System.Runtime.CompilerServices
open Fxcel.Core
open Fxcel.Core.Common

/// <summary>Excel Workbook</summary>
[<Struct;IsReadOnly;NoComparison;>]
type Workbook internal (workbook: MicrosoftWorkbook, status: DisposeStatus, worksheets: ResizeArray<Worksheet>) =
  interface IDisposable with
    member __.Dispose() = __.dispose()
  
  interface IEnumerable<Worksheet> with
    member __.GetEnumerator() = (worksheets :> IEnumerable<Worksheet>).GetEnumerator()
  
  interface IEnumerable with
    member __.GetEnumerator() = (worksheets :> IEnumerable).GetEnumerator()
  
  /// <summary></summary>
  [<ComponentModel.DataAnnotations.Range(1, 512, ErrorMessage= "Value for {0} must be between {0} and {1}")>]
  member __.Item with get (index: int) = worksheets.[index - 1]

  /// <summary></summary>
  member __.name with get() = workbook.Name
  
  /// <summary></summary>
  member __.activate () = workbook.Activate()
  /// <summary></summary>
  member __.save () = workbook.Save()
  /// <summary></summary>
  member __.save_as (filepath: string) = workbook.SaveAs(filepath)

  /// <summary></summary>
  member __.dispose() =
    if not status.Disposed then
      Com.release' workbook
      status.Disposed <- true
