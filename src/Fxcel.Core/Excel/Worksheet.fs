namespace Fxcel.Core.Excel

open System
open System.Runtime.CompilerServices
open Fxcel.Core
open Fxcel.Core.Common

/// <summary>Excel Worksheet</summary>
[<IsReadOnly;Struct;>]
type Worksheet internal (worksheet: MicrosoftWorksheet, status: DisposeStatus) =
  interface IDisposable with
    member __.Dispose() = __.dispose()

  member __.Name with get() : string = worksheet.Name and set(name) = worksheet.Name <- name

  member __.dispose() =
    if not status.Disposed then
      Com.release' worksheet
      status.Disposed <- true
