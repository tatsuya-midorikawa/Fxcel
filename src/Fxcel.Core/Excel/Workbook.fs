﻿namespace Fxcel.Core.Excel

open System
open System.Runtime.CompilerServices
open Fxcel.Core
open Fxcel.Core.Common

/// <summary>Excel Workbook</summary>
[<IsReadOnly;Struct;>]
type Workbook internal (workbook: MicrosoftWorkbook, status: DisposeStatus) =
  interface IDisposable with
    member __.Dispose() = __.dispose()
  /// <summary></summary>
  member __.name with get() = workbook.Name
  /// <summary></summary>
  member __.dispose() =
    if not status.Disposed then
      Com.release' workbook
      status.Disposed <- true
