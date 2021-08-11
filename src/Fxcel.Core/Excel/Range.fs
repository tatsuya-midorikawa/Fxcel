namespace Fxcel.Core.Excel

open System
open System.Runtime.CompilerServices
open Fxcel.Core
open Fxcel.Core.Common

/// <summary>Excel Range</summary>
[<IsReadOnly;Struct;>]
type Range internal (range: MicrosoftRange, status: DisposeStatus) =
  /// <summary></summary>
  member __.value with get() : obj = range.Value(10) and set(name) = range.Value(10) <- name
