namespace Fxcel.Core.Excel

open System
open System.Runtime.CompilerServices
open Fxcel.Core
open Fxcel.Core.Common

/// <summary>Excel Range</summary>
[<IsReadOnly;Struct;>]
type Range internal (range: MicrosoftRange, status: DisposeStatus) =
  member __.Value with get() : obj = range.Value() and set(name) = range.Value() <- name
