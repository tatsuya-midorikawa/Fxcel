namespace Fxcel.Core

open System
open System.Buffers
open System.Runtime.InteropServices
open Microsoft.FSharp.NativeInterop

module Common =
  type internal MicrosoftExcel = Microsoft.Office.Interop.Excel.Application
  type internal MicrosoftWorkbook = Microsoft.Office.Interop.Excel.Workbook
  type internal MicrosoftWorksheet = Microsoft.Office.Interop.Excel.Worksheet
  type internal MicrosoftRange = Microsoft.Office.Interop.Excel.Range
  type internal DisposeStatus = { mutable Disposed: bool }

  [<Measure>] type handle
  [<Measure>] type id

  [<Literal>]
  let internal s_ok = 0x00000000
  [<Literal>]
  let internal e_fail = 0x80004005
  [<Literal>]
  let internal gw_hwnd_next = 2
  [<Literal>]
  let internal wm_close = 0x0010

  [<Unverifiable>]
  let inline internal stackalloc<'T when 'T: unmanaged>(count: int) = NativePtr.stackalloc<'T>(count: int)

  let inline rent'<'T> (length: int)= ArrayPool<'T>.Shared.Rent(length)
  let inline return'<'T> (array: array<'T>, clear: bool)= ArrayPool<'T>.Shared.Return(array, clear)

  let inline to_handle (h: int) = LanguagePrimitives.Int32WithMeasure<handle> h
  let inline to_id (id: int) = LanguagePrimitives.Int32WithMeasure<id> id
  let to_enum<'T when 'T: enum<int>> (n: int) = LanguagePrimitives.EnumOfValue<int, 'T> n

