namespace Fxcel.Core

open System
open System.Buffers
open System.Runtime.InteropServices

module Common =
  [<Measure>] type handle
  [<Measure>] type id

  let internal s_ok = 0x00000000
  let internal e_fail = 0x80004005
  let internal gw_hwnd_next = 2
  let internal wm_close = 0x0010

  let rent'<'T> (length: int)= ArrayPool<'T>.Shared.Rent(length)
  let return'<'T> (array: array<'T>, clear: bool)= ArrayPool<'T>.Shared.Return(array, clear)

  let to_handle (h: int) = LanguagePrimitives.Int32WithMeasure<handle> h
  let to_id (id: int) = LanguagePrimitives.Int32WithMeasure<id> id
