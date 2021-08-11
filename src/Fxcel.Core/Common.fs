namespace Fxcel.Core

open System
open System.Buffers
open System.Runtime.InteropServices

module Common =
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

  let inline rent'<'T> (length: int)= ArrayPool<'T>.Shared.Rent(length)
  let inline return'<'T> (array: array<'T>, clear: bool)= ArrayPool<'T>.Shared.Return(array, clear)

  let inline to_handle (h: int) = LanguagePrimitives.Int32WithMeasure<handle> h
  let inline to_id (id: int) = LanguagePrimitives.Int32WithMeasure<id> id
