namespace Fxcel.Core

open System
open System.Buffers
open System.Runtime.InteropServices

module Com =
  /// <summary>GUIDからCOMインスタンスを生成する.</summary>
  let new'<'T> (cls_id: Guid) = Type.GetTypeFromCLSID(cls_id) |> Activator.CreateInstance :?> 'T
  
  /// <summary>COMオブジェクトを解放する.</summary>
  let release' (com: obj) = if com <> null then while 0 < Marshal.ReleaseComObject(com) do () done
