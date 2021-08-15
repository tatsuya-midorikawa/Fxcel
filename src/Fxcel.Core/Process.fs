namespace Fxcel.Core

open System.Diagnostics
open System.Runtime.InteropServices.ComTypes
open Microsoft.Office.Interop.Excel
open Fxcel.Core.Common
open Fxcel.Core.Natives

module Process =
  /// <summary>Window HandleからProcess Idを取得する.</summary>
  let inline get_pid (hwnd: int<handle>) =
    let mutable pid = 0
    Win32.get_window_thread_process_id(int hwnd, &pid) |> ignore
    pid |> to_id
    
  /// <summary>Process IdからWindow Handleを取得する.</summary>
  let inline get_hwnd (pid: int<id>) =
    let rec loop (pid': int<id>) (hwnd': int<handle>) =
      match hwnd' with
      | 0<handle> -> hwnd'
      | _ ->
        let hwnd = int hwnd'
        if Win32.get_parent hwnd = 0 && Win32.is_window_visible hwnd <> 0 && pid' = get_pid hwnd' then hwnd'
        else loop pid (to_handle (Win32.get_window (int hwnd', gw_hwnd_next)))
      
    loop pid (Win32.find_window (null, null) |> to_handle)

  /// <summary>Excel Processを列挙する.</summary>
  let inline enumerate () = System.Diagnostics.Process.GetProcessesByName "Excel"
  
  /// <summary>対象のプロセスを終了する.</summary>
  let inline kill (hwnd: int<handle>) =
    try
      let pid = get_pid hwnd
      Win32.send_message (int hwnd, wm_close, 0n, 0n) |> ignore
      Process.GetProcessById(int pid).Kill(true)
    with _ -> ()

  /// <summary>Excel Processにアタッチする.</summary>
  let inline attach (hwnd: int<handle>) =
    let rec loop (table': IRunningObjectTable) (monikers': IEnumMoniker) (fetchedMonikers': nativeint) =
      #if DEBUG
      printfn $"frame count= %d{StackTrace().FrameCount}"
      #endif

      let container : array<IMoniker> = [| null |]
      match monikers'.Next(1, container, fetchedMonikers') with
      | 0 ->
        let mutable com = null
        if table'.GetObject(container.[0], &com) = s_ok then
          match com with
          // TODO: Fxcel.Core.Excel.Applicationでwrap
          | :? Workbook as wb -> wb
          | _ ->
            Com.release' com
            loop table' monikers' fetchedMonikers'
        else
          Com.release' com
          loop table' monikers' fetchedMonikers'
      | _ -> 
        raise (exn "The HWND is not found.")
        
    let mutable table : IRunningObjectTable = null
    let mutable monikers : IEnumMoniker = null

    if Win32.get_running_object_table(0, &table) <> 0 || table = null then
      raise (exn "Running object table is not found.")

    table.EnumRunning &monikers
    monikers.Reset ()

    loop table monikers 0n
