namespace Fxcel.Core

#if DEBUG
open System.Diagnostics
#endif

open System.Runtime.InteropServices
open System.Runtime.InteropServices.ComTypes
open Microsoft.Office.Interop.Excel
open Fxcel.Core.Common
open Fxcel.Core.Natives

module Process =
  /// <summary>Excel Processを列挙する.</summary>
  let enumerate () = System.Diagnostics.Process.GetProcessesByName "Excel"

  /// <summary>Excel Processにアタッチする.</summary>
  let attach (hwnd: int<handle>) =
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
          // TODO: Fxcel.Core.ExcelApplicationでwrap
          | :? Workbook as wb -> wb
          | _ ->
            release' com
            loop table' monikers' fetchedMonikers'
        else
          release' com
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

  /// <summary>Window HandleからProcess Idを取得する.</summary>
  let get_pid (hwnd: int<handle>) =
    let mutable pid = 0
    Win32.get_window_thread_process_id(int hwnd, &pid) |> ignore
    pid |> to_id
    
  /// <summary>Process IdからWindow Handleを取得する.</summary>
  let get_hwnd (pid: int<id>) =
    let rec loop (pid': int<id>) (hwnd': int<handle>) =
      match hwnd' with
      | 0<handle> -> hwnd'
      | _ ->
        let hwnd = int hwnd'
        if Win32.get_parent hwnd = 0 && Win32.is_window_visible hwnd <> 0 && pid' = get_pid hwnd' then hwnd'
        else loop pid (to_handle (Win32.get_window (int hwnd', gw_hwnd_next)))
      
    loop pid (Win32.find_window (null, null) |> to_handle)
