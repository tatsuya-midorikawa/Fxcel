#r "nuget: System.Runtime.InteropServices, 4.3.0"
#r @"..\bin\Debug\net5.0\Fxcel.Core.dll"
#r @"..\obj\Debug\net5.0\Interop.Microsoft.Office.Core.dll"
#r @"..\obj\Debug\net5.0\Interop.Microsoft.Office.Interop.Excel.dll"

open Fxcel.Core
open Fxcel.Core.Common

let hwnd = Process.get_hwnd 33332<id>
printfn $"hwnd= %d{hwnd}"

let pid = Process.get_pid hwnd
printfn $"pid= %d{pid}"

let com = Process.attach hwnd
printfn $"com= %A{com}"

Com.release' com

Process.kill hwnd
