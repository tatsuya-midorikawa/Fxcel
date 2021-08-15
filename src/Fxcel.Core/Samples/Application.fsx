﻿#r "nuget: System.Runtime.InteropServices, 4.3.0"
#r @"..\bin\Debug\net5.0\Fxcel.Core.dll"
#r @"..\obj\Debug\net5.0\Interop.Microsoft.Office.Core.dll"
#r @"..\obj\Debug\net5.0\Interop.Microsoft.Office.Interop.Excel.dll"

open Fxcel.Core.Excel
open System
open System.Runtime.CompilerServices

let main() =
  use excel = Excel.new'()
  excel.set_display_alerts false
  excel.set_visible true
  excel.input_box(prompt= "Test") |> printfn "%A"
  excel.get_save_as_filename() |> printfn "%A"

  0

main()
