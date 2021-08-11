#r "nuget: System.Runtime.InteropServices, 4.3.0"
#r @"..\bin\Debug\net5.0\Fxcel.Core.dll"
#r @"..\obj\Debug\net5.0\Interop.Microsoft.Office.Core.dll"
#r @"..\obj\Debug\net5.0\Interop.Microsoft.Office.Interop.Excel.dll"

open Fxcel.Core
open System
open System.Runtime.CompilerServices

type MicrosoftExcel = Microsoft.Office.Interop.Excel.Application
let excel = Com.new'<MicrosoftExcel> Interop.excel'id

printfn $"%A{excel}"

Com.release' excel
