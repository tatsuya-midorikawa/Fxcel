open System
open Fxcel.Core
open Fxcel.Core.Excel
open Fxcel.Core.Common

[<EntryPoint>]
let main argv =
  //use excel = Excel.new' ()
  //printfn "%A" excel
  //let com = Com.new'<MicrosoftExcel> Interop.excel'id

  try
    let com = Type.GetTypeFromCLSID(Interop.excel'id) |> Activator.CreateInstance :?> MicrosoftExcel
    //let com = Type.GetTypeFromCLSID(Interop.excel'id) |> Activator.CreateInstance :?> ExcelInterop.IApplication
    
    //let com' = com.Application
    let books = com.Workbooks
    //let com'' = com.Application
    //Com.release' com''
    Com.release' books
    //Com.release' com'
    Com.release' com
    printfn "ok"
  with
  | e -> printfn $"{e.Message}"
  

  0 // return an integer exit code
