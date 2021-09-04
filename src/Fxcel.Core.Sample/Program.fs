open System
//open Fxcel.Core
//open Fxcel.Core.Excel
//open Fxcel.Core.Common

open Fxcel.Core.Interop

let inline (?) (lhs: obj) (rhs: string) =
  $"{lhs.GetType().ToString()}:{rhs}"

let inline (?<-) (lhs: obj) (rhs: string) ([<ParamArray>] args) =
  $"{lhs.GetType().ToString()}:{rhs}, {args}"
//let inline (?) (lhs: int) (rhs: int) = lhs + rhs

[<EntryPoint>]
let main argv =
  //use excel = Excel.new' ()
  //printfn "%A" excel
  //let com = Com.new'<MicrosoftExcel> Interop.excel'id

  //try
  //  let com = Type.GetTypeFromCLSID(Interop.excel'id) |> Activator.CreateInstance :?> MicrosoftExcel
  //  //let com = Type.GetTypeFromCLSID(Interop.excel'id) |> Activator.CreateInstance :?> ExcelInterop.IApplication
  //  
  //  //let com' = com.Application
  //  let books = com.Workbooks
  //  //let com'' = com.Application
  //  //Com.release' com''
  //  Com.release' books
  //  //Com.release' com'
  //  Com.release' com
  //  printfn "ok"
  //with
  //| e -> printfn $"{e.Message}"
  
  //let a = ( ? ) 0 ""

  //let (.@) (lhs: int) (rhs: int) = lhs + rhs
  //(0 .@ 0) |> printfn "%d"

  (0?aaa <- "bbb", "ccc") |> printfn "%s"
  




  let app = XlApplication()



  0 // return an integer exit code
