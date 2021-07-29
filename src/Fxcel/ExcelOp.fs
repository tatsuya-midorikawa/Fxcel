namespace Fxcel

open Midoliy.Office.Interop

[<AutoOpen>]
module ExcelOp =
  type ExcelOpBuilder (excel: IExcelApplication) =
    member __.Yield (_: unit) = excel
    member __.Zero() = ()
    [<CustomOperation("set")>]
    member __.Set(excel: IExcelApplication, mode: Calculation) = excel.Calculation <- mode; excel
    [<CustomOperation("set")>]
    member __.Set(excel: IExcelApplication, visibiliy: AppVisibility) = excel.Visibility <- visibiliy; excel

  let excel'op excel = ExcelOpBuilder excel
