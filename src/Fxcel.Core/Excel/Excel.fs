namespace Fxcel.Core.Excel

open Fxcel.Core
open Fxcel.Core.Common

module Excel = 
  let private new_app () =
    let com = Com.new'<MicrosoftExcel> Interop.excel'id
    new Application (com, { Disposed= false }, ResizeArray<Workbook>(), ResizeArray<Workbook>())
  
  // TODO:
  /// <summary></summary>
  let new' () =
    let excel = new_app () 
    excel.blank_workbook() |> ignore
    excel

  // TODO:
  /// <summary></summary>
  let open' (file: string) =
    let excel = new_app () 
    file |> excel.open_file |> ignore
    excel

  // TODO:
  /// <summary></summary>
  let from' (template: string) =
    let excel = new_app () 
    template |> excel.create_from |> ignore
    excel
