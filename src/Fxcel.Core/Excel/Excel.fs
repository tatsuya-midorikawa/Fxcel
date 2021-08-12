namespace Fxcel.Core.Excel

open Fxcel.Core
open Fxcel.Core.Common

module Excel = 
  // TODO:
  /// <summary></summary>
  let create () =
    let com = Com.new'<MicrosoftExcel> Interop.excel'id
    let excel = new Application (com, { Disposed= false }, ResizeArray<Workbook>())
    excel.blank_workbook() |> ignore
    excel
