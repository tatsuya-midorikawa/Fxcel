namespace Fxcel.Core.Excel

open Fxcel.Core
open Fxcel.Core.Common

module Excel = 
  // TODO:
  /// <summary></summary>
  let create () =
    let excel = Com.new'<MicrosoftExcel> Interop.excel'id
    new Application (excel, { Disposed= false }, ResizeArray<Workbook>())
