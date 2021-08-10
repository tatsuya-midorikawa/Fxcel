namespace Fxcel.Core

module Excel =
  type internal MicrosoftExcel = Microsoft.Office.Interop.Excel.Application
  
  /// <summary></summary>
  type Application () = 
    let excel = Com.new'<MicrosoftExcel> Interop.excel'id

    do
      excel.IgnoreRemoteRequests <- true
      excel.DisplayAlerts <- false
      excel.Visible <- false
