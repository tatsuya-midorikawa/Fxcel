namespace Fxcel.Core.Excel

open System
open System.IO
open System.Runtime.CompilerServices
open Fxcel.Core
open Fxcel.Core.Common
open Fxcel.Core.Excel.Constant

module Application =
  type internal MicrosoftCalculation = Microsoft.Office.Interop.Excel.XlCalculation
  
  let internal to_nullable<'T when 'T: struct and 'T: (new: unit -> 'T) and 'T :> ValueType> (value: Option<'T>) = match value with Some value -> Nullable value | None -> Nullable()
  let internal unwrap<'T when 'T: not struct and 'T: null> (value: Option<'T>) = match value with Some value -> value | None -> null

  /// <summary></summary>
  [<Flags>]
  type InputBoxType = Formula = 0 | Number = 1 | String = 2 | Boolean = 4 | RangeObject = 8 | Error = 16 | Array = 64

/// <summary>Excel Application</summary>
[<IsReadOnly;Struct;>]
type Application internal (excel: MicrosoftExcel, status: DisposeStatus, workbooks: ResizeArray<Workbook>) =
  interface IDisposable with
    member __.Dispose() = __.dispose()
    
  /// <summary></summary>
  [<ComponentModel.DataAnnotations.Range(1, 255, ErrorMessage= "Value for {0} must be between {0} and {1}")>]
  member __.Item with get (index: int) = workbooks.[index - 1]
  /// <summary></summary>
  member __.Item with get (name: string) = workbooks |> Seq.find (fun wb -> wb.name = name)
  
  /// <summary></summary>
  member __.window_handle with get () : int<handle> = excel.Hwnd |> to_handle
  /// <summary></summary>
  member __.ignore_remote_requests with get () : bool = excel.IgnoreRemoteRequests
  member __.set_ignore_remote_requests (ignore_remote_requests: bool) = excel.IgnoreRemoteRequests <- ignore_remote_requests
  /// <summary></summary>
  member __.display_alerts with get () : bool = excel.DisplayAlerts
  member __.set_display_alerts (display_alerts: bool) = excel.DisplayAlerts <- display_alerts
  /// <summary></summary>
  member __.visible with get () : bool = excel.Visible
  member __.set_visible (visible: bool) = excel.Visible <- visible
  /// <summary></summary>
  member __.calculation with get () : Calculation = excel.Calculation |> (int >> to_enum<Calculation>)
  member __.set_calculations (calculation: Calculation) = excel.Calculation  <- calculation |> (int >> to_enum<Application.MicrosoftCalculation>)
  /// <summary></summary>
  member __.active_workbook 
    with get () : Workbook = 
      let book = new Workbook(excel.ActiveWorkbook, { Disposed = false })
      workbooks.Add(book)
      book

  /// <summary></summary>
  member __.blank_workbook () =
    let book = new Workbook (excel.Workbooks.Add (), { Disposed = false })
    workbooks.Add(book)
    book
  /// <summary></summary>
  member __.open_file (file: string) =
    let book = new Workbook (Path.GetFullPath(file) |> excel.Workbooks.Open, { Disposed = false })
    workbooks.Add(book)
    book
  /// <summary></summary>
  member __.create_from (template: string) =
    let book = new Workbook (excel.Workbooks.Add (Path.GetFullPath(template)), { Disposed = false })
    workbooks.Add(book)
    book

  /// <summary>Quit excel application.</summary>
  member __.quit () = excel.Quit()
  /// <summary>Operation undo.</summary>
  member __.undo () = excel.Undo()
  /// <summary>Run excel macro.</summary>
  member __.run (macro: string, ?arg1: obj, ?arg2: obj, ?arg3: obj, ?arg4: obj, ?arg5: obj, ?arg6: obj, ?arg7: obj, ?arg8: obj, ?arg9: obj, ?arg10: obj, ?arg11: obj, ?arg12: obj, ?arg13: obj, ?arg14: obj, ?arg15: obj) = 
    match (arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15) with
    | (None, None, None, None, None, None, None, None, None, None, None, None, None, None, None) -> excel.Run(macro, arg1)
    | (Some arg1, None, None, None, None, None, None, None, None, None, None, None, None, None, None) -> excel.Run(macro, arg1)
    | (Some arg1, Some arg2, None, None, None, None, None, None, None, None, None, None, None, None, None) -> excel.Run(macro, arg1, arg2)
    | (Some arg1, Some arg2, Some arg3, None, None, None, None, None, None, None, None, None, None, None, None) -> excel.Run(macro, arg1, arg2, arg3)
    | (Some arg1, Some arg2, Some arg3, Some arg4, None, None, None, None, None, None, None, None, None, None, None) -> excel.Run(macro, arg1, arg2, arg3, arg4)
    | (Some arg1, Some arg2, Some arg3, Some arg4, Some arg5, None, None, None, None, None, None, None, None, None, None) -> excel.Run(macro, arg1, arg2, arg3, arg4, arg5)
    | (Some arg1, Some arg2, Some arg3, Some arg4, Some arg5, Some arg6, None, None, None, None, None, None, None, None, None) -> excel.Run(macro, arg1, arg2, arg3, arg4, arg5, arg6)
    | (Some arg1, Some arg2, Some arg3, Some arg4, Some arg5, Some arg6, Some arg7, None, None, None, None, None, None, None, None) -> excel.Run(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7)
    | (Some arg1, Some arg2, Some arg3, Some arg4, Some arg5, Some arg6, Some arg7, Some arg8, None, None, None, None, None, None, None) -> excel.Run(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8)
    | (Some arg1, Some arg2, Some arg3, Some arg4, Some arg5, Some arg6, Some arg7, Some arg8, Some arg9, None, None, None, None, None, None) -> excel.Run(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9)
    | (Some arg1, Some arg2, Some arg3, Some arg4, Some arg5, Some arg6, Some arg7, Some arg8, Some arg9, Some arg10, None, None, None, None, None) -> excel.Run(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10)
    | (Some arg1, Some arg2, Some arg3, Some arg4, Some arg5, Some arg6, Some arg7, Some arg8, Some arg9, Some arg10, Some arg11, None, None, None, None) -> excel.Run(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11)
    | (Some arg1, Some arg2, Some arg3, Some arg4, Some arg5, Some arg6, Some arg7, Some arg8, Some arg9, Some arg10, Some arg11, Some arg12, None, None, None) -> excel.Run(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12)
    | (Some arg1, Some arg2, Some arg3, Some arg4, Some arg5, Some arg6, Some arg7, Some arg8, Some arg9, Some arg10, Some arg11, Some arg12, Some arg13, None, None) -> excel.Run(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13)
    | (Some arg1, Some arg2, Some arg3, Some arg4, Some arg5, Some arg6, Some arg7, Some arg8, Some arg9, Some arg10, Some arg11, Some arg12, Some arg13, Some arg14, None) -> excel.Run(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14)
    | (Some arg1, Some arg2, Some arg3, Some arg4, Some arg5, Some arg6, Some arg7, Some arg8, Some arg9, Some arg10, Some arg11, Some arg12, Some arg13, Some arg14, Some arg15) -> excel.Run(macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15)
    | _ -> raise (NotSupportedException())
    
  /// <summary>Show input box.</summary>
  member __.input_box(prompt: string, ?title: string, ?default'input: string, ?xpos: int, ?ypos: int, ?help'filepath: string, ?help'context'id: int, ?type': Application.InputBoxType) =
    match (title, default'input, xpos, ypos, help'filepath, help'context'id, type') with
    | (Some title, None, None, None, None, None, None) -> excel.InputBox(Prompt= prompt, Title= title)
    | (None, Some default'input, None, None, None, None, None) -> excel.InputBox(Prompt= prompt, Default= default'input)
    | (None, None, Some xpos, None, None, None, None) -> excel.InputBox(Prompt= prompt, Left= xpos)
    | (None, None, None, Some ypos, None, None, None) -> excel.InputBox(Prompt= prompt, Top= ypos)
    | (None, None, None, None, Some help'filepath, None, None) -> excel.InputBox(Prompt= prompt, HelpFile= help'filepath)
    | (None, None, None, None, None, Some help'context'id, None) -> excel.InputBox(Prompt= prompt, HelpContextID= help'context'id)
    | (None, None, None, None, None, None, Some type') -> excel.InputBox(Prompt= prompt, Type= type')
    | (Some title, Some default'input, None, None, None, None, None) -> excel.InputBox(Prompt= prompt, Title= title, Default= default'input)
    | (Some title, None, Some xpos, None, None, None, None) -> excel.InputBox(Prompt= prompt, Title= title, Left= xpos)
    | (Some title, None, None, Some ypos, None, None, None) -> excel.InputBox(Prompt= prompt, Title= title, Top= ypos)
    | (Some title, None, None, None, Some help'filepath, None, None) -> excel.InputBox(Prompt= prompt, Title= title, HelpFile= help'filepath)
    | (Some title, None, None, None, None, Some help'context'id, None) -> excel.InputBox(Prompt= prompt, Title= title, HelpContextID= help'context'id)
    | (Some title, None, None, None, None, None, Some type') -> excel.InputBox(Prompt= prompt, Title= title, Type= type')
    | (Some title, Some default'input, Some xpos, None, None, None, None) -> excel.InputBox(Prompt= prompt, Title= title, Default= default'input, Left= xpos)
    | (Some title, Some default'input, None, Some ypos, None, None, None) -> excel.InputBox(Prompt= prompt, Title= title, Default= default'input, Top= ypos)
    | (Some title, Some default'input, Some xpos, Some ypos, None, None, None) -> excel.InputBox(Prompt= prompt, Title= title, Default= default'input, Left= xpos, Top= ypos)
    | (Some title, Some default'input, Some xpos, None, None, None, Some type') -> excel.InputBox(Prompt= prompt, Title= title, Default= default'input, Left= xpos, Type= type')
    | (Some title, Some default'input, None, Some ypos, None, None, Some type') -> excel.InputBox(Prompt= prompt, Title= title, Default= default'input, Top= ypos, Type= type')
    | (Some title, Some default'input, Some xpos, Some ypos, None, None, Some type') -> excel.InputBox(Prompt= prompt, Title= title, Default= default'input, Left= xpos, Top= ypos, Type= type')
    | (Some title, Some default'input, Some xpos, None, Some help'filepath, None, None) -> excel.InputBox(Prompt= prompt, Title= title, Default= default'input, Left= xpos, HelpFile= help'filepath)
    | (Some title, Some default'input, None, Some ypos, Some help'filepath, None, None) -> excel.InputBox(Prompt= prompt, Title= title, Default= default'input, Top= ypos, HelpFile= help'filepath)
    | (Some title, Some default'input, Some xpos, Some ypos, Some help'filepath, None, None) -> excel.InputBox(Prompt= prompt, Title= title, Default= default'input, Left= xpos, Top= ypos, HelpFile= help'filepath)
    | (Some title, Some default'input, Some xpos, None, Some help'filepath, None, Some type') -> excel.InputBox(Prompt= prompt, Title= title, Default= default'input, Left= xpos, HelpFile= help'filepath, Type= type')
    | (Some title, Some default'input, None, Some ypos, Some help'filepath, None, Some type') -> excel.InputBox(Prompt= prompt, Title= title, Default= default'input, Top= ypos, HelpFile= help'filepath, Type= type')
    | (Some title, Some default'input, Some xpos, Some ypos, Some help'filepath, None, Some type') -> excel.InputBox(Prompt= prompt, Title= title, Default= default'input, Left= xpos, Top= ypos, HelpFile= help'filepath, Type= type')
    | (Some title, Some default'input, Some xpos, None, Some help'filepath, Some help'context'id, None) -> excel.InputBox(Prompt= prompt, Title= title, Default= default'input, Left= xpos, HelpFile= help'filepath, HelpContextID= help'context'id)
    | (Some title, Some default'input, None, Some ypos, Some help'filepath, Some help'context'id, None) -> excel.InputBox(Prompt= prompt, Title= title, Default= default'input, Top= ypos, HelpFile= help'filepath, HelpContextID= help'context'id)
    | (Some title, Some default'input, Some xpos, Some ypos, Some help'filepath, Some help'context'id, None) -> excel.InputBox(Prompt= prompt, Title= title, Default= default'input, Left= xpos, Top= ypos, HelpFile= help'filepath, HelpContextID= help'context'id)
    | (Some title, Some default'input, Some xpos, None, Some help'filepath, Some help'context'id, Some type') -> excel.InputBox(Prompt= prompt, Title= title, Default= default'input, Left= xpos, HelpFile= help'filepath, HelpContextID= help'context'id, Type= type')
    | (Some title, Some default'input, None, Some ypos, Some help'filepath, Some help'context'id, Some type') -> excel.InputBox(Prompt= prompt, Title= title, Default= default'input, Top= ypos, HelpFile= help'filepath, HelpContextID= help'context'id, Type= type')
    | (Some title, Some default'input, Some xpos, Some ypos, Some help'filepath, Some help'context'id, Some type') -> excel.InputBox(Prompt= prompt, Title= title, Default= default'input, Left= xpos, Top= ypos, HelpFile= help'filepath, HelpContextID= help'context'id, Type= type')
    | (Some title, None, None, None, Some help'filepath, Some help'context'id, None) -> excel.InputBox(Prompt= prompt, Title= title, HelpFile= help'filepath, HelpContextID= help'context'id)
    | (Some title, None, None, None, Some help'filepath, Some help'context'id, Some type') -> excel.InputBox(Prompt= prompt, Title= title, HelpFile= help'filepath, HelpContextID= help'context'id, Type= type')
    | (Some title, None, None, None, Some help'filepath, None, Some type') -> excel.InputBox(Prompt= prompt, Title= title, HelpFile= help'filepath, Type= type')
    | (None, Some default'input, None, None, Some help'filepath, Some help'context'id, None) -> excel.InputBox(Prompt= prompt, Default= default'input, HelpFile= help'filepath, HelpContextID= help'context'id)
    | (None, Some default'input, None, None, Some help'filepath, Some help'context'id, Some type') -> excel.InputBox(Prompt= prompt, Default= default'input, HelpFile= help'filepath, HelpContextID= help'context'id, Type= type')
    | (None, Some default'input, None, None, Some help'filepath, None, Some type') -> excel.InputBox(Prompt= prompt, Default= default'input, HelpFile= help'filepath, Type= type')
    | (None, None, None, None, Some help'filepath, Some help'context'id, None) -> excel.InputBox(Prompt= prompt, HelpFile= help'filepath, HelpContextID= help'context'id)
    | (None, None, None, None, Some help'filepath, Some help'context'id, Some type') -> excel.InputBox(Prompt= prompt, HelpFile= help'filepath, HelpContextID= help'context'id, Type= type')
    | (None, None, None, None, Some help'filepath, None, Some type') -> excel.InputBox(Prompt= prompt, HelpFile= help'filepath, Type= type')
    | (Some title, Some default'input, None, None, Some help'filepath, Some help'context'id, None) -> excel.InputBox(Prompt= prompt, Title= title, Default= default'input, HelpFile= help'filepath, HelpContextID= help'context'id)
    | (Some title, Some default'input, None, None, Some help'filepath, Some help'context'id, Some type') -> excel.InputBox(Prompt= prompt, Title= title, Default= default'input, HelpFile= help'filepath, HelpContextID= help'context'id, Type= type')
    | (Some title, Some default'input, None, None, Some help'filepath, None, Some type') -> excel.InputBox(Prompt= prompt, Title= title, Default= default'input, HelpFile= help'filepath, Type= type')
    | (None, None, None, None, None, None, None) -> excel.InputBox(Prompt= prompt)
    | _ -> raise (NotSupportedException())
    |> unbox<string>


  /// <summary></summary>
  member __.dispose () =
    if not status.Disposed then
      // 子要素を解放
      workbooks |> Seq.iter (fun wb -> wb.dispose())
      // 自分自身を解放
      if not __.visible then
        __.quit ()
        Process.kill __.window_handle
      Com.release' excel
      // 後始末
      status.Disposed <- true
      GC.Collect()
