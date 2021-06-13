namespace Fxcel

open Midoliy.Office.Interop

[<AutoOpen>]
module Function =

  [<Struct>]
  type Handle = { Name: string; Hwnd: int }

  let private isNullOrEmpty value = System.String.IsNullOrEmpty(value)

  /// <summary>起動しているExcelプロセスを列挙する.</summary>
  let enumerate () = 
    Excel.EnumerateProcess()
    |> Array.map (fun p -> { Name = p.MainWindowTitle; Hwnd = int p.MainWindowHandle })
    
  /// <summary>起動しているExcelプロセスを表示する.</summary>
  let show () =
    let ps = enumerate()
    ps 
    |> Array.iteri (fun i handle -> 
      printfn $"[{i}] Active workbook= {handle.Name}"
      use excel = Excel.Attach handle.Hwnd
      excel |> Seq.iteri (fun j wb -> printfn $"  workbook({j+1})= {wb.Name}"))
    ps

  /// <summary>handleがExcelの場合アタッチする.</summary>
  let attach (handle: Handle) = Excel.Attach handle.Hwnd
  
  /// <summary>プログラムをExcelからデタッチする.</summary>
  let detach (excel: IExcelApplication) = excel.Dispose()

  /// <summary>指定したindexの位置にあるWorkbookを取得する.</summary>
  let workbook (index: int) (excel: IExcelApplication) =
    if index <= 0 then
      raise (exn "index は 1 以上で指定してください")
    excel.[index]
    
  /// <summary>指定したindexの位置にあるWorksheetを取得する.</summary>
  let worksheet (target: obj) (book: IWorkbook) =
    match target with
    | :? string as name -> if isNullOrEmpty name then book.[1] else book.[name]
    | :? int as index -> if index <= 0 then book.[1] else book.[index]
    | _ -> book.[1]

  /// <summary>Cell/Range/Row/Columnなど、Valueプロパティを持つインスタンスに対して値を取得する</summary>
  let inline get (cell: ^a) = (^a: (member get_Value: unit -> obj) cell)

  /// <summary>
  /// Cell/Range/Row/Columnなど、Valueプロパティを持つインスタンスに対して値を取得する.
  /// float型と互換性がない値の場合、例外が発生する.
  /// </summary>
  let inline number (cell: ^a) = get cell |> System.Convert.ToDouble

  /// <summary>
  /// Cell/Range/Row/Columnなど、Valueプロパティを持つインスタンスに対して値を取得する.
  /// string型と互換性がない値の場合、例外が発生する.
  /// </summary>
  let inline str (cell: ^a) = get cell |> System.Convert.ToString
  
  /// <summary>
  /// Cell/Range/Row/Columnなど、Valueプロパティを持つインスタンスに対して値を取得する.
  /// DateTime型と互換性がない値の場合、例外が発生する.
  /// </summary>
  let inline date (cell: ^a) = get cell |> System.Convert.ToDateTime

  /// <summary>Cell/Range/Row/Columnなど、Valueプロパティを持つインスタンスに対して値を設定する.</summary>
  let inline set value (cell: ^a) = (^a: (member set_Value: obj -> unit) cell, value)

  /// <summary>Cell/Range/Row/Columnなど、Formulaプロパティを持つインスタンスに対して値を設定する.</summary>
  let inline fx value (cell: ^a) =
    (^a: (member set_Formula: obj -> unit) cell, if (string value).StartsWith("=") then value else $"={value}")
