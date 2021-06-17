namespace Fxcel

open System.IO
open System.Collections.Generic
open Midoliy.Office.Interop

[<AutoOpen>]
type Excel () =
  static member cells (row: IExcelRow) : seq<IExcelRange> =
    let enumerator = row.GetEnumerator()
    seq {
      while enumerator.MoveNext() do
        yield enumerator.Current
    }

  static member cells (column: IExcelColumn) : seq<IExcelRange> =
    let enumerator = column.GetEnumerator()
    seq {
      while enumerator.MoveNext() do
        yield enumerator.Current
    }

[<AutoOpen>]
module Function =
  type Color = System.Drawing.Color
  type Pattern = Midoliy.Office.Interop.Pattern
  
  [<Struct>]
  type Handle = { Name: string; Hwnd: int }

  let private isNullOrEmpty value = System.String.IsNullOrEmpty(value)
  
  let private getExcelPath (path: string) =
    let extension = Path.GetExtension path
    match extension with
    | ".xls" | ".xlsx" | ".xlsm" -> path
    | _ -> $"{path}.xlsx"

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
  
  /// <summary>空のワークブックを新規作成する.</summary>
  let create () = Excel.BlankWorkbook()

  /// <summary>テンプレートファイルからワークブックを新規作成する.</summary>
  let createFrom (template: string) = Excel.CreateFrom(getExcelPath template)

  /// <summary>既存のワークブックを開く.</summary>
  let open' (filepath: string) = Excel.Open(getExcelPath filepath)

  /// <summary>指定したindexの位置にあるWorkbookを取得する.</summary>
  let workbook (index: int) (excel: IExcelApplication) =
    if index <= 0 then raise (exn "index は 1 以上で指定してください")
    else excel.[index]
    
  /// <summary>指定したindexの位置にあるWorksheetを取得する.</summary>
  let worksheet (target: obj) (book: IWorkbook) =
    match target with
    | :? string as name -> if isNullOrEmpty name then book.[1] else book.[name]
    | :? int as index -> if index <= 0 then book.[1] else book.[index]
    | _ -> book.[1]
  
  /// <summary>
  /// WorkbookやWorksheet, Cellなどを選択する.
  /// activate関数で選択した場合, 単一選択となる.
  /// </summary>
  let inline activate (target: ^T) = (^T: (member Activate: unit -> unit) target)
  
  /// <summary>
  /// WorkbookやWorksheet, Cellなどを選択する.
  /// select関数で選択した場合, 複数選択となる.
  /// 複数選択を解除したい場合, activate関数を特定のWorkbookやWorksheet, Cellに対して呼び出す.
  /// </summary>
  let inline select (target: ^T) = (^T: (member Select: unit -> unit) target)
  
  /// <summary>Cellなどからアドレス文字列を取得する</summary>
  let inline address (cell: ^Cell) = (^Cell: (member get_Address: unit -> string) cell)

  /// <summary>Cellなどから値を取得する</summary>
  let inline value (cell: ^Cell) = (^Cell: (member get_Value: unit -> obj) cell)
  
  /// <summary>Rangeなどの範囲選択したCellから値を取得し配列情報に変換する.</summary>
  let inline values (range: ^Range) = 
    let vs = value range
    if vs.GetType() = typeof<obj[,]> then
      let xs = vs :?> obj[,]
      xs.[*,*]
    else
      Array2D.init 1 1 (fun i j -> vs)
      
  /// <summary>Rangeなどの範囲選択した場所から行単位で列挙する.</summary>
  let inline rows (range: IExcelRange) : seq<IExcelRow> =
    let enumerator = range.Rows.GetEnumerator()
    seq {
      while enumerator.MoveNext() do
        yield enumerator.Current
    }

  /// <summary>Rangeなどの範囲選択した場所から列単位で列挙する.</summary>
  let inline columns (range: IExcelRange) : seq<IExcelColumn> =
    let enumerator = range.Columns.GetEnumerator()
    seq {
      while enumerator.MoveNext() do
        yield enumerator.Current
    }
    
  /// <summary>Applies the given function to each element of the collection.</summary>
  let iter = Seq.iter

  /// <summary>
  /// Applies the given function to each element of the collection.
  /// The integer passed to the function indicates the index of element.
  /// The index starts at 1.
  /// </summary>
  let inline iteri (action: int -> ^T -> unit) (source: seq< ^T>) =
    let mutable i = 1
    for x in source do
      action i x
      i <- i + 1

  /// <summary>gets関数で取得した配列の長さ情報を取得する.</summary>
  /// <return>(row数, column数)</return>
  let len (range: #obj[,]) = (Array2D.length1 range, Array2D.length2 range)

  /// <summary>
  /// Cellなどから値を取得する.
  /// int型と互換性がない値の場合, 例外が発生する.
  /// </summary>
  let inline integer (cell: ^Cell) = value cell |> System.Convert.ToInt32

  /// <summary>
  /// Cellなどから値を取得する.
  /// float型と互換性がない値の場合, 例外が発生する.
  /// </summary>
  let inline number (cell: ^Cell) = value cell |> System.Convert.ToDouble

  /// <summary>
  /// Cellなどから値を取得する.
  /// decimal型と互換性がない値の場合, 例外が発生する.
  /// </summary>
  let inline money (cell: ^Cell) = value cell |> System.Convert.ToDecimal

  /// <summary>
  /// Cellなどから値を取得する.
  /// string型と互換性がない値の場合, 例外が発生する.
  /// </summary>
  let inline str (cell: ^Cell) = value cell |> System.Convert.ToString
  
  /// <summary>
  /// Cellなどから値を取得する.
  /// DateTime型と互換性がない値の場合, 例外が発生する.
  /// </summary>
  let inline date (cell: ^Cell) = value cell |> System.Convert.ToDateTime

  /// <summary>Cellなどに値を設定する.</summary>
  let inline set value (cell: ^Cell) = (^Cell: (member set_Value: obj -> unit) cell, value)

  /// <summary>Cellなどに関数を設定する.</summary>
  let inline fx value (cell: ^Cell) = (^Cell: (member set_Formula: obj -> unit) cell, if (string value).StartsWith("=") then value else $"={value}")
  
  /// <summary>Cellなどに背景色を設定する.</summary>
  let inline bgcolor (color: Color) (cell: IExcelRange) = cell.Interior.Color <- color

  /// <summary>Cellなどに背景色パターンを設定する.</summary>
  let inline bgpattern (pattern: Pattern) (cell: IExcelRange) = cell.Interior.Pattern <- pattern

