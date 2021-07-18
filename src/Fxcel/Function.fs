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

  static member cellsi (row: IExcelRow) : seq<(int * IExcelRange)> =
    let enumerator = row.GetEnumerator()
    let mutable i = 0
    seq {
      while enumerator.MoveNext() do
        i <- i + 1
        yield (i, enumerator.Current)
    }

  static member cellsi (column: IExcelColumn) : seq<(int * IExcelRange)> =
    let enumerator = column.GetEnumerator()
    let mutable i = 0
    seq {
      while enumerator.MoveNext() do
        i <- i + 1
        yield (i, enumerator.Current)
    }

[<AutoOpen>]
module Function =
  type Color = System.Drawing.Color
  type Pattern = Midoliy.Office.Interop.Pattern
  type DeleteShiftDirection = Midoliy.Office.Interop.DeleteShiftDirection
  type AppVisibility = Midoliy.Office.Interop.AppVisibility

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

  /// <summary>対象のセル, 行, 列などを削除する</summary>
  let inline delete (direction: DeleteShiftDirection) (range: ^Range) = (^Range: (member Delete: DeleteShiftDirection -> bool) range, direction)
  
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

  /// <summary>Rangeなどの範囲選択した場所から行単位で列挙する(index付き).</summary>
  let inline rowsi (range: IExcelRange) : seq<(int * IExcelRow)> =
    let enumerator = range.Rows.GetEnumerator()
    let mutable i = 0
    seq {
      while enumerator.MoveNext() do
        i <- i + 1
        yield (i, enumerator.Current)
    }

  /// <summary>Rangeなどの範囲選択した場所から列単位で列挙する.</summary>
  let inline columns (range: IExcelRange) : seq<IExcelColumn> =
    let enumerator = range.Columns.GetEnumerator()
    seq {
      while enumerator.MoveNext() do
        yield enumerator.Current
    }

  /// <summary>Rangeなどの範囲選択した場所から列単位で列挙する(index付き).</summary>
  let inline columnsi (range: IExcelRange) : seq<(int * IExcelColumn)> =
    let enumerator = range.Columns.GetEnumerator()
    let mutable i = 0
    seq {
      while enumerator.MoveNext() do
        i <- i + 1
        yield (i, enumerator.Current)
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
  
  let inline cast<'T> (value: obj) =
    match typeof<'T> with
    | t when t = typeof<bool> -> System.Convert.ToBoolean value |> box :?> 'T
    | t when t = typeof<int8> -> System.Convert.ToSByte value |> box :?> 'T
    | t when t = typeof<int16> -> System.Convert.ToInt16 value |> box :?> 'T
    | t when t = typeof<int> -> System.Convert.ToInt32 value |> box :?> 'T
    | t when t = typeof<int64> -> System.Convert.ToInt64 value |> box :?> 'T
    | t when t = typeof<uint8> -> System.Convert.ToByte value |> box :?> 'T
    | t when t = typeof<uint16> -> System.Convert.ToUInt16 value |> box :?> 'T
    | t when t = typeof<uint> -> System.Convert.ToUInt32 value |> box :?> 'T
    | t when t = typeof<uint64> -> System.Convert.ToUInt64 value |> box :?> 'T
    | t when t = typeof<float> -> System.Convert.ToDouble value |> box :?> 'T
    | t when t = typeof<float32> -> System.Convert.ToSingle value |> box :?> 'T
    | t when t = typeof<decimal> -> System.Convert.ToDecimal value |> box :?> 'T
    | t when t = typeof<System.DateTime> -> System.Convert.ToDateTime value |> box :?> 'T
    | t when t = typeof<string> -> System.Convert.ToString value |> box :?> 'T
    | t when t = typeof<obj> -> value :?> 'T
    | _ -> value :?> 'T
    
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


