open Fxcel
open System

let (| Even | Odd |) value = 
  if value % 2 = 0 then Even else Odd

[<Measure>]
type kg

try
  use excel = create()
  //excel.Visibility <- AppVisibility.Visible

  let sheet = excel |> workbook(1) |> worksheet(1)
  
  sheet.["A1:A3"] |> set 100 
  sheet.["B1:B3"] |> fx "SUM(A1:A3)"
  let a = sheet.["A1:A3"] |> getsfx
  let b = sheet.["B1:B3"] |> getsfx
  
  sheet |> get'column(1) |> address |> printfn "%s"
  sheet |> get'columns(1, 3) |> address |> printfn "%s"
  sheet |> get'row(1) |> address |> printfn "%s"
  sheet |> get'rows(1, 3) |> address |> printfn "%s"

  "A" |> column'number |> printfn "%i"     // 1
  "J" |> column'number |> printfn "%i"     // 10
  "DX" |> column'number |> printfn "%i"   // 128

  //sheet.["A1"] |> current'columns |> address |> printfn "%s"
  //sheet.["A1:B3"] |> current'columns |> address |> printfn "%s"
  //sheet.["A1"] |> current'rows |> address |> printfn "%s"
  //sheet.["A1:B3"] |> current'rows |> address |> printfn "%s"

  //sheet.["B1:B3"] |> set 200
  //sheet.["C2"] |> set 200
  //let a = sheet.["C2"] |> getfx
  //let ax = sheet.["A1:B3"] |> gets
  //let h = ax |> head
  //let l = ax |> last
  //let orig = sheet.["A1:B3"] 
  //let v: obj = sheet.["A1:B3"] |> get
  //let v': obj = sheet.["A1:B3"] |> gets |> head
  //let v'': int = sheet.["A1:B3"] |> gets<int> |> head
  //excel |> workbook(1) |> worksheet(1) |> select
  //sheet.["A2,B4"] |> activate

  excel'op excel {
    set calc'auto
    set visibility'hidden
  }
  |> ignore

  cell'op {
    //copy sheet.["A1:A3"]
    //paste sheet.["B1"] paste'mode
    ////delete sheet.["A1:A3"] delete'mode

    //paste sheet.["B1"] { paste'mode with Paste = paste'values; SkipBlanks = true }
    //insert sheet.["B1"] insert'mode
    //insert sheet.["B1"] { insert'mode with Shift = shift'down; Origin= origin'below }
    ////delete sheet.["A1"] delete'mode
    ////delete sheet.["A1"] { delete'mode with Shift= shift'up }

    //set sheet.["A1"] sheet.["B1"]
    //set sheet.["A1"] 100

    //fx sheet.["A1"] sheet.["B1"]
    //fx sheet.["A1"] "SUM(A1:B3)"
    width sheet.["A1"] 100
    height sheet.["A1"] 100

    //fit'height sheet.["A1"]
    //fit'width sheet.["A1"]
    //set sheet.["A1"] theme'accent1
    //set sheet.["A1"] "a"
    //set sheet.["A1"] v'center
    //set sheet.["A1"] h'center

    //merge sheet.["C1:D3"] true
    //merge sheet.["E1:F3"] false
    //unmerge sheet.["C1:D3"]
    //unmerge sheet.["E1:F3"]

    set sheet.["A1"] DateTime.Now
    format sheet.["A1"] "(日付)yyyy-MM-dd"
  }

  ////sheet |> saveAs "sample.xlsx"
  //let book = excel |> workbook(1)
  //book.Close()


  //ruledline sheet.["B2:C5"] {
  //  top { border with Color= Color.Red }
  //  left { border with Color= Color.Orange; Weight= weight'thick }
  //  right { border with LineStyle= linestyle'dashdot }
  //  bottom { border with Weight= weight'medium }
  //  horizontal { border with Color= Color.Blue; Weight= weight'medium }
  //  vertical { border with Color= rgb (0, 128, 255); Weight= weight'thin }

  //  // growing と falling は値がExcel内部で共有されているため、設定値は後勝ちする。
  //  growing { border with Weight= weight'hairline }
  //  falling { border with Weight= weight'thick }
  //}
  //|> ignore
  
  ////sheet.["A1"] |> set "サンプルテキスト"
  ////font sheet.["A1"] {
  ////  name "あんずもじ"
  ////  size 24.0
  ////  color Color.Blue
  ////  bold truec
  ////  strikethrough true
  ////}
  ////|> ignore

  //font sheet.["A1:A3"] {
  //  set Color.Orange
  //  set ( rgb(0, 128, 255) )
  //  set { r= 0; g= 128; b= 255; }
  //  set underline'double
  //  set style'normal
  //  set (style'normal ||| style'strikethrough ||| style'shadow)

  //  // フォントの指定
  //  set "メイリオ"
  //  // フォントサイズの設定
  //  set 16.0
  //  // 下線の設定
  //  underline underline'double

  //  // フォント色の設定
  //  color Color.Orange
  //  // or
  //  color ( rgb(0, 128, 255) )
  //  // or
  //  color { r= 0; g= 128; b= 255; }


  //  // フォントスタイルの設定
  //  style style'normal
  //  // スタイルを複数選択する場合は以下のように指定する.
  //  style (style'normal ||| style'strikethrough ||| style'shadow)
  //  // style を利用しなくとも各種スタイルをひとつずつ ON/OFF 可能
  //  bold true
  //  italic true
  //  shadow true
  //  outline true
  //  strikethrough true
  //  subscript true
  //  superscript true

  //}
  //|> ignore

  //sheet.["A1:B3"] |> address |> printfn "%s"
  //10 |> colname |> printfn "%s"
  //128 |> colname |> printfn "%s"


  //sheet.["A1"].Rows |> delete dd'up

  //// rows関数を利用して, 1行ずつ取得する
  //for (index, row) in sheet.["A1:B3"] |> rowsi do
  //  //if index % 2 = 0 then
  //  //  column |> bgpattern Pattern.Checker
  //  //else
  //  //  column |> bgpattern Pattern.CrissCross

  //  // 各cell毎に何か処理をする
  //  for cell in row do
  //    printf $"{cell |> get} "
  //  printfn ""

  //// columns関数を利用して, 1行ずつ取得する
  //for (index, column) in sheet.["A1:B3"] |> columnsi do
  //  //if index % 2 = 0 then
  //  //  column |> bgpattern Pattern.Checker
  //  //else
  //  //  column |> bgpattern Pattern.CrissCross

  //  // 各cell毎に何か処理をする
  //  for cell in column do
  //    printf $"{cell |> get} "
  //  printfn ""

with
  _ -> ()

//[<EntryPoint>]
//let main argv =
//  use excel = create()
//  excel.Visibility <- AppVisibility.Visible

//  let sheet = excel |> workbook(1) |> worksheet(1)
//  sheet.["A1:A3"] |> set 100 
//  sheet.["B1:B3"] |> set 200
//  sheet.["C1"] |> fx "SUM(A1:B1)"
  
//  sheet.["A1:B3"] |> fx "COUNT(A1:B3)"
//  sheet.["C2"] |> set 200
//  let a = sheet.["C2"] |> getfx
//  let ax = sheet.["A1:B3"] |> gets
//  let h = ax |> head
//  let l = ax |> last
//  let orig = sheet.["A1:B3"] 
//  let v: obj = sheet.["A1:B3"] |> get
//  let v': obj = sheet.["A1:B3"] |> gets |> head
//  let v'': int = sheet.["A1:B3"] |> gets<int> |> head

//  // columns関数を利用して, 1行ずつ取得する
//  for (index, column) in sheet.["A1:B3"] |> columnsi do
//    if index % 2 = 0 then
//      column |> bgcolor Color.Blue
//    else
//      column |> bgcolor Color.Red

//    // 各cell毎に何か処理をする
//    for cell in column do
//      printf $"{cell |> get} "
//    printfn ""

//  //sheet |> saveAs @"D:\OneDrive\デスクトップ\foo.xlsx"


//  //use excel = open' @"D:\OneDrive\デスクトップ\foo.xlsx"
//  //excel.Visibility <- AppVisibility.Visible

//  //let sheet = excel |> workbook(1) |> worksheet(1)
//  //sheet.["B1:B3"] |> set 200
//  //sheet |> save

//  0




//let read () = System.Console.ReadLine()
//let toInt (s: string) = System.Convert.ToInt32(s)

////printf "アタッチするExcelを指定してください。---> "
////let ps = show()
////let index = read() |> toInt

////let ps = enumerate()
////let app = attach ps.[0]
////let sheet = app |> workbook(1) |> worksheet(1)

////sheet.["A1:A3"] |> set 100
////sheet.["B1"] |> fx "SUM(A1:A3)"
////printfn "%A" (get sheet.["B1"])

////sheet.["A1:A3"]
////|> address
////|> printfn "%s"

//let ps = enumerate()
//let app = attach ps.[0]
//let sheet = app |> workbook(1) |> worksheet(1)

////sheet.["A1:A3"] |> set 100
////sheet.["B1"] |> fx "SUM(A1:A3)"

////sheet.["A1:B3"]
////|> gets
////|> iteri (fun i j row -> printfn $"[%d{i}, %d{j}] {row}")

////sheet.["A1:B3"]
////|> rows
////|> iter (fun row -> printfn $"%A{row}")

////sheet.["A1:B3"]
////|> rowsi
////|> iter (fun (i, row) ->  printfn $"[%d{i}] %A{row}")

////sheet.["A1:B3"]
////|> columns
////|> iter (fun col -> printfn $"%A{col}")

////sheet.["A1:B3"]
////|> columnsi
////|> iter (fun (i, col) -> printfn $"[%d{i}] %A{col}")


//////let enumerator = sheet.["A1:B3"].GetEnumerator()
//////printfn $"%A{enumerator.Current}"
//////enumerator.MoveNext() |> ignore
//////let v = enumerator.Current
//////let t = v.GetType()
//////printfn $"%A{t}"

////let range = sheet.["A1:B3"]
////printfn $"Height= {range.Height}, Width= {range.Width}"
////printfn $"RowHeight= {range.RowHeight}, ColumnWidth= {range.ColumnWidth}"

////let rs = range.Rows |> Seq.toArray
////let rs2 = range |> get :?> obj[,]
////printfn $"%A{rs2.[1, 1]}"




////let range = sheet.["A1:B3"]
////range |> rows |> iter (fun row -> 
////  row |> cells |> iter (fun cell ->
////    printf $"%A{cell |> integer} ")
////  printfn "")

////printfn ""

////range |> columns |> iter (fun col -> 
////  col |> cells |> iter (fun cell ->
////    printf $"%A{cell |> integer} ")
////  printfn "")
  


//let range = sheet.["A1:B3"]
////range |> rows |> iteri (fun i row ->
////  if i % 2 = 0 then row |> bgcolor Color.Azure
////  else row |> bgpattern Pattern.Gray16)


//// 罫線の設定
//ruledline sheet.["B2:C5"] {
//  top (border { color Color.Red })
//  left (border { color Color.Red; weight thick })
//  right (border { style dashdot })
//  bottom (border { weight medium })
//  horizontal (border { color Color.Blue; weight thick })
//  vertical (border { color Color.Green; weight thick })
//  // growing と falling は値がExcel内部で共有されているため、設定値は後勝ちする。
//  growing (border { color Color.Red })
//  top (border { style lineNone })
//}
//|> ignore

//sheet.["A1"] |> set "サンプルテキスト"
//font sheet.["A1"] {
//  name "あんずもじ"
//  size 24.0
//  color Color.Blue
//  bold true
//  strikethrough true
//}
//|> ignore


////showFonts()


////range.Rows
////|> Seq.iter (fun row -> 
////  row |> Seq.iter (fun cell -> printf $"%A{cell.Address} ")
////  printfn "")

//detach app

////let range = sheet.["A1:B3"] |> gets
////let (x, y) = range |> len
////for i = 0 to x - 1 do
////  for j = 0 to y - 1 do
////    printf $"{range.[i, j]} "
////  printfn ""

////detach app


