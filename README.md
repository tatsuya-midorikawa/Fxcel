# Fxcel - Excel operations library  

![Fxcel](https://raw.githubusercontent.com/tatsuya-midorikawa/Fxcel/main/assets/fxcel.png)  


## 🔷 What's this?  

- Fxcel は F# で簡単に Excel の COM 操作をするためのライブラリです。  
  - C# 向けの Excel COM 操作ライブラリである ***[Midoliy.Office.Interop.Excel](https://github.com/Midoliy/Midoliy.Office.Interop.Excel)*** のラッパーライブラリとなります。
- .NET 5.0 以上の環境をサポートしています。  
- 主に F# Script や F# Interactive での利用を想定して設計をしていますが、Console アプリや Desktop アプリでも問題なく利用可能です。  
- COM を利用するため Excel のインストールが必要です。  

---

## 🔷 Get started  

### ◼◻ F# Interactiveで利用する

#### 1. FSIを起動する  

```powershell
dotnet fsi
```

#### 2. Fxcel を読み込む

Fxcel を nuget から読み込みます。

```fsharp
#r "nuget: Fxcel, 0.0.13";;
open Fxcel;;
```  

### ◼◻ F# Scriptで利用する

#### 1. **.fsx** ファイルを作成する  

まずはコーディングを始めるために **main.fsx** を作成して、VSCode で開きましょう。  

```powershell
mkdir D:/work
cd D:/work
new-item main.fsx
code D:/work
```

#### 2. Fxcel を読み込む

**main.fsx** に Fxcel を利用するためのコードを追加します。

```fsharp
#r "nuget: Fxcel"
open Fxcel
```  

### ◼◻ F# プロジェクトで利用する

#### 1. 新規プロジェクトを作成する  

```powershell
mkdir D:/work
cd D:/work
dotnet new console -lang=F# -o=FxcelSample
``` 
#### 2. Fxcel を読み込む 

```powershell
cd D:/work/FxcelSample
dotnet add package Fxcel
``` 

---

## 🔷 Reference for F# Interactive

### ◼◻ 起動中のExcelプロセス一覧をターミナルに表示しつつ取得する / `show ()`

```fsharp
let processList = show ();;
```

### ◼◻ 起動中のExcelプロセス一覧を取得する / `enumerate ()`

```fsharp
let processList = enumerate ();;
```

### ◼◻ 起動中のExcelプロセスにアタッチする / `attach (excel: Handle)`

```fsharp
let processList = enumerate ();;
let excel = processList.[0] |> attach;;
```

### ◼◻ アタッチ済みのExcelプロセスをデタッチする / `detach (excel: IExcelApplication)`

```fsharp
let processList = enumerate ();;
let excel = processList.[0] |> attach;;

// do somethings

excel |> detach;;
```

---

## 🔷 Reference  

### ◼◻ Excelワークブックを新規作成する / `create ()`

```fsharp
[<EntryPoint>]
let main argv =
  use excel = create ()
```

### ◼◻ 既存のExcelワークブックをテンプレートとして新規Excelワークブックを作成する / `create (template: string)`

```fsharp
[<EntryPoint>]
let main argv =
  use excel = create "C:/work/sample.xlsx"
```

### ◼◻ 既存のExcelワークブックを開く / `open' (filepath: string)`

```fsharp
[<EntryPoint>]
let main argv =
  use excel = open' "C:/work/sample.xlsx"
```

### ◼◻ Excelワークブックを名前を付けて保存する / `saveAs (filepath: string) excelObject`

```fsharp
[<EntryPoint>]
let main argv =
  // Excelワークブックを新規作成
  use excel = create()
  // 先頭のワークブックを取得する
  let book = excel |> workbook(1)
  
  // do somethings

  // 名前を付けて保存
  book |> saveAs "C:/work/sample.xlsx"
```

### ◼◻ Excelワークブックを上書き保存する / `save excelObject`

```fsharp
[<EntryPoint>]
let main argv =
  // 既存のExcelワークブックを開く
  use excel = open' "C:/work/sample.xlsx"
  // 先頭のワークブックを取得する
  let book = excel |> workbook(1)

  // do somethings

  // 上書き保存する
  book |> save
```

### ◼◻ Excelワークブックオブジェクトを取得する / `workbook (index: int) (excel: IExcelApplication)`

```fsharp
[<EntryPoint>]
let main argv =
  use excel = open' "C:/work/sample.xlsx"

  // ワークブックオブジェクトを取得する
  //   -> index は 1 始まりであることに注意する
  let book = excel |> workbook(1)
```

### ◼◻ Excelワークシートオブジェクトを取得する / `worksheet (index: int | string) (workbook: IWrokbook)`

```fsharp
[<EntryPoint>]
let main argv =
  use excel = open' "C:/work/sample.xlsx"

  // ワークシートオブジェクトを取得する
  //   -> index は 1 始まりであることに注意する
  let sheet = excel |> workbook(1) |> worksheet(1)

  // シート名を指定して取得することもできる
  let sheet = excel |> workbook(1) |> worksheet("Sheet1")
```

### ◼◻ Excelワークシートを新規追加する / `newsheet (book: IWorkbook)`

```fsharp
[<EntryPoint>]
let main argv =
  use excel = open' "C:/work/sample.xlsx"
  let sheet = excel |> workbook(1) |> newsheet
```

### ◼◻ Excel Cellオブジェクトを取得 / `sheet.[address]`

```fsharp
[<EntryPoint>]
let main argv =
  use excel = open' "C:/work/sample.xlsx"
  let sheet = excel |> workbook(1) |> worksheet(1)

  // Cellオブジェクトをアドレス形式で取得
  let cell = sheet.["A1"]
  // CellオブジェクトをR1C1形式で取得
  let cell = sheet.[1, 1]
```

### ◼◻ Excel Rangeオブジェクトを取得 / `sheet.[address]`

```fsharp
[<EntryPoint>]
let main argv =
  use excel = open' "C:/work/sample.xlsx"
  let sheet = excel |> workbook(1) |> worksheet(1)

  // Rangeオブジェクトをアドレス形式で取得
  let cell = sheet.["A1:B3"]
  // Cellオブジェクトを2つのアドレスを指定して取得
  let cell = sheet.["A1", "B3"]
```

### ◼◻ Excel Rangeを行ごとに列挙する / `rows (range: IExcelRange) / rowsi (range: IExcelRange)`

```fsharp
[<EntryPoint>]
let main argv =
  use excel = open' "C:/work/sample.xlsx"
  let sheet = excel |> workbook(1) |> worksheet(1)

  // rows関数を利用して, 1行ずつ取得する
  for row in sheet.["A1:B3"] |> rows do
    // 各Cell毎に何か処理をする
    for cell in row do
      // do somethings


  // rowsi関数を利用して, インデックス付きで1行ずつ取得する
  //   -> index は 1 始まりであることに注意する
  for (index, row) in sheet.["A1:B3"] |> rowsi do
    // 各Cell毎に何か処理をする
    for cell in row do
      // do somethings
```

### ◼◻ Excel Rangeを列ごとに列挙する / `columns (range: IExcelRange)`

```fsharp
[<EntryPoint>]
let main argv =
  use excel = open' "C:/work/sample.xlsx"
  let sheet = excel |> workbook(1) |> worksheet(1)

  // columns関数を利用して, 1行ずつ取得する
  for column in sheet.["A1:B3"] |> columns do
    // 各Cell毎に何か処理をする
    for cell in column do
      // do somethings


  // columnsi関数を利用して, インデックス付きで1行ずつ取得する
  //   -> index は 1 始まりであることに注意する
  for column in sheet.["A1:B3"] |> columns do
    // 各Cell毎に何か処理をする
    for cell in column do
      // do somethings
```

### ◼◻ Excel Cellオブジェクトから値を取得する / `get (cell: IExcelRange) / get<'T> (cell: IExcelRange)`

```fsharp
[<EntryPoint>]
let main argv =
  use excel = open' "C:/work/sample.xlsx"
  let sheet = excel |> workbook(1) |> worksheet(1)

  // Cellオブジェクトから値を取得する
  let v: obj = sheet.["A1"] |> get

  // Cellオブジェクトから値を指定した型で取得する
  //   -> 指定した型と互換性がない場合, System.InvalidCastException
  let v: int = sheet.["A1"] |> get<int>

  // 複数要素がある場合は先頭要素の値を取得する.
  //   -> 以下の場合 sheet.["A1"] の値が得られる.
  let v: obj = sheet.["A1:B3"] |> get

  // 複数要素がある場合は先頭要素の値を指定した型で取得する.
  //   -> 以下の場合 sheet.["A1"] の値が得られる.
  //   -> 指定した型と互換性がない場合, System.InvalidCastException
  let v: int = sheet.["A1:B3"] |> get<int>
```

### ◼◻ Excel Rangeオブジェクトから値を取得する / `gets (range: IExcelRange) / gets<'T> (range: IExcelRange)`

```fsharp
[<EntryPoint>]
let main argv =
  use excel = open' "C:/work/sample.xlsx"
  let sheet = excel |> workbook(1) |> worksheet(1)

  // Rangeオブジェクトから値を取得する
  let vs: obj [,]  = sheet.["A1:A3"] |> gets

  // Rangeオブジェクトから値を指定した型で取得する
  //   -> 指定した型と互換性がない場合, System.InvalidCastException
  let vs: int [,]  = sheet.["A1:A3"] |> gets<int>

  // Rangeオブジェクトから先頭要素の値を取得する
  //   -> 以下の場合 sheet.["A1"] の値が得られる.
  let v: obj = sheet.["A1:B3"] |> gets |> head

  // Rangeオブジェクトから先頭要素の値を指定した型で取得する
  //   -> 以下の場合 sheet.["A1"] の値が得られる.
  //   -> 指定した型と互換性がない場合, System.InvalidCastException
  let v: int = sheet.["A1:B3"] |> gets<int> |> head

  // 複数要素がある場合は最終要素の値を取得する.
  //   -> 以下の場合 sheet.["B3"] の値が得られる.
  let v: obj = sheet.["A1:B3"] |> gets |> last

  // 複数要素がある場合は最終要素の値を指定した型で取得する.
  //   -> 以下の場合 sheet.["B3"] の値が得られる.
  //   -> 指定した型と互換性がない場合, System.InvalidCastException
  let v: int = sheet.["A1:B3"] |> gets<int> |> last
```

### ◼◻ Excel Cellオブジェクトから関数を取得する / `getfx (cell: IExcelRange)`

```fsharp
[<EntryPoint>]
let main argv =
  use excel = open' "C:/work/sample.xlsx"
  let sheet = excel |> workbook(1) |> worksheet(1)

  // Cellオブジェクトから関数を取得する
  let fn: string = sheet.["A1"] |> getfx
```

### ◼◻ Excel Rnageオブジェクトから関数を取得する / `getsfx (range: IExcelRange)`

```fsharp
[<EntryPoint>]
let main argv =
  use excel = open' "C:/work/sample.xlsx"
  let sheet = excel |> workbook(1) |> worksheet(1)

  // Rangeオブジェクトから関数を取得する
  let fns: string [,] = sheet.["A1:A3"] |> getsfx
```

### ◼◻ Excel Cell / Rangeオブジェクトに値を設定する / `set (value: obj) (target: IExcelRange)`

```fsharp
[<EntryPoint>]
let main argv =
  use excel = open' "C:/work/sample.xlsx"
  let sheet = excel |> workbook(1) |> worksheet(1)

  // 対象オブジェクトに値を設定する
  sheet.["A1"] |> set 100
  sheet.["A1:B3"] |> set 100
```

### ◼◻ Excel Cell / Rangeオブジェクトに関数を設定する / `fx (func: string) (target: IExcelRange)`

```fsharp
[<EntryPoint>]
let main argv =
  use excel = open' "C:/work/sample.xlsx"
  let sheet = excel |> workbook(1) |> worksheet(1)

  // 対象オブジェクトに値を設定する
  sheet.["A1"] |> fx "SUM(A2:A5)"
  sheet.["A1:B3"] |> fx "COUNT(A1:B3)"
```

### ◼◻ Excel Cell / Rangeオブジェクトなどに背景色を設定する / `bgcolor (color: Color) (target: IExcelRange)`

```fsharp
[<EntryPoint>]
let main argv =
  use excel = open' "C:/work/sample.xlsx"
  let sheet = excel |> workbook(1) |> worksheet(1)

  // 対象オブジェクトの背景色を設定する
  sheet.["A1"] |> bgcolor Color.Red
  sheet.["B1:B3"] |> bgcolor Color.Blue
```

### ◼◻ Excel Cell / Rangeオブジェクトなどに背景パターンを設定する / `bgpattern (pattern: Pattern) (target: IExcelRange)`

```fsharp
[<EntryPoint>]
let main argv =
  use excel = open' "C:/work/sample.xlsx"
  let sheet = excel |> workbook(1) |> worksheet(1)

  // 対象オブジェクトの背景パターンを設定する
  sheet.["A1"] |> bgpattern Pattern.Checker
  sheet.["B1:B3"] |> bgpattern Pattern.CrissCross
```

### ◼◻ 罫線を設定する / `ruledline (target: IExcelRange)` コンピュテーション式

#### 📑 `ruledline` で利用できるカスタムオペレーション

| operation name | description |
| --- | --- |
| `top (border)` | 最上部の横罫線. |
| `bottom (border)` | 最下部の横罫線. |
| `left (border)` | 最左部の縦罫線. |
| `right (border)` | 最右部の縦罫線. |
| `horizontal (border)` | 中間部の横罫線. |
| `vertical (border)` | 中間部の縦罫線. |
| `growing (border)` | 左下から右上に向けての罫線. 色や太さの設定は `falling` と共有. |
| `falling (border)` | 左上から右下に向けての罫線. 色や太さの設定は `growing` と共有. |

#### 📑 `Border` に設定できる値

| operation name | description | values |
| --- | --- | --- |
| `LineStyle` | 罫線のスタイル. | `linestyle'none` / `linestyle'dot` / `linestyle'double` / `linestyle'dash` / `linestyle'continuous` / `linestyle'dashdot` / `linestyle'dashdotdot` / `linestyle'slant`|
| `Weight` | 罫線の太さ. | `weight'medium` / `weight'hairline` / `weight'thin` / `weight'thick` |
| `Color` | 罫線の色. | `Color.Red` / `Color.Orange` / `Color.Blue` / `rgb(r, g, b)` and more... |

```fsharp
[<EntryPoint>]
let main argv =
  use excel = open' "C:/work/sample.xlsx"
  let sheet = excel |> workbook(1) |> worksheet(1)

  // 罫線を設定する
  ruledline sheet.["B2:C5"] {
    // 各 Border の値は with を利用して指定する.
    top { border with Color= Color.Red }
    left { border with Color= Color.Orange; Weight= weight'thick }
    right { border with LineStyle= linestyle'dashdot }
    bottom { border with Weight= weight'medium }
    horizontal { border with Color= Color.Blue; Weight= weight'medium }
    vertical { border with Color= rgb (0, 128, 255); Weight= weight'thin }

    // growing と falling は値がExcel内部で共有されているため、設定値は後勝ちする.
    growing { border with Weight= weight'hairline }
    falling { border with Weight= weight'thick }
  }
  |> ignore
```

### ◼◻ フォントを設定する / `font (target: IExcelRange)` コンピュテーション式

#### 📑 `font` で利用できるカスタムオペレーション

| operation name | description | values |
| --- | --- | --- |
| `name (name: string)` | フォント名. | `游ゴシック` / `メイリオ` / `consolas` and more... |
| `size (size: float)` | フォントサイズ. | `8.0` / `10.5` / `24.0` and more... |
| `style (style: FontStyle)` | フォントスタイル. `Flags` なので複数まとめて指定可能. | `style'normal` / `style'bold` / `style'italic'` / `style'shadow` / `style'strikethrough` / `style'subscript` / `style'superscript` / `style'singleUnderline` / `style'doubleUnderline` |
| `color (value: Color)` | フォント色. | `Color.Red` / `Color.Orange` / `Color.Blue` and more... |
| `color (value: RGB)` | フォント色. | `{ r= 0; g= 128; b= 255; }` |
| `underline (style: Underline)` | 下線. | `underline'none` / `underline'double` / `underline'doubleAccounting` / `underline'single` / `underline'singleAccounting` |
| `bold (on: bool)` | 太字. | `true` or `false` |
| `italic (on: bool)` | イタリック体. | `true` or `false` |
| `shadow (on: bool)` | フォント影. | `true` or `false` |
| `outline (on: bool)` | アウトラインフォント. | `true` or `false` |
| `strikethrough (on: bool)` | 打ち消し線. | `true` or `false` |
| `subscript (on: bool)` | 下付き文字にする. | `true` or `false` |
| `superscript (on: bool)` | 上付き文字にする. | `true` or `false` |

```fsharp
[<EntryPoint>]
let main argv =
  use excel = open' "C:/work/sample.xlsx"
  let sheet = excel |> workbook(1) |> worksheet(1)

  // フォントを設定する
  font sheet.["A1:A3"] {
    // フォントの指定
    name "メイリオ"
    // フォントサイズの設定
    size 16.0
    // 下線の設定
    underline underline'double

    // フォント色の設定
    color Color.Orange
    // or
    color ( rgb(0, 128, 255) )
    // or
    color { r= 0; g= 128; b= 255; }

    // フォントスタイルの設定
    style style'normal
    // スタイルを複数選択する場合は以下のように指定する.
    style (style'normal ||| style'strikethrough ||| style'shadow)
    // style を利用しなくとも各種スタイルをひとつずつ ON/OFF 可能
    bold true
    italic true
    shadow true
    outline true
    strikethrough true
    subscript true
    superscript true
  }
  |> ignore
```

### ◼◻ Excel Cell / Range オブジェクトなどを操作する（コピー・ペースト・挿入・削除） / `op ()` コンピュテーション式


#### 📑 `op` で利用できるカスタムオペレーション

| operation name | description |
| --- | --- |
| `copy (target: IExcelRange)` | 対象をクリップボードにコピーする. |
| `paste (target: IExcelRange, pasteMode: PasteMode)` | 対象にクリップボードの値を貼り付ける. |
| `insert (target: IExcelRange, insertMode: InsertMode)` | 対象にクリップボードの値を挿入する. |
| `delete (target: IExcelRange, deleteMode: DeleteMode)` | 対象を削除する. |

#### 📑 `PasteMode` の要素

| name | description | values |
| --- | --- | --- |
| `Paste` | 貼り付け方式. / `default: paste'all` | `paste'values` / `paste'comments` / `paste'formulas` / `paste'formats` / `paste'all` / `paste'validation` / `paste'exceptBorders` / `paste'colmnWidths` / `paste'formulasAndNumberFormats` / `paste'valuesAndNumberFormats` / `paste'allUsingSourceTheme` / `paste'allMergingConditionalFormats` |
| `Op` | 貼り付け時の演算方法. / `default: op'none`| `op'none` / `op'add` / `op'sub` / `op'mul` / `op'div` |
| `SkipBlanks` | 空白セルを無視するか. / `default: false` | `true` or `false` |
| `SkipBlanks` | 行列を入れ替えるか. / `default: false` | `true` or `false` |

#### 📑 `InsertMode` の要素

| name | description | values |
| --- | --- | --- |
| `Shift` | 挿入後に他のセルをどうシフト移動するか. / `default: shift'down` | `shift'right` / `shift'down` |
| `Origin` | 書式をコピーしてくる方向. / `default: origin'right / origin'below` | `origin'left` / `origin'above` / `origin'right` / `origin'below` |

#### 📑 `DeleteMode` の要素

| name | description | values |
| --- | --- | --- |
| `Shift` | 削除後に他のセルをどうシフト移動するか. | `shift'left` / `shift'up` |


```fsharp
[<EntryPoint>]
let main argv =
  use excel = open' "C:/work/sample.xlsx"
  let sheet = excel |> workbook(1) |> worksheet(1)
  
  op {
    // A1 をクリップボードにコピー
    copy sheet.["A1"]
    // 範囲コピーも可能
    copy sheet.["A1:A3"]

    // クリップボードのデータを B1 に貼り付け
    paste sheet.["B1"] paste'mode
    paste sheet.["B1"] { paste'mode with Paste= paste'values }
    paste sheet.["B1"] { paste'mode with SkipBlanks= true }
    paste sheet.["B1"] { paste'mode with Paste= paste'values; SkipBlanks= true }
    // 範囲貼り付けも可能
    paste sheet.["B1:B3"] paste'mode

    // クリップボードのデータを C1 に挿入
    insert sheet.["C1"] insert'mode
    insert sheet.["C1"] { insert'mode with Shift= shift'down }
    insert sheet.["C1"] { insert'mode with Origin= origin'below }
    insert sheet.["C1"] { insert'mode with Shift= shift'right; Origin= origin'below }
    // 範囲挿入も可能
    insert sheet.["C1:C3"] insert'mode

    // A1 のデータを削除する
    delete sheet.["A1"] delete'mode
    delete sheet.["A1"] { delete'mode with Shift= shift'up }
    // 範囲削除も可能
    delete sheet.["A1:A3"] delete'mode
  }
```

### ◼◻ Excel Cell / Range オブジェクトなどを削除する / `delete (direction: DeleteShiftDirection) (target: ^Range)`

#### 📑 `DeleteShiftDirection`

| value | description |
| --- | --- |
| `shift'left` | 削除後, 左方向へシフト. |
| `shift'up` | 削除後, 上方向へシフト. |

```fsharp
[<EntryPoint>]
let main argv =
  use excel = open' "C:/work/sample.xlsx"
  let sheet = excel |> workbook(1) |> worksheet(1)
  
  // 対象を削除する
  sheet.["A1"] |> delete shift'up
  sheet.["A1:A3"] |> delete shift'left
```

---

## 🔷 Utility  

### ◼◻ 数値をカラム名に変換する / `colname (index: int)`

```fsharp
let name = 1 |> colname     // A
let name = 10 |> colname    // J
let name = 128 |> colname   // DX
```

### ◼◻ 対象の Range オブジェクトからアドレスを取得する / `address (target: IExcelRange)`

```fsharp
let adds = sheet.["A1"] |> address      // $A$1
let adds = sheet.["A1:B3"] |> address   // $A$1:$B$3
```

### ◼◻ 対象の Excel オブジェクトを選択する / `activate (target: ^T)`

```fsharp
// Workbookを選択状態にする
excel |> workbook(1) |> activate

// Worksheetを選択状態にする
excel |> workbook(1) |> worksheet(1) |> activate

// Cellを選択状態にする
sheet.["B1"] |> activate
sheet.["A1:B3"] |> activate
```

### ◼◻ 対象の Excel オブジェクトを選択する / `select (target: ^T)`

```fsharp
// Worksheet(1)を選択状態にする
excel |> workbook(1) |> worksheet(1) |> select

// Cellを選択状態にする
sheet.["B1"] |> select
sheet.["D1:E3"] |> select
```

---

## 🔷 TIPS  

### ◼◻ `try-finally` の利用  

例外処理を施していない場合 Excel COM オブジェクトが適切に解放されず, プロセス上に残ってしまう恐れがあります.  
`try-finally` (または `try-with`) と `use` を併用することで Excel COM オブジェクトの解放漏れを防げます.  

```fsharp
try
  // use を利用する.
  use excel = create ()

  // do somethings

finally
  ()
```  

また, F# Interactive で利用する場合, `attach` したあとは必ず `detach` する必要があります.  

```powershell
let ps = enumerate ();;
let excel = ps.[0] |> attach;;

# do somethings

excel |> detach;;
```
