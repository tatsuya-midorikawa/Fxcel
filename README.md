# Fxcel - Excel operations library  

![Fxcel](https://raw.githubusercontent.com/tatsuya-midorikawa/Fxcel/main/assets/fxcel.png)  


## What's this?  

- Fxcel は F# で簡単に Excel の COM 操作をするためのライブラリです。  
  - C# 向けの Excel COM 操作ライブラリである ***[Midoliy.Office.Interop.Excel](https://github.com/Midoliy/Midoliy.Office.Interop.Excel)*** のラッパーライブラリとなります。
- .NET 5.0 以上の環境をサポートしています。  
- 主に F# Script や F# Interactive での利用を想定して設計をしていますが、Console アプリや Desktop アプリでも問題なく利用可能です。  
- COM を利用するため Excel のインストールが必要です。  


## Get started  

### 1. F# Scriptで利用する

#### 1-1. **.fsx** ファイルを作成する  

まずはコーディングを始めるために **main.fsx** を作成して、VSCode で開きましょう。  

```powershell
mkdir D:/work
cd D:/work
new-item main.fsx
code D:/work
```

#### 1-2. Fxcel を読み込む

**main.fsx** に Fxcel を利用するためのコードを追加します。

```fsharp
#r "nuget: Fxcel"
open Fxcel
```  

### 2. F# プロジェクトで利用する

#### 2-1. 新規プロジェクトを作成する  

```powershell
mkdir D:/work
cd D:/work
dotnet new console -lang=F# -o=FxcelSample
``` 
#### 2-2. Fxcel を読み込む 

```powershell
cd D:/work/FxcelSample
dotnet add package Fxcel
``` 


## Reference  

### Excelワークブックを新規作成する

```fsharp
[<EntryPoint>]
let main argv =
  use excel = create()
```

### 既存のExcelワークブックを開く

```fsharp
[<EntryPoint>]
let main argv =
  use excel = open' "C:/work/sample.xlsx"
  excel |> workbook(1) |> saveAs "C:/work/sample.xlsx"
```

### Excelワークブックを名前を付けて保存する

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

### Excelワークブックを上書き保存する

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

### Excelワークブックオブジェクトを取得する

```fsharp
[<EntryPoint>]
let main argv =
  use excel = open' "C:/work/sample.xlsx"

  // ワークブックオブジェクトを取得する
  //   -> index は 1 始まりであることに注意する
  let book = excel |> workbook(1)
```

### Excelワークシートオブジェクトを取得する

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
