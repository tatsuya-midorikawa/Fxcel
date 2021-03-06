# Fxcel - Excel operations library  

![Fxcel](https://raw.githubusercontent.com/tatsuya-midorikawa/Fxcel/main/assets/fxcel.png)  


## ð· What's this?  

- Fxcel ã¯ F# ã§ç°¡åã« Excel ã® COM æä½ãããããã®ã©ã¤ãã©ãªã§ãã  
  - C# åãã® Excel COM æä½ã©ã¤ãã©ãªã§ãã ***[Midoliy.Office.Interop.Excel](https://github.com/Midoliy/Midoliy.Office.Interop.Excel)*** ã®ã©ããã¼ã©ã¤ãã©ãªã¨ãªãã¾ãã
- .NET 5.0 ä»¥ä¸ã®ç°å¢ããµãã¼ããã¦ãã¾ãã  
- ä¸»ã« F# Script ã F# Interactive ã§ã®å©ç¨ãæ³å®ãã¦è¨­è¨ããã¦ãã¾ãããConsole ã¢ããªã Desktop ã¢ããªã§ãåé¡ãªãå©ç¨å¯è½ã§ãã  
- COM ãå©ç¨ãããã Excel ã®ã¤ã³ã¹ãã¼ã«ãå¿è¦ã§ãã  

---

## ð· Get started  

### â¼â» F# Interactiveã§å©ç¨ãã

#### 1. FSIãèµ·åãã  

```powershell
dotnet fsi
```

#### 2. Fxcel ãèª­ã¿è¾¼ã

Fxcel ã nuget ããèª­ã¿è¾¼ã¿ã¾ãã

```fsharp
#r "nuget: Fxcel, 0.0.21";;
open Fxcel;;
```  

### â¼â» F# Scriptã§å©ç¨ãã

#### 1. **.fsx** ãã¡ã¤ã«ãä½æãã  

ã¾ãã¯ã³ã¼ãã£ã³ã°ãå§ããããã« **main.fsx** ãä½æãã¦ãVSCode ã§éãã¾ãããã  

```powershell
mkdir D:/work
cd D:/work
new-item main.fsx
code D:/work
```

#### 2. Fxcel ãèª­ã¿è¾¼ã

**main.fsx** ã« Fxcel ãå©ç¨ããããã®ã³ã¼ããè¿½å ãã¾ãã

```fsharp
#r "nuget: Fxcel"
open Fxcel
```  

### â¼â» F# ãã­ã¸ã§ã¯ãã§å©ç¨ãã

#### 1. æ°è¦ãã­ã¸ã§ã¯ããä½æãã  

```powershell
mkdir D:/work
cd D:/work
dotnet new console -lang=F# -o=FxcelSample
``` 
#### 2. Fxcel ãèª­ã¿è¾¼ã 

```powershell
cd D:/work/FxcelSample
dotnet add package Fxcel
``` 

---

## ð· Reference for F# Interactive

### â¼â» èµ·åä¸­ã®Excelãã­ã»ã¹ä¸è¦§ãã¿ã¼ããã«ã«è¡¨ç¤ºãã¤ã¤åå¾ãã<br>`show ()`

```fsharp
let processList = show ();;
```

### â¼â» èµ·åä¸­ã®Excelãã­ã»ã¹ä¸è¦§ãåå¾ãã<br>`enumerate ()`

```fsharp
let processList = enumerate ();;
```

### â¼â» èµ·åä¸­ã®Excelãã­ã»ã¹ã«ã¢ã¿ãããã<br>`attach (excel: Handle)`

```fsharp
let processList = enumerate ();;
let excel = processList.[0] |> attach;;
```

### â¼â» ã¢ã¿ããæ¸ã¿ã®Excelãã­ã»ã¹ããã¿ãããã<br>`detach (excel: IExcelApplication)`

```fsharp
let processList = enumerate ();;
let excel = processList.[0] |> attach;;

// do somethings

excel |> detach;;
```

---

## ð· Reference  

### â¼â» Workbookãæ°è¦ä½æãã<br>`create (): IExcelApplication`

```fsharp
[<EntryPoint>]
let main argv =
  use excel = create ()
```

### â¼â» æ¢å­Workbookããã³ãã¬ã¼ãã¨ãã¦æ°è¦Workbookãä½æãã<br>`create (template: string): IExcelApplication`

```fsharp
[<EntryPoint>]
let main argv =
  use excel = create "C:/work/sample.xlsx"
```

### â¼â» æ¢å­Workbookãéã<br>`open' (filepath: string): IExcelApplication`

```fsharp
[<EntryPoint>]
let main argv =
  use excel = open' "C:/work/sample.xlsx"
```

### â¼â» Workbookãååãä»ãã¦ä¿å­ãã<br>`saveAs (filepath: string) (excelObject: ^ExcelObject): unit`

```fsharp
[<EntryPoint>]
let main argv =
  // Workbookãæ°è¦ä½æã, ExcelApplicationãåå¾.
  use excel = create()

  // (1) Workbookãå©ç¨ãã¦, ååãä»ãã¦ä¿å­.
  let book = excel |> workbook(1)
  // do somethings
  book |> saveAs "C:/work/sample.xlsx"

  // (2) Worksheetãå©ç¨ãã¦, ååãä»ãã¦ä¿å­.
  let sheet = excel |> workbook(1) |> worksheet(1)
  // do somethings
  sheet |> saveAs "C:/work/sample.xlsx"
```

### â¼â» Workbookãä¸æ¸ãä¿å­ãã<br>`save (excelObject: ^ExcelObject): unit`

```fsharp
[<EntryPoint>]
let main argv =
  // æ¢å­ã®Excelã¯ã¼ã¯ããã¯ãéã.
  use excel = open' "C:/work/sample.xlsx"

  // (1) Workbookãå©ç¨ãã¦, ä¸æ¸ãä¿å­.
  let book = excel |> workbook(1)
  // do somethings
  book |> save

  // (2) Worksheetãå©ç¨ãã¦, ä¸æ¸ãä¿å­.
  let sheet = excel |> workbook(1) |> worksheet(1)
  // do somethings
  sheet |> save
```

### â¼â» Workbookãåå¾ãã<br>`workbook (index: int) (excel: IExcelApplication): IWorkbook`

```fsharp
[<EntryPoint>]
let main argv =
  use excel = open' "C:/work/sample.xlsx"

  // indexãæå®ãã¦Workbookãåå¾.
  //   -> index ã¯ 1 å§ã¾ãã§ãããã¨ã«æ³¨æãã.
  let book = excel |> workbook(1)
```

### â¼â» Worksheetãåå¾ãã<br>`worksheet (index: int | string) (workbook: IWrokbook): IWorksheet`

```fsharp
[<EntryPoint>]
let main argv =
  use excel = open' "C:/work/sample.xlsx"

  // (1) indexãæå®ãã¦Worksheetãåå¾.
  //   -> index ã¯ 1 å§ã¾ãã§ãããã¨ã«æ³¨æãã.
  let sheet = excel |> workbook(1) |> worksheet(1)

  // (2) sheet nameãæå®ãã¦åå¾.
  let sheet = excel |> workbook(1) |> worksheet("Sheet1")
```

### â¼â» Worksheetãæ°è¦è¿½å ãã<br>`newsheet (book: IWorkbook): IWorksheet`

```fsharp
[<EntryPoint>]
let main argv =
  use excel = open' "C:/work/sample.xlsx"
  let sheet = excel |> workbook(1) |> newsheet
```

### â¼â» IExcelRangeãªãã¸ã§ã¯ããåå¾ãã<br>`sheet.[address]: IExcelRange`

```fsharp
[<EntryPoint>]
let main argv =
  use excel = open' "C:/work/sample.xlsx"
  let sheet = excel |> workbook(1) |> worksheet(1)

  // (1) ã¢ãã¬ã¹å½¢å¼ã§åå¾.
  let cell = sheet.["A1"]
  // (2) R1C1å½¢å¼ã§åå¾.
  let cell = sheet.[1, 1]
  // (3) ç¯å²ãã¢ãã¬ã¹å½¢å¼ã§åå¾.
  let range = sheet.["A1:B3"]
  // (4) ç¯å²ãå§ç¹ã»ã«ã¢ãã¬ã¹ã¨çµç¹ã»ã«ã¢ãã¬ã¹ãæå®ãã¦åå¾.
  let range = sheet.["A1", "B3"]
```

### â¼â» IWorksheetããè¡ãåå¾ãã<br>`get'row (index: int) (sheet: IWorksheet): IExcelRow`<br>`get'rows (begin': int, end': int) (sheet: IWorksheet): IExcelRows`

```fsharp
[<EntryPoint>]
let main argv =
  use excel = open' "C:/work/sample.xlsx"
  let sheet = excel |> workbook(1) |> worksheet(1)

  // 1è¡åå¾.
  let r = sheet |> get'row(1)       // $1:$1
  // è¤æ°è¡åå¾.
  let r = sheet |> get'rows(1, 3)   // $1:$3
```

### â¼â» IWorksheetããåãåå¾ãã<br>`get'column (index: int) (sheet: IWorksheet): IExcelRow`<br>`get'columns (begin': int, end': int) (sheet: IWorksheet): IExcelRows`

```fsharp
[<EntryPoint>]
let main argv =
  use excel = open' "C:/work/sample.xlsx"
  let sheet = excel |> workbook(1) |> worksheet(1)

  // 1ååå¾.
  let c = sheet |> get'column(1)       // $A:$A
  // è¤æ°ååå¾.
  let c = sheet |> get'columns(1, 3)   // $A:$C
```

### â¼â» IExcelRangeã®è¡å¨ä½ãåå¾ãã<br>`current'rows (range: IExcelRange): IExcelRows`

```fsharp
[<EntryPoint>]
let main argv =
  use excel = open' "C:/work/sample.xlsx"
  let sheet = excel |> workbook(1) |> worksheet(1)

  // è¡å¨ä½ãåå¾.
  let r = sheet.["A1"] |> current'rows      // $1:$1
  let r = sheet.["A1:B3"] |> current'rows   // $1:$3
```

### â¼â» IExcelRangeã®åå¨ä½ãåå¾ãã<br>`current'columns (range: IExcelRange): IExcelColumns`

```fsharp
[<EntryPoint>]
let main argv =
  use excel = open' "C:/work/sample.xlsx"
  let sheet = excel |> workbook(1) |> worksheet(1)

  // åå¨ä½ãåå¾.
  let r = sheet.["A1"] |> current'columns      // $A:$A
  let r = sheet.["A1:B3"] |> current'columns   // $A:$B
```

### â¼â» ç¯å²ãã¼ã¿ãè¡ãã¨ã«åæãã<br>`rows (range: IExcelRange): seq<IExcelRow>`<br>`rowsi (range: IExcelRange): seq<int * IExcelRow>`

```fsharp
[<EntryPoint>]
let main argv =
  use excel = open' "C:/work/sample.xlsx"
  let sheet = excel |> workbook(1) |> worksheet(1)

  // rowsé¢æ°ãå©ç¨ãã¦, 1è¡ãã¤åå¾.
  for row in sheet.["A1:B3"] |> rows do
    // åCellæ¯ã«ä½ãå¦ç.
    for cell in row do
      // do somethings


  // rowsié¢æ°ãå©ç¨ãã¦, ã¤ã³ããã¯ã¹ä»ãã§1è¡ãã¤åå¾.
  //   -> index ã¯ 1 å§ã¾ãã§ãããã¨ã«æ³¨æ.
  for (index, row) in sheet.["A1:B3"] |> rowsi do
    // åCellæ¯ã«ä½ãå¦ç.
    for cell in row do
      // do somethings
```

### â¼â» ç¯å²ãã¼ã¿ãåãã¨ã«åæãã<br>`columns (range: IExcelRange): seq<IExcelColumn>`<br>`columnsi (range: IExcelRange): seq<int * IExcelColumn>`

```fsharp
[<EntryPoint>]
let main argv =
  use excel = open' "C:/work/sample.xlsx"
  let sheet = excel |> workbook(1) |> worksheet(1)

  // columnsé¢æ°ãå©ç¨ãã¦, 1è¡ãã¤åå¾.
  for column in sheet.["A1:B3"] |> columns do
    // åCellæ¯ã«ä½ãå¦ç.
    for cell in column do
      // do somethings


  // columnsié¢æ°ãå©ç¨ãã¦, ã¤ã³ããã¯ã¹ä»ãã§1è¡ãã¤åå¾.
  //   -> index ã¯ 1 å§ã¾ãã§ãããã¨ã«æ³¨æ.
  for (index, column) in sheet.["A1:B3"] |> columnsi do
    // åCellæ¯ã«ä½ãå¦ç.
    for cell in column do
      // do somethings
```

### â¼â» å¤ãåå¾ãã<br>`get (cell: IExcelRange): obj`<br>`get<'T> (cell: IExcelRange): 'T`

```fsharp
[<EntryPoint>]
let main argv =
  use excel = open' "C:/work/sample.xlsx"
  let sheet = excel |> workbook(1) |> worksheet(1)

  // å¤ãåå¾.
  let v: obj = sheet.["A1"] |> get

  // å¤ãåä»ãã§åå¾.
  //   -> æå®ããåã¨äºææ§ããªãå ´å, System.InvalidCastException.
  let v: int = sheet.["A1"] |> get<int>

  // è¤æ°è¦ç´ ãããå ´åã¯åé ­è¦ç´ ã®å¤ã®ã¿åå¾.
  //   -> ä»¥ä¸ã®å ´å sheet.["A1"] ã®å¤ãå¾ããã.
  let v: obj = sheet.["A1:B3"] |> get

  // è¤æ°è¦ç´ ãããå ´åã¯åé ­è¦ç´ ã®åä»ãã®å¤ã®ã¿åå¾.
  //   -> ä»¥ä¸ã®å ´å sheet.["A1"] ã®å¤ãå¾ããã.
  //   -> æå®ããåã¨äºææ§ããªãå ´å, System.InvalidCastException.
  let v: int = sheet.["A1:B3"] |> get<int>
```

### â¼â» å¤ãéåãã¼ã¿ã§åå¾ãã<br>`gets (range: IExcelRange): obj [,]`<br>`gets<'T> (range: IExcelRange): 'T [,]`

```fsharp
[<EntryPoint>]
let main argv =
  use excel = open' "C:/work/sample.xlsx"
  let sheet = excel |> workbook(1) |> worksheet(1)

  // å¤ãéåãã¼ã¿ã¨ãã¦åå¾.
  let vs: obj [,]  = sheet.["A1:A3"] |> gets

  // å¤ãåä»ãã®éåãã¼ã¿ã¨ãã¦åå¾.
  //   -> æå®ããåã¨äºææ§ããªãå ´å, System.InvalidCastException.
  let vs: int [,]  = sheet.["A1:A3"] |> gets<int>

  // åå¾ããéåãã¼ã¿ããåé ­è¦ç´ ã®å¤ãåå¾.
  //   -> ä»¥ä¸ã®å ´å sheet.["A1"] ã®å¤ãå¾ããã.
  let v: obj = sheet.["A1:B3"] |> gets |> head

  // åå¾ããåä»ãã®éåãã¼ã¿ããåé ­è¦ç´ ã®å¤ãåå¾.
  //   -> ä»¥ä¸ã®å ´å sheet.["A1"] ã®å¤ãå¾ããã.
  //   -> æå®ããåã¨äºææ§ããªãå ´å, System.InvalidCastException.
  let v: int = sheet.["A1:B3"] |> gets<int> |> head

  // åå¾ããéåãã¼ã¿ããæ«å°¾è¦ç´ ã®å¤ãåå¾.
  //   -> ä»¥ä¸ã®å ´å sheet.["B3"] ã®å¤ãå¾ããã.
  let v: obj = sheet.["A1:B3"] |> gets |> last

  // åå¾ããåä»ãã®éåãã¼ã¿ããæ«å°¾è¦ç´ ã®å¤ãåå¾.
  //   -> ä»¥ä¸ã®å ´å sheet.["B3"] ã®å¤ãå¾ããã.
  //   -> æå®ããåã¨äºææ§ããªãå ´å, System.InvalidCastException.
  let v: int = sheet.["A1:B3"] |> gets<int> |> last
```

### â¼â» é¢æ°ãåå¾ãã<br>`getfx (cell: IExcelRange): string`

```fsharp
[<EntryPoint>]
let main argv =
  use excel = open' "C:/work/sample.xlsx"
  let sheet = excel |> workbook(1) |> worksheet(1)

  // é¢æ°ãåå¾.
  let fn: string = sheet.["A1"] |> getfx
```

### â¼â» é¢æ°ãéåãã¼ã¿ã§åå¾ãã<br>`getsfx (range: IExcelRange): string [,]`

```fsharp
[<EntryPoint>]
let main argv =
  use excel = open' "C:/work/sample.xlsx"
  let sheet = excel |> workbook(1) |> worksheet(1)

  // é¢æ°ãéåãã¼ã¿ã§åå¾.
  let fns: string [,] = sheet.["A1:A3"] |> getsfx
```

### â¼â» å¤ãè¨­å®ãã<br>`set (value: obj) (target: IExcelRange): unit`

```fsharp
[<EntryPoint>]
let main argv =
  use excel = open' "C:/work/sample.xlsx"
  let sheet = excel |> workbook(1) |> worksheet(1)

  // å¤ãè¨­å®.
  sheet.["A1"] |> set 100
  sheet.["A1:B3"] |> set 100
```

### â¼â» é¢æ°ãè¨­å®ãã<br>`fx (func: string) (target: IExcelRange): unit`

```fsharp
[<EntryPoint>]
let main argv =
  use excel = open' "C:/work/sample.xlsx"
  let sheet = excel |> workbook(1) |> worksheet(1)

  // é¢æ°ãè¨­å®.
  sheet.["A1"] |> fx "SUM(A2:A5)"
  sheet.["A1:B3"] |> fx "COUNT(A1:B3)"
```

### â¼â» èæ¯è²ãè¨­å®ãã<br>`bgcolor (color: Color) (target: IExcelRange): unit`

```fsharp
[<EntryPoint>]
let main argv =
  use excel = open' "C:/work/sample.xlsx"
  let sheet = excel |> workbook(1) |> worksheet(1)

  // èæ¯è²ãè¨­å®.
  sheet.["A1"] |> bgcolor Color.Red
  sheet.["B1:B3"] |> bgcolor Color.Blue
  sheet.["C1"] |> bgcolor (rgb(0, 128, 255))
```

### â¼â» èæ¯ãã¿ã¼ã³ãè¨­å®ãã<br>`bgpattern (pattern: Pattern) (target: IExcelRange): unit`

| arg name | values |
| --- | --- |
| pattern | `pattern'none`<br>`pattern'auto`<br>`pattern'up`<br>`pattern'down`<br>`pattern'vertical`<br>`pattern'horizontal`<br>`pattern'lightUp`<br>`pattern'lightDown`<br>`pattern'lightVertical`<br>`pattern'lightHorizontal`<br>`pattern'gray8`<br>`pattern'gray16`<br>`pattern'gray25`<br>`pattern'gray50`<br>`pattern'gray75`<br>`pattern'semigray75`<br>`pattern'solid`<br>`pattern'checker`<br>`pattern'grid`<br>`pattern'crisscross`<br>`pattern'linearGradient`<br>`pattern'rectangularGradient` |

```fsharp
[<EntryPoint>]
let main argv =
  use excel = open' "C:/work/sample.xlsx"
  let sheet = excel |> workbook(1) |> worksheet(1)

  // èæ¯ãã¿ã¼ã³ãè¨­å®.
  sheet.["A1"] |> bgpattern pattern'checker
  sheet.["B1:B3"] |> bgpattern pattern'crisscross
```

### â¼â» ç½«ç·ãè¨­å®ãã<br>`ruledline (target: IExcelRange): IBorders` ã³ã³ãã¥ãã¼ã·ã§ã³å¼

#### ð `ruledline` ã§å©ç¨ã§ããã«ã¹ã¿ã ãªãã¬ã¼ã·ã§ã³

| operation name | description |
| --- | --- |
| `top (border)` | æä¸é¨ã®æ¨ªç½«ç·. |
| `bottom (border)` | æä¸é¨ã®æ¨ªç½«ç·. |
| `left (border)` | æå·¦é¨ã®ç¸¦ç½«ç·. |
| `right (border)` | æå³é¨ã®ç¸¦ç½«ç·. |
| `horizontal (border)` | ä¸­éé¨ã®æ¨ªç½«ç·. |
| `vertical (border)` | ä¸­éé¨ã®ç¸¦ç½«ç·. |
| `growing (border)` | å·¦ä¸ããå³ä¸ã«åãã¦ã®ç½«ç·. è²ãå¤ªãã®è¨­å®ã¯ `falling` ã¨å±æ. |
| `falling (border)` | å·¦ä¸ããå³ä¸ã«åãã¦ã®ç½«ç·. è²ãå¤ªãã®è¨­å®ã¯ `growing` ã¨å±æ. |

#### ð `Border` ã«è¨­å®ã§ããå¤

| operation name | description | values |
| --- | --- | --- |
| `LineStyle` | ç½«ç·ã®ã¹ã¿ã¤ã«.<br>**default: `linestyle'continuous`** | `linestyle'none`<br>`linestyle'dot`<br>`linestyle'double`<br>`linestyle'dash`<br>`linestyle'continuous`<br>`linestyle'dashdot`<br>`linestyle'dashdotdot`<br>`linestyle'slant`|
| `Weight` | ç½«ç·ã®å¤ªã.<br>**default: `weight'medium`** | `weight'medium`<br>`weight'hairline`<br>`weight'thin`<br>`weight'thick` |
| `Color` | ç½«ç·ã®è².<br>**default: `Color.Black`** | `Color.Red`<br>`Color.Orange`<br>`Color.Blue`<br>`rgb(r, g, b)`<br>and more... |

```fsharp
[<EntryPoint>]
let main argv =
  use excel = open' "C:/work/sample.xlsx"
  let sheet = excel |> workbook(1) |> worksheet(1)

  // ç½«ç·ãè¨­å®.
  ruledline sheet.["B2:C5"] {
    // å Border ã®å¤ã¯ with ãå©ç¨ãã¦æå®.
    top { border with Color= Color.Red }
    left { border with Color= Color.Orange; Weight= weight'thick }
    right { border with LineStyle= linestyle'dashdot }
    bottom { border with Weight= weight'medium }
    horizontal { border with Color= Color.Blue; Weight= weight'medium }
    vertical { border with Color= rgb (0, 128, 255); Weight= weight'thin }

    // growing ã¨ falling ã¯å¤ãExcelåé¨ã§å±æããã¦ãããããè¨­å®å¤ã¯å¾åã¡ãã.
    growing { border with Weight= weight'hairline }
    falling { border with Weight= weight'thick }
  }
  |> ignore
```

### â¼â» ãã©ã³ããè¨­å®ãã / `font (target: IExcelRange): IRangeFont` ã³ã³ãã¥ãã¼ã·ã§ã³å¼

#### ð `font` ã§å©ç¨ã§ããã«ã¹ã¿ã ãªãã¬ã¼ã·ã§ã³

| operation name | description | values |
| --- | --- | --- |
| `set (fontName: string)`<br>`name (fontName: string)` | ãã©ã³ãå. | `æ¸¸ã´ã·ãã¯`<br>`ã¡ã¤ãªãª`<br>`consolas`<br>and more... |
| `set (size: float)`<br>`size (size: float)` | ãã©ã³ããµã¤ãº. | `8.0`<br>`10.5`<br>`24.0`<br>and more... |
| `set (style: FontStyle)` | ãã©ã³ãã¹ã¿ã¤ã«. `Flags` ãªã®ã§è¤æ°ã¾ã¨ãã¦æå®å¯è½. | `style'normal`<br>`style'bold`<br>`style'italic'`<br>`style'shadow`<br>`style'strikethrough`<br>`style'subscript`<br>`style'superscript`<br>`style'singleUnderline`<br>`style'doubleUnderline` |
| `set (value: Color)` | ãã©ã³ãè². | `Color.Red`<br>`Color.Orange`<br>`Color.Blue`<br>and more... |
| `set (value: RGB)` | ãã©ã³ãè². | `rgb(r: int, g: int, b: int)`<br>`{ r= 0; g= 128; b= 255; }` |
| `set (style: Underline)` | ä¸ç·. | `underline'none`<br>`underline'double`<br>`underline'doubleAccounting`<br>`underline'single`<br>`underline'singleAccounting` |
| `bold (on: bool)` | å¤ªå­. | `true` or `false` |
| `italic (on: bool)` | ã¤ã¿ãªãã¯ä½. | `true` or `false` |
| `shadow (on: bool)` | ãã©ã³ãå½±. | `true` or `false` |
| `outline (on: bool)` | ã¢ã¦ãã©ã¤ã³ãã©ã³ã. | `true` or `false` |
| `strikethrough (on: bool)` | æã¡æ¶ãç·. | `true` or `false` |
| `subscript (on: bool)` | ä¸ä»ãæå­ã«ãã. | `true` or `false` |
| `superscript (on: bool)` | ä¸ä»ãæå­ã«ãã. | `true` or `false` |

```fsharp
[<EntryPoint>]
let main argv =
  use excel = open' "C:/work/sample.xlsx"
  let sheet = excel |> workbook(1) |> worksheet(1)

  // ãã©ã³ããè¨­å®.
  font sheet.["A1:A3"] {
    // ãã©ã³ãã®æå®.
    set "ã¡ã¤ãªãª"  // or
    name "ã¡ã¤ãªãª"
    // ãã©ã³ããµã¤ãºã®è¨­å®.
    set 16.0  // or
    size 16.0
    // ä¸ç·ã®è¨­å®.
    set underline'double

    // ãã©ã³ãè²ã®è¨­å®.
    set Color.Orange          // or
    set ( rgb(0, 128, 255) )  // or
    set { r= 0; g= 128; b= 255; }

    // ãã©ã³ãã¹ã¿ã¤ã«ã®è¨­å®.
    set style'normal
    // ã¹ã¿ã¤ã«ãè¤æ°é¸æããå ´åã¯ä»¥ä¸ã®ããã«æå®.
    set (style'normal ||| style'strikethrough ||| style'shadow)
    // style ãç´æ¥æå®ããªãã¨ãåç¨®ã¹ã¿ã¤ã«ãã²ã¨ã¤ãã¤ ON/OFF å¯è½.
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

### â¼â» IExcelApplicationãªãã¸ã§ã¯ããæä½ãã<br>`excel'op (excel: IExcelApplication)` ã³ã³ãã¥ãã¼ã·ã§ã³å¼

#### ð `excel'op` ã§å©ç¨ã§ããã«ã¹ã¿ã ãªãã¬ã¼ã·ã§ã³

| operation name | description | values |
| --- | --- | --- |
| `set (mode: Calculation)` | Excelã®åè¨ç®å¶å¾¡ãè¨­å®ãã.<br>**default: `calc'manual`** | `calc'auto`<br>`calc'manual`<br>`calc'semiauto` |
| `set (visibility: AppVisibility)` | Excelã®è¡¨ç¤ºç¶æãè¨­å®ãã.<br>**default: `visibility'hidden`** | `visibility'visible`<br>`visibility'hidden` |

```fsharp
[<EntryPoint>]
let main argv =
  use excel = open' "C:/work/sample.xlsx"
  
  excel'op excel {
    set calc'semiauto
    set visibility'visible
  }
```

### â¼â» IExcelRangeãªãã¸ã§ã¯ããæä½ããï¼ã³ãã¼ã»ãã¼ã¹ãã»æ¿å¥ã»åé¤ãªã©ï¼<br>`sheet'op (sheet): IWorksheet)` ã³ã³ãã¥ãã¼ã·ã§ã³å¼

#### ð `sheet'op` ã§å©ç¨ã§ããã«ã¹ã¿ã ãªãã¬ã¼ã·ã§ã³

| operation name | description | values |
| --- | --- | --- |
| `copy (target: string)` | å¯¾è±¡ãã¯ãªãããã¼ãã«ã³ãã¼ãã. | - |
| `paste (target: string, pasteMode: PasteMode)` | å¯¾è±¡ã«ã¯ãªãããã¼ãã®å¤ãè²¼ãä»ãã. | - |
| `insert (target: string, insertMode: InsertMode)` | å¯¾è±¡ã«ã¯ãªãããã¼ãã®å¤ãæ¿å¥ãã. | - |
| `delete (target: string, deleteMode: DeleteMode)` | å¯¾è±¡ãåé¤ãã. | - |
| `set (target: string, value: obj)` | å¯¾è±¡ã«å¤ãè¨­å®ãã. | - |
| `set (target: string, color: Color)` | å¯¾è±¡ã®èæ¯è²ãè¨­å®ãã. | `Color.Red`<br>`Color.Orange`<br>`Color.Blue`<br>and more... |
| `set (target: string, theme: ThemeColor)` | å¯¾è±¡ã®èæ¯è²ããã¼ãã«ã©ã¼ã§è¨­å®ãã. | `theme'background1`<br>`theme'background2`<br>`theme'foreground1`<br>`theme'foreground2`<br>`theme'accent1`<br>`theme'accent2`<br>`theme'accent3`<br>`theme'accent4`<br>`theme'accent5`<br>`theme'accent6`<br> |
| `set (target: string, pattern: Pattern)` | å¯¾è±¡ã®èæ¯ãã¿ã¼ã³ãè¨­å®ãã. | `pattern'none`<br>`pattern'auto`<br>`pattern'up`<br>`pattern'down`<br>`pattern'vertical`<br>`pattern'horizontal`<br>`pattern'lightUp`<br>`pattern'lightDown`<br>`pattern'lightVertical`<br>`pattern'lightHorizontal`<br>`pattern'gray8`<br>`pattern'gray16`<br>`pattern'gray25`<br>`pattern'gray50`<br>`pattern'gray75`<br>`pattern'semigray75`<br>`pattern'solid`<br>`pattern'checker`<br>`pattern'grid`<br>`pattern'crisscross`<br>`pattern'linearGradient`<br>`pattern'rectangularGradient` |
| `set (target: string, halign: HorizontalAlignment)` |æå­ã®æ°´å¹³ä½ç½®ãè¨­å®ãã. | `h'right`<br>`h'left`<br>`h'center`<br>`h'justify`<br>`h'distributed`<br>`h'general`<br>`h'fill`<br>`h'centerAcrossSelection` |
| `set (target: string, valign: VerticalAlignment)` |æå­ã®åç´ä½ç½®ãè¨­å®ãã. | `v'top`<br>`v'bottom`<br>`v'center `<br>`v'justify`<br>`v'distributed` |
| `fx (target: string, formula: string)` | å¯¾è±¡ã«é¢æ°ãè¨­å®ãã. | - |
| `width (target: string, length: int)` | å¯¾è±¡ã®åå¹ãptåä½ã§è¨­å®ãã. | - |
| `height (target: string, length: int)` | å¯¾è±¡ã®è¡é«ãptåä½ã§è¨­å®ãã. | - |
| `fit'width (target: string)` | å¯¾è±¡ã®åå¹ãèªåèª¿æ´ãã. | - |
| `fit'height (target: string)` | å¯¾è±¡ã®è¡é«ãèªåèª¿æ´ãã. | - |
| `merge (target: string, across: bool)` | ã»ã«ãçµåãã. | `true`: ç¯å²åã®ã»ã«ãè¡ãã¨ã«çµå.<br>`false`: ç¯å²åãã¹ã¦ã®ã»ã«ã1ã¤ã«çµå. |
| `unmerge (target: string)` | ã»ã«ã®çµåãè§£é¤ãã. | - |
| `wrap (target: string)` | æãè¿ãã¦å¨ä½ãè¡¨ç¤ºãã. | - |
| `unwrap (target: string)` | æãè¿ãã¦å¨ä½ãè¡¨ç¤ºãè§£é¤ãã. | - |
| `shrink (target: string)` | ç¸®å°ãã¦å¨ä½ãè¡¨ç¤ºãã. | - |
| `unshrink (target: string)` | ç¸®å°ãã¦å¨ä½ãè¡¨ç¤ºãè§£é¤ãã. | - |
| `orientation (target: string, angle: int)` | æå­ã®æ¹åãè¨­å®ãã. | -90Â° ~ 90Â° |
| `format (target: string, format: string)` | ã»ã«ã®å¤ã®è¡¨ç¤ºå½¢å¼ãè¨­å®ãã. | - |

#### ð `PasteMode` ã®è¦ç´ 

| name | description | values |
| --- | --- | --- |
| `Paste` | è²¼ãä»ãæ¹å¼.<br>**default: `paste'all`** | `paste'values`<br>`paste'comments`<br>`paste'formulas`<br>`paste'formats`<br>`paste'all`<br>`paste'validation`<br>`paste'exceptBorders`<br>`paste'colmnWidths`<br>`paste'formulasAndNumberFormats`<br> `paste'valuesAndNumberFormats`<br>`paste'allUsingSourceTheme`<br>`paste'allMergingConditionalFormats` |
| `Op` | è²¼ãä»ãæã®æ¼ç®æ¹æ³.<br>**default: `op'none`** | `op'none`<br>`op'add`<br>`op'sub`<br>`op'mul`<br>`op'div` |
| `SkipBlanks` | ç©ºç½ã»ã«ãç¡è¦ããã.<br>**default: `false`** | `true` or `false` |
| `Transpose` | è¡åãå¥ãæ¿ããã.<br>**default: `false`** | `true` or `false` |

#### ð `InsertMode` ã®è¦ç´ 

| name | description | values |
| --- | --- | --- |
| `Shift` | æ¿å¥å¾ã«ä»ã®ã»ã«ãã©ãã·ããç§»åããã.<br>**default: `shift'down`** | `shift'right`<br>`shift'down` |
| `Origin` | æ¸å¼ãã³ãã¼ãã¦ããæ¹å.<br>**default: `origin'right`, `origin'below`** | `origin'left`<br>`origin'above`<br>`origin'right`<br>`origin'below` |

#### ð `DeleteMode` ã®è¦ç´ 

| name | description | values |
| --- | --- | --- |
| `Shift` | åé¤å¾ã«ä»ã®ã»ã«ãã©ãã·ããç§»åããã. | `shift'left`<br>`shift'up` |


```fsharp
[<EntryPoint>]
let main argv =
  use excel = open' "C:/work/sample.xlsx"
  let sheet = excel |> workbook(1) |> worksheet(1)
  
  sheet'op sheet {
    // A1 ãã¯ãªãããã¼ãã«ã³ãã¼.
    copy "A1"
    // ç¯å²ã³ãã¼ãå¯è½.
    copy "A1:A3"

    // ã¯ãªãããã¼ãã®ãã¼ã¿ã B1 ã«è²¼ãä»ã.
    paste "B1" paste'mode
    paste "B1" { paste'mode with Paste= paste'values }
    paste "B1" { paste'mode with Op= op'add }
    paste "B1" { paste'mode with SkipBlanks= true }
    paste "B1" { paste'mode with Transpose= true }
    paste "B1" { paste'mode with Paste= paste'values; SkipBlanks= true }
    paste "B1" { paste'mode with Paste= paste'values; Transpose= true }
    // ç¯å²è²¼ãä»ããå¯è½.
    paste "B1:B3" paste'mode

    // ã¯ãªãããã¼ãã®ãã¼ã¿ã C1 ã«æ¿å¥.
    insert "C1" insert'mode
    insert "C1" { insert'mode with Shift= shift'down }
    insert "C1" { insert'mode with Origin= origin'below }
    insert "C1" { insert'mode with Shift= shift'right; Origin= origin'below }
    // ç¯å²æ¿å¥ãå¯è½.
    insert "C1:C3" insert'mode

    // A1 ã®ãã¼ã¿ãåé¤ãã.
    delete "A1" delete'mode
    delete "A1" { delete'mode with Shift= shift'up }
    // ç¯å²åé¤ãå¯è½.
    delete "A1:A3" delete'mode

    // A1 ã«å¤ãè¨­å®.
    set "A1" 100
    set "A1" sheet.["B1"]

    // A1 ã«é¢æ°ãè¨­å®.
    fx "A1" "SUM(A1:B3)"
    fx "A1" sheet.["B1"]
    
    // åå¹ã 100 ã«è¨­å®.
    width "A1" 100
    width "A1:B3" 100
    // è¡é«ã 100 ã«è¨­å®.
    height "A1" 100
    height "A1:B3" 100

    // åå¹ãèªåèª¿æ´.
    fit'width "A1"
    fit'width "A1:B3"
    // è¡é«ãèªåèª¿æ´.
    fit'height "A1"
    fit'height "A1:B3"

    // èæ¯è²ãè¨­å®.
    set "A1" Color.Blue
    // èæ¯è²ããã¼ãã«ã©ã¼ã§è¨­å®.
    set "A1" theme'accent1
    // èæ¯ãã¿ã¼ã³ãè¨­å®.
    set "A1" pattern'horizontal

    // ã»ã«ãçµå.
    merge "C1:D3" true
    merge "E1:F3" false
    // ã»ã«ã®çµåãè§£é¤.
    unmerge "C1:D3"
    unmerge "E1:F3"

    // æãè¿ãã¦å¨ä½ãè¡¨ç¤º.
    wrap "A1"
    // æãè¿ãã¦å¨ä½ãè¡¨ç¤ºãè§£é¤.
    unwrap "A1"

    // ç¸®å°ãã¦å¨ä½ãè¡¨ç¤º.
    shrink "A1"
    // ç¸®å°ãã¦å¨ä½ãè¡¨ç¤ºãè§£é¤.
    unshrink "A1"

    // æå­ã®æ¹åãè¨­å®.
    orientation "A1" -90
    orientation "A1" 0
    orientation "A1" 90

    // è¡¨ç¤ºå½¢å¼ãè¨­å®.
    format "A1" "(æ¥ä»)yyyy-MM-dd"
  }
```

### â¼â» ãã£ã¼ããã°ã©ããæ¿å¥ãã<br>`chart'op (sheet: IWorksheet)` ã³ã³ãã¥ãã¼ã·ã§ã³å¼

#### ð `chart'op` ã§å©ç¨ã§ããã«ã¹ã¿ã ãªãã¬ã¼ã·ã§ã³

| operation name | description | values |
| --- | --- | --- |
| `select (target: string)` | ãã£ã¼ãã§å©ç¨ãããã¼ã¿ã®ç¯å²ãé¸æãã. | - |
| `size (size: (int<cols> * int<rows>))` | ãã£ã¼ãã®ãµã¤ãºãã»ã«æ°ã§æå®ãã. (å¹ * é«ã). | - |
| `position (position: string)` | ãã£ã¼ããæ¿å¥ããä½ç½®ãæå®ãã. | - |
| `add (recipe: ChartRecipe, ?newLayout: bool)` | ãã£ã¼ããè¿½å ãã. | `newLayout`: åçæ¸å¼è¨­å®è¦åãä½¿ç¨ãã (default: true). |

#### ð `ChartRecipe` ã®è¦ç´ 

| name | description |
| --- | --- |
| `columnClustered` | éåç¸¦æ£ |
| `columnStacked` | ç©ã¿ä¸ãç¸¦æ£ |
| `columnStacked100` | 100% ç©ã¿ä¸ãç¸¦æ£ |
| `barStacked` | ç©ã¿ä¸ãæ¨ªæ£ |
| `barStacked100` | 100% ç©ã¿ä¸ãæ¨ªæ£ |
| `column3d` | 3-D ç¸¦æ£ |
| `columnClustered3d` | 3-D éåç¸¦æ£ |
| `coneCol` | 3-D åéåç¸¦æ£ |
| `coneColClustered` | éååéå ç¸¦æ£ |
| `cylinderCol` | 3-D åæ±å ç¸¦æ£ |
| `cylinderColClustered` | éååéå ç¸¦æ£ |
| `pyramidCol` | 3-D ãã©ãããåç¸¦æ£ |
| `pyramidColClustered` | éåãã©ãããå ç¸¦æ£ |
| `columnStacked3d` | 3-D ç©ã¿ä¸ãç¸¦æ£ |
| `columnStacked3d100` | 3-D 100% ç©ã¿ä¸ãç¸¦æ£ |
| `cylinderColStacked` | ç©ã¿ä¸ãåéå ç¸¦æ£ |
| `cylinderColStacked100` | 100% ç©ã¿ä¸ãåæ±å ç¸¦æ£ |
| `cylinderBarStacked` | ç©ã¿ä¸ãåæ±å æ¨ªæ£ |
| `cylinderBarStacked100` | 100% ç©ã¿ä¸ãåæ±å æ¨ªæ£ |
| `coneColStacked` | ç©ã¿ä¸ãåéå ç¸¦æ£ |
| `coneColStacked100` | 100% ç©ã¿ä¸ãåéå ç¸¦æ£ |
| `coneBarStacked` | ç©ã¿ä¸ãåéå æ¨ªæ£ |
| `coneBarStacked100` | 100% ç©ã¿ä¸ãåéå æ¨ªæ£ |
| `pyramidColStacked` | ç©ã¿ä¸ããã©ãããå ç¸¦æ£ |
| `pyramidColStacked100` | 100% ç©ã¿ä¸ããã©ãããå ç¸¦æ£ |
| `pyramidBarStacked` | ç©ã¿ä¸ããã©ãããå æ¨ªæ£ |
| `pyramidBarStacked100` | 100% ç©ã¿ä¸ããã©ãããå æ¨ªæ£ |
| `barStacked3d` | 3-D ç©ã¿ä¸ãæ¨ªæ£ |
| `barStacked3d100` | 3-D 100% ç©ã¿ä¸ãæ¨ªæ£ |
| `barClustered` | éåæ¨ªæ£ |
| `barClustered3d` | 3-D éåæ¨ªæ£ |
| `cylinderBarClustered` | éååæ±å æ¨ªæ£ |
| `coneBarClustered` | éååéå æ¨ªæ£ |
| `pyramidBarClustered` | éåãã©ãããå æ¨ªæ£ |
| `lineStacked` | ç©ã¿ä¸ãæãç· |
| `lineStacked100` | 100% ç©ã¿ä¸ãæãç· |
| `lineMarkersStacked100` | ãã¼ã«ã¼ä»ã 100% ç©ã¿ä¸ãæãç· |
| `line` | æãç· |
| `lineMarkers` | ãã¼ã«ã¼ä»ãæãç· |
| `lineMarkersStacked` | ãã¼ã«ã¼ä»ãç©ã¿ä¸ãæãç· |
| `pieOfPie` | è£å©åã°ã©ãä»ãå |
| `barOfPie` | è£å©ç¸¦æ£ã°ã©ãä»ãå |
| `doughnut` | ãã¼ãã |
| `doughnutExploded` | åå²ãã¼ãã |
| `pie` | å |
| `pieExploded` | åå²å |
| `pie3d` | 3-D å |
| `pieExploded3d` | åå² 3-D å |
| `xyScatter` | æ£å¸å³ |
| `xyScatterSmooth` | å¹³æ»ç·ä»ãæ£å¸å³ |
| `xyScatterSmoothNoMarkers` | å¹³æ»ç·ä»ãæ£å¸å³ï¼ãã¼ã¿ ãã¼ã«ã¼ãªãï¼ |
| `xyScatterLines` | æãç·ä»ãæ£å¸å³ |
| `xyScatterLinesNoMarkers` | æãç·ä»ãæ£å¸å³ï¼ãã¼ã¿ ãã¼ã«ã¼ãªãï¼ |
| `area` | é¢ |
| `areaStacked` | ç©ã¿ä¸ãé¢ |
| `areaStacked100` | 100% ç©ã¿ä¸ãé¢ |
| `area3d` | 3-D é¢ |
| `areaStacked3d` | 3-D ç©ã¿ä¸ãé¢ |
| `areaStacked1003d` | 3-D 100% ç©ã¿ä¸ãé¢ |
| `radar` | ã¬ã¼ãã¼ |
| `radarMarkers` | ãã¼ã¿ ãã¼ã«ã¼ä»ãã¬ã¼ãã¼ |
| `radarFilled` | å¡ãã¤ã¶ãã¬ã¼ãã¼ |
| `surface` | 3-D è¡¨é¢ |
| `surfaceWireframe` | 3-D è¡¨é¢ï¼ã¯ã¤ã¤ã¼ãã¬ã¼ã ï¼ |
| `surfaceTopView` | è¡¨é¢ï¼ããããã¥ã¼ï¼ |
| `surfaceTopViewWireframe` | è¡¨é¢ï¼ããããã¥ã¼ã»ã¯ã¤ã¤ã¼ãã¬ã¼ã ï¼ |
| `line3d` | 3-D æãç· |
| `bubble` | ããã« |
| `bubble3dEffect` | 3-D å¹æä»ãããã« |
| `stockHLC` | é«å¤-å®å¤-çµå¤ |
| `stockOHLC` | å§å¤-é«å¤-å®å¤-çµå¤ |
| `stockVHLC` | åºæ¥é«-é«å¤-å®å¤-çµå¤ |
| `stockVOHLC` | åºæ¥é«-å§å¤-é«å¤-å®å¤-çµå¤ |
| `sunburst` | ãµã³ãã¼ã¹ã |


```fsharp
[<EntryPoint>]
let main argv =
  use excel = open' "C:/work/sample.xlsx"
  let sheet = excel |> workbook(1) |> worksheet(1)
  
  sheet.["A1:F1"] |> set [| 100; 120; 110; 100; 200; 180; |]

  chart'op sheet {
    select "A1:F1"
    position "A2"
    size (6<cols>, 10<rows>)
    add ChartRecipe.line
  } |> ignore

```

### â¼â» IExcelRangeãåé¤ãã<br>`delete (direction: DeleteShiftDirection) (target: IExcelRange): unit`

#### ð `DeleteShiftDirection`

| value | description |
| --- | --- |
| `shift'left` | åé¤å¾, å·¦æ¹åã¸ã·ãã. |
| `shift'up` | åé¤å¾, ä¸æ¹åã¸ã·ãã. |

```fsharp
[<EntryPoint>]
let main argv =
  use excel = open' "C:/work/sample.xlsx"
  let sheet = excel |> workbook(1) |> worksheet(1)
  
  // å¯¾è±¡ãåé¤.
  sheet.["A1"] |> delete shift'up
  sheet.["A1:A3"] |> delete shift'left
```

---

## ð· Utility  

### â¼â» æ°å¤ãã«ã©ã åã«å¤æãã<br>`column'name (index: int): string`

```fsharp
let name = 1 |> column'name     // A
let name = 10 |> column'name    // J
let name = 128 |> column'name   // DX
```

### â¼â» ã«ã©ã åãã¤ã³ããã¯ã¹ã«å¤æãã<br>`column'number (column: string): int`

```fsharp
let number = "A" |> column'number     // 1
let number = "J" |> column'number     // 10
let number = "DX" |> column'number    // 128
```

### â¼â» IExcelRangeããã¢ãã¬ã¹ãåå¾ãã<br>`address (target: IExcelRange): string`

```fsharp
let adds = sheet.["A1"] |> address      // $A$1
let adds = sheet.["A1:B3"] |> address   // $A$1:$B$3
```

### â¼â» ExcelObjectãé¸æãã<br>`activate (target: ^T): unit`

```fsharp
// Workbookãé¸æç¶æã«ãã.
excel |> workbook(1) |> activate

// Worksheetãé¸æç¶æã«ãã.
excel |> workbook(1) |> worksheet(1) |> activate

// Cellãé¸æç¶æã«ãã.
sheet.["B1"] |> activate
sheet.["A1:B3"] |> activate
```

### â¼â» ExcelObjectãé¸æãã<br>`select (target: ^T): unit`

```fsharp
// Worksheet(1)ãé¸æç¶æã«ãã.
excel |> workbook(1) |> worksheet(1) |> select

// Cellãé¸æç¶æã«ãã.
sheet.["B1"] |> select
sheet.["D1:E3"] |> select
```

---

## ð· TIPS  

### â¼â» `try-with` ã®å©ç¨  

ä¾å¤å¦çãæ½ãã¦ããªãå ´å Excel COM ãªãã¸ã§ã¯ããé©åã«è§£æ¾ããã, ãã­ã»ã¹ä¸ã«æ®ã£ã¦ãã¾ãæããããã¾ã.  
`try-with` (ã¾ãã¯ `try-with-finally`) ã¨ `use` ãä½µç¨ãããã¨ã§ Excel COM ãªãã¸ã§ã¯ãã®è§£æ¾æ¼ããé²ãã¾ã.  

```fsharp
try
  // use ãå©ç¨ãã.
  use excel = create ()

  // do somethings

with
  _ -> ()
```  

ã¾ã, F# Interactive ã§å©ç¨ããå ´å, `attach` ãããã¨ã¯å¿ã `detach` ããå¿è¦ãããã¾ã.  

```powershell
let ps = enumerate ();;
let excel = ps.[0] |> attach;;

# do somethings

excel |> detach;;
```

### â¼â» `LangVersion` ã®æå®  

**.NET 5** ãå©ç¨ãã¦ããå ´å, ãã¾ãã³ã³ãã¥ãã¼ã·ã§ã³å¼ãåä½ããªãå¯è½æ§ãããã¾ã.  
ãã®å ´åã¯ `LangVersion` ã« **preview** ãæå®ããããã«ãã¦ãã ãã.

```xml
<!-- .fsproj ãæ¸ãæããå ´å -->
<PropertyGroup>
  <LangVersion>preview</LangVersion>
</PropertyGroup>
```

```powershell
# dotnet fsi ã³ãã³ãã«ãªãã·ã§ã³ãæå®ããå ´å
dotnet fsi --langversion:preview
```
