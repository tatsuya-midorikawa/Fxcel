namespace Fxcel

open Midoliy.Office.Interop
open System.Drawing

[<AutoOpen>]
module SheetOp =
  type PasteMode = { Paste: PasteType; Op: PasteOperation; SkipBlanks: bool; Transpose: bool; }
  type InsertMode = { Shift: InsertShiftDirection; Origin: InsertFormatOrigin; }
  type DeleteMode = { Shift: DeleteShiftDirection; }

  type SheetOpBuilder (sheet: IWorksheet) =
    member __.Yield (_: unit) = ()
    member __.Zero() = ()
    [<CustomOperation("copy")>]
    member __.Copy(_: unit, target: string) =
      sheet.[target].Copy() |> ignore
    [<CustomOperation("copy")>]
    member __.Copy(_: unit, leftTop: string, rightBottom: string) = 
      sheet.[$"%s{leftTop}:%s{rightBottom}"].Copy() |> ignore
    [<CustomOperation("copy")>]
    member __.Copy(_: unit, (r, c): (int<row> * int<col>)) = 
      sheet.[int r, int c].Copy() |> ignore
    [<CustomOperation("copy")>]
    member __.Copy(_: unit, (ltR, ltC): (int<row> * int<col>), (rbR, rbC): (int<row> * int<col>)) = 
      sheet.[$"%s{column'name (int ltC)}%i{ltR}:%s{column'name (int rbC)}%i{rbR}"].Copy() |> ignore
    [<CustomOperation("copy")>]
    member __.Copy(_: unit, (c, r): (string * int<row>)) = 
      sheet.[$"%s{c}%i{r}"].Copy() |> ignore
    [<CustomOperation("copy")>]
    member __.Copy(_: unit, (ltC, ltR): (string * int<row>), (rbC, rbR): (string * int<row>)) = 
      sheet.[$"%s{ltC}%i{ltR}:%s{rbC}%i{rbR}"].Copy() |> ignore

    [<CustomOperation("paste")>]
    member __.Paste(_: unit, target: string, pasteMode: PasteMode) =
      sheet.[target].Paste(pasteMode.Paste, pasteMode.Op, pasteMode.SkipBlanks, pasteMode.Transpose) |> ignore
    [<CustomOperation("paste")>]
    member __.Paste(_: unit, leftTop: string, rightBottom: string, pasteMode: PasteMode) = 
      sheet.[$"%s{leftTop}:%s{rightBottom}"].Paste(pasteMode.Paste, pasteMode.Op, pasteMode.SkipBlanks, pasteMode.Transpose) |> ignore
    [<CustomOperation("paste")>]
    member __.Paste(_: unit, (r, c): (int<row> * int<col>), pasteMode: PasteMode) = 
      sheet.[int r, int c].Paste(pasteMode.Paste, pasteMode.Op, pasteMode.SkipBlanks, pasteMode.Transpose) |> ignore
    [<CustomOperation("paste")>]
    member __.Paste(_: unit, (ltR, ltC): (int<row> * int<col>), (rbR, rbC): (int<row> * int<col>), pasteMode: PasteMode) = 
      sheet.[$"%s{column'name (int ltC)}%i{ltR}:%s{column'name (int rbC)}%i{rbR}"].Paste(pasteMode.Paste, pasteMode.Op, pasteMode.SkipBlanks, pasteMode.Transpose) |> ignore
    [<CustomOperation("paste")>]
    member __.Paste(_: unit, (c, r): (string * int<row>), pasteMode: PasteMode) = 
      sheet.[$"%s{c}%i{r}"].Paste(pasteMode.Paste, pasteMode.Op, pasteMode.SkipBlanks, pasteMode.Transpose) |> ignore
    [<CustomOperation("paste")>]
    member __.Paste(_: unit, (ltC, ltR): (string * int<row>), (rbC, rbR): (string * int<row>), pasteMode: PasteMode) = 
      sheet.[$"%s{ltC}%i{ltR}:%s{rbC}%i{rbR}"].Paste(pasteMode.Paste, pasteMode.Op, pasteMode.SkipBlanks, pasteMode.Transpose) |> ignore

    [<CustomOperation("insert")>]
    member __.Insert(_: unit, target: string, insertMode: InsertMode) =
      sheet.[target].Insert(insertMode.Shift, insertMode.Origin) |> ignore
    [<CustomOperation("insert")>]
    member __.Insert(_: unit, leftTop: string, rightBottom: string, insertMode: InsertMode) = 
      sheet.[$"%s{leftTop}:%s{rightBottom}"].Insert(insertMode.Shift, insertMode.Origin) |> ignore
    [<CustomOperation("insert")>]
    member __.Insert(_: unit, (r, c): (int<row> * int<col>), insertMode: InsertMode) = 
      sheet.[int r, int c].Insert(insertMode.Shift, insertMode.Origin) |> ignore
    [<CustomOperation("insert")>]
    member __.Insert(_: unit, (ltR, ltC): (int<row> * int<col>), (rbR, rbC): (int<row> * int<col>), insertMode: InsertMode) = 
      sheet.[$"%s{column'name (int ltC)}%i{ltR}:%s{column'name (int rbC)}%i{rbR}"].Insert(insertMode.Shift, insertMode.Origin) |> ignore
    [<CustomOperation("insert")>]
    member __.Insert(_: unit, (c, r): (string * int<row>), insertMode: InsertMode) = 
      sheet.[$"%s{c}%i{r}"].Insert(insertMode.Shift, insertMode.Origin) |> ignore
    [<CustomOperation("insert")>]
    member __.Insert(_: unit, (ltC, ltR): (string * int<row>), (rbC, rbR): (string * int<row>), insertMode: InsertMode) = 
      sheet.[$"%s{ltC}%i{ltR}:%s{rbC}%i{rbR}"].Insert(insertMode.Shift, insertMode.Origin) |> ignore

    [<CustomOperation("delete")>]
    member __.Delete(_: unit, target: string, deleteMode: DeleteMode) =
      sheet.[target].Delete(deleteMode.Shift) |> ignore
    [<CustomOperation("delete")>]
    member __.Delete(_: unit, leftTop: string, rightBottom: string, deleteMode: DeleteMode) = 
      sheet.[$"%s{leftTop}:%s{rightBottom}"].Delete(deleteMode.Shift) |> ignore
    [<CustomOperation("delete")>]
    member __.Delete(_: unit, (r, c): (int<row> * int<col>), deleteMode: DeleteMode) = 
      sheet.[int r, int c].Delete(deleteMode.Shift) |> ignore
    [<CustomOperation("delete")>]
    member __.Delete(_: unit, (ltR, ltC): (int<row> * int<col>), (rbR, rbC): (int<row> * int<col>), deleteMode: DeleteMode) = 
      sheet.[$"%s{column'name (int ltC)}%i{ltR}:%s{column'name (int rbC)}%i{rbR}"].Delete(deleteMode.Shift) |> ignore
    [<CustomOperation("delete")>]
    member __.Delete(_: unit, (c, r): (string * int<row>), deleteMode: DeleteMode) = 
      sheet.[$"%s{c}%i{r}"].Delete(deleteMode.Shift) |> ignore
    [<CustomOperation("delete")>]
    member __.Delete(_: unit, (ltC, ltR): (string * int<row>), (rbC, rbR): (string * int<row>), deleteMode: DeleteMode) = 
      sheet.[$"%s{ltC}%i{ltR}:%s{rbC}%i{rbR}"].Delete(deleteMode.Shift) |> ignore

    [<CustomOperation("set")>]
    member __.Set(_: unit, target: string, value: IExcelRange) = 
      sheet.[target].Value <- value.Value
    [<CustomOperation("set")>]
    member __.Set(_: unit, target: string, value: obj) = 
      sheet.[target].Value <- value
    [<CustomOperation("set")>]
    member __.Set(_: unit, leftTop: string, rightBottom: string, value: obj) = 
      sheet.[$"%s{leftTop}:%s{rightBottom}"].Value <- value
    [<CustomOperation("set")>]
    member __.Set(_: unit, (r, c): (int<row> * int<col>), value: obj) = 
      sheet.[int r, int c].Value <- value
    [<CustomOperation("set")>]
    member __.Set(_: unit, (ltR, ltC): (int<row> * int<col>), (rbR, rbC): (int<row> * int<col>), value: obj) = 
      sheet.[$"%s{column'name (int ltC)}%i{ltR}:%s{column'name (int rbC)}%i{rbR}"].Value <- value
    [<CustomOperation("set")>]
    member __.Set(_: unit, (c, r): (string * int<row>), value: obj) = 
      sheet.[$"%s{c}%i{r}"].Value <- value
    [<CustomOperation("set")>]
    member __.Set(_: unit, (ltC, ltR): (string * int<row>), (rbC, rbR): (string * int<row>), value: obj) = 
      sheet.[$"%s{ltC}%i{ltR}:%s{rbC}%i{rbR}"].Value <- value

    [<CustomOperation("set")>]
    member __.Set(_: unit, target: string, color: Color) = 
      sheet.[target].Interior.Color <- color
    [<CustomOperation("set")>]
    member __.Set(_: unit, leftTop: string, rightBottom: string, color: Color) = 
      sheet.[$"%s{leftTop}:%s{rightBottom}"].Interior.Color <- color
    [<CustomOperation("set")>]
    member __.Set(_: unit, (r, c): (int<row> * int<col>), color: Color) = 
      sheet.[int r, int c].Interior.Color <- color
    [<CustomOperation("set")>]
    member __.Set(_: unit, (ltR, ltC): (int<row> * int<col>), (rbR, rbC): (int<row> * int<col>), color: Color) = 
      sheet.[$"%s{column'name (int ltC)}%i{ltR}:%s{column'name (int rbC)}%i{rbR}"].Interior.Color <- color
    [<CustomOperation("set")>]
    member __.Set(_: unit, (c, r): (string * int<row>), color: Color) = 
      sheet.[$"%s{c}%i{r}"].Interior.Color <- color
    [<CustomOperation("set")>]
    member __.Set(_: unit, (ltC, ltR): (string * int<row>), (rbC, rbR): (string * int<row>), color: Color) = 
      sheet.[$"%s{ltC}%i{ltR}:%s{rbC}%i{rbR}"].Interior.Color <- color

    [<CustomOperation("set")>]
    member __.Set(_: unit, target: string, theme: ThemeColor) = 
      sheet.[target].Interior.ThemeColor <- theme
    [<CustomOperation("set")>]
    member __.Set(_: unit, leftTop: string, rightBottom: string, theme: ThemeColor) = 
      sheet.[$"%s{leftTop}:%s{rightBottom}"].Interior.ThemeColor <- theme
    [<CustomOperation("set")>]
    member __.Set(_: unit, (r, c): (int<row> * int<col>), theme: ThemeColor) = 
      sheet.[int r, int c].Interior.ThemeColor <- theme
    [<CustomOperation("set")>]
    member __.Set(_: unit, (ltR, ltC): (int<row> * int<col>), (rbR, rbC): (int<row> * int<col>), theme: ThemeColor) = 
      sheet.[$"%s{column'name (int ltC)}%i{ltR}:%s{column'name (int rbC)}%i{rbR}"].Interior.ThemeColor <- theme
    [<CustomOperation("set")>]
    member __.Set(_: unit, (c, r): (string * int<row>), theme: ThemeColor) = 
      sheet.[$"%s{c}%i{r}"].Interior.ThemeColor <- theme
    [<CustomOperation("set")>]
    member __.Set(_: unit, (ltC, ltR): (string * int<row>), (rbC, rbR): (string * int<row>), theme: ThemeColor) = 
      sheet.[$"%s{ltC}%i{ltR}:%s{rbC}%i{rbR}"].Interior.ThemeColor <- theme

    [<CustomOperation("set")>]
    member __.Set(_: unit, target: string, pattern: Pattern) = 
      sheet.[target].Interior.Pattern <- pattern
    [<CustomOperation("set")>]
    member __.Set(_: unit, leftTop: string, rightBottom: string, pattern: Pattern) = 
      sheet.[$"%s{leftTop}:%s{rightBottom}"].Interior.Pattern <- pattern
    [<CustomOperation("set")>]
    member __.Set(_: unit, (r, c): (int<row> * int<col>), pattern: Pattern) = 
      sheet.[int r, int c].Interior.Pattern <- pattern
    [<CustomOperation("set")>]
    member __.Set(_: unit, (ltR, ltC): (int<row> * int<col>), (rbR, rbC): (int<row> * int<col>), pattern: Pattern) = 
      sheet.[$"%s{column'name (int ltC)}%i{ltR}:%s{column'name (int rbC)}%i{rbR}"].Interior.Pattern <- pattern
    [<CustomOperation("set")>]
    member __.Set(_: unit, (c, r): (string * int<row>), pattern: Pattern) = 
      sheet.[$"%s{c}%i{r}"].Interior.Pattern <- pattern
    [<CustomOperation("set")>]
    member __.Set(_: unit, (ltC, ltR): (string * int<row>), (rbC, rbR): (string * int<row>), pattern: Pattern) = 
      sheet.[$"%s{ltC}%i{ltR}:%s{rbC}%i{rbR}"].Interior.Pattern <- pattern

    [<CustomOperation("set")>]
    member __.Set(_: unit, target: string, halign: HorizontalAlignment) =
      sheet.[target].HorizontalAlignment <- halign
    [<CustomOperation("set")>]
    member __.Set(_: unit, leftTop: string, rightBottom: string, halign: HorizontalAlignment) = 
      sheet.[$"%s{leftTop}:%s{rightBottom}"].HorizontalAlignment <- halign
    [<CustomOperation("set")>]
    member __.Set(_: unit, (r, c): (int<row> * int<col>), halign: HorizontalAlignment) = 
      sheet.[int r, int c].HorizontalAlignment <- halign
    [<CustomOperation("set")>]
    member __.Set(_: unit, (ltR, ltC): (int<row> * int<col>), (rbR, rbC): (int<row> * int<col>), halign: HorizontalAlignment) = 
      sheet.[$"%s{column'name (int ltC)}%i{ltR}:%s{column'name (int rbC)}%i{rbR}"].HorizontalAlignment <- halign
    [<CustomOperation("set")>]
    member __.Set(_: unit, (c, r): (string * int<row>), halign: HorizontalAlignment) = 
      sheet.[$"%s{c}%i{r}"].HorizontalAlignment <- halign
    [<CustomOperation("set")>]
    member __.Set(_: unit, (ltC, ltR): (string * int<row>), (rbC, rbR): (string * int<row>), halign: HorizontalAlignment) = 
      sheet.[$"%s{ltC}%i{ltR}:%s{rbC}%i{rbR}"].HorizontalAlignment <- halign

    [<CustomOperation("set")>]
    member __.Set(_: unit, target: string, valign: VerticalAlignment) =
      sheet.[target].VerticalAlignment <- valign
    [<CustomOperation("set")>]
    member __.Set(_: unit, leftTop: string, rightBottom: string, valign: VerticalAlignment) = 
      sheet.[$"%s{leftTop}:%s{rightBottom}"].VerticalAlignment <- valign
    [<CustomOperation("set")>]
    member __.Set(_: unit, (r, c): (int<row> * int<col>), valign: VerticalAlignment) = 
      sheet.[int r, int c].VerticalAlignment <- valign
    [<CustomOperation("set")>]
    member __.Set(_: unit, (ltR, ltC): (int<row> * int<col>), (rbR, rbC): (int<row> * int<col>), valign: VerticalAlignment) = 
      sheet.[$"%s{column'name (int ltC)}%i{ltR}:%s{column'name (int rbC)}%i{rbR}"].VerticalAlignment <- valign
    [<CustomOperation("set")>]
    member __.Set(_: unit, (c, r): (string * int<row>), valign: VerticalAlignment) = 
      sheet.[$"%s{c}%i{r}"].VerticalAlignment <- valign
    [<CustomOperation("set")>]
    member __.Set(_: unit, (ltC, ltR): (string * int<row>), (rbC, rbR): (string * int<row>), valign: VerticalAlignment) = 
      sheet.[$"%s{ltC}%i{ltR}:%s{rbC}%i{rbR}"].VerticalAlignment <- valign

    [<CustomOperation("fx")>]
    member __.Fx(_: unit, target: string, value: IExcelRange) =
      sheet.[target].Formula <- value.Formula
    [<CustomOperation("fx")>]
    member __.Fx(_: unit, target: string, value: string) =
      sheet.[target].Formula <- if (string value).StartsWith("=") then value else $"={value}"
    [<CustomOperation("fx")>]
    member __.Fx(_: unit, leftTop: string, rightBottom: string, value: string) = 
      sheet.[$"%s{leftTop}:%s{rightBottom}"].Formula <- if (string value).StartsWith("=") then value else $"={value}"
    [<CustomOperation("fx")>]
    member __.Fx(_: unit, (r, c): (int<row> * int<col>), value: string) = 
      sheet.[int r, int c].Formula <- if (string value).StartsWith("=") then value else $"={value}"
    [<CustomOperation("fx")>]
    member __.Fx(_: unit, (ltR, ltC): (int<row> * int<col>), (rbR, rbC): (int<row> * int<col>), value: string) = 
      sheet.[$"%s{column'name (int ltC)}%i{ltR}:%s{column'name (int rbC)}%i{rbR}"].Formula <- if (string value).StartsWith("=") then value else $"={value}"
    [<CustomOperation("fx")>]
    member __.Fx(_: unit, (c, r): (string * int<row>), value: string) = 
      sheet.[$"%s{c}%i{r}"].Formula <- if (string value).StartsWith("=") then value else $"={value}"
    [<CustomOperation("fx")>]
    member __.Fx(_: unit, (ltC, ltR): (string * int<row>), (rbC, rbR): (string * int<row>), value: string) = 
      sheet.[$"%s{ltC}%i{ltR}:%s{rbC}%i{rbR}"].Formula <- if (string value).StartsWith("=") then value else $"={value}"

    [<CustomOperation("width")>]
    member __.SetWidth(_: unit, target: string, length: int) =
      sheet.[target].ColumnWidth <- length
    [<CustomOperation("width")>]
    member __.SetWidth(_: unit, leftTop: string, rightBottom: string, length: int) = 
      sheet.[$"%s{leftTop}:%s{rightBottom}"].ColumnWidth <- length
    [<CustomOperation("width")>]
    member __.SetWidth(_: unit, (r, c): (int<row> * int<col>), length: int) = 
      sheet.[int r, int c].ColumnWidth <- length
    [<CustomOperation("width")>]
    member __.SetWidth(_: unit, (ltR, ltC): (int<row> * int<col>), (rbR, rbC): (int<row> * int<col>), length: int) = 
      sheet.[$"%s{column'name (int ltC)}%i{ltR}:%s{column'name (int rbC)}%i{rbR}"].ColumnWidth <- length
    [<CustomOperation("width")>]
    member __.SetWidth(_: unit, (c, r): (string * int<row>), length: int) = 
      sheet.[$"%s{c}%i{r}"].ColumnWidth <- length
    [<CustomOperation("width")>]
    member __.SetWidth(_: unit, (ltC, ltR): (string * int<row>), (rbC, rbR): (string * int<row>), length: int) = 
      sheet.[$"%s{ltC}%i{ltR}:%s{rbC}%i{rbR}"].ColumnWidth <- length

    [<CustomOperation("height")>]
    member __.SetHeight(_: unit, target: string, length: int) =
      sheet.[target].RowHeight <- length
    [<CustomOperation("height")>]
    member __.SetHeight(_: unit, leftTop: string, rightBottom: string, length: int) = 
      sheet.[$"%s{leftTop}:%s{rightBottom}"].RowHeight <- length
    [<CustomOperation("height")>]
    member __.SetHeight(_: unit, (r, c): (int<row> * int<col>), length: int) = 
      sheet.[int r, int c].RowHeight <- length
    [<CustomOperation("height")>]
    member __.SetHeight(_: unit, (ltR, ltC): (int<row> * int<col>), (rbR, rbC): (int<row> * int<col>), length: int) = 
      sheet.[$"%s{column'name (int ltC)}%i{ltR}:%s{column'name (int rbC)}%i{rbR}"].RowHeight <- length
    [<CustomOperation("height")>]
    member __.SetHeight(_: unit, (c, r): (string * int<row>), length: int) = 
      sheet.[$"%s{c}%i{r}"].RowHeight <- length
    [<CustomOperation("height")>]
    member __.SetHeight(_: unit, (ltC, ltR): (string * int<row>), (rbC, rbR): (string * int<row>), length: int) = 
      sheet.[$"%s{ltC}%i{ltR}:%s{rbC}%i{rbR}"].RowHeight <- length

    [<CustomOperation("fit'width")>]
    member __.FitWidth(_: unit, target: string) =
      sheet.[target].EntireColumn.AutoFit()
    [<CustomOperation("fit'width")>]
    member __.FitWidth(_: unit, leftTop: string, rightBottom: string) = 
      sheet.[$"%s{leftTop}:%s{rightBottom}"].EntireColumn.AutoFit()
    [<CustomOperation("fit'width")>]
    member __.FitWidth(_: unit, (r, c): (int<row> * int<col>)) = 
      sheet.[int r, int c].EntireColumn.AutoFit()
    [<CustomOperation("fit'width")>]
    member __.FitWidth(_: unit, (ltR, ltC): (int<row> * int<col>), (rbR, rbC): (int<row> * int<col>)) = 
      sheet.[$"%s{column'name (int ltC)}%i{ltR}:%s{column'name (int rbC)}%i{rbR}"].EntireColumn.AutoFit()
    [<CustomOperation("fit'width")>]
    member __.FitWidth(_: unit, (c, r): (string * int<row>)) = 
      sheet.[$"%s{c}%i{r}"].EntireColumn.AutoFit()
    [<CustomOperation("fit'width")>]
    member __.FitWidth(_: unit, (ltC, ltR): (string * int<row>), (rbC, rbR): (string * int<row>)) = 
      sheet.[$"%s{ltC}%i{ltR}:%s{rbC}%i{rbR}"].EntireColumn.AutoFit()

    [<CustomOperation("fit'height")>]
    member __.FitHeight(_: unit, target: string) =
      sheet.[target].EntireRow.AutoFit()
    [<CustomOperation("fit'height")>]
    member __.FitHeight(_: unit, leftTop: string, rightBottom: string) = 
      sheet.[$"%s{leftTop}:%s{rightBottom}"].EntireRow.AutoFit()
    [<CustomOperation("fit'height")>]
    member __.FitHeight(_: unit, (r, c): (int<row> * int<col>)) = 
      sheet.[int r, int c].EntireRow.AutoFit()
    [<CustomOperation("fit'height")>]
    member __.FitHeight(_: unit, (ltR, ltC): (int<row> * int<col>), (rbR, rbC): (int<row> * int<col>)) = 
      sheet.[$"%s{column'name (int ltC)}%i{ltR}:%s{column'name (int rbC)}%i{rbR}"].EntireRow.AutoFit()
    [<CustomOperation("fit'height")>]
    member __.FitHeight(_: unit, (c, r): (string * int<row>)) = 
      sheet.[$"%s{c}%i{r}"].EntireRow.AutoFit()
    [<CustomOperation("fit'height")>]
    member __.FitHeight(_: unit, (ltC, ltR): (string * int<row>), (rbC, rbR): (string * int<row>)) = 
      sheet.[$"%s{ltC}%i{ltR}:%s{rbC}%i{rbR}"].EntireRow.AutoFit()

    [<CustomOperation("merge")>]
    member __.Merge(_: unit, target: string, across: bool) =
      sheet.[target].Merge(across)
    [<CustomOperation("merge")>]
    member __.Merge(_: unit, leftTop: string, rightBottom: string, across: bool) = 
      sheet.[$"%s{leftTop}:%s{rightBottom}"].Merge(across)
    [<CustomOperation("merge")>]
    member __.Merge(_: unit, (r, c): (int<row> * int<col>), across: bool) = 
      sheet.[int r, int c].Merge(across)
    [<CustomOperation("merge")>]
    member __.Merge(_: unit, (ltR, ltC): (int<row> * int<col>), (rbR, rbC): (int<row> * int<col>), across: bool) = 
      sheet.[$"%s{column'name (int ltC)}%i{ltR}:%s{column'name (int rbC)}%i{rbR}"].Merge(across)
    [<CustomOperation("merge")>]
    member __.Merge(_: unit, (c, r): (string * int<row>), across: bool) = 
      sheet.[$"%s{c}%i{r}"].Merge(across)
    [<CustomOperation("merge")>]
    member __.Merge(_: unit, (ltC, ltR): (string * int<row>), (rbC, rbR): (string * int<row>), across: bool) = 
      sheet.[$"%s{ltC}%i{ltR}:%s{rbC}%i{rbR}"].Merge(across)

    [<CustomOperation("unmerge")>]
    member __.UnMerge(_: unit, target: string) =
      sheet.[target].UnMerge()
    [<CustomOperation("unmerge")>]
    member __.UnMerge(_: unit, leftTop: string, rightBottom: string) = 
      sheet.[$"%s{leftTop}:%s{rightBottom}"].UnMerge()
    [<CustomOperation("unmerge")>]
    member __.UnMerge(_: unit, (r, c): (int<row> * int<col>)) = 
      sheet.[int r, int c].UnMerge()
    [<CustomOperation("unmerge")>]
    member __.UnMerge(_: unit, (ltR, ltC): (int<row> * int<col>), (rbR, rbC): (int<row> * int<col>)) = 
      sheet.[$"%s{column'name (int ltC)}%i{ltR}:%s{column'name (int rbC)}%i{rbR}"].UnMerge()
    [<CustomOperation("unmerge")>]
    member __.UnMerge(_: unit, (c, r): (string * int<row>)) = 
      sheet.[$"%s{c}%i{r}"].UnMerge()
    [<CustomOperation("unmerge")>]
    member __.UnMerge(_: unit, (ltC, ltR): (string * int<row>), (rbC, rbR): (string * int<row>)) = 
      sheet.[$"%s{ltC}%i{ltR}:%s{rbC}%i{rbR}"].UnMerge()

    [<CustomOperation("wrap")>]
    member __.WrapText(_: unit, target: string) =
      sheet.[target].WrapText <- true
    [<CustomOperation("wrap")>]
    member __.WrapText(_: unit, leftTop: string, rightBottom: string) = 
      sheet.[$"%s{leftTop}:%s{rightBottom}"].WrapText <- true
    [<CustomOperation("wrap")>]
    member __.WrapText(_: unit, (r, c): (int<row> * int<col>)) = 
      sheet.[int r, int c].WrapText <- true
    [<CustomOperation("wrap")>]
    member __.WrapText(_: unit, (ltR, ltC): (int<row> * int<col>), (rbR, rbC): (int<row> * int<col>)) = 
      sheet.[$"%s{column'name (int ltC)}%i{ltR}:%s{column'name (int rbC)}%i{rbR}"].WrapText <- true
    [<CustomOperation("wrap")>]
    member __.WrapText(_: unit, (c, r): (string * int<row>)) = 
      sheet.[$"%s{c}%i{r}"].WrapText <- true
    [<CustomOperation("wrap")>]
    member __.WrapText(_: unit, (ltC, ltR): (string * int<row>), (rbC, rbR): (string * int<row>)) = 
      sheet.[$"%s{ltC}%i{ltR}:%s{rbC}%i{rbR}"].WrapText <- true

    [<CustomOperation("unwrap")>]
    member __.UnWrapText(_: unit, target: string) =
      sheet.[target].WrapText <- false
    [<CustomOperation("unwrap")>]
    member __.UnWrapText(_: unit, leftTop: string, rightBottom: string) = 
      sheet.[$"%s{leftTop}:%s{rightBottom}"].WrapText <- false
    [<CustomOperation("unwrap")>]
    member __.UnWrapText(_: unit, (r, c): (int<row> * int<col>)) = 
      sheet.[int r, int c].WrapText <- false
    [<CustomOperation("unwrap")>]
    member __.UnWrapText(_: unit, (ltR, ltC): (int<row> * int<col>), (rbR, rbC): (int<row> * int<col>)) = 
      sheet.[$"%s{column'name (int ltC)}%i{ltR}:%s{column'name (int rbC)}%i{rbR}"].WrapText <- false
    [<CustomOperation("unwrap")>]
    member __.UnWrapText(_: unit, (c, r): (string * int<row>)) = 
      sheet.[$"%s{c}%i{r}"].WrapText <- false
    [<CustomOperation("unwrap")>]
    member __.UnWrapText(_: unit, (ltC, ltR): (string * int<row>), (rbC, rbR): (string * int<row>)) = 
      sheet.[$"%s{ltC}%i{ltR}:%s{rbC}%i{rbR}"].WrapText <- false

    [<CustomOperation("shrink")>]
    member __.ShrinkToFit(_: unit, target: string) =
      sheet.[target].ShrinkToFit <- true
    [<CustomOperation("shrink")>]
    member __.ShrinkToFit(_: unit, leftTop: string, rightBottom: string) = 
      sheet.[$"%s{leftTop}:%s{rightBottom}"].ShrinkToFit <- true
    [<CustomOperation("shrink")>]
    member __.ShrinkToFit(_: unit, (r, c): (int<row> * int<col>)) = 
      sheet.[int r, int c].ShrinkToFit <- true
    [<CustomOperation("shrink")>]
    member __.ShrinkToFit(_: unit, (ltR, ltC): (int<row> * int<col>), (rbR, rbC): (int<row> * int<col>)) = 
      sheet.[$"%s{column'name (int ltC)}%i{ltR}:%s{column'name (int rbC)}%i{rbR}"].ShrinkToFit <- true
    [<CustomOperation("shrink")>]
    member __.ShrinkToFit(_: unit, (c, r): (string * int<row>)) = 
      sheet.[$"%s{c}%i{r}"].ShrinkToFit <- true
    [<CustomOperation("shrink")>]
    member __.ShrinkToFit(_: unit, (ltC, ltR): (string * int<row>), (rbC, rbR): (string * int<row>)) = 
      sheet.[$"%s{ltC}%i{ltR}:%s{rbC}%i{rbR}"].ShrinkToFit <- true

    [<CustomOperation("unshrink")>]
    member __.UnShrinkToFit(_: unit, target: string) =
      sheet.[target].ShrinkToFit <- false
    [<CustomOperation("unshrink")>]
    member __.UnShrinkToFit(_: unit, leftTop: string, rightBottom: string) = 
      sheet.[$"%s{leftTop}:%s{rightBottom}"].ShrinkToFit <- false
    [<CustomOperation("unshrink")>]
    member __.UnShrinkToFit(_: unit, (r, c): (int<row> * int<col>)) = 
      sheet.[int r, int c].ShrinkToFit <- false
    [<CustomOperation("unshrink")>]
    member __.UnShrinkToFit(_: unit, (ltR, ltC): (int<row> * int<col>), (rbR, rbC): (int<row> * int<col>)) = 
      sheet.[$"%s{column'name (int ltC)}%i{ltR}:%s{column'name (int rbC)}%i{rbR}"].ShrinkToFit <- false
    [<CustomOperation("unshrink")>]
    member __.UnShrinkToFit(_: unit, (c, r): (string * int<row>)) = 
      sheet.[$"%s{c}%i{r}"].ShrinkToFit <- false
    [<CustomOperation("unshrink")>]
    member __.UnShrinkToFit(_: unit, (ltC, ltR): (string * int<row>), (rbC, rbR): (string * int<row>)) = 
      sheet.[$"%s{ltC}%i{ltR}:%s{rbC}%i{rbR}"].ShrinkToFit <- false

    [<CustomOperation("orientation")>]
    member __.Orientation(_: unit, target: string, angle: int) =
      sheet.[target].Orientation <- angle
    [<CustomOperation("orientation")>]
    member __.Orientation(_: unit, leftTop: string, rightBottom: string, angle: int) = 
      sheet.[$"%s{leftTop}:%s{rightBottom}"].Orientation <- angle
    [<CustomOperation("orientation")>]
    member __.Orientation(_: unit, (r, c): (int<row> * int<col>), angle: int) = 
      sheet.[int r, int c].Orientation <- angle
    [<CustomOperation("orientation")>]
    member __.Orientation(_: unit, (ltR, ltC): (int<row> * int<col>), (rbR, rbC): (int<row> * int<col>), angle: int) = 
      sheet.[$"%s{column'name (int ltC)}%i{ltR}:%s{column'name (int rbC)}%i{rbR}"].Orientation <- angle
    [<CustomOperation("orientation")>]
    member __.Orientation(_: unit, (c, r): (string * int<row>), angle: int) = 
      sheet.[$"%s{c}%i{r}"].Orientation <- angle
    [<CustomOperation("orientation")>]
    member __.Orientation(_: unit, (ltC, ltR): (string * int<row>), (rbC, rbR): (string * int<row>), angle: int) = 
      sheet.[$"%s{ltC}%i{ltR}:%s{rbC}%i{rbR}"].Orientation <- angle

    [<CustomOperation("format")>]
    member __.Format(_: unit, target: string, format: string) =
      sheet.[target].Format <- format
    [<CustomOperation("format")>]
    member __.Format(_: unit, leftTop: string, rightBottom: string,format: string) = 
      sheet.[$"%s{leftTop}:%s{rightBottom}"].Format <- format
    [<CustomOperation("format")>]
    member __.Format(_: unit, (r, c): (int<row> * int<col>), format: string) = 
      sheet.[int r, int c].Format <- format
    [<CustomOperation("format")>]
    member __.Format(_: unit, (ltR, ltC): (int<row> * int<col>), (rbR, rbC): (int<row> * int<col>), format: string) = 
      sheet.[$"%s{column'name (int ltC)}%i{ltR}:%s{column'name (int rbC)}%i{rbR}"].Format <- format
    [<CustomOperation("format")>]
    member __.Format(_: unit, (c, r): (string * int<row>), format: string) = 
      sheet.[$"%s{c}%i{r}"].Format <- format
    [<CustomOperation("format")>]
    member __.Format(_: unit, (ltC, ltR): (string * int<row>), (rbC, rbR): (string * int<row>), format: string) = 
      sheet.[$"%s{ltC}%i{ltR}:%s{rbC}%i{rbR}"].Format <- format

    [<CustomOperation("select")>]
    member __.Select(_: unit, target: string) =
      sheet.[target].Select()
    [<CustomOperation("select")>]
    member __.Select(_: unit, leftTop: string, rightBottom: string) = 
      sheet.[$"%s{leftTop}:%s{rightBottom}"].Select()
    [<CustomOperation("select")>]
    member __.Select(_: unit, (r, c): (int<row> * int<col>)) = 
      sheet.[int r, int c].Select()
    [<CustomOperation("select")>]
    member __.Select(_: unit, (ltR, ltC): (int<row> * int<col>), (rbR, rbC): (int<row> * int<col>)) = 
      sheet.[$"%s{column'name (int ltC)}%i{ltR}:%s{column'name (int rbC)}%i{rbR}"].Select()
    [<CustomOperation("select")>]
    member __.Select(_: unit, (c, r): (string * int<row>)) = 
      sheet.[$"%s{c}%i{r}"].Select()
    [<CustomOperation("select")>]
    member __.Select(_: unit, (ltC, ltR): (string * int<row>), (rbC, rbR): (string * int<row>)) = 
      sheet.[$"%s{ltC}%i{ltR}:%s{rbC}%i{rbR}"].Select()

  let sheet'op sheet = SheetOpBuilder sheet
  let paste'mode = { Paste= PasteType.All; Op= PasteOperation.None; SkipBlanks= false; Transpose= false; }
  let insert'mode = { Shift= InsertShiftDirection.Down; Origin= InsertFormatOrigin.FromRightOrBelow; }
  let delete'mode = { Shift= DeleteShiftDirection.Left; }

