namespace Fxcel

open System.Drawing
open Midoliy.Office.Interop

[<AutoOpen>]
module Font =
  let showFonts() = FontFamily.Families |> Array.iter (fun font -> printfn $" - {font.Name}")

  type FontBuilder (range: IExcelRange) =
    member __.Yield (_: unit) = range.Font
    member __.Zero() = ()
    [<CustomOperation("name")>]
    member __.SetName(current: IRangeFont, name: string) = current.Name <- name; current
    [<CustomOperation("size")>]
    member __.SetSize(current: IRangeFont, size: float) = current.Size <- size; current
    [<CustomOperation("style")>]
    member __.SetStyle(current: IRangeFont, style: FontStyle) = current.Style <- style; current
    [<CustomOperation("color")>]
    member __.SetColor(current: IRangeFont, color: Color) = current.Color <- color; current
    [<CustomOperation("rgb")>]
    member __.SetRGB(current: IRangeFont, color: RGB) = current.Color <- Color.FromArgb(color.r, color.g, color.b); current
    [<CustomOperation("underline")>]
    member __.SetUnderline(current: IRangeFont, underline: Underline) = current.Underline <- underline; current
    [<CustomOperation("bold")>]
    member __.SetBold(current: IRangeFont, on: bool) = current.Bold <- on; current
    [<CustomOperation("italic")>]
    member __.SetItalic(current: IRangeFont, on: bool) = current.Italic <- on; current
    [<CustomOperation("shadow")>]
    member __.SetShadow(current: IRangeFont, on: bool) = current.Shadow <- on; current
    [<CustomOperation("outline")>]
    member __.SetOutlineFont(current: IRangeFont, on: bool) = current.OutlineFont <- on; current
    [<CustomOperation("strikethrough")>]
    member __.SetStrikethrough(current: IRangeFont, on: bool) = current.Strikethrough <- on; current
    [<CustomOperation("subscript")>]
    member __.SetSubscript(current: IRangeFont, on: bool) = current.Subscript <- on; current
    [<CustomOperation("superscript")>]
    member __.SetSuperscript(current: IRangeFont, on: bool) = current.Superscript <- on; current

  let font x = FontBuilder x
