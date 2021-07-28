namespace Fxcel

open System.Drawing
open Midoliy.Office.Interop

type RGB = { r: int; g: int; b: int }
type ThemeColor = { theme: Midoliy.Office.Interop.ThemeColor; tint: Tint }

[<AutoOpen>]
module Constant =
  let rgb (r, g, b) = Color.FromArgb(r, g, b)

  let weight'medium = BorderWeight.Medium
  let weight'hairline = BorderWeight.Hairline
  let weight'thin = BorderWeight.Thin
  let weight'thick = BorderWeight.Thick
  
  let linestyle'none = LineStyle.None
  let linestyle'dot = LineStyle.Dot
  let linestyle'double = LineStyle.Double
  let linestyle'dash = LineStyle.Dash
  let linestyle'continuous = LineStyle.Continuous
  let linestyle'dashdot = LineStyle.DashDot
  let linestyle'dashdotdot = LineStyle.DashDotDot
  let linestyle'slant = LineStyle.SlantDashDot

  let theme'bg1 = Midoliy.Office.Interop.ThemeColor.Background1
  let theme'bg2 = Midoliy.Office.Interop.ThemeColor.Background2
  let theme'fg1 = Midoliy.Office.Interop.ThemeColor.Foreground1
  let theme'fg2 = Midoliy.Office.Interop.ThemeColor.Foreground2
  let theme'accent1 = Midoliy.Office.Interop.ThemeColor.Accent1
  let theme'accent2 = Midoliy.Office.Interop.ThemeColor.Accent2
  let theme'accent3 = Midoliy.Office.Interop.ThemeColor.Accent3
  let theme'accent4 = Midoliy.Office.Interop.ThemeColor.Accent4
  let theme'accent5 = Midoliy.Office.Interop.ThemeColor.Accent5
  let theme'accent6 = Midoliy.Office.Interop.ThemeColor.Accent6

  let tint'dark50 = Tint.Dark50
  let tint'dark25 = Tint.Dark25
  let tint'defultTint = Tint.Default
  let tint'light40 = Tint.Light40
  let tint'light60 = Tint.Light60
  let tint'light80 = Tint.Light80
  
  let style'normal = FontStyle.None
  let style'bold = FontStyle.Bold
  let style'italic = FontStyle.Italic
  let style'shadow = FontStyle.Shadow
  let style'strikethrough = FontStyle.Strikethrough
  let style'subscript = FontStyle.Subscript
  let style'superscript = FontStyle.Superscript
  let style'singleUnderline = FontStyle.SingleUnderline
  let style'doubleUnderline = FontStyle.DoubleUnderline

  let underline'none = Underline.None
  let underline'double = Underline.Double
  let underline'doubleAccounting = Underline.DoubleAccounting
  let underline'single = Underline.Single
  let underline'singleAccounting = Underline.SingleAccounting

  let shift'left = DeleteShiftDirection.Left
  let shift'up = DeleteShiftDirection.Up
  let shift'right = InsertShiftDirection.Right
  let shift'down = InsertShiftDirection.Down

  let paste'values = PasteType.Values
  let paste'comments = PasteType.Comments
  let paste'formulas = PasteType.Formulas
  let paste'formats = PasteType.Formats
  let paste'all = PasteType.All
  let paste'validation = PasteType.Validation
  let paste'exceptBorders = PasteType.AllExceptBorders
  let paste'colmnWidths = PasteType.ColumnWidths
  let paste'formulasAndNumberFormats = PasteType.FormulasAndNumberFormats
  let paste'valuesAndNumberFormats = PasteType.ValuesAndNumberFormats
  let paste'allUsingSourceTheme = PasteType.AllUsingSourceTheme
  let paste'allMergingConditionalFormats = PasteType.AllMergingConditionalFormats
  
  let op'none = PasteOperation.None
  let op'add = PasteOperation.Add
  let op'sub = PasteOperation.Subtract
  let op'mul = PasteOperation.Multiply
  let op'div = PasteOperation.Divide

  let origin'left = InsertFormatOrigin.FromLeftOrAbove
  let origin'above = InsertFormatOrigin.FromLeftOrAbove
  let origin'right = InsertFormatOrigin.FromRightOrBelow
  let origin'below = InsertFormatOrigin.FromRightOrBelow
