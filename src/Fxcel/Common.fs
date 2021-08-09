namespace Fxcel

open System.Drawing
open Midoliy.Office.Interop

type RGB = { r: int; g: int; b: int }
type ThemeColor = { theme: Midoliy.Office.Interop.ThemeColor; tint: Tint }
[<Measure>] type range
[<Measure>] type cell
[<Measure>] type row
[<Measure>] type col
[<Measure>] type rows
[<Measure>] type cols

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

  let theme'background1 = ThemeColor.Background1
  let theme'background2 = ThemeColor.Background2
  let theme'foreground1 = ThemeColor.Foreground1
  let theme'foreground2 = ThemeColor.Foreground2
  let theme'accent1 = ThemeColor.Accent1
  let theme'accent2 = ThemeColor.Accent2
  let theme'accent3 = ThemeColor.Accent3
  let theme'accent4 = ThemeColor.Accent4
  let theme'accent5 = ThemeColor.Accent5
  let theme'accent6 = ThemeColor.Accent6

  let tint'dark50 = Tint.Dark50
  let tint'dark25 = Tint.Dark25
  let tint'defultTint = Tint.Default
  let tint'light40 = Tint.Light40
  let tint'light60 = Tint.Light60
  let tint'light80 = Tint.Light80
  
  let pattern'none = Pattern.None
  let pattern'auto = Pattern.Automatic
  let pattern'up = Pattern.Up
  let pattern'down = Pattern.Down
  let pattern'vertical = Pattern.Vertical
  let pattern'horizontal = Pattern.Horizontal
  let pattern'lightUp = Pattern.LightUp
  let pattern'lightDown = Pattern.LightDown
  let pattern'lightVertical = Pattern.LightVertical
  let pattern'lightHorizontal = Pattern.LightHorizontal
  let pattern'gray8 = Pattern.Gray8
  let pattern'gray16 = Pattern.Gray16
  let pattern'gray25 = Pattern.Gray25
  let pattern'gray50 = Pattern.Gray50
  let pattern'gray75 = Pattern.Gray75
  let pattern'semigray75 = Pattern.SemiGray75
  let pattern'solid = Pattern.Solid
  let pattern'checker = Pattern.Checker
  let pattern'grid = Pattern.Grid
  let pattern'crisscross = Pattern.CrissCross
  let pattern'linearGradient = Pattern.LinearGradient
  let pattern'rectangularGradient = Pattern.RectangularGradient

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
  
  let calc'auto = Calculation.Auto
  let calc'manual = Calculation.Manual
  let calc'semiauto = Calculation.Semiauto

  let visibility'visible = AppVisibility.Visible
  let visibility'hidden = AppVisibility.Hidden
  
  let h'right = HorizontalAlignment.Right
  let h'left = HorizontalAlignment.Left
  let h'center = HorizontalAlignment.Center
  let h'justify = HorizontalAlignment.Justify
  let h'distributed = HorizontalAlignment.Distributed
  let h'general = HorizontalAlignment.General
  let h'fill = HorizontalAlignment.Fill
  let h'centerAcrossSelection = HorizontalAlignment.CenterAcrossSelection

  let v'top = VerticalAlignment.Top
  let v'bottom = VerticalAlignment.Bottom
  let v'center = VerticalAlignment.Center
  let v'justify = VerticalAlignment.Justify
  let v'distributed = VerticalAlignment.Distributed
