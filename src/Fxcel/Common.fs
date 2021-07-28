namespace Fxcel

open System.Drawing
open Midoliy.Office.Interop

type RGB = { r: int; g: int; b: int }
type ThemeColor = { theme: Midoliy.Office.Interop.ThemeColor; tint: Tint }

[<AutoOpen>]
module Constant =
  let medium = BorderWeight.Medium
  let hairline = BorderWeight.Hairline
  let thin = BorderWeight.Thin
  let thick = BorderWeight.Thick
  
  let lineNone = LineStyle.None
  let dot = LineStyle.Dot
  let double = LineStyle.Double
  let dash = LineStyle.Dash
  let continuous = LineStyle.Continuous
  let dashdot = LineStyle.DashDot
  let dashdotdot = LineStyle.DashDotDot
  let slant = LineStyle.SlantDashDot

  let bg1 = Midoliy.Office.Interop.ThemeColor.Background1
  let bg2 = Midoliy.Office.Interop.ThemeColor.Background2
  let fg1 = Midoliy.Office.Interop.ThemeColor.Foreground1
  let fg2 = Midoliy.Office.Interop.ThemeColor.Foreground2
  let accent1 = Midoliy.Office.Interop.ThemeColor.Accent1
  let accent2 = Midoliy.Office.Interop.ThemeColor.Accent2
  let accent3 = Midoliy.Office.Interop.ThemeColor.Accent3
  let accent4 = Midoliy.Office.Interop.ThemeColor.Accent4
  let accent5 = Midoliy.Office.Interop.ThemeColor.Accent5
  let accent6 = Midoliy.Office.Interop.ThemeColor.Accent6

  let dark50 = Tint.Dark50
  let dark25 = Tint.Dark25
  let defultTint = Tint.Default
  let light40 = Tint.Light40
  let light60 = Tint.Light60
  let light80 = Tint.Light80
  
  let fs'normal = FontStyle.None
  let fs'bold = FontStyle.Bold
  let fs'italic = FontStyle.Italic
  let fs'shadow = FontStyle.Shadow
  let fs'strikethrough = FontStyle.Strikethrough
  let fs'subscript = FontStyle.Subscript
  let fs'superscript = FontStyle.Superscript
  let fs'singleUnderline = FontStyle.SingleUnderline
  let fs'doubleUnderline = FontStyle.DoubleUnderline

  let ul'none = Underline.None
  let ul'double = Underline.Double
  let ul'doubleAccounting = Underline.DoubleAccounting
  let ul'single = Underline.Single
  let ul'singleAccounting = Underline.SingleAccounting
