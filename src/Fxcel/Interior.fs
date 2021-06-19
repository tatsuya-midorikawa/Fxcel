namespace Fxcel

open System.Drawing
open Midoliy.Office.Interop

[<AutoOpen>]
module Interior =
  let medium = BorderWeight.Medium
  let hairline = BorderWeight.Hairline
  let thin = BorderWeight.Thin
  let thick = BorderWeight.Thick
  
  let none = LineStyle.None
  let dot = LineStyle.Dot
  let double = LineStyle.Double
  let dash = LineStyle.Dash
  let continuous = LineStyle.Continuous
  let dashdot = LineStyle.DashDot
  let dashdotdot = LineStyle.DashDotDot
  let slant = LineStyle.SlantDashDot

  type Border = { Style: LineStyle; Weight: BorderWeight; Color: Color }
  type RuledLineBuilder (range: IExcelRange) =
    member private __.Update(target: IBorders, index: BordersIndex, border: Border) =
      target.[index].LineStyle <- border.Style
      target.[index].Weight <- border.Weight
      target.[index].Color <- border.Color
      target
    
    member __.Yield (_: unit) = range.Borders
    member __.Zero() = ()
    [<CustomOperation("growing")>]
    member __.SetRuledLineUp(current: IBorders, border: Border) = __.Update(current, BordersIndex.DiagonalUp, border)
    [<CustomOperation("falling")>]
    member __.SetRuledLineDown(current: IBorders, border: Border) = __.Update(current, BordersIndex.DiagonalDown, border)
    [<CustomOperation("top")>]
    member __.SetRuledLineTop(current: IBorders, border: Border) = __.Update(current, BordersIndex.EdgeTop, border)
    [<CustomOperation("bottom")>]
    member __.SetRuledLineBottom(current: IBorders, border: Border) = __.Update(current, BordersIndex.EdgeBottom, border)
    [<CustomOperation("left")>]
    member __.SetRuledLineLeft(current: IBorders, border: Border) = __.Update(current, BordersIndex.EdgeLeft, border)
    [<CustomOperation("right")>]
    member __.SetRuledLineRight(current: IBorders, border: Border) = __.Update(current, BordersIndex.EdgeRight, border)
    [<CustomOperation("horizontal")>]
    member __.SetRuledLineHorizontal(current: IBorders, border: Border) = __.Update(current, BordersIndex.InsideHorizontal, border)
    [<CustomOperation("vertical")>]
    member __.SetRuledLineVertical(current: IBorders, border: Border) = __.Update(current, BordersIndex.InsideVertical, border)

  type BorderBuilder () =
    member __.Yield (_: unit) = { Style= LineStyle.Continuous; Weight= BorderWeight.Medium; Color= Color.Black }
    member __.Zero() = ()
    [<CustomOperation("style")>]
    member __.SetStyle(current: Border, style: LineStyle) = { current with Style = style }
    [<CustomOperation("weight")>]
    member __.SetWeight(current: Border, weight: BorderWeight) = { current with Weight = weight }
    [<CustomOperation("color")>]
    member __.SetColor(current: Border, color: Color) = { current with Color = color }

  let ruledline x = RuledLineBuilder x
  let border = BorderBuilder()
