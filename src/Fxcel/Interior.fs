namespace Fxcel

open System.Drawing
open Midoliy.Office.Interop

[<AutoOpen>]
module Interior =

  type RuledLineBuilder (range: IExcelRange) =
    member private __.Update(target: IBorders, index: BordersIndex, border: Border) =
      match border.LineStyle with
      | LineStyle.None ->
        target.[index].LineStyle <- border.LineStyle
      | _ ->
        target.[index].LineStyle <- border.LineStyle
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


  let ruledline x = RuledLineBuilder x
