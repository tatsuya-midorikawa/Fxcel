namespace Fxcel

open System.Drawing
open Midoliy.Office.Interop

[<AutoOpen>]
module Border =

  type Border = { Style: LineStyle; Weight: BorderWeight; Color: Color }

  type BorderBuilder () =
    member __.Yield (_: unit) = { Style= LineStyle.Continuous; Weight= BorderWeight.Medium; Color= Color.Black }
    member __.Zero() = ()
    [<CustomOperation("style")>]
    member __.SetStyle(current: Border, style: LineStyle) = { current with Style = style }
    [<CustomOperation("weight")>]
    member __.SetWeight(current: Border, weight: BorderWeight) = { current with Weight = weight }
    [<CustomOperation("color")>]
    member __.SetColor(current: Border, color: Color) = { current with Color = color }
    [<CustomOperation("rgb")>]
    member __.SetRGB(current: Border, color: RGB) = { current with Color = Color.FromArgb(color.r, color.g, color.b) }

  let border = BorderBuilder()
  let solidline = { Style= LineStyle.Continuous; Weight= BorderWeight.Medium; Color= Color.Black }
  let dotline = { Style= LineStyle.Dot; Weight= BorderWeight.Medium; Color= Color.Black }
  let dashline = { Style= LineStyle.Dash; Weight= BorderWeight.Medium; Color= Color.Black }
  let doubleline = { Style= LineStyle.Double; Weight= BorderWeight.Medium; Color= Color.Black }
