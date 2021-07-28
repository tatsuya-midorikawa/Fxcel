namespace Fxcel

open System.Drawing
open Midoliy.Office.Interop

[<AutoOpen>]
module Border =
  type Border = { LineStyle: LineStyle; Weight: BorderWeight; Color: Color }

  let border = { LineStyle= LineStyle.Continuous; Weight= BorderWeight.Medium; Color= Color.Black }
  let solidline = { LineStyle= LineStyle.Continuous; Weight= BorderWeight.Medium; Color= Color.Black }
  let dotline = { LineStyle= LineStyle.Dot; Weight= BorderWeight.Medium; Color= Color.Black }
  let dashline = { LineStyle= LineStyle.Dash; Weight= BorderWeight.Medium; Color= Color.Black }
  let doubleline = { LineStyle= LineStyle.Double; Weight= BorderWeight.Medium; Color= Color.Black }
