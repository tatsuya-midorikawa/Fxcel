namespace Fxcel

open Midoliy.Office.Interop
open System.Drawing

[<AutoOpen>]
module ChartOp =

  type ChartOpBuilder (sheet: IWorksheet) =
    [<DefaultValue>] val mutable width : int<cols>
    [<DefaultValue>] val mutable height : int<rows>
    [<DefaultValue>] val mutable position : string

    member __.Yield (_: unit) = ()
    member __.Zero() = ()
    /// <summary>チャートで利用するデータ範囲を選択する.</summary>
    /// <param name="target">データの範囲</param>
    [<CustomOperation("select")>]
    member __.Select(_: unit, target: string) = sheet.[target].Select()

    /// <summary></summary>
    /// <param name="size">チャートのサイズ. (width * height)</param>
    [<CustomOperation("size")>]
    member __.Size(_: unit, size: (int<cols> * int<rows>)) = 
      let w, h = size
      __.width <- w
      __.height <- h
      
    /// <summary></summary>
    /// <param name="position">チャートの挿入位置をアドレス形式で指定.</param>
    [<CustomOperation("position")>]
    member __.Position(_: unit, position: string) = __.position <- position

    /// <summary>
    /// チャートを追加する.
    /// </summary>
    /// <param name="recipe">追加するチャートの種類</param>
    /// <param name="newLayout">動的書式設定規則を使用する. default: true</param>
    [<CustomOperation("add")>]
    member __.Add(_: unit, recipe: ChartRecipe, ?newLayout: bool) =
      let isDefault = System.String.IsNullOrWhiteSpace
      let newLayout = match newLayout with Some v -> v | None -> true
      match (int __.width, int __.height, isDefault(__.position)) with
      | (w, h, true) when 0 < w && 0 < h -> sheet.Shapes.AddChart(recipe, sheet.[$"A1:%s{column'name w}1"], sheet.[$"A1:A%d{h}"], newLayout)
      | (w, h, false) when 0 < w && 0 < h -> sheet.Shapes.AddChart(recipe, sheet.[__.position], sheet.[$"A1:%s{column'name w}1"], sheet.[$"A1:A%d{h}"], newLayout)
      | _ -> sheet.Shapes.AddChart(recipe, newLayout)

  let chart'op (sheet: IWorksheet) = ChartOpBuilder sheet
